#通用重试动作：执行动作后轮询识别，直到目标出现或消失
#pipeline侧接口约定（节点action填Custom，custom_action填下面的动作名）：
#  RepeatUntilFoundAction     param: {"action": "Click", "action_param": {}, "wait_nodes": ["目标节点"]}
#                             反复执行内置动作，直到wait_nodes里任一节点命中
#  RepeatUntilNotFoundAction  param: {"custom_action": "PopupKeys", "custom_action_param": {}, "wait_node": "目标节点"}
#                             反复执行自定义动作，直到wait_node不命中
#  action和custom_action二选一；repeat_count省略或<=0时默认3；interval_ms省略或0时默认3000
#  interval_ms窗口内每500ms截图识别一次；目标状态满足后立即返回True，耗尽次数返回False

import json
import time

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction

_POLL_MS = 500  #每轮动作后的等待窗口内，截图识别的间隔
_INNER_NODE = "__RepeatUntilActionInner"  #MaaEnd使用的临时动作节点名
_DEFAULT_REPEAT = 3
_DEFAULT_INTERVAL = 3000


def _load_param(argv):
    try:
        param = json.loads(argv.custom_action_param or "{}")
    except (TypeError, json.JSONDecodeError) as e:
        print(f"[RepeatUntil] 参数不是有效JSON: {e}")
        return None
    if not isinstance(param, dict):
        print("[RepeatUntil] 参数必须是JSON对象")
        return None
    return param


def _build_action(param):
    action = param.get("action")
    custom_action = param.get("custom_action")
    if bool(action) == bool(custom_action):
        print("[RepeatUntil] action和custom_action必须二选一")
        return None

    if action:
        if not isinstance(action, str):
            print("[RepeatUntil] action必须是字符串")
            return None
        action_param = param.get("action_param") or {}
        if not isinstance(action_param, dict):
            print("[RepeatUntil] action_param必须是JSON对象")
            return None
        return {"type": action, "param": action_param}

    if not isinstance(custom_action, str):
        print("[RepeatUntil] custom_action必须是字符串")
        return None
    custom_action_param = param.get("custom_action_param")
    if custom_action_param is None:
        custom_action_param = {}
    if not isinstance(custom_action_param, dict):
        print("[RepeatUntil] custom_action_param必须是JSON对象")
        return None
    return {
        "type": "Custom",
        "param": {
            "custom_action": custom_action,
            "custom_action_param": custom_action_param,
        },
    }


def _parse(argv, until_found):
    param = _load_param(argv)
    if param is None:
        return None

    action = _build_action(param)
    if action is None:
        return None

    if until_found:
        waits = param.get("wait_nodes")
        if (
            not isinstance(waits, list)
            or not waits
            or any(not isinstance(name, str) or not name for name in waits)
        ):
            print("[RepeatUntil] wait_nodes必须是非空字符串数组")
            return None
    else:
        wait_node = param.get("wait_node")
        if not isinstance(wait_node, str) or not wait_node:
            print("[RepeatUntil] wait_node必须是非空字符串")
            return None
        waits = [wait_node]

    repeat = param.get("repeat_count", _DEFAULT_REPEAT)
    interval = param.get("interval_ms", _DEFAULT_INTERVAL)
    if isinstance(repeat, bool) or not isinstance(repeat, int):
        print("[RepeatUntil] repeat_count必须是整数")
        return None
    if isinstance(interval, bool) or not isinstance(interval, int):
        print("[RepeatUntil] interval_ms必须是整数")
        return None
    if repeat <= 0:
        repeat = _DEFAULT_REPEAT
    if interval < 0:
        print("[RepeatUntil] interval_ms不能为负数")
        return None
    if interval == 0:
        interval = _DEFAULT_INTERVAL

    return action, waits, repeat, interval


def _run_loop(context, box, action, waits, repeat, interval, until_found):
    controller = context.tasker.controller
    override = {
        _INNER_NODE: {
            "pre_delay": 0,
            "post_delay": 0,
            "rate_limit": 0,
            "action": action,
        }
    }

    for round_no in range(1, repeat + 1):
        if context.tasker.stopping:
            print("[RepeatUntil] 任务正在停止，中止重试")
            return False

        detail = context.run_action(_INNER_NODE, box, "", override)
        if detail is None or not detail.success:
            print(f"[RepeatUntil] 第{round_no}轮动作执行失败，继续等待/重试")

        deadline = time.monotonic() + interval / 1000
        while True:
            if context.tasker.stopping:
                return False
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            time.sleep(min(_POLL_MS / 1000, remaining))
            if context.tasker.stopping:
                return False

            controller.post_screencap().wait()
            try:
                image = controller.cached_image
            except RuntimeError:
                print("[RepeatUntil] 取截图失败，跳过本轮识别")
                continue

            if until_found:
                for name in waits:
                    result = context.run_recognition(name, image)
                    if result is not None and result.hit:
                        print(f"[RepeatUntil] 第{round_no}轮后「{name}」已出现")
                        return True
            else:
                result = context.run_recognition(waits[0], image)
                if result is None or not result.hit:
                    print(f"[RepeatUntil] 第{round_no}轮后「{waits[0]}」已消失")
                    return True

    print(f"[RepeatUntil] {repeat}轮执行完毕，目标状态未满足")
    return False


@AgentServer.custom_action("RepeatUntilFoundAction")  #重复动作直到任一目标节点出现
class RepeatUntilFoundAction(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        parsed = _parse(argv, until_found=True)
        if parsed is None:
            return False  #参数不合法
        action, waits, repeat, interval = parsed
        return _run_loop(context, argv.box, action, waits, repeat, interval, True)


@AgentServer.custom_action("RepeatUntilNotFoundAction")  #重复动作直到目标节点消失
class RepeatUntilNotFoundAction(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        parsed = _parse(argv, until_found=False)
        if parsed is None:
            return False  #参数不合法
        action, waits, repeat, interval = parsed
        return _run_loop(context, argv.box, action, waits, repeat, interval, False)
