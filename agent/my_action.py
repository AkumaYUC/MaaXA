import json
import ctypes
import re

from maa.agent.agent_server import AgentServer
from maa.custom_action import CustomAction
from maa.context import Context


@AgentServer.custom_action("my_action_111")  #此为自定义动作模版
class MyCustomAction(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        print("my_action_111 is running!")

        return True



@AgentServer.custom_action("Win32BringToFront")  #将指定窗口置顶
class Win32BringToFront(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        pattern = json.loads(argv.custom_action_param or "{}").get("window_regex", "")  #拉取命令
        if not pattern:
            return False  #没配置window_regex参数，直接返回False


        user32 = ctypes.windll.user32  #api调入入口
        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)  #回调函数类型定义
        def enum_windows_proc(hwnd, lparam):
            length = user32.GetWindowTextLengthW(hwnd)  #获取窗口标题长度
            if length > 0:
                buffer = ctypes.create_unicode_buffer(length + 1)  #创建缓冲区
                user32.GetWindowTextW(hwnd, buffer, length + 1)  #获取窗口标题
                title = buffer.value
                if re.search(pattern, title) and user32.IsWindowVisible(hwnd):  #匹配正则表达式且窗口可见
                    found.append(hwnd)  #匹配成功，加入found列表
                    return False  #停止找窗口
                
            return True  #继续找窗口
        user32.EnumWindows(enum_windows_proc, None)  #开始遍历所有窗口
        if not found:
            return False  #没找到匹配的窗口（有可能是没打开），返回False

        hwnd = found[0]  #取第一个匹配的窗口句柄
        if user32.IsIconic(hwnd):  #判断窗口是否最小化
            user32.ShowWindow(hwnd, 9)  #如果最小化，则恢复窗口，9 = SW_RESTORE
        else:
            user32.ShowWindow(hwnd, 5)  #如果不是最小化，则显示窗口，5 = SW_SHOW

        user32.keybd_event(0x12, 0, 0, 0)  #模拟按下Alt键
        user32.SetForegroundWindow(hwnd)  #将窗口置顶
        user32.keybd_event(0x12, 0, 2, 0)  #模拟释放Alt键

        return True  #成功将窗口置顶，返回True