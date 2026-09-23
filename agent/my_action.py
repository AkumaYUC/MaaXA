import json
import ctypes
import re
import time

from maa.agent.agent_server import AgentServer
from maa.custom_action import CustomAction
from maa.context import Context

from common import _params  #pipeline 参数解析兜底（写坏打日志返回False，不在agent进程里抛异常）


@AgentServer.custom_action("Win32BringToFront")  #将指定窗口置顶
class Win32BringToFront(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)  #拉取参数；参数写坏打日志返回False，不在agent进程里抛异常
        if param is None:
            return False
        pattern = param.get("window_regex", "")  #要匹配的窗口标题正则
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

        #只救「真最小化」（缩到任务栏时截图只能截到桌面，任务必挂）；窗口大小一律不主动改。
        #模板和 ROI 是按「最大化窗口」的尺寸标定的，任何主动还原/最大化都可能改掉窗口尺寸
        #导致归一化截图变化、模板全挂（踩坑清单#26 同源）。最大化与否由用户自己保持。
        iconic_at_entry = bool(user32.IsIconic(hwnd))
        if iconic_at_entry:
            user32.ShowWindow(hwnd, 9)  #9 = SW_RESTORE；最大化后被最小化的窗口会还原回最大化
            time.sleep(0.1)

        #本动作原先一行日志都没有，窗口出问题只能靠 on_error 截图反推，补上状态快照
        from ctypes import wintypes
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        print(f"[Win32BringToFront] hwnd={hwnd} iconic_at_entry={iconic_at_entry} "
              f"zoomed_now={bool(user32.IsZoomed(hwnd))} "
              f"rect={rect.left},{rect.top},{rect.right},{rect.bottom} "
              f"size={rect.right - rect.left}x{rect.bottom - rect.top}")

        if user32.GetForegroundWindow() == hwnd:  #已经在前台就到此为止，多余动作只会引入副作用（同 _activate_window 原则）
            return True

        #前台锁是间歇性的：单次 SetForegroundWindow 时灵时不灵，实测「要点好几次才能启动任务」。
        #改为最多 3 轮阶梯重试，任一轮到位即返回，全部落空才报 False 让 pipeline 走 on_error。
        #不用 Alt 绕锁——Alt 会让金蝶进菜单加速键模式，紧跟的 Enter 被当成激活菜单项（踩坑清单#26）。
        our_tid = ctypes.windll.kernel32.GetCurrentThreadId()
        target_tid = user32.GetWindowThreadProcessId(hwnd, None)  #目标窗口所属线程

        def activate(attach_tid):
            #把「当前拥有前台资格的线程」和「目标窗口线程」一起挂到自己的输入队列上再置顶。
            #Windows 的前台资格绑定在最近收到输入的线程上，借它的队列才能绕过前台锁。
            if attach_tid:
                user32.AttachThreadInput(our_tid, attach_tid, True)
            user32.AttachThreadInput(our_tid, target_tid, True)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            user32.AttachThreadInput(our_tid, target_tid, False)  #用完立刻解绑，挂着不放会连带卡住输入
            if attach_tid:
                user32.AttachThreadInput(our_tid, attach_tid, False)

        for round_no in range(1, 4):  #最多 3 轮
            fg_hwnd = user32.GetForegroundWindow()  #当前前台窗口（每轮重取，它会被上一轮改变）
            if fg_hwnd == hwnd:
                break  #已经到位，不必再动

            if round_no <= 2:
                #前两轮借「当前前台窗口的线程」的资格，比只挂目标线程更有效
                activate(user32.GetWindowThreadProcessId(fg_hwnd, None) if fg_hwnd else 0)
            else:
                #第 3 轮兜底：SwitchToThisWindow 是 Alt+Tab 的内部实现，不受前台锁限制
                #（它是 API 调用不是模拟按键，不会像 Alt 那样把金蝶带进菜单加速键模式）
                user32.SwitchToThisWindow(hwnd, True)

            time.sleep(0.15)  #给窗口切换一点时间

        #旧版无脑返回 True 装成功，前台实际被 Windows 间歇性拒绝（日志 327 次 Failed to ensure foreground）。
        #现在真的到位才报成功；结果连两侧句柄一起打日志，失败时不用靠 on_error 截图反推。
        at_front = user32.GetForegroundWindow() == hwnd
        print(f"[Win32BringToFront] 置顶{'成功' if at_front else '失败'} "
              f"foreground={user32.GetForegroundWindow()} target={hwnd}")
        return at_front