import json
import ctypes
import re
import time

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

        #借目标线程的输入队列绕系统前台锁（AttachThreadInput），不用 Alt——Alt 会让金蝶进菜单加速键模式
        #（踩坑清单#26：Alt 后紧跟 Enter 被当成激活菜单项，实测弹出带「关闭」按钮的窗口）
        target_tid = user32.GetWindowThreadProcessId(hwnd, None)  #目标窗口所属线程
        our_tid = ctypes.windll.kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(our_tid, target_tid, True)  #挂上，SetForegroundWindow 才不会被前台锁拒绝
        user32.BringWindowToTop(hwnd)
        ok = user32.SetForegroundWindow(hwnd)
        user32.AttachThreadInput(our_tid, target_tid, False)  #用完立刻解绑，挂着不放会连带卡住输入
        time.sleep(0.1)  #给窗口切换一点时间

        #旧版无脑返回 True 装成功，前台实际被 Windows 间歇性拒绝（日志 327 次 Failed to ensure foreground）
        #任务开头只拉这一次，装成功=后面全程裸奔。真的到位了才报成功，失败让 pipeline 走 on_error
        return bool(ok) and user32.GetForegroundWindow() == hwnd