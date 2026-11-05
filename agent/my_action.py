#将指定窗口置顶并确保前台（Win32BringToFront）
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名）：
#  Win32BringToFront  param: {"window_regex": "金蝶云星空"}
#                     前置条件：窗口必须已经存在（找不到窗口返回False，让pipeline走on_error）
#                     选窗规则：类名与标题都参与评分 —— 金蝶主窗口类名里带 Kingdee.BOS.XPF.App.exe，
#                     比标题可靠（标题带「金蝶」二字的笔记软件会被同一正则误命中）；
#                     类名命中正则的优先 → 标题整体匹配的次之 → 面积大的再次之
#  返回False的两种情形：①找不到窗口 ②窗口还原不了（多因目标以管理员身份运行，UIPI 静默拦截，
#  日志里会把两侧进程的完整性级别打出来，并给出处理办法）——失败当场报，不拖到识别超时
#
#合成点击（SynthClick）
#pipeline侧接口约定（节点action填Custom，custom_action填SynthClick）：
#  SynthClick  param: {"window_regex": "金蝶云星空", "activate_regex": "列表过滤", "x": 892, "y": 589}
#                     前置条件：窗口必须已经存在且可见（找不到返回False，让pipeline走on_error）
#                     window_regex：基准窗口 —— 图像坐标按它的客户区短边换算（ROI 就是按它标定的）
#                     activate_regex：点击前要抢到前台的窗口（弹窗类节点必须给），省略=基准窗口自己
#                     x/y：1216x720 图像坐标；省略时用前序识别框 argv.box 的中心
#                     行为：抢前台 → SetCursorPos → move_ms → 左键按下 → hold_ms → 抬起 → settle_ms
#                     全程不复刻 maa 的「每次点击前把主窗拉前台」，也不碰目标窗口以外的窗口
#                     move_ms/hold_ms/settle_ms：可选，默认 120/80/150（毫秒），写错就用默认值
#                     存在理由：maa 的 Seize 输入法在 touch_down/touch_up 里各把主窗拉一次前台，
#                     而金蝶的弹窗是独立顶层窗 —— 弹窗被推成「非前台」后，WPF 会把这一次点击当成
#                     「激活点击」吃掉，「确定」按钮连点 64 次毫无反应就是这么来的（已实测定案）
#  返回False的三种情形：①找不到窗口（没开/标题变了）②点位最上层不是目标窗口（有覆盖层遮挡，
#  点了也是白点、还可能点到别的窗口上）③SendInput 没发出去 —— 都是当场报，不拖到识别超时
#
#自定义识别（节点recognition填Custom、custom_recognition填下面的识别名；和动作是两个互不相干的槽）
#  ErrorPopupOpen     param: 可省 ；{"title": "错误提示", "class_part": "Kingdee.BOS.XPF.App.exe"} 可覆盖
#                     金蝶的错误弹窗（粘贴异常就是它）还挂着时命中，关了就命中不了
#                     判据走 Win32 窗口枚举、不看像素：模板匹配会被 DPI/缩放差异打死
#                     （粘贴异常.png 实测 0.607921 < 阈值 0.70，6 次尝试 0 命中，根因见知识库改动记忆 2026-11-05），
#                     小字 OCR 同样不可靠（「粘贴异常」现场只有约 42x12 像素，踩坑清单 #32）
#                     title 要和弹窗标题栏文字精确一致（金蝶所有错误弹窗共用「错误提示」这个标题，
#                     所以本识别不区分正文写的是粘贴异常还是别的报错：先让 pipeline 把它关掉，
#                     要不要重粘交给下一步的字段校验去判断，别让弹窗一直杵在那儿挡路）
#                     class_part 是防误伤的：别的程序也可能有叫「错误提示」的窗口，类名不含金蝶的一律不算
#                     命中返回占位框(0,0,1,1)，只配 PopupKeys / ClickKey 这类不需要目标的动作，别配 Click+target
#                     配 inverse: true 就是「弹窗已经关了」判据（关闭校验节点用）

import json
import ctypes
import re
import time

from ctypes import wintypes  #RECT 等结构体定义（必须显式 import，import ctypes 不会带出 wintypes）

from maa.agent.agent_server import AgentServer
from maa.custom_action import CustomAction
from maa.custom_recognition import CustomRecognition
from maa.context import Context

from common import _params  #pipeline 参数解析兜底（写坏打日志返回False，不在agent进程里抛异常）

_ERROR_POPUP_TITLE = "错误提示"  #金蝶错误弹窗的标题栏文字（粘贴异常弹窗就是它，现场 FindWindowW 实测命中）
_ERROR_POPUP_CLASS = "Kingdee.BOS.XPF.App.exe"  #金蝶 WPF 窗口类名里嵌的可执行文件名，用来把金蝶的弹窗和别人的窗口分开

_IMAGE_SHORT_SIDE = 720  #图像空间短边：maa 把窗口客户区按短边归一化成 1216x720，pipeline 的 ROI 都按它标定
_MIN_CLIENT_SIDE = 200  #客户区任一边小于此值就认为几何不可信（窗口最小化时是 237x39 的假几何）
_DEFAULT_MOVE_MS = 120  #移动到按下之间等多久（给 XPF 建立 hover/命中态）
_DEFAULT_HOLD_MS = 80  #按下到抬起之间按住多久
_DEFAULT_SETTLE_MS = 150  #抬起之后等多久再让下一个节点抓图，避免抓到「按下态」的中间帧
_DPI_READY = False  #进程级 DPI 感知只能设一次，设过就不再试；线程级的每次调用都设
_MOUSEEVENTF_LEFTDOWN = 0x0002  #左键按下
_MOUSEEVENTF_LEFTUP = 0x0004  #左键抬起


class _MOUSEINPUT(ctypes.Structure):
    #SendInput 的鼠标输入项
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class _INPUT(ctypes.Structure):
    #SendInput 的联合体：本动作只发鼠标，故只声明鼠标项
    class _INPUTUNION(ctypes.Union):
        _fields_ = [("mi", _MOUSEINPUT)]

    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]


def _integrity_rid(pid):
    #读进程令牌的完整性级别 RID：12288=HIGH(管理员)、8192=MEDIUM(普通)、4096=LOW、0=UNTRUSTED
    #只用于把「窗口还原不了」翻译成人话，任何一步失败返回-1，绝不影响动作主流程
    try:
        k32 = ctypes.windll.kernel32
        adv = ctypes.windll.advapi32
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        k32.CloseHandle.restype = wintypes.BOOL
        adv.OpenProcessToken.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.POINTER(wintypes.HANDLE)]
        adv.OpenProcessToken.restype = wintypes.BOOL
        adv.GetTokenInformation.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        adv.GetTokenInformation.restype = wintypes.BOOL
        adv.GetSidSubAuthorityCount.argtypes = [ctypes.c_void_p]
        adv.GetSidSubAuthorityCount.restype = ctypes.POINTER(ctypes.c_ubyte)
        adv.GetSidSubAuthority.argtypes = [ctypes.c_void_p, wintypes.DWORD]
        adv.GetSidSubAuthority.restype = ctypes.POINTER(wintypes.DWORD)

        process = k32.OpenProcess(0x1000, False, pid)  #0x1000 = PROCESS_QUERY_LIMITED_INFORMATION，跨完整性级别也能打开
        if not process:
            return -1
        try:
            token = wintypes.HANDLE()
            if not adv.OpenProcessToken(wintypes.HANDLE(process), 0x0008, ctypes.byref(token)):  #0x0008 = TOKEN_QUERY
                return -1
            try:
                need = wintypes.DWORD()
                adv.GetTokenInformation(token, 25, None, 0, ctypes.byref(need))  #25 = TokenIntegrityLevel，先探所需长度
                buf = ctypes.create_string_buffer(max(need.value, 16))
                if not adv.GetTokenInformation(token, 25, buf, need.value, ctypes.byref(need)):
                    return -1
                sid = ctypes.cast(buf, ctypes.POINTER(ctypes.c_void_p))[0]  #TOKEN_MANDATORY_LABEL.Label.Sid
                count = adv.GetSidSubAuthorityCount(sid)[0]
                if count <= 0:
                    return -1
                return adv.GetSidSubAuthority(sid, count - 1)[0]
            finally:
                k32.CloseHandle(token)
        finally:
            k32.CloseHandle(process)
    except Exception:  #诊断性质：环境不配合就放弃，绝不让诊断把动作搞挂
        return -1


def _integrity_label(rid):
    #把RID翻成日志里能看懂的说法
    if rid == 12288:
        return "HIGH(管理员)"
    if rid == 8192:
        return "MEDIUM(普通)"
    if rid == 4096:
        return "LOW"
    if rid == 0:
        return "UNTRUSTED"
    if rid < 0:
        return "UNKNOWN(读不到)"
    return f"rid={rid}"


def _find_window(user32, pattern):
    #按正则枚举可见顶层窗口，返回 (选中的hwnd, 候选清单)；没找到返回 (0, [])
    #评分键 (cls_match, fullmatch, area)：类名也命中正则的最优先 —— EnumWindows 枚举顺序无保证，
    #旧版取 found[0] 等于抽签，曾经抽到标题带「金蝶」的 Obsidian 笔记窗口
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)  #EnumWindows 回调必须返回 4 字节 BOOL；c_bool 是 1 字节，x64 下属于未定义行为
    def enum_windows_proc(hwnd, lparam):
        length = user32.GetWindowTextLengthW(hwnd)  #获取窗口标题长度
        if length > 0:
            buffer = ctypes.create_unicode_buffer(length + 1)  #创建缓冲区
            user32.GetWindowTextW(hwnd, buffer, length + 1)  #获取窗口标题
            title = buffer.value
            if re.search(pattern, title) and user32.IsWindowVisible(hwnd):  #匹配正则表达式且窗口可见
                cls_buf = ctypes.create_unicode_buffer(256)  #类名缓冲区
                user32.GetClassNameW(hwnd, cls_buf, 256)  #窗口类名
                rect = wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)  #窗口面积（最小化时是小几何）
                found.append({
                    "hwnd": hwnd,
                    "cls": cls_buf.value,
                    "area": area,
                    "fullmatch": re.fullmatch(pattern, title) is not None,  #标题整体匹配（而非被长标题蹭到）
                    "cls_match": re.search(pattern, cls_buf.value) is not None,  #类名也命中正则
                })
        return True  #收全候选再评分，不能命中就停
    user32.EnumWindows(enum_windows_proc, None)  #开始遍历所有窗口
    if not found:
        return 0, []
    best = max(found, key=lambda item: (item["cls_match"], item["fullmatch"], item["area"]))
    return best["hwnd"], found


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
        k32 = ctypes.windll.kernel32  #进程/线程相关

        hwnd, candidates = _find_window(user32, pattern)
        if not hwnd:
            return False  #没找到匹配的窗口（有可能是没打开），返回False

        if len(candidates) > 1:
            #正则命中的窗口不止一个时把候选都打出来：选错窗口这类问题靠这行日志定位
            for item in candidates:
                print(f"[Win32BringToFront] 候选窗口 hwnd={item['hwnd']} cls_hit={item['cls_match']} "
                      f"fullmatch={item['fullmatch']} area={item['area']} cls={item['cls'][:48]}")

        #模板和 ROI 是按「最大化窗口」的尺寸标定的，任何主动还原/最大化都可能改掉窗口尺寸
        iconic_at_entry = bool(user32.IsIconic(hwnd))
        if iconic_at_entry:
            user32.ShowWindow(hwnd, 9)  #9 = SW_RESTORE；最大化后被最小化的窗口会还原回最大化
            time.sleep(0.1)

        #ShowWindow 对更高完整性级别的目标会被 UIPI 静默拒绝：不报错、什么都不发生。
        #窗口若还原不了，ScreenDC 抓到的就是 158x26 的最小化几何再被放大成的假图，
        #后面的识别只会全部超时 —— 与其拖 20 秒，不如当场把原因说清楚并返回False
        if user32.IsIconic(hwnd):
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))  #窗口所属进程
            rid = _integrity_rid(pid.value)
            mine = _integrity_rid(k32.GetCurrentProcessId())  #本进程（agent）的完整性级别
            print(f"[Win32BringToFront] 窗口还原失败：仍处于最小化（hwnd={hwnd}）。"
                  f"目标进程完整性={_integrity_label(rid)}，本进程完整性={_integrity_label(mine)}；"
                  f"低级别进程改不了高级别进程的窗口（UIPI），这不是重试能解决的。"
                  f"若这个窗口是金蝶客户端：退出后不要用桌面/开始菜单快捷方式重启"
                  f"（快捷方式走 Shell.exe，管理员账户下会自动提权），"
                  f"直接运行 Kingdee.BOS.XPF.App.exe 即可")
            return False

        #本动作原先一行日志都没有，窗口出问题只能靠 on_error 截图反推，补上状态快照
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        print(f"[Win32BringToFront] hwnd={hwnd} iconic_at_entry={iconic_at_entry} "
              f"zoomed_now={bool(user32.IsZoomed(hwnd))} "
              f"rect={rect.left},{rect.top},{rect.right},{rect.bottom} "
              f"size={rect.right - rect.left}x{rect.bottom - rect.top}")

        if user32.GetForegroundWindow() == hwnd:  #已经在前台就到此为止，多余动作只会引入副作用（同 _activate_window 原则）
            return True

        #改为最多 3 轮阶梯重试，任一轮到位即返回，全部落空才报 False 让 pipeline 走 on_error。
        our_tid = k32.GetCurrentThreadId()
        target_tid = user32.GetWindowThreadProcessId(hwnd, None)  #目标窗口所属线程

        def activate(attach_tid):
            #把「当前拥有前台资格的线程」和「目标窗口线程」一起挂到自己的输入队列上再置顶。
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


def _prepare_user32(user32):
    #给本模块用到的 user32 函数补上 argtypes/restype：64 位下 hwnd 超过 int 范围时，
    #不声明就会被 ctypes 截断成 32 位（现有代码靠 hwnd 数值小侥幸没炸，新代码不赌这个）
    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.GetClientRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.RECT)]
    user32.GetClientRect.restype = wintypes.BOOL
    user32.IsIconic.argtypes = [ctypes.c_void_p]
    user32.IsIconic.restype = wintypes.BOOL
    user32.IsWindowVisible.argtypes = [ctypes.c_void_p]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.ClientToScreen.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
    user32.ClientToScreen.restype = wintypes.BOOL
    user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
    user32.SetCursorPos.restype = wintypes.BOOL
    user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
    user32.GetCursorPos.restype = wintypes.BOOL
    user32.WindowFromPoint.argtypes = [wintypes.POINT]
    user32.WindowFromPoint.restype = ctypes.c_void_p
    user32.GetAncestor.argtypes = [ctypes.c_void_p, wintypes.UINT]
    user32.GetAncestor.restype = ctypes.c_void_p
    user32.BringWindowToTop.argtypes = [ctypes.c_void_p]
    user32.BringWindowToTop.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [ctypes.c_void_p]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
    user32.AttachThreadInput.restype = wintypes.BOOL
    user32.SwitchToThisWindow.argtypes = [ctypes.c_void_p, wintypes.BOOL]  #无返回值，restype 不用动
    user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(_INPUT), ctypes.c_int]
    user32.SendInput.restype = wintypes.UINT


def _ensure_dpi_aware(user32):
    #让本线程按物理像素读窗口坐标。agent 进程默认是 DPI UNAWARE，150% 缩放的机器上
    #GetClientRect 会返回逻辑像素（1942 会读成 1294），和真实屏幕像素对不上就会静默点偏。
    #线程级设置随时可调、进程级只能设一次，两个都试，失败不抛（老系统没这些 API）
    try:
        user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(-4))  #-4 = PER_MONITOR_AWARE_V2
    except Exception:  #诊断性质：环境不配合就认了，尺寸对不上时日志会露馅
        pass
    global _DPI_READY
    if not _DPI_READY:
        try:
            user32.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
            user32.SetProcessDpiAwarenessContext.restype = wintypes.BOOL
            user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        except Exception:
            pass
        _DPI_READY = True


def _bring_to_front(user32, k32, hwnd):
    #把目标窗口抢到前台并复查。挂到「当前前台窗口线程」+「目标窗口线程」的输入队列上再置顶，
    #用完立刻解绑（挂着不放会连带卡住输入）。第 3 轮用 SwitchToThisWindow 兜底 ——
    #它是 Alt+Tab 的内部实现，不受前台锁限制，而且不像模拟 Alt 会把金蝶带进菜单加速键模式
    our_tid = k32.GetCurrentThreadId()
    target_tid = user32.GetWindowThreadProcessId(hwnd, None)
    for round_no in range(1, 4):
        if user32.GetForegroundWindow() == hwnd:
            return True
        if round_no <= 2:
            attach_tid = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
            if attach_tid:
                user32.AttachThreadInput(our_tid, attach_tid, True)
            if target_tid:
                user32.AttachThreadInput(our_tid, target_tid, True)
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            if target_tid:
                user32.AttachThreadInput(our_tid, target_tid, False)
            if attach_tid:
                user32.AttachThreadInput(our_tid, attach_tid, False)
        else:
            user32.SwitchToThisWindow(hwnd, True)
        time.sleep(0.15)  #给窗口切换一点时间
    return user32.GetForegroundWindow() == hwnd


def _as_ms(value, default):
    #毫秒参数的宽松解析：没给/给了非整数就用默认值，别让一个笔误把动作搞成不点击
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return default
    if value < 0:
        return default
    return int(value)


@AgentServer.custom_action("SynthClick")  #合成鼠标点击（先抢弹窗前台再点击）
class SynthClick(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)  #拉取参数；参数写坏打日志返回False，不在agent进程里抛异常
        if param is None:
            return False
        pattern = param.get("window_regex", "")  #基准窗口：图像坐标按它的客户区换算
        if not pattern:
            print("[SynthClick] 必须给window_regex参数（ROI是按这个窗口的客户区标定的）")
            return False
        activate_pattern = param.get("activate_regex") or pattern  #要点到前台的窗口，默认就是基准窗口
        move_ms = _as_ms(param.get("move_ms"), _DEFAULT_MOVE_MS)  #移动到按下之间的等待
        hold_ms = _as_ms(param.get("hold_ms"), _DEFAULT_HOLD_MS)  #按住时长
        settle_ms = _as_ms(param.get("settle_ms"), _DEFAULT_SETTLE_MS)  #抬起后的收尾等待

        user32 = ctypes.windll.user32
        k32 = ctypes.windll.kernel32
        _prepare_user32(user32)
        _ensure_dpi_aware(user32)

        base_hwnd, _ = _find_window(user32, pattern)
        if not base_hwnd:
            print(f"[SynthClick] 找不到基准窗口（window_regex={pattern}）：窗口没开，或者标题对不上")
            return False
        target_hwnd, _ = _find_window(user32, activate_pattern)
        if not target_hwnd:
            print(f"[SynthClick] 找不到要点窗口（activate_regex={activate_pattern}）：弹窗还没出现就被点，或者标题对不上")
            return False
        if user32.IsIconic(base_hwnd) or user32.IsIconic(target_hwnd):
            #最小化时客户区是 200x40 一类的假几何，换算出来的点毫无意义；还原是前置节点的职责
            print(f"[SynthClick] 窗口处于最小化（base={base_hwnd} target={target_hwnd}），"
                  f"客户区几何不可信，拒绝点击；请让前置的置顶节点先还原窗口")
            return False

        #本动作存在的全部理由：点击前必须让「弹窗自己」在前台。maa 的 Seize 输入法在
        #touch_down/touch_up 里各把主窗拉一次前台，弹窗被推成非前台后，WPF 会把这一次点击
        #当成「激活点击」吃掉 —— 「确定」连点 64 次毫无反应就是这么来的
        if not _bring_to_front(user32, k32, target_hwnd):
            fg = user32.GetForegroundWindow()
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(target_hwnd, ctypes.byref(pid))
            print(f"[SynthClick] 抢前台失败：foreground={fg} target={target_hwnd}（{_integrity_label(_integrity_rid(pid.value))}），"
                  f"本进程={_integrity_label(_integrity_rid(k32.GetCurrentProcessId()))}；"
                  f"目标级别更高时是 UIPI 拦截，低级别进程拉不动高级别窗口，这不是重试能解决的")
            return False

        image_x = param.get("x")
        image_y = param.get("y")
        if image_x is None or image_y is None:
            box = argv.box  #节点自己的识别框（Custom 动作拿得到），省略 x/y 时用框中心
            if box is None or box.w <= 0 or box.h <= 0:
                print("[SynthClick] 既没给x/y，前序识别框也是空的；请在参数里显式给图像坐标")
                return False
            image_x = box.x + box.w / 2
            image_y = box.y + box.h / 2

        #图像→客户区：maa 把客户区按短边归一到 720（本机 1942x1150 → 1216x720），pipeline 的 ROI 都按它标定
        client_rect = wintypes.RECT()
        user32.GetClientRect(base_hwnd, ctypes.byref(client_rect))
        client_w = client_rect.right - client_rect.left
        client_h = client_rect.bottom - client_rect.top
        if min(client_w, client_h) < _MIN_CLIENT_SIDE:
            print(f"[SynthClick] 基准窗口客户区只有 {client_w}x{client_h}，几何不可信（窗口没还原完），拒绝点击")
            return False
        scale = min(client_w, client_h) / _IMAGE_SHORT_SIDE
        client_x = int(round(image_x * scale))
        client_y = int(round(image_y * scale))
        if not (0 <= client_x < client_w and 0 <= client_y < client_h):
            print(f"[SynthClick] 图像坐标({image_x},{image_y})换算后落在客户区外：client=({client_x},{client_y})，"
                  f"客户区={client_w}x{client_h}；参数或识别框给错了")
            return False
        origin = wintypes.POINT(0, 0)
        if not user32.ClientToScreen(base_hwnd, ctypes.byref(origin)):
            print("[SynthClick] ClientToScreen失败，拿不到客户区原点")
            return False
        screen_x = origin.x + client_x
        screen_y = origin.y + client_y

        #落点自查：那一点最上层必须是目标窗口。ScreenDC 抓的是屏幕像素，被别的窗口盖住时
        #识别照样命中、点击却落到别人身上 —— 与其静默点歪，不如当场报出来
        hit_hwnd = user32.WindowFromPoint(wintypes.POINT(screen_x, screen_y))
        hit_root = user32.GetAncestor(hit_hwnd, 2) if hit_hwnd else 0  #2 = GA_ROOT
        target_root = user32.GetAncestor(target_hwnd, 2)
        if hit_root != target_root:
            print(f"[SynthClick] 落点被别的窗口挡着：image=({image_x:.0f},{image_y:.0f}) screen=({screen_x},{screen_y}) "
                  f"最上层hwnd={hit_hwnd} 目标hwnd={target_hwnd}；点了也是点到别人身上，拒绝盲点")
            return False

        #两步定位：先退 2px 再到位，保证窗口一定收到一次「位置变化」的移动消息（上一次点完光标就在原地时，
        #原地 SetCursorPos 不会产生新的 WM_MOUSEMOVE，XPF 的 hover/命中态建立不起来）
        user32.SetCursorPos(screen_x - 2, screen_y - 2)
        time.sleep(0.03)
        if not user32.SetCursorPos(screen_x, screen_y):
            print(f"[SynthClick] SetCursorPos失败：({screen_x},{screen_y})")
            return False
        time.sleep(move_ms / 1000)  #留出 hover 建立时间
        now = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(now))
        if now.x != screen_x or now.y != screen_y:
            #光标被抢或被别的程序挪走时，后续按下就落到别的控件上，宁可失败也不能乱点
            print(f"[SynthClick] 光标没到位：期望({screen_x},{screen_y}) 实际({now.x},{now.y})，拒绝点击")
            return False

        if not _send_mouse(user32, _MOUSEEVENTF_LEFTDOWN):
            print("[SynthClick] 左键按下没发出去（SendInput返回0）：目标进程完整性更高、或处于安全桌面时会被拦")
            return False
        time.sleep(hold_ms / 1000)
        if not _send_mouse(user32, _MOUSEEVENTF_LEFTUP):
            print("[SynthClick] 左键抬起没发出去，可能已经在别的窗口上按着左键，请手动点一下桌面复位")
            return False
        time.sleep(settle_ms / 1000)

        still_front = user32.GetForegroundWindow() == target_hwnd
        print(f"[SynthClick] 点击完成 image=({image_x:.0f},{image_y:.0f}) client=({client_x},{client_y}) "
              f"screen=({screen_x},{screen_y}) target={target_hwnd} 截图基准={client_w}x{client_h} "
              f"点击后仍在前台={still_front}")
        return True


def _send_mouse(user32, flags):
    #发一个「不带坐标」的鼠标事件：位置由 SetCursorPos 定位，按下/抬起只按当前光标位置生效。
    #绝对坐标那套（MOUSEEVENTF_ABSOLUTE + 0-65535 归一化）在 DPI 缩放下容易静默算偏，这里不用
    item = _INPUT()
    item.type = 0  #0 = INPUT_MOUSE
    item.mi.dx = 0
    item.mi.dy = 0
    item.mi.mouseData = 0
    item.mi.dwFlags = flags
    item.mi.time = 0
    item.mi.dwExtraInfo = None
    return user32.SendInput(1, ctypes.byref(item), ctypes.sizeof(_INPUT)) == 1


@AgentServer.custom_recognition("ErrorPopupOpen")  #金蝶错误弹窗（粘贴异常就是它）还挂着吗
class ErrorPopupOpen(CustomRecognition):

    def analyze(
        self,
        context: Context,
        argv: CustomRecognition.AnalyzeArg,
    ):

        param = _params(argv, "custom_recognition_param", tag="ErrorPopupOpen")
        if param is None:
            return None  #参数写坏，不命中，让pipeline接着评估next里的下一个候选

        title = param.get("title", _ERROR_POPUP_TITLE)
        class_part = param.get("class_part", _ERROR_POPUP_CLASS)
        if not isinstance(title, str) or not title:
            print(f"[ErrorPopupOpen] title不是非空字符串: {title!r}")
            return None
        if not isinstance(class_part, str) or not class_part:
            print(f"[ErrorPopupOpen] class_part不是非空字符串: {class_part!r}")
            return None

        user32 = ctypes.windll.user32
        _prepare_user32(user32)  #补上argtypes/restype，64位下hwnd超过int范围时不会被截断

        #re.escape：title 是给人照着标题栏抄的，里面万一有括号、点这类正则元字符，不能当成模式去解释
        hwnd, candidates = _find_window(user32, re.escape(title))
        if not candidates:
            print(f"[ErrorPopupOpen] 没有标题为「{title}」的可见窗口，不命中")
            return None

        #_find_window 的评分只管「类名是否也命中同一个正则」，管不住别的程序也叫这个名字，
        #所以这里自己再筛一遍：标题必须整体相等（fullmatch），类名必须带金蝶的标识
        mine = [item for item in candidates if item["fullmatch"] and class_part in item["cls"]]
        if not mine:
            print(f"[ErrorPopupOpen] 标题「{title}」的窗口不属于金蝶（候选: "
                  f"{[(item['hwnd'], item['fullmatch'], item['cls'][:40]) for item in candidates]}），不命中")
            return None

        #candidates 是 EnumWindows 的顺序，就是从顶到下的 Z 序；模态弹窗一定在最上面，取第一个即可。
        #真同时挂着好几个错误弹窗，关掉一个之后 inverse 校验还会看到剩下的，pipeline 自己会再兜一轮
        best = mine[0]
        print(f"[ErrorPopupOpen] 命中「{title}」hwnd={best['hwnd']} cls={best['cls'][:48]} "
              f"同类候选{len(mine)}个")
        return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1), detail={"hwnd": best["hwnd"], "title": title})

