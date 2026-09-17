#调拨单填写自定义动作
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名）：
#  ReadTransferExcel  param: {"path": "D:/xxx/调拨单.xlsx"}  读Excel建单据队列，放流程最开头
#                     path是单据文件路径，GUI拖拽单据时由interface.json的option传进来（调试也可直接写死）
#                     单据里每个要填的值（含库存组织）都由填表人写进Excel，agent原样透传、不做任何转换，
#                     所以代码侧不含任何公司专有数据 —— 换一家公司只用改Excel，不用改代码
#                     文件归档由GUI在任务全部成功后负责（复制副本进归档目录），agent这边不碰文件
#  NextOrder          param: 无  取下一张单（第一张也靠它取）；全部填完返回False触发on_error，可接收尾节点
#  PasteOrderField    param: {"field": "source_org", "select_all": true}  当前单单头字段 → 剪贴板 → Ctrl+V粘贴到焦点输入框
#                     select_all为true时先Ctrl+A全选，粘贴直接覆盖输入框里原有的文字
#field可选值：source_org调出库存组织 / target_org调入库存组织 / transfer_type调拨类型 / remark备注
#  PasteOrderColumn   param: {"column": "material_codes", "method": "ctrl_v"}  当前单指定列拼成多行文本 → 剪贴板 → 粘贴
#                     前置pipeline节点必须已把光标点在目标列第一个单元格，金蝶收到多行粘贴会自动扩展行
#                     method可选: "ctrl_v"(默认)Ctrl+V粘贴 / "clipboard_only"只复制到剪贴板不粘贴（配合pipeline右键粘贴）
#                     first_only为true时只取第一行（仓库仓位这类整单一致的列，粘完再点金蝶的批量填充铺开）
#                     focus_window填窗口标题时，粘贴前先把该窗口置前台（维度数据录入这类独立弹窗必须加，
#                     框架每次抓图都会把金蝶主窗口拉回前台，焦点得在发Ctrl+V前抢回来，否则粘到主窗口表格里）
#column可选值：material_codes物料编码 / quantities数量 / source_warehouses调出仓库 / source_locations调出仓位
#             target_warehouses调入仓库 / target_locations调入仓位
#  PasteText          param: {"text": "直接调拨单列表"}  固定文本 → 剪贴板 → Ctrl+A全选后Ctrl+V覆盖粘贴（搜索框用）
#  PressCtrlHome      param: 无  发送Ctrl+Home，把表格光标移回首行首列（物料行数多时用来置顶）
#                     只发按键不点鼠标，焦点必须已经在表格里，前一个节点不能是点按钮之类会抢焦点的动作
#  PopupKeys          param: {"window": "维度数据录入", "key": [40, 13], "key_delay": 1500}  抢弹窗焦点后依次发送key里的按键
#                     key是虚拟键码数组（13回车 / 27ESC / 9Tab / 38↑ / 40↓），逐个按下再松开，不是组合键
#                     key_delay是每个按键之后等待的毫秒数，默认150；下拉列表要等后台读数据就填1500~2000
#                     字段名跟pipeline的ClickKey保持一致（都叫key、都收数组），别写成keys
#                     window要和弹窗标题栏文字精确一致（内部走FindWindowW精确匹配，不是正则）
#                     金蝶的查找字段常要两个回车（第一个提交字段值、第二个才按默认按钮），这时填 "key": [13, 13]
#                     弹窗内选下拉项用 ↓ + 回车 比鼠标双击稳得多（Click在识别框内随机取点，双击经常判不出来）
#                     独立顶层弹窗必须用这个动作发按键：框架每次抓图都会把金蝶主窗口拉回前台，
#                     抢焦点和发按键之间隔一个pipeline节点焦点就丢了，键会打到主窗口上去
#                     找不到窗口返回False（不乱发按键），让pipeline走on_error重来
#自定义识别（节点recognition填Custom、custom_recognition填名字；和上面的动作是两个互不相干的槽）：
#  OrderRowCount      param: {"equal": 1}  当前单的物料行数等于equal时命中，否则不命中
#                     行数直接读agent里的队列，跟界面无关，所以命中与否是瞬间确定的，
#                     不像弹窗判据那样有"还没渲染出来就被兜底候选抢跑"的时序竞态，
#                     可以放心和兜底候选写在同一个next列表里（本命中在前、兜底在后）
#                     命中时返回占位框(0,0,1,1)，只配ClickKey这类不需要目标的动作，别配Click+target
#循环结构参考：读取Excel → NextOrder → 点新增 → 填单/保存/提交 → NextOrder（还有单回到填单，没有了on_error收尾）

import ctypes
import json
import time

import pyperclip

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction
from maa.custom_recognition import CustomRecognition

import excel_reader

#agent是独立进程，下面的模块级变量在同一进程的所有Action之间共享，用来记单据队列
_orders = []  #解析后的单据队列
_current = -1  #当前单据下标，-1 = 还没开始取

#PasteOrderColumn的column参数 → OrderItem字段名
_COLUMN_FIELDS = {
    "material_codes": "material_code",
    "quantities": "quantity",
    "source_warehouses": "source_warehouse",
    "source_locations": "source_location",
    "target_warehouses": "target_warehouse",
    "target_locations": "target_location",
}

#PasteOrderField的field参数 → TransferOrder字段名
_ORDER_FIELDS = {
    "source_org": "source_org",
    "target_org": "target_org",
    "remark": "remark",
    "transfer_type": "transfer_type",
}


#Ctrl+A / Ctrl+V / Ctrl+Home / 置前台用到的虚拟键码（_ctrl_a、_ctrl_v、_ctrl_home、_focus_window发按键用）
_VK_CONTROL = 0x11
_VK_A = 0x41  #字母A的虚拟键码
_VK_V = 0x56  #字母V的虚拟键码
_VK_HOME = 0x24  #Home键的虚拟键码
_VK_MENU = 0x12  #Alt键的虚拟键码（置前台前按一下，绕开系统的前台锁）
_KEYEVENTF_KEYUP = 0x0002  #按键释放标志


@AgentServer.custom_action("PasteText")  #粘贴固定文本（pipeline调用）
class PasteText(CustomAction):
    def run(self, context, argv):
        text = json.loads(argv.custom_action_param or "{}").get("text", "")
        if not text:
            return False
        _ctrl_a()  #搜索框可能有残留文字，先全选覆盖
        pyperclip.copy(text)
        _ctrl_v()
        return True


@AgentServer.custom_action("PasteOrderField")  #当前单指定字段 → 文本 → 剪贴板 → Ctrl+V粘贴到焦点单元格
class PasteOrderField(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not 0 <= _current < len(_orders):
            print("[PasteOrderField] 没有当前单据（先执行ReadTransferExcel和NextOrder）")
            return False  #队列还没就绪

        param = json.loads(argv.custom_action_param or "{}")  #拉取参数
        field = param.get("field", "")
        if field not in _ORDER_FIELDS:
            print(f"[PasteOrderField] 未知field: {field!r}，可选: {list(_ORDER_FIELDS)}")
            return False  #field拼错了

        value = getattr(_orders[_current], field)  #取当前单这一字段的值
        if param.get("select_all"):  #输入框默认有文字时先全选，粘贴直接覆盖
            _ctrl_a()
        pyperclip.copy(value)
        _ctrl_v()
        print(f"[PasteOrderField] 已粘贴{field}: {value!r}")
        return True  #粘贴成功


def _ctrl_a():
    #发送Ctrl+A，全选当前焦点控件里的内容
    user32 = ctypes.windll.user32
    user32.keybd_event(_VK_CONTROL, 0, 0, 0)  #按下Ctrl
    user32.keybd_event(_VK_A, 0, 0, 0)  #按下A
    time.sleep(0.05)
    user32.keybd_event(_VK_A, 0, _KEYEVENTF_KEYUP, 0)  #松开A
    user32.keybd_event(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)  #松开Ctrl


def _ctrl_v():
    #发送Ctrl+V，把剪贴板内容粘贴进当前有焦点的控件（keybd_event用法同my_action的Win32BringToFront）
    user32 = ctypes.windll.user32
    user32.keybd_event(_VK_CONTROL, 0, 0, 0)  #按下Ctrl
    user32.keybd_event(_VK_V, 0, 0, 0)  #按下V
    time.sleep(0.05)
    user32.keybd_event(_VK_V, 0, _KEYEVENTF_KEYUP, 0)  #松开V
    user32.keybd_event(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)  #松开Ctrl
    time.sleep(0.1)  #给金蝶一点响应时间


def _ctrl_home():
    #发送Ctrl+Home，把表格光标移回首行首列（keybd_event能保证Ctrl一直按住，pipeline的ClickKey传数组是逐个点击）
    user32 = ctypes.windll.user32
    user32.keybd_event(_VK_CONTROL, 0, 0, 0)  #按下Ctrl
    user32.keybd_event(_VK_HOME, 0, 0, 0)  #按下Home
    time.sleep(0.05)
    user32.keybd_event(_VK_HOME, 0, _KEYEVENTF_KEYUP, 0)  #松开Home
    user32.keybd_event(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0)  #松开Ctrl
    time.sleep(0.1)  #给金蝶一点响应时间


def _focus_window(title):
    #按标题把窗口拉到前台。框架每次ScreenDC抓图都会ensure foreground把金蝶主窗口拉回前台，
    #「维度数据录入」这类独立顶层弹窗的焦点会被顺手抢走，所以发按键前必须自己抢回来
    #置顶必须和发按键在同一个动作里完成，中间插任何pipeline节点都会再抓一次图、焦点又丢
    user32 = ctypes.windll.user32
    hwnd = user32.FindWindowW(None, title)  #按窗口标题精确找，找不到返回0
    if not hwnd:
        return False  #弹窗没开，或者标题和传进来的不一致
    user32.keybd_event(_VK_MENU, 0, 0, 0)  #按下Alt，绕开系统前台锁（同my_action的Win32BringToFront）
    ok = user32.SetForegroundWindow(hwnd)  #把弹窗置前台
    user32.keybd_event(_VK_MENU, 0, _KEYEVENTF_KEYUP, 0)  #松开Alt
    time.sleep(0.1)  #给窗口切换一点时间
    return bool(ok)  #置顶成功与否


def _activate_window(title):
    #把弹窗激活，但全程不碰Alt。_focus_window靠按一下Alt绕系统前台锁，而Alt单击会让金蝶进菜单加速键模式，
    #紧跟的Enter就被当成「激活菜单项」——实测会弹出一个带「关闭」按钮的窗口，而目标弹窗根本没关
    #所以发按键的路径必须用这个：已经在前台就什么都不做，需要切换时借目标线程的输入队列绕前台锁
    user32 = ctypes.windll.user32
    hwnd = user32.FindWindowW(None, title)  #按窗口标题精确找，找不到返回0
    if not hwnd:
        return False  #弹窗没开，或者标题和传进来的不一致
    if user32.GetForegroundWindow() == hwnd:
        return True  #已经是前台窗口，多余的置顶动作只会引入副作用

    target_tid = user32.GetWindowThreadProcessId(hwnd, None)  #弹窗所属的线程
    our_tid = ctypes.windll.kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(our_tid, target_tid, True)  #挂到目标线程的输入队列上，SetForegroundWindow才不会被拒
    user32.BringWindowToTop(hwnd)
    ok = user32.SetForegroundWindow(hwnd)
    user32.AttachThreadInput(our_tid, target_tid, False)  #用完立刻解绑，挂着不放会连带卡住输入
    time.sleep(0.1)  #给窗口切换一点时间
    return bool(ok)  #激活成功与否


@AgentServer.custom_action("ReadTransferExcel")  #读取调拨单Excel，解析成单据队列
class ReadTransferExcel(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        global _orders, _current

        param = json.loads(argv.custom_action_param or "{}")  #拉取参数
        path = param.get("path", "")  #单据文件路径，GUI拖拽单据时由interface.json的option传进来
        if not path:
            print("[ReadTransferExcel] 没给path，不知道该读哪个文件")
            return False  #没给文件路径

        try:
            _orders = excel_reader.read_orders(path)  #读文件并校验，格式不对会抛ExcelFormatError
        except excel_reader.ExcelFormatError as e:
            print(f"[ReadTransferExcel] {e}")  #把具体原因打出来（缺哪列/哪行数据坏）
            return False
        except Exception as e:
            print(f"[ReadTransferExcel] 读取Excel失败: {e}")
            return False

        _current = -1  #重置队列，第一张单由NextOrder来取
        total_rows = sum(len(o.items) for o in _orders)
        print(f"[ReadTransferExcel] 解析成功: {len(_orders)}张单据 / 共{total_rows}行物料")
        return True  #读取成功


@AgentServer.custom_action("NextOrder")  #推进到下一张单据，没有更多单时返回False（触发on_error走收尾）
class NextOrder(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        global _current

        _current += 1  #指针后移一张
        if _current >= len(_orders):
            print("[NextOrder] 所有单据已处理完毕")
            return False  #没有更多单了，让on_error接手（比如接结束流程的节点）

        order = _orders[_current]
        print(f"[NextOrder] 当前第{order.index}张单，{len(order.items)}行物料")
        return True  #取到单了，继续填单流程


@AgentServer.custom_action("PasteOrderColumn")  #当前单指定列 → 多行文本 → 剪贴板 → Ctrl+V粘贴到焦点单元格
class PasteOrderColumn(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        if not 0 <= _current < len(_orders):
            print("[PasteOrderColumn] 没有当前单据（先执行ReadTransferExcel和NextOrder）")
            return False  #队列还没就绪

        param = json.loads(argv.custom_action_param or "{}")
        column = param.get("column", "")  #拉取列名参数
        field = _COLUMN_FIELDS.get(column)
        if not field:
            print(f"[PasteOrderColumn] 未知column: {column!r}，可选: {list(_COLUMN_FIELDS)}")
            return False  #column拼错了

        order = _orders[_current]
        lines = [getattr(item, field) for item in order.items]  #取当前单这一列的所有行
        if param.get("first_only"):  #整单四个仓库仓位列的值都一样，只粘第一行，剩下的交给金蝶的批量填充
            lines = lines[:1]
        if not any(lines):
            print(f"[PasteOrderColumn] {column}整列为空，跳过粘贴")
            return True  #此列无数据，无需粘贴
        pyperclip.copy("\r\n".join(lines))  #用CRLF拼接，和Excel复制出来的格式一致（金蝶块粘贴只认CRLF切行）

        method = param.get("method", "ctrl_v")
        if method == "clipboard_only":
            print(f"[PasteOrderColumn] 已复制{column}到剪贴板: {len(lines)}行（等待右键粘贴）")
            return True  #只复制不粘贴，pipeline侧负责右键+点粘贴

        focus_window = param.get("focus_window")  #独立弹窗要先把焦点抢回来，否则Ctrl+V会打到主窗口的表格上
        if focus_window and not _focus_window(focus_window):
            print(f"[PasteOrderColumn] 没找到窗口「{focus_window}」，跳过粘贴")
            return False  #弹窗没开就别乱粘，返回False让pipeline走on_error重来

        _ctrl_v()
        print(f"[PasteOrderColumn] 已粘贴{column}: {len(lines)}行")
        return True  #粘贴成功


@AgentServer.custom_action("PressCtrlHome")  #发送Ctrl+Home，把表格光标移回首行（pipeline置顶节点调用）
class PressCtrlHome(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        _ctrl_home()
        print("[PressCtrlHome] 已发送Ctrl+Home")
        return True  #按键已发出，是否真的置顶交给pipeline的校验节点判断


@AgentServer.custom_action("PopupKeys")  #把独立弹窗抢到前台后发按键（弹窗内回车确定、ESC取消）
class PopupKeys(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        param = json.loads(argv.custom_action_param or "{}")
        title = param.get("window", "")
        keys = param.get("key") or []  #虚拟键码数组，逐个按下再松开
        if not title or not keys:
            print(f"[PopupKeys] 缺少window或key参数: {param!r}")
            return False  #参数不全，不乱发按键

        if not _activate_window(title):
            print(f"[PopupKeys] 没找到窗口「{title}」或激活失败")
            return False  #弹窗没开就别发按键，让pipeline走on_error重来

        user32 = ctypes.windll.user32
        key_delay = param.get("key_delay", 150)  #每个按键之后等待的毫秒数
        for vk in keys:
            user32.keybd_event(vk, 0, 0, 0)  #按下
            time.sleep(0.05)
            user32.keybd_event(vk, 0, _KEYEVENTF_KEYUP, 0)  #松开
            time.sleep(key_delay / 1000)  #给金蝶响应时间，也把连续的两个键隔开
        print(f"[PopupKeys] 已向「{title}」发送按键: {keys}，间隔{key_delay}ms")
        return True  #按键已发出，弹窗有没有关掉交给pipeline的校验节点判断


@AgentServer.custom_recognition("OrderRowCount")  #按当前单的物料行数分流（单物料/多物料走不同分支）
class OrderRowCount(CustomRecognition):

    def analyze(
        self,
        context: Context,
        argv: CustomRecognition.AnalyzeArg,
    ):

        if not 0 <= _current < len(_orders):
            print("[OrderRowCount] 没有当前单据（先执行ReadTransferExcel和NextOrder）")
            return None  #队列还没就绪，不命中

        param = json.loads(argv.custom_recognition_param or "{}")
        want = param.get("equal")
        if want is None:
            print(f"[OrderRowCount] 缺少equal参数: {param!r}")
            return None  #参数不全，不命中

        rows = len(_orders[_current].items)
        if rows != want:
            print(f"[OrderRowCount] 当前单{rows}行，不等于{want}，不命中")
            return None  #行数不匹配，让pipeline接着评估next列表里的下一个候选

        print(f"[OrderRowCount] 当前单{rows}行，命中")
        return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1), detail={"rows": rows})  #占位框，动作是ClickKey不需要真目标
