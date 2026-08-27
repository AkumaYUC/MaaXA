#调拨单填写自定义动作
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名）：
#  ReadTransferExcel  param: {"path": "D:/xxx/调拨单.xlsx"}  读Excel建单据队列，放流程最开头
#  NextOrder          param: 无  取下一张单（第一张也靠它取）；全部填完返回False触发on_error，可接收尾节点
#  PasteOrderField    param: {"field": "source_org", "select_all": true}  当前单单头字段 → 剪贴板 → Ctrl+V粘贴到焦点输入框
#                     select_all为true时先Ctrl+A全选，粘贴直接覆盖输入框里原有的文字
#field可选值：source_org调出库存组织 / target_org调入库存组织 / transfer_type调拨类型 / remark备注
#  PasteOrderColumn   param: {"column": "material_codes", "method": "ctrl_v"}  当前单指定列拼成多行文本 → 剪贴板 → 粘贴
#                     前置pipeline节点必须已把光标点在目标列第一个单元格，金蝶收到多行粘贴会自动扩展行
#                     method可选: "ctrl_v"(默认)Ctrl+V粘贴 / "clipboard_only"只复制到剪贴板不粘贴（配合pipeline右键粘贴）
#column可选值：material_codes物料编码 / quantities数量 / source_warehouses调出仓库 / source_locations调出仓位
#             target_warehouses调入仓库 / target_locations调入仓位
#  PasteText          param: {"text": "直接调拨单列表"}  固定文本 → 剪贴板 → Ctrl+A全选后Ctrl+V覆盖粘贴（搜索框用）
#  PressCtrlHome      param: 无  发送Ctrl+Home，把表格光标移回首行首列（物料行数多时用来置顶）
#                     只发按键不点鼠标，焦点必须已经在表格里，前一个节点不能是点按钮之类会抢焦点的动作
#循环结构参考：读取Excel → NextOrder → 点新增 → 填单/保存/提交 → NextOrder（还有单回到填单，没有了on_error收尾）

import ctypes
import json
import time

import pyperclip

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction

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


#Ctrl+A / Ctrl+V / Ctrl+Home用到的虚拟键码（_ctrl_a、_ctrl_v、_ctrl_home发按键用）
_VK_CONTROL = 0x11
_VK_A = 0x41  #字母A的虚拟键码
_VK_V = 0x56  #字母V的虚拟键码
_VK_HOME = 0x24  #Home键的虚拟键码
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


@AgentServer.custom_action("ReadTransferExcel")  #读取调拨单Excel，解析成单据队列
class ReadTransferExcel(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        global _orders, _current

        path = json.loads(argv.custom_action_param or "{}").get("path", "")  #拉取Excel路径参数
        if not path:
            print("[ReadTransferExcel] 缺少path参数")
            return False  #没配置path，直接返回False

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
        if not any(lines):
            print(f"[PasteOrderColumn] {column}整列为空，跳过粘贴")
            return True  #此列无数据，无需粘贴
        pyperclip.copy("\r\n".join(lines))  
        #用CRLF拼接，和Excel复制出来的格式一致（金蝶块粘贴只认CRLF切行）


        method = param.get("method", "ctrl_v")
        if method == "clipboard_only":
            print(f"[PasteOrderColumn] 已复制{column}到剪贴板: {len(lines)}行（等待右键粘贴）")
            return True  #只复制不粘贴，pipeline侧负责右键+点粘贴

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
