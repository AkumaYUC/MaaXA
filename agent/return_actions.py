#销售退换货单据任务自定义动作
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名）：
#  ReadBillNumbers    param: {"numbers": "{单据编码}"}  把用户输入的单据编码切成队列，放流程最开头
#                     numbers由interface.json的input型option传进来（用户在GUI里敲，不用上传文件）；
#                     换行/逗号/分号/顿号/空格都算分隔符 —— 批量时一行一个编码、或逗号隔开都行；
#                     队列内去重保序、且大小写不敏感（重复编码只跑一遍：样例单业务方已经推过一次，
#                     重复下推=造重复单据，这是本流程唯一会污染真实数据的地方）
#  NextBill           param: 无  取下一张单（第一张也靠它取）；全部跑完返回False触发on_error，收尾节点接WriteBillResult
#  PasteBillNumber    param: {"select_all": true}  当前单号 → 剪贴板 → Ctrl+V粘到焦点输入框（列表搜索框用）
#                     select_all默认就是true：搜索框里常有上一次搜索的残留，先Ctrl+A全选再粘、直接覆盖
#                     前置pipeline节点必须已把光标点在目标输入框里（粘的是焦点，不是眼里看到的位置）
#  VerifyCurrentBill  识别 param: 无  焦点字段的内容和当前单号一致时命中，不一致/读不到都不命中
#                     粘完单号后用它回读校验「真的粘上去了」：agent打「已粘贴」不代表粘上了，金蝶会吞键
#                     不命中时pipeline接着评估next里的兜底候选（重新粘贴 / 报错收尾）
#                     命中的框是占位框(0,0,1,1)，只配ClickKey这类不需要目标的动作
#  CaptureBillNumber  param: {"slot": "return_no", "pattern": "^[A-Z]{2,8}\\d{4,}$"}  读焦点字段 → 存进当前单
#                     slot可选值：return_no销售退货单号 / outbound_no销售出库单号，同一个slot重复读会覆盖前值
#                     pattern是形状校验，只挡「根本不是单号」的内容（标签文字/空值/乱码）：
#                     读到的若也是单号形状（比如焦点落到源单号列），这里区分不出来，靠pipeline保证焦点点对
#                     读法=剪贴板哨兵+Ctrl+A+Ctrl+C精确读取，不走OCR，没有识别错误和置信度问题
#                     前置pipeline节点必须已把光标点/选中在目标字段上
#  WriteBillResult    param: {"path": "D:/xxx/结果.txt", "require": ["return_no", "outbound_no"]}
#                     收尾落盘：一行一张单，制表符分列，UTF-8带BOM（记事本/Excel打开中文都不乱码）
#                     首行表头：源单号/销售退货单号/销售出库单号；整份覆盖写、不追加，同一批重跑不会翻倍
#                     require是可选的校验项，默认两个号都要有；配件单这类确实没有出库单时写 {"require": ["return_no"]}
#                     写空数组[]也是按默认两项校验，不是「不校验」
#                     缺号时照样把文件写出来，但返回False让任务变红 —— 结果不完整不能装作没事
#循环结构参考：ReadBillNumbers → NextBill → 搜索单号/下推/填单 → CaptureBillNumber(退货单号)
#             → CaptureBillNumber(出库单号) → NextBill（还有单就回到搜索，没有了走on_error收尾）
#             → WriteBillResult
#底层复用：Ctrl组合键发送从transfer_actions导入，不在这里再写一份keybd_event

import re
from dataclasses import dataclass
from pathlib import Path

import pyperclip

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction
from maa.custom_recognition import CustomRecognition

from common import _params  #pipeline 参数解析兜底（写坏打日志返回None，不在agent进程里抛异常）
from transfer_actions import _chord, _ctrl_a, _ctrl_v  #Ctrl组合键发送（和调拨单共用一套实现）

#agent是独立进程，下面的模块级变量在同一进程的所有Action之间共享，用来记单据队列
_bills = []  #待处理单据队列（用户输入的单据编码）
_current = -1  #当前单据下标，-1 = 还没开始取

#CaptureBillNumber的slot参数白名单（键=参数值，值=中文名，日志和结果文件表头都用它）
_SLOTS = {
    "return_no": "销售退货单号",
    "outbound_no": "销售出库单号",
}

#读焦点字段前先塞进剪贴板的哨兵：Ctrl+C要是没生效就会原样读回它，一眼看出「复制根本没发生」
#没有这道的话，读到的会是上一次粘贴留下的旧值 —— 校验会「假通过」，比不做校验还坏
_SENTINEL = "__MAAXA_SENTINEL__"

#单号形状：2~8个大写字母 + 4位以上数字（如 AB202401010001 这种「字母前缀+日期+流水号」）
_DEFAULT_PATTERN = r"^[A-Z]{2,8}\d{4,}$"

#用户输入的分隔符：空白（含换行）/半角逗号/全角逗号/半角分号/全角分号/顿号
_SEPARATOR = re.compile(r"[\s,，;；、]+")

_VK_C = 0x43  #字母C的虚拟键码


@dataclass
class BillResult:
    #一张销售退换货单据的处理结果（源单号由用户输入，两个下游单号由流程回读填进来）
    bill_no: str  #源单号：用户输入的ERP退货单据编号，如 AB202401010001
    return_no: str = ""  #销售退货单号：负数量那批下推出来的
    outbound_no: str = ""  #销售出库单号：正数量那批下推出来的


def _current_bill():
    #取当前单据。ReadBillNumbers / NextBill 还没跑过（或已经取完）时返回None
    if not 0 <= _current < len(_bills):
        return None
    return _bills[_current]


def _norm(text):
    #归一化：全角转半角 + 转大写 + 去掉首尾空白。单号比对和落盘都走它
    #剪贴板读回来可能是None而不是异常：pyperclip底层的GlobalLock失败时判错条件不成立，直接返回None；
    #这里当成空串交给调用方判失败，不然None会一路抛穿maa的ctypes回调 —— 识别类抛异常框架当成「命中」
    if not isinstance(text, str):
        return ""
    #全角字符区(U+FF01~U+FF5E)与ASCII码相差0xFEE0，表意空格U+3000换成普通空格；
    #OA那边实测剪贴板拿到的是全角逗号/全角括号，不归一化会出现「明明一致却判不一致」；
    #转大写是因为金蝶单号本身就是大写：手输时敲成小写要能认出来是同一张，
    #否则「ab2024…」和「AB2024…」会被当成两张单各推一次 —— 那就是重复下推造重复单据
    chars = []
    for ch in text:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            chars.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            chars.append(" ")
        else:
            chars.append(ch)
    return "".join(chars).strip().upper()


def _read_focus_text(tag):
    #读当前焦点控件里的文字：先往剪贴板放哨兵，再Ctrl+A全选、Ctrl+C复制，最后读剪贴板。
    #读回哨兵说明Ctrl+C没生效（焦点不在输入框上），返回None让调用方判失败
    #tag是调用方的类名，日志前缀统一用[类名]——本函数被读取类和识别类共用，不能写死一个名字
    try:
        pyperclip.copy(_SENTINEL)
    except Exception as e:
        print(f"[{tag}] 剪贴板写入失败: {e}")
        return None
    _ctrl_a()  #全选，保证复制的是一整个字段，而不是光标后面那半截
    _chord(_VK_C)  #Ctrl+C复制（内部收尾等100ms，给金蝶写剪贴板的时间）
    try:
        text = pyperclip.paste()
    except Exception as e:
        print(f"[{tag}] 剪贴板读取失败: {e}")
        return None
    if text == _SENTINEL:
        print(f"[{tag}] 剪贴板哨兵原样读回，Ctrl+C没生效（焦点可能不在输入框上）")
        return None
    return _norm(text)


@AgentServer.custom_action("ReadBillNumbers")  #把用户输入的单据编码切成待处理队列
class ReadBillNumbers(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        global _bills, _current

        _bills = []  #先清空：本动作是流程开头，参数就算写坏了也不能让上一轮的队列留着继续跑
        _current = -1

        param = _params(argv)  #拉取参数
        if param is None:
            return False
        raw = param.get("numbers", "")  #用户敲进来的原文，GUI的input型option传进来
        if not isinstance(raw, str):
            print(f"[ReadBillNumbers] numbers必须是字符串: {raw!r}")
            return False  #参数类型不对

        parts = [part for part in _SEPARATOR.split(raw) if part]
        if not parts:
            print("[ReadBillNumbers] 没给单据编码")
            return False  #空输入，别拿空队列往下跑

        seen = set()
        for part in parts:
            bill_no = _norm(part)
            if bill_no in seen:  #重复输入只跑一遍（样例单已经推过，重复下推=造重复单据）
                print(f"[ReadBillNumbers] 重复编码已跳过: {bill_no}")
                continue
            seen.add(bill_no)
            _bills.append(BillResult(bill_no=bill_no))

        print(f"[ReadBillNumbers] 已建队列: {len(_bills)}张单据")
        for index, bill in enumerate(_bills, start=1):
            print(f"[ReadBillNumbers]   第{index}张: {bill.bill_no}")
        return True  #队列就绪


@AgentServer.custom_action("NextBill")  #推进到下一张单据，没有更多单时返回False（触发on_error走收尾）
class NextBill(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        global _current

        _current += 1  #指针后移一张
        if _current >= len(_bills):
            print("[NextBill] 所有单据已处理完毕")
            return False  #没有更多单了，让on_error接手（接WriteBillResult收尾）

        bill = _bills[_current]
        print(f"[NextBill] 当前第{_current + 1}/{len(_bills)}张单: {bill.bill_no}")
        return True  #取到单了，继续跑这张单的流程


@AgentServer.custom_action("PasteBillNumber")  #当前单号 → 剪贴板 → Ctrl+A后Ctrl+V粘贴到焦点输入框
class PasteBillNumber(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        bill = _current_bill()
        if bill is None:
            print("[PasteBillNumber] 没有当前单据（先执行ReadBillNumbers和NextBill）")
            return False  #队列还没就绪

        param = _params(argv)  #拉取参数
        if param is None:
            return False
        select_all = param.get("select_all")  #只有显式写false才不全选，不写/写null都按默认全选走
        if select_all is None or select_all:
            _ctrl_a()  #默认全选覆盖：搜索框里常有上一次搜索留下的残留
        try:
            pyperclip.copy(bill.bill_no)
        except Exception as e:
            print(f"[PasteBillNumber] 剪贴板写入失败: {e}")
            return False  #剪贴板被别的程序占着（Windows上它是全局锁），粘不出去就别装作粘上了
        _ctrl_v()
        print(f"[PasteBillNumber] 已粘贴单号: {bill.bill_no}")
        return True  #粘贴已发出，有没有粘上交给VerifyCurrentBill回读判定


@AgentServer.custom_action("CaptureBillNumber")  #读焦点字段 → 存进当前单（剪贴板精确读取，不走OCR）
class CaptureBillNumber(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        bill = _current_bill()
        if bill is None:
            print("[CaptureBillNumber] 没有当前单据（先执行ReadBillNumbers和NextBill）")
            return False  #队列还没就绪

        param = _params(argv)  #拉取参数
        if param is None:
            return False
        slot = param.get("slot", "")
        if slot not in _SLOTS:
            print(f"[CaptureBillNumber] 未知slot: {slot!r}，可选: {list(_SLOTS)}")
            return False  #slot拼错了

        text = _read_focus_text("CaptureBillNumber")  #读不到返回None；读到空字段返回空串，下面按同样口径处理
        if not text:
            print(f"[CaptureBillNumber] 没读到{_SLOTS[slot]}的内容，不记")
            return False  #读不到就别存，免得把空值当成「这张单没有这个号」

        pattern = param.get("pattern", _DEFAULT_PATTERN)  #形状校验用的正则，默认就是单号形状
        if not isinstance(pattern, str):
            print(f"[CaptureBillNumber] pattern必须是字符串: {pattern!r}")
            return False  #正则写坏了，不猜
        try:
            matched = re.fullmatch(pattern, text) is not None
        except re.error as e:
            print(f"[CaptureBillNumber] pattern不是有效正则: {e}")
            return False  #正则写坏了，不猜
        if not matched:
            print(f"[CaptureBillNumber] {_SLOTS[slot]}形状不符（读到 {text!r}，pattern={pattern}）")
            return False  #形状不对：十有八九焦点点错了字段，宁可任务红

        setattr(bill, slot, text)
        print(f"[CaptureBillNumber] 已记录{_SLOTS[slot]}: {text}")
        return True  #记录成功


@AgentServer.custom_action("WriteBillResult")  #收尾落盘：一行一张单的结果文件（整份覆盖写）
class WriteBillResult(CustomAction):

    def run(
        self,
        context: Context,
        argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)  #拉取参数
        if param is None:
            return False
        path = param.get("path", "")  #结果文件路径，由pipeline给
        if not isinstance(path, str) or not path:
            print(f"[WriteBillResult] path必须是非空字符串: {path!r}")
            return False  #没给路径或给错了（比如直接写了个数字）
        if not _bills:
            print("[WriteBillResult] 队列是空的，没有可写的结果")
            return False  #压根没跑过ReadBillNumbers，别写一份空文件出来

        require = param.get("require") or list(_SLOTS)  #哪些号必须读到，默认两个都要
        if not isinstance(require, list) or any(not isinstance(slot, str) or slot not in _SLOTS for slot in require):
            print(f"[WriteBillResult] require只能填{list(_SLOTS)}里的值: {require!r}")
            return False  #require写坏了（不是list、元素不是字符串、或slot名字不对），不猜

        lines = ["\t".join(["源单号"] + list(_SLOTS.values()))]  #首行表头
        for bill in _bills:
            lines.append("\t".join([bill.bill_no] + [getattr(bill, slot) for slot in _SLOTS]))

        try:
            #先在内存里编码成bytes再落盘：write_text是先打开文件截断、再逐块编码，
            #万一单号里有编码不了的字符，上一份跑出来的结果文件会被清成0字节 —— 结果文件不能这么丢；
            #utf-8-sig带BOM，记事本/Excel打开中文不乱码；写bytes不做换行翻译，\r\n原样落盘
            data = ("\r\n".join(lines) + "\r\n").encode("utf-8-sig")
            Path(path).write_bytes(data)
        except (OSError, ValueError) as e:
            print(f"[WriteBillResult] 写结果文件失败: {e}")
            return False  #写不进去就得报出来（目录不存在、路径带非法字符、编码失败都走这里）

        print(f"[WriteBillResult] 已写入{len(_bills)}张单的结果: {path}")
        missing = [
            f"{bill.bill_no}缺{_SLOTS[slot]}"
            for bill in _bills
            for slot in require
            if not getattr(bill, slot)
        ]
        if missing:
            print(f"[WriteBillResult] 有{len(missing)}个号没读到，结果不完整: {missing}")
            return False  #文件照写（数据不能丢），但任务变红提醒人来看
        return True  #收尾完成


@AgentServer.custom_recognition("VerifyCurrentBill")  #回读校验：焦点字段内容 == 当前单号时命中
class VerifyCurrentBill(CustomRecognition):

    def analyze(
        self,
        context: Context,
        argv: CustomRecognition.AnalyzeArg,
    ):

        bill = _current_bill()
        if bill is None:
            print("[VerifyCurrentBill] 没有当前单据（先执行ReadBillNumbers和NextBill）")
            return None  #队列还没就绪，不命中

        text = _read_focus_text("VerifyCurrentBill")
        if not text:
            print("[VerifyCurrentBill] 焦点字段读不到内容，不命中")
            return None  #读不到就不命中，让pipeline接着评估next里的兜底候选

        if text != bill.bill_no:
            print(f"[VerifyCurrentBill] 字段里是 {text!r}，当前单号是 {bill.bill_no!r}，不一致，不命中")
            return None  #没粘上或粘错了，让pipeline重新粘或收尾

        print(f"[VerifyCurrentBill] 当前单号 {bill.bill_no} 已在焦点字段里，命中")
        return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1), detail={"bill_no": bill.bill_no})  #占位框
