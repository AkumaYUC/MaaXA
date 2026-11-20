#调拨单Excel读取器（严格模式，两套单据共用一套校验骨架）
#表头必须精确匹配模板列名，缺列/重复列直接报错 —— 填ERP单据的场景下猜错列比报错危险得多
#分隔规则：数据行之间空一行 = 分隔两张单据
#单头字段（组织/申请人这类整单只有一个值的列）由填表人写在首行，本模块原样读取、不做任何转换（不推导、不映射、不补零、不改大小写）
#所以这份代码里不含任何公司专有数据，换一家公司只用改Excel，不用碰代码
#
#模板（pipeline侧 ReadTransferExcel 的 template 参数，缺省 transfer 保持旧写法可用）：
#  transfer     直接调拨单
#  requisition  调拨申请单
#每张单的列写在 TEMPLATES 里，一行一列：Excel表头 → 内部字段名 → 报错用中文名 → 是否单头字段 → 是否必填
#新增一张单据只在这张表里加行，下面的校验/分组/组装逻辑不用动（三套平行字典的写法已被否掉：
#字段名在 COLUMNS / _ORDER_LEVEL / _REQUIRED 里各写一遍，改一处漏两处正是这类代码的出事方式）

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl

#调拨类型的合法值，必须与金蝶下拉项一字不差
_TRANSFER_TYPES = ("组织内调拨", "跨组织调拨")


def _c(excel, key, label, order=False, required=False):
    #一列的声明：Excel表头 / 内部字段名 / 报错用中文名 / 整单以首行为准 / 首行必填
    return {"excel": excel, "key": key, "label": label, "order": order, "required": required}


TEMPLATES = {
    "transfer": [  #直接调拨单：组织是单头字段，调拨类型由两个组织是否相同推导
        _c("物料编码", "material_code", "物料编码"),
        _c("数量", "quantity", "数量"),
        _c("调出仓库", "source_warehouse", "调出仓库", required=True),
        _c("调出仓位", "source_location", "调出仓位"),
        _c("调入仓库", "target_warehouse", "调入仓库", required=True),
        _c("调入仓位", "target_location", "调入仓位"),
        _c("调出库存组织", "source_org", "调出库存组织", order=True, required=True),
        _c("调入库存组织", "target_org", "调入库存组织", order=True, required=True),
        _c("备注", "remark", "备注", order=True),  #默认为*
    ],
    "requisition": [  #调拨申请单：组织是明细列（金蝶里每行都有，但整单填同一个值），申请组织/申请人才是单头
        _c("物料编码", "material_code", "物料编码"),
        _c("数量", "quantity", "数量"),
        _c("调出组织", "source_org", "调出组织", order=True, required=True),
        _c("调出仓库", "source_warehouse", "调出仓库"),
        _c("调出仓位", "source_location", "调出仓位"),
        _c("调入组织", "target_org", "调入组织", order=True, required=True),
        _c("调入仓库", "target_warehouse", "调入仓库"),
        _c("调入仓位", "target_location", "调入仓位"),
        _c("申请组织", "apply_org", "申请组织", order=True, required=True),
        _c("调拨类型", "transfer_type", "调拨类型", order=True, required=True),
        _c("调拨申请类型", "apply_type", "调拨申请类型", order=True),  #组合框下拉，每个账套选项不同；留空=保留金蝶默认
        _c("申请人", "applicant", "申请人", order=True, required=True),
        _c("联系电话", "phone", "联系电话", order=True),
        _c("通讯地址", "address", "通讯地址", order=True),
        _c("联系人", "contact", "联系人", order=True),
        _c("需求日期（当天）", "demand_date", "需求日期", order=True),  #留空 = 当天，由代码补
    ],
}


@dataclass
class OrderItem:
    #一行物料明细（对应Excel一行）。单头字段也挂在行上，因为Excel里它们就写在某一行的单元格里，
    #整单取哪个值由模板的 order 标记决定；本模板没涉及的列留空串
    material_code: str = ""  #物料编码
    quantity: str = ""  #数量（如 33）
    source_warehouse: str = ""  #调出仓库
    source_location: str = ""  #调出仓位
    target_warehouse: str = ""  #调入仓库
    target_location: str = ""  #调入仓位
    source_org: str = ""  #调出组织（直接调拨单表头叫「调出库存组织」，申请单叫「调出组织」，同一个字段）
    target_org: str = ""  #调入组织
    remark: str = ""  #备注（直接调拨单专用，默认*）
    apply_org: str = ""  #申请组织（申请单专用）
    transfer_type: str = ""  #调拨类型（申请单读Excel原文，直接调拨单由组织推导）
    apply_type: str = ""  #调拨申请类型（申请单专用，组合框；留空=保留金蝶默认）
    applicant: str = ""  #申请人（申请单专用）
    phone: str = ""  #联系电话（申请单专用）
    address: str = ""  #通讯地址（申请单专用）
    contact: str = ""  #联系人（申请单专用）
    demand_date: str = ""  #需求日期（申请单专用，留空补当天）
    row_index: int = 0  #Excel原始行号，报错定位用


@dataclass
class TransferOrder:
    #一张调拨单（空行分隔出的一组连续数据行）
    items: list  #这张单的所有物料行
    index: int = 0  #第几张单（从1开始）
    template: str = "transfer"  #本单来自哪套模板
    source_org: str = ""  #调出组织（用户填的原文，原样透传）
    target_org: str = ""  #调入组织（用户填的原文，原样透传）
    remark: str = ""  #单据备注，直接调拨单用，默认为*
    apply_org: str = ""  #申请组织
    transfer_type: str = ""  #调拨类型
    apply_type: str = ""  #调拨申请类型（申请单专用，组合框）
    applicant: str = ""  #申请人
    phone: str = ""  #联系电话
    address: str = ""  #通讯地址
    contact: str = ""  #联系人
    demand_date: str = ""  #需求日期


class ExcelFormatError(Exception):
    #Excel文件不存在或格式不符合预期
    pass


def _norm_cell(value):
    #单元格值 → 规范化字符串：None → 空串；整数值的浮点（如33.0）→ "33"；日期 → 2026/11/9；其余去首尾空白
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))  #33.0 → "33"，不处理的话粘贴到金蝶就成33.0了
    if isinstance(value, (datetime, date)):  #Excel里填成日期格式的单元格，openpyxl给的是datetime对象
        return f"{value.year}/{value.month}/{value.day}"  #str()出来会是"2026-11-09 00:00:00"，金蝶日期框不认
    return str(value).strip()


def _today_text():
    #当天日期，写成金蝶日期框认的 2026/11/9（月、日不补零，和界面显示一致）
    d = date.today()
    return f"{d.year}/{d.month}/{d.day}"


def _template_keys(cols):
    #列声明表 → 这张模板用到的 (Excel表头→字段名) 与 (字段名→中文名) 两张查表
    return {c["excel"]: c["key"] for c in cols}, {c["key"]: c["label"] for c in cols}


def _build_order(index, items, cols, template):
    #组装单据：order标记的字段以首行为准，其余行留空是正常填法，填了就必须和首行一模一样
    first = items[0]
    order_cols = [c for c in cols if c["order"]]
    order = TransferOrder(items=items, index=index, template=template)

    for c in order_cols:
        setattr(order, c["key"], getattr(first, c["key"]))

    if template == "transfer":
        #备注取整单第一个非空（直接调拨单没写备注就默认*；申请单模板里没有备注列，保持空串）
        order.remark = next((it.remark for it in items if it.remark), "") or "*"
        #这两个值必须与金蝶「调拨类型」下拉项一字不差，文案有变时这里和管道都要同步改
        order.transfer_type = "组织内调拨" if order.source_org == order.target_org else "跨组织调拨"

    if template == "requisition" and not order.demand_date:
        order.demand_date = _today_text()  #Excel列名写着（当天），留空就是今天

    #必填看所有列：直接调拨单的仓库是明细列但整单一致（order=False, required=True），
    #只遍历 order_cols 会把它漏掉 —— 旧版是直接检查 first.source_warehouse 的，别把这条校验写没了
    for c in cols:
        if c["required"] and not getattr(first, c["key"]):
            raise ExcelFormatError(f"第 {first.row_index} 行：{c['label']}为空，每张单的首行必须填")

    if template == "requisition" and order.transfer_type not in _TRANSFER_TYPES:
        #模版单元格印着「组织内调拨/跨组织调拨」两个候选，照原样粘进金蝶必然失败，在这里就拦下
        raise ExcelFormatError(
            f"第 {first.row_index} 行：调拨类型「{order.transfer_type}」不是合法值，只能填 {' 或 '.join(_TRANSFER_TYPES)}")

    for item in items[1:]:  #其余行留空正常，填了就必须和首行一致；同一组织写成两种写法也会在这里被拦下
        for c in order_cols:
            value = getattr(item, c["key"])
            if value and value != getattr(order, c["key"]):
                raise ExcelFormatError(
                    f"第 {item.row_index} 行：{c['label']}「{value}」与首行「{getattr(order, c['key'])}」不一致，一张单只能有一个{c['label']}")
    return order


def read_orders(file_path, template="transfer"):
    #读取Excel，返回按空行分组的单据列表，格式不对抛ExcelFormatError
    if template not in TEMPLATES:
        raise ExcelFormatError(f"未知模板: {template!r}，可选: {list(TEMPLATES)}")
    cols = TEMPLATES[template]
    header_map, _ = _template_keys(cols)

    path = Path(file_path)
    if not path.exists():
        raise ExcelFormatError(f"文件不存在: {path}")  #文件都没得读，直接报错

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)  #data_only=True：读公式的计算结果而不是公式本身
    try:
        ws = wb.active  #取第一个工作表
        rows = list(ws.iter_rows(values_only=True))  #一次性读出所有行
    finally:
        wb.close()

    if not rows:
        raise ExcelFormatError("Excel是空的")

    #── 严格表头校验：标准列名必须全部找到，且不允许重复 ──
    header = rows[0]  #第一行是表头
    col_idx = {}  #内部字段名 → 列号
    for i, text in enumerate(header):
        name = _norm_cell(text)
        if name in header_map:  #这一列是标准列
            key = header_map[name]
            if key in col_idx:
                raise ExcelFormatError(f"表头第 {i + 1} 列「{name}」重复出现")  #同名列出现两次没法确定用哪个，报错
            col_idx[key] = i
    missing = [c["excel"] for c in cols if c["key"] not in col_idx]
    if missing:
        #缺列直接报错并把实际表头打出来，方便对着改Excel
        raise ExcelFormatError(f"表头缺少列: {', '.join(missing)}（实际读到的表头: {[str(t) for t in header]}）")

    #── 数据行逐行校验 + 按空行分组 ──
    orders = []
    current = []  #当前正在攒的单据

    for row_no, row in enumerate(rows[1:], start=2):  #从第2行开始，row_no是Excel实际行号
        values = {key: _norm_cell(row[col_idx[key]]) for key in header_map.values()}  #按列名取值，列在Excel里的先后顺序无所谓

        if not any(values.values()):  #整行为空 = 单据分隔符
            if current:
                orders.append(_build_order(len(orders) + 1, current, cols, template))  #收掉当前攒的单
                current = []
            continue

        item = OrderItem(**values, row_index=row_no)  #values 的键就是 OrderItem 的字段名（见列声明），整包递过去

        #逐行校验，报错带Excel行号，方便对着表改
        if not item.material_code:
            raise ExcelFormatError(f"第 {row_no} 行：物料编码为空")
        digits = item.quantity.replace(".", "", 1)  #允许小数，不允许其他字符
        if not digits.isdigit():
            raise ExcelFormatError(f"第 {row_no} 行：数量不是有效数字（读到 {item.quantity!r}）")
        if int(digits) == 0:  #"0"/"0.0"/"000" 都拦掉，0数量的明细录进ERP只是一行垃圾
            raise ExcelFormatError(f"第 {row_no} 行：数量必须大于0（读到 {item.quantity!r}）")

        current.append(item)

    if current:  #最后一组后面没有空行也要收尾
        orders.append(_build_order(len(orders) + 1, current, cols, template))

    if not orders:
        raise ExcelFormatError("没有读到任何数据行")
    return orders
