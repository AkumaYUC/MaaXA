#调拨单Excel读取器（严格模式）
#表头必须精确匹配标准列名，缺列/重复列直接报错 —— 填ERP单据的场景下猜错列比报错危险得多
#分隔规则：数据行之间空一行 = 分隔两张单据

from dataclasses import dataclass
from pathlib import Path

import openpyxl

#标准表头（Excel第一行必须精确包含这些列名）→ 内部字段名
COLUMNS = {
    "物料编码": "material_code",
    "数量": "quantity",
    "调出仓库": "source_warehouse",
    "调出仓位": "source_location",
    "调入仓库": "target_warehouse",
    "调入仓位": "target_location",
    "备注": "remark",  #默认为*
}

ORG_RULES = [
    ("委外仓", "103"),  #调出/调入仓库尾缀 → 库存组织
    ("电子仓", "103"),
    ("研发仓", "102"),
]


def derive_org_code(warehouse_name):
    #按仓库尾缀推导库存组织，找不到返回空串
    for suffix, org_code in ORG_RULES:
        if warehouse_name.endswith(suffix):
            return org_code
    return ""


def _build_order(index, items):
    #组装单据：组织代码整单推导必须一致（金蝶组织是单头字段），备注取第一个非空（没有默认*）
    first = items[0]
    source_org = derive_org_code(first.source_warehouse)
    target_org = derive_org_code(first.target_warehouse)
    if not source_org:
        raise ExcelFormatError(f"第 {first.row_index} 行：调出仓库「{first.source_warehouse}」匹配不到库存组织")
    if not target_org:
        raise ExcelFormatError(f"第 {first.row_index} 行：调入仓库「{first.target_warehouse}」匹配不到库存组织")

    for item in items[1:]:  #其余行推导的组织必须和首行一致，混用直接报错
        s = derive_org_code(item.source_warehouse)
        t = derive_org_code(item.target_warehouse)
        if s != source_org:
            raise ExcelFormatError(f"第 {item.row_index} 行：调出仓库「{item.source_warehouse}」推导组织{s or '空'}与首行{source_org}不一致，一张单只能有一个调出库存组织")
        if t != target_org:
            raise ExcelFormatError(f"第 {item.row_index} 行：调入仓库「{item.target_warehouse}」推导组织{t or '空'}与首行{target_org}不一致，一张单只能有一个调入库存组织")

    remark = next((it.remark for it in items if it.remark), "")  #取整单第一个非空备注
    transfer_type = "组织内调拨" if source_org == target_org else "跨组织调拨"  #推导组织是否一致
    return TransferOrder(
        index=index,
        items=items,
        source_org=source_org,
        target_org=target_org,
        remark=remark or "*",  #整单都没备注默认填*
        transfer_type=transfer_type,  #调拨类型，两组织一致：组织内调拨；不一致：跨组织调拨
    )


@dataclass
class OrderItem:
    #一行物料明细（对应Excel一行）
    material_code: str  #物料编码（如 01-002-01565）
    quantity: str  #数量（如 33）
    source_warehouse: str  #调出仓库
    source_location: str  #调出仓位
    target_warehouse: str  #调入仓库
    target_location: str  #调入仓位
    remark: str            #备注
    row_index: int  #Excel原始行号，报错定位用


@dataclass
class TransferOrder:
    #一张调拨单（空行分隔出的一组连续数据行）
    source_org: str  #调出库存组织（按仓库尾缀推导，委外仓/电子仓为103，研发仓为102）
    target_org: str  #调入库存组织（按仓库尾缀推导，委外仓/电子仓为103，研发仓为102）
    remark: str  #单据备注，默认为*，单据第一行有文字用文字
    index: int  #第几张单（从1开始）
    items: list  #这张单的所有物料行
    transfer_type: str  #调拨类型，两组织一致：组织内调拨；不一致：跨组织调拨


class ExcelFormatError(Exception):
    #Excel文件不存在或格式不符合预期
    pass


def _norm_cell(value):
    #单元格值 → 规范化字符串：None → 空串；整数值的浮点（如33.0）→ "33"；其余去首尾空白
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))  #33.0 → "33"，不处理的话粘贴到金蝶就成33.0了
    return str(value).strip()


def read_orders(file_path):
    #读取Excel，返回按空行分组的单据列表，格式不对抛ExcelFormatError
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
        if name in COLUMNS:  #这一列是标准列
            key = COLUMNS[name]
            if key in col_idx:
                raise ExcelFormatError(f"表头第 {i + 1} 列「{name}」重复出现")  #同名列出现两次没法确定用哪个，报错
            col_idx[key] = i
    missing = [name for name, key in COLUMNS.items() if key not in col_idx]
    if missing:
        #缺列直接报错并把实际表头打出来，方便对着改Excel
        raise ExcelFormatError(f"表头缺少列: {', '.join(missing)}（实际读到的表头: {[str(t) for t in header]}）")

    #── 数据行逐行校验 + 按空行分组 ──
    orders = []
    current = []  #当前正在攒的单据

    for row_no, row in enumerate(rows[1:], start=2):  #从第2行开始，row_no是Excel实际行号
        values = [_norm_cell(row[col_idx[key]]) for key in COLUMNS.values()]  #按标准列顺序取值

        if not any(values):  #整行为空 = 单据分隔符
            if current:
                orders.append(_build_order(len(orders) + 1, current))  #收掉当前攒的单
                current = []
            continue

        item = OrderItem(
            material_code=values[0],
            quantity=values[1],
            source_warehouse=values[2],
            source_location=values[3],
            target_warehouse=values[4],
            target_location=values[5],
            remark=values[6],
            row_index=row_no,
        )

        #逐行校验，报错带Excel行号，方便对着表改
        if not item.material_code:
            raise ExcelFormatError(f"第 {row_no} 行：物料编码为空")
        if not item.quantity.replace(".", "", 1).isdigit():  #允许小数，不允许其他字符
            raise ExcelFormatError(f"第 {row_no} 行：数量不是有效数字（读到 {item.quantity!r}）")

        current.append(item)

    if current:  #最后一组后面没有空行也要收尾
        orders.append(_build_order(len(orders) + 1, current))

    if not orders:
        raise ExcelFormatError("没有读到任何数据行")
    return orders
