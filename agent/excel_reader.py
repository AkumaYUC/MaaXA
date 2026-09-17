#调拨单Excel读取器（严格模式）
#表头必须精确匹配标准列名，缺列/重复列直接报错 —— 填ERP单据的场景下猜错列比报错危险得多
#分隔规则：数据行之间空一行 = 分隔两张单据
#库存组织由填表人直接写在Excel里，本模块原样读取、不做任何转换（不推导、不映射、不补零、不改大小写）
#所以这份代码里不含任何公司专有数据，换一家公司只用改Excel，不用碰代码

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
    "调出库存组织": "source_org",
    "调入库存组织": "target_org",
    "备注": "remark",  #默认为*
}


def _build_order(index, items):
    #组装单据：库存组织以首行为准（金蝶里是单头字段），备注取第一个非空（没有默认*）
    first = items[0]
    source_org = first.source_org
    target_org = first.target_org

    if not source_org:
        raise ExcelFormatError(f"第 {first.row_index} 行：调出库存组织为空，每张单的首行必须填")
    if not target_org:
        raise ExcelFormatError(f"第 {first.row_index} 行：调入库存组织为空，每张单的首行必须填")

    for item in items[1:]:  #其余行留空是正常填法，忽略；填了就必须和首行一模一样，同一组织写成两种写法也会在这里被拦下
        if item.source_org and item.source_org != source_org:
            raise ExcelFormatError(f"第 {item.row_index} 行：调出库存组织「{item.source_org}」与首行「{source_org}」不一致，一张单只能有一个调出库存组织")
        if item.target_org and item.target_org != target_org:
            raise ExcelFormatError(f"第 {item.row_index} 行：调入库存组织「{item.target_org}」与首行「{target_org}」不一致，一张单只能有一个调入库存组织")

    remark = next((it.remark for it in items if it.remark), "")  #取整单第一个非空备注
    transfer_type = "组织内调拨" if source_org == target_org else "跨组织调拨"  #这两个值必须与金蝶「调拨类型」下拉项一字不差，文案有变时这里和管道都要同步改
    return TransferOrder(
        index=index,
        items=items,
        source_org=source_org,  #原样透传，粘进金蝶的就是用户填的那串字
        target_org=target_org,
        remark=remark or "*",  #整单都没备注默认填*
        transfer_type=transfer_type,
    )


@dataclass
class OrderItem:
    #一行物料明细（对应Excel一行）
    material_code: str  #物料编码
    quantity: str  #数量（如 33）
    source_warehouse: str  #调出仓库
    source_location: str  #调出仓位
    target_warehouse: str  #调入仓库
    target_location: str  #调入仓位
    source_org: str  #本行填的调出库存组织（整单以首行为准，其余行留空正常）
    target_org: str  #本行填的调入库存组织（整单以首行为准，其余行留空正常）
    remark: str            #备注
    row_index: int  #Excel原始行号，报错定位用


@dataclass
class TransferOrder:
    #一张调拨单（空行分隔出的一组连续数据行）
    source_org: str  #调出库存组织（用户填的原文，原样透传）
    target_org: str  #调入库存组织（用户填的原文，原样透传）
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
        values = {key: _norm_cell(row[col_idx[key]]) for key in COLUMNS.values()}  #按列名取值，列在Excel里的先后顺序无所谓

        if not any(values.values()):  #整行为空 = 单据分隔符
            if current:
                orders.append(_build_order(len(orders) + 1, current))  #收掉当前攒的单
                current = []
            continue

        item = OrderItem(
            material_code=values["material_code"],
            quantity=values["quantity"],
            source_warehouse=values["source_warehouse"],
            source_location=values["source_location"],
            target_warehouse=values["target_warehouse"],
            target_location=values["target_location"],
            source_org=values["source_org"],
            target_org=values["target_org"],
            remark=values["remark"],
            row_index=row_no,
        )

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
        orders.append(_build_order(len(orders) + 1, current))

    if not orders:
        raise ExcelFormatError("没有读到任何数据行")
    return orders
