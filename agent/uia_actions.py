#UIA 后台字段填写自定义动作（不抢前台、不动用户鼠标键盘）
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名）：
#  前置：GUI 连接设置里键盘方式必须是 PostMessage（Win32 控制单元），鼠标保持 Seize 不动 ——
#        键盘和鼠标是两套独立配置，互不影响
#  UiaSetCell         param: {"field": "material_code", "row": 0, "value": "01-036-00262", "expect": "绝缘纸"}
#                     按金蝶字段标识直接给明细表格某个格子写值（UIA ValuePattern，零鼠标零焦点）
#                     field填金蝶内部字段名，与客户端UIA的AutomationId、网页端data-field、WebAPI字段名三者同源，
#                     所以不受列顺序/横向滚动/分辨率影响（踩坑清单#18/#28/#37在这条路上不存在）
#                     field可选值见下方 _CELL_FIELDS 白名单（写成别的一律拒绝，白名单是必须的）
#                     row是明细行号，0=第一行，省略=0；超出界面实际行数直接返回False，不会顺着往下写
#                     value给了就写这个固定值；没给就从当前单据队列按 field 取（复用Excel契约，pipeline不用改Excel）
#                     **空值（没给value且队列里该字段也是空）视为「本行无需填这一列」，记日志并返回True**，
#                     不算失败 —— 否则Excel里备注这类允许留空的列会白白触发on_error
#                     expect填了就在写完后回读校验，不一致返回False走on_error；省略=不校验
#                     物料编码这类基础资料字段金蝶自己会解析并带出名称规格，后面的列不用手填
#  UiaPickList        param: {"field": "source_warehouse", "row": 0}
#                     仓库/仓位这类"点击弹查找列表再选一行"的字段，整段封装成一个动作：
#                     ① UIA SetFocus 到格子（不移动鼠标）→ ② 逐字符输入过滤词
#                     → ③ 等列表出来并数出目标行 → ④ UIA 选中该行 + 点「确定」（全程零输入注入）
#                     过滤词两种给法，**优先用 text_field**（每张单的仓库/仓位不同，写死只能跑一种单）：
#                       text_field: 从当前单据取，值填 _COLUMN_ATTR_BY_FIELD 里的键名
#                                   通常与 field 同名（field=source_warehouse 就填 text_field=source_warehouse）
#                       text:        固定关键词，只在测试/单一时段场景用
#                     两者的值都来自 Excel 模板的对应列（直接调拨单模板有 调出仓库/调出仓位/调入仓库/调入仓位 四列）
#                     row是明细行号，0=第一行，省略=0（同一列每行各有一个格子，多行单据必须逐行调）
#                     text是包含匹配：填关键词片段能命中完整名称；多行命中时取最靠前一行，
#                     最终校验用**归一化等值**（剥掉"名称(编码)"里的编码后缀后要求相等），
#                     所以"选到隔壁行"会被拒，不会静默写错仓库
#                     选行只认 SelectionItemPattern.Select()：拿不到或返回失败就直接失败，
#                     绝不拿 SetFocus 顶替 —— 焦点≠选中，选中项没变就点「确定」会选到别的行
#                     列表没出来/没有命中行/选完回读不符都返回False，让pipeline走on_error重试
#                     前置：本动作会先关掉残留的旧查找列表（否则句柄基线会把新弹窗也排除掉，
#                     造成跨重试死锁）；只关"带DataItem行且有成对确定+取消按钮"的金蝶弹窗，不碰主窗口
#                     注意：本动作只服务"数据格子的查找列表"，不碰维度数据录入那类独立弹窗
#  UiaReadCell        param: {"field": "source_warehouse", "row": 0, "expect": "<要核对的关键词>"}
#                     回读明细格子当前已提交的值，代替OCR做校验（小字文案OCR不准，#32已判死）
#                     判据三选一，按优先级：not_expect > nonzero > expect
#                     not_expect：值**不等于**它才通过（判"粘贴生效"用：默认是0，粘上去就不是0）
#                     nonzero：值不是0（能识别金蝶的 "5.0000000000" 补零写法；空值算不通过）
#                     expect：值包含它才通过（判"有没有仓库名"这类）
#                     都不填 = 只读回值打日志，不校验
#                     match_any为true时改成"任一关键字命中"，配keys数组用
#  UiaFillColumn      param: {"field": "quantity"}  按Excel顺序**逐行**把一列填满（UIA直写，不碰剪贴板）
#                     用途：物料列靠Ctrl+V粘贴自动扩行之后，数量列逐格写值 ——
#                     数量列不吃Ctrl+V（踩坑#10）、右键块粘贴又要动鼠标，UIA逐格写是唯一后台办法
#                     只认"界面上已经存在且有内容的行"：金蝶粘完多行物料后末尾会自动多一行空行，
#                     空行不参与（否则数量和Excel对不上）；行号取格子Name里的真实"Row N"
#                     Excel里该列为空的行会跳过（不写空值，免得清掉已有内容）
#                     前置：物料必须已经粘好（行已存在），否则返回False
#                     行数一律按**锚点列（物料编码）**判定，不看目标列自己有没有值 ——
#                     本动作的用途就是把空列填满，用目标列判行会永远得出"一行都没有"
#  UiaColumnFilled   param: {"field": "quantity", "nonzero": true}  整列每一行都填了时命中
#                     比只读首行（UiaReadCell row:0）强的对账：逐行检查锚点列圈出的每一行
#                     适用范围：数量列这种"每行都必须有值"的列；
#                     仓库/仓位那种整单同值只填首行的列**不要用**（它们本来就只在首行有值）
#                     nonzero:true 时额外要求非0（数量默认0，粘上物料后不该是0）
#                     anchor可选：用哪一列圈出行集合，默认material_code
#                     有空格子/0值就不命中，交给调用方重填或走收尾；命中返回占位框
#  UiaRowCount        param: {"field": "material_code"}  界面已填行数 == Excel队列行数时命中（对账）
#                     这是"填数量前先确认行数对不对"的识别策略：
#                     金蝶末尾的空行已排除，所以它比数UIA节点序更贴近真实可填行数
#                     不命中 → 让pipeline走兜底（重粘物料 / 报错收尾），避免"少粘几行还往下填"
#                     inverse:true 时反过来（行数不一致才命中），配自环重试用
#                     命中返回占位框(0,0,1,1)，只配分流用
#  UiaSetHeaderField  param: {"field": "target_org", "value": "<组织编码>", "expect": "<组织名片段>"}
#                     给**单头**字段写值（组织/调拨类型/备注/申请人…），UIA ValuePattern，零鼠标零按键注入。
#                     存在的理由就是后台：原先单头走 PasteOrderField（Ctrl+V，底层 keybd_event 全局注入），
#                     前台是用户程序时按键会打进用户的窗口里 —— 后台模式不能用它，只能走 UIA
#                     field 用 _HEADER_FIELDS 的键名；label 可覆盖默认标签（申请单叫「调出组织」不是「调出库存组织」）
#                     value 省略时从当前单据队列取（Excel 契约不变）；值为空就什么都不动、返回True
#                     ⚠️组织/申请人这类基础资料字段**别配 expect**：11-16 probe7 实测编码→名称的
#                     解析发生在**保存提交时**，写后立刻回读还是编码原样，expect 反而误判失败；
#                     外键校验放到保存之后用 UiaHeaderFilled 做
#                     定位靠标签几何配对（单头的 aid 全是模板部件名 PART_Editor，不唯一，没法按 aid 找）；
#                     同名标签出现多次（多张单据页同开）一律拒绝写入，和 _detail_grid 同一原则
#                     ⚠️写通道（2026-11-16 probe4/7 实测）：外层控件的 ValuePattern **只读** ——
#                     SetValue 抛 COMError 0x80040200（UIA_E_INVALIDOPERATION），但同一个 pattern 读值正常；
#                     真正能写的是它内层的 EditControl（KDPopupBaseEdit/ButtonEdit/KDDatePicker 都有，
#                     IsReadOnly=False）。本动作先走内层 Edit，找不到才退外层。
#                     编码→外键解析**已实测成立**（组织编码→保存后显示组织名称、人员编码→显示姓名），
#                     解析发生在**保存提交时**，写后立刻回读可能还是编码原样 —— 校验要放到保存之后
#  UiaPickCombo       param: {"field": "apply_type", "value": "<下拉里真实存在的选项>"}  组合框下拉选择（全程后台）
#                     适用：调拨类型 / 调拨申请类型 这类 KDComboBoxEdit 字段 ——
#                     外层 ValuePattern.SetValue 抛 INVALIDOPERATION（probe4）、又没有可写的内层 Edit，
#                     唯一通道是"展开下拉→点中目标项"。
#                     提交机制（2026-11-18 probe58/59/60 三轮对照实测）：
#                       ExpandCollapsePattern.Expand() 能后台展开（下拉项直接进主窗树，ListItemControl）；
#                       提交 = Select() 高亮 → PostMessage Enter 到主窗（probe60 S3 路线，全程零鼠标）。
#                       踩过的两个假提交：只 Select 不回车（只高亮）；Select+DoDefaultAction 打旧对象
#                       （下拉重渲染后引用失效）—— 都表现为"调用成功、回读不变"，别走回头路。
#                     value 省略时从当前单据队列取（Excel 契约不变）；值为空=不动、保留金蝶默认、返回True；
#                     现值已等于目标值直接过（不展开不点击，probe59 实测金蝶默认值恰好正确时别白点）。
#                     ⚠️下拉项**每个公司/账套不一样**（同一字段在这套账套里是哪些选项，完全由账套配置决定，
#                     代码里不写死任何具体业务词，一律从 Excel 传进来），
#                     所以匹配必须是**等值**不是包含，找不到就失败并在日志里列出全部可选项——
#                     宁可停下让用户改 Excel，也绝不"差不多"选一个静默填错。
#                     提交后**重新配对再回读**（控件可能重建，旧引用读的是缓存），等值才算成功。
#  UiaPickFlex        param: {"field": "source_location", "row": 0, "keyword": "某仓位名称", "expect": "名称片段"}
#                     弹性域（仓位/库位这类"多维值集组合"字段）的后台填写：走金蝶自己的
#                     「维度数据录入」弹窗，而不是往格子里塞字符串。
#                     ⚠️这类字段**不接受**任何形态的直写：名称、编码、编码+名称、仓库/120 全部实测被吞
#                       （格子内层编辑器收得到文本，金蝶从不提交解析）—— 所以只能驱动弹窗。
#                     ⚠️前置：**同行的仓库列必须先填好**。仓库为空时仓位格是禁用态，
#                       里面连编辑控件都没有（只有一个 TextBlock），点了也不会有弹窗。
#                     keyword 省略时从当前单据队列按 field 取（Excel 契约不变，填名称即可，不用编码）；
#                       keyword 为空 = 本行无需填这一列，记日志返回True（和 UiaSetCell 空值语义一致）
#                     expect 填了就在提交后要求格子回读包含它，省略=只要求非空
#                     流程（每一步都有判据，不靠"调用没抛异常"下结论）：
#                       ①格子必须在视口内 —— 不在就用明细网格的 ScrollPattern 横向滚过去
#                       ②PostMessage 点格子中心进编辑态 → 再点放大镜候选位 → 等弹窗出现
#                         放大镜是金蝶自绘的，**不在 UIA 树里**（没有 Invoke/Expand/Legacy 元素），
#                         所以只能"点"；但点是 PostMessage 到主窗，不碰用户鼠标键盘
#                       ③支路一（优先）：清空输入框 → 焦点进内层 TextBox → 逐字 WM_CHAR 输入 keyword
#                         → 金蝶弹「全模糊匹配」下拉 → 双击唯一命中行 → 输入框回显编码
#                         ⚠️SetValue 不触发下拉，只有真实字符消息触发（实测）
#                       ④支路二（下拉没出现时兜底）：SetValue(keyword) → Invoke 内层 PART_Item →
#                         等「值集资料选择列表」→ 包含匹配找唯一行 → SelectionItem.Select →
#                         提交（先 Enter 后双击兜底，判据=输入框回显编码）；没命中点 bNext 翻页，
#                         翻完仍没有 = 失败（绝不瞎选第一行）
#                         ⚠️这条列表查询实测会整批返回空（会话/账套相关），所以才加支路一
#                       ⑤Invoke 弹窗「确定」→ 回读格子
#                     坐标鲁棒性：放大镜落点**全部相对格子自身 UIA 矩形按比例算**，
#                       列顺序、列宽、DPI、分辨率变了 rect 自己会变，候选点跟着变；
#                       再配"候选位逐个试 + 弹窗出现才算命中"，点错候选只是多试一次，不会写错数据
#                     ⚠️全程不隐藏弹窗/列表窗口：实测隐藏会留下关不掉的僵尸顶层窗，
#                       后续再触发就永远等不到新弹窗（跨重试死锁）
#                     ⚠️弹窗会短暂抢前台（单列约 20-30s），动作结束用 _fg_restore 还给用户窗口
#  UiaSetForm         param: {"form": "requisition"}  切换明细字段白名单（transfer|requisition）
#                     字段标识跨单据类型不通用（踩坑清单#38）：直接调拨单 FSRC*/FDEST* 一族、
#                     申请单 FSTOCK* 一族。每个任务 pipeline 的**第一个节点**必须先调它，
#                     之后所有 Uia* 明细动作按这套 aid 找格子。
#                     不调本动作 = transfer（直接调拨单老行为，一行都不用改）
#                     ⚠️同一 GUI 进程里两种单混跑会互相踩白名单 —— 一个任务只服务一种单据
#                     form=requisition 时组织两列（source_org/target_org）进 _CELL_FIELDS：
#                     它们是**明细列每行必填**，Excel 整单同值只写首行 → 配 UiaFillColumn broadcast:true
#  UiaNavigate        param: {"text": "调拨申请单", "wait_label": "申请组织", "timeout_ms": 15000}
#                     任意页 → 顶部搜索直达该单据的编辑页（2026-11-17 probe17 实测成立）：
#                     ① 顶部搜索框（y<80 的可见 EditControl）SetFocus + ValuePattern 写关键词
#                     ② PostMessage WM_KEYDOWN/UP VK_RETURN 到主窗口 → 金蝶直接开出「调拨申请单 - 新增」编辑页
#                     ③ 轮询等 wait_label（该单编辑页上的单头标签）可见才算到位
#                     首页磁贴**不要走**：磁贴是纯 TextBlock，Invoke 没有、Legacy DoDefaultAction 抛
#                     托管异常（0x80131627），2026-11-17 probe15/16 判死
#                     wait_label 省略 = 只发导航不等结果（返回 True 不代表页已就位，慎用）
#  UiaInvokeButton    param: {"name": "保存", "post_wait_ms": 3000}
#                     按名字 Invoke 工具栏按钮（保存/提交/列表/新增…），零鼠标零坐标（§2.9-2 的正式封装）。
#                     先按 MenuItemControl 找名 —— 金蝶工具栏 保存/提交/审核 是 MenuItem，
#                     真正带 InvokePattern 的是它内层 Button（aid=PART_DefaultButt，Name 为空），
#                     直接按 ButtonControl 找名会漏（11-16 probe2b 就漏过"保存"）；
#                     MenuItem 没有再按 ButtonControl 找（列表/退出/选单 这类是普通 Button）
#                     ⚠️只点非 offscreen 的：编辑页上 是/否/确定/取消 大多是隐藏预置按钮，
#                     按名+可见双闸过滤，绝不点隐藏对话框的按钮
#                     post_wait_ms 给界面跳转留时间（保存→列表约 2-3s），省略=3000
#自定义识别（节点recognition填Custom、custom_recognition填名字）：
#  UiaCellHasValue    param: {"field": "source_locations", "row": 0}  当前行该字段有值时命中（跳过填写用）
#                     与 OrderColumnEmpty 同一用途，但读的是界面真实值而不是Excel队列，
#                     所以能兜住"填了但没填进去"的情况；命中返回占位框(0,0,1,1)，只配分流
#  UiaHeaderFilled    param: {"field": "target_org", "expect": "<组织名片段>"}  单头字段已有值时命中（分流/收尾校验）
#                     定位方式与 UiaSetHeaderField 完全一致；expect 是"包含"判据，省略=只要求非空
#                     命中返回占位框(0,0,1,1)，只配分流用
#设计约束（为什么这样切）：
#  - Context 只能读 node_data、不能回写，所以"数出的行数"没法从agent传回pipeline；
#    与其让pipeline去猜，不如把"数行+选行+确认"整段收进一个动作里（KISS）
#  - 所有字段名走白名单，pipeline里写错名字立刻报错，不会静默写到别处；
#    取单据字段用显式 lambda 表而不是 getattr —— 后者能被 "__class__" 这类串利用
#  - 全程不产生系统级输入注入（不发 keybd_event / SendInput / 不动光标）：
#    只走 UIA 调用 + 向金蝶自己窗口 PostMessage，用户前台程序收不到任何输入
#  - 行号：金蝶每个格子的 Name 自带 "Row N, Column M:" 前缀，用它核对真实行号；
#    光看UIA节点顺序会被虚拟化/滚动骗到（未实现的行不在树里）

import ctypes
import functools
import sys
import threading
import time

from ctypes import wintypes  #RECT 等结构体（import ctypes 不会带出 wintypes）

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction
from maa.custom_recognition import CustomRecognition

from common import _params  #pipeline 参数解析兜底（写坏打日志返回False，不在agent进程里抛异常）

#UiaSetCell / UiaReadCell 能碰的明细字段白名单（键名=参数名，值=金蝶内部字段标识）
#键名沿用 PasteOrderColumn 的列名风格，pipeline 侧换动作时不用换脑子
#
#⚠️值必须是**单据编辑页**的 AutomationId，不能拿列表页的那套 —— 两者不一样，
#  而且用错了不会报错，只会"找不到格子"。旧值 FSTOCKID/FSTOCKINID/FSTOCKLOCID/
#  FSTOCKLOCINID/FNOTE 是从「直接调拨单列表」页的表头抄的（那页显示 FSrcStockId.FName 这类
#  带 .FName 后缀的字段），单据编辑页里根本不存在：2026-11-16 实测 7 个旧 id 全部★不存在★，
#  于是所有仓库/仓位动作都在 _cell_node 第一步就 return False。
#  这就是 11-10 记的"仓库仓位四列 SetValue 无效"的真相 —— 不是金蝶拒绝写，
#  是字段名压根没命中过，那次实验连格子都没找到。
#下面的值逐个在单据编辑页实测存在（每个都"共3个可见1个"，3=残留的旧单据页签）：
#⚠️2026-11-16 晚订正：上面那段注释的归因已被只读探针推翻 ——
#  FSTOCKID 一族不是列表页名，是**调拨申请单编辑页的正确标识**（FSTOCKORGID C9/FSTOCKID C10/
#  FSTOCKLOCID C11/FSTOCKORGINID C12/FSTOCKINID C13/FSTOCKLOCINID C15/FNOTE C19 全部实测存在）。
#  真正的规则是**字段标识跨单据类型不通用**（踩坑清单#38）：直接调拨单 FSRC*/FDEST* 一族、
#  申请单 FSTOCK* 一族、列表页第三套。所以白名单分两套，由 UiaSetForm 在任务开头切换：
#  transfer    直接调拨单（默认，不调 UiaSetForm 时行为与从前完全一致）
#  requisition 调拨申请单（组织两列=明细列每行必填，按编码写→保存解析外键，probe4/7 实测成立；
#              仓库列按编码直写；仓位两列=弹性域，直写一律被吞，必须配 UiaPickFlex 走维度弹窗）
_CELL_FIELDS_TRANSFER = {
    "material_code": "FMATERIALID",  #物料编码（Column 2）
    "quantity": "FQTY",  #数量（Column 7）
    "note": "FNOTEENTRY",  #明细备注（Column 20，内层 ButtonEdit 只吃纯文本）
    "source_warehouse": "FSRCSTOCKID",  #调出仓库（Column 10，查找类，配 UiaPickList）
    "source_location": "FSRCSTOCKLOCID",  #调出仓位（Column 11，查找类）
    "target_warehouse": "FDESTSTOCKID",  #调入仓库（Column 12，查找类）
    "target_location": "FDESTSTOCKLOCID",  #调入仓位（Column 13，查找类）
}

_CELL_FIELDS_REQUISITION = {
    "material_code": "FMATERIALID",  #物料编码（申请单在 Column 1；aid 寻址不受列号影响）
    "quantity": "FQTY",  #数量（Column 7）
    "note": "FNOTE",  #明细备注（申请单编辑页真名；FNOTEENTRY 在这边不存在）
    "source_org": "FSTOCKORGID",  #调出组织（明细列每行必填，编码→外键实测成立）
    "target_org": "FSTOCKORGINID",  #调入组织（同上）
    "source_warehouse": "FSTOCKID",  #调出仓库（编码直写，配 UiaSetCell）
    "source_location": "FSTOCKLOCID",  #调出仓位（弹性域，配 UiaPickFlex；直写会被吞）
    "target_warehouse": "FSTOCKINID",  #调入仓库（编码直写，配 UiaSetCell）
    "target_location": "FSTOCKLOCINID",  #调入仓位（弹性域，配 UiaPickFlex；直写会被吞）
}

_FORM_FIELDS = {
    "transfer": _CELL_FIELDS_TRANSFER,
    "requisition": _CELL_FIELDS_REQUISITION,
}

#当前生效的明细白名单（UiaSetForm 切换内容；默认 transfer = 直接调拨单老行为）
_CELL_FIELDS = dict(_CELL_FIELDS_TRANSFER)

#调出/调入库存组织**不在这里**：单据编辑页里没有任何含 ORG 的 AutomationId（实测），
#明细格子靠 aid 定位这条路对单头不成立 —— 单头字段的 aid 是模板部件名 'PART_Editor'，
#一张单里几十个控件共用，根本不唯一。所以单头只能"按界面标签的几何位置找它右边那个编辑框"，
#见下面的 _HEADER_FIELDS + _header_editor。
#（2026-11-16 更正：这里早先写过"组织两列继续走 PasteOrderField，别配成 UIA"，
# 那句是错的 —— 当时探针只按 ControlTypeName=='EditControl' 筛，把金蝶自绘的
# KDPopupBaseEdit/KDComboBoxEdit/ButtonEdit 全滤掉了，于是得出"单头 0 个可写编辑框"。
# 放宽到 ClassName 含 Edit/ComboBox/Spin/Date 后，实测每个单头字段都找得到、
# 而且都带 ValuePattern。）

#只能用 UiaPickList 走查找列表的字段（白名单，挡"用 UiaSetCell 硬写查找字段"这种误配）
#⚠️这套闸只对**直接调拨单**成立（那边名称/编码直写都没验通，唯一验证过的填法是前台弹窗链路）。
# 调拨申请单实测：**按编码直写解析成立**（仓库编码 → 金蝶解析成仓库名称），
# 名称直写被吞 —— 所以 requisition 族这边闸是空的，UiaSetCell 直写编码就是正路。
#  代价是 Excel 仓库列**必须填编码**（填名称会被金蝶静默吞成空，pipeline 用 UiaReadCell
#  nonzero 把关，吞了会当场失败而不是带病保存）。
#  模糊匹配弹窗（人工输入关键词弹列表选行）后台能开弹窗但**不交付**：过滤词语义和手输不一致
#  （probe30）、Enter 提交回填从未可靠验证（probe28/29）、多页签污染下有误写风险 —— 用户口径
#  "避免出错就不弄"。UiaPickList 保留给直接调拨单侧的后续实验。
#⚠️2026-11-19 复测订正：上面"仓库/仓位四列都只认编码"这句**只对仓库成立**。
#  仓位是弹性域（多维值集组合），没有任何单一字符串能解析成它 —— 名称和编码都被吞不是"写法不对"，
#  是这类字段压根不接受直写。正路是驱动金蝶自己的「维度数据录入」弹窗（UiaPickFlex）。
#  所以 requisition 族把两个仓位列放回闸里：UiaSetCell 会拒写它们，逼着配 UiaPickFlex。
#  transfer 族四个列本来就在闸里（那边弹窗触发仍未验证，UiaPickFlex 只在 requisition 交付）。
_LOOKUP_FIELDS_TRANSFER = frozenset((
    "source_warehouse", "source_location", "target_warehouse", "target_location",
))
_LOOKUP_FIELDS_REQUISITION = frozenset((
    "source_location", "target_location",  #仓位=弹性域，只能走 UiaPickFlex
))

#当前生效的查找类闸（UiaSetForm 原地切内容；默认 transfer = 直接调拨单老行为）
_LOOKUP_FIELDS = set(_LOOKUP_FIELDS_TRANSFER)

#单头字段（整单只有一个值的那些：组织/调拨类型/备注/申请人…）→ 界面上的标签文本。
#键名与 excel_reader 的字段名一一对应，pipeline 里 UiaSetCell 和 UiaSetHeaderField 用同一套名字。
#标签文本是**直接调拨单**编辑页实测到的可见 TextControl.Name（2026-11-16 12:5x 实测）。
#⚠️调拨申请单的标签不一样：同一个字段那边叫「调出组织」而不是「调出库存组织」
# （excel_reader.py:71 的注释写明了），所以下面标了"申请单待实测"的几项，
#  等做 TransferRequisitionForm 时按申请单编辑页的真实标签再对一遍；
#  对不上时的表现是"找不到标签"直接报错，不会写错地方，可以安全试。
#查找类字段（组织）写进去之后金蝶要把编码解析成外键，能不能解析成功**尚未实测**：
#所以配 UiaSetHeaderField 时**必须写 expect**（期望回读到的名称片段，如 组织编码 → 组织名称），
#否则金蝶只把编码当纯文本显示、底下没绑上组织时，动作会"写入成功"但其实单据是错的。
_HEADER_FIELDS = {
    "source_org": "调出库存组织",  # 申请单待实测：调出组织
    "target_org": "调入库存组织",  # 申请单待实测：调入组织
    "transfer_type": "调拨类型",
    "remark": "备注",
    "apply_org": "申请组织",  # 申请单专用，待实测
    "applicant": "申请人",  # 申请单专用，待实测
    "phone": "联系电话",  # 申请单专用，待实测
    "address": "通讯地址",  # 申请单专用，待实测
    "contact": "联系人",  # 申请单专用，待实测
    "demand_date": "需求日期",  # 申请单专用，待实测
    "apply_type": "调拨申请类型",  # 申请单专用，KDComboBoxEdit（probe58 实测标签几何配对成立）
}

#单头编辑框的类名特征。金蝶单头是自绘控件，UIA ControlTypeName 只有 EditControl/ComboBoxControl
#两种，但 ClassName 是 KDPopupBaseEdit / KDComboBoxEdit / ButtonEdit / KDDatePicker /
#KDPopupCalcEdit / KDNumEditorDisplayer 这一堆 —— 判据必须是"类名里带 Edit/ComboBox/Spin/Date"，
#只按 ControlTypeName=='EditControl' 筛会一个都不剩（2026-11-16 我就这么误判成"单头 0 个可写框"）。
#注意还要**排掉 aid=='PART_Editor' 的内层 TextBox**：它是上面那个复合控件的内部部件，
#矩形和父控件几乎重合，留着会让同一个字段配出两个候选，分不清该写哪个。
_HEADER_CLASS_HINTS = ("Edit", "ComboBox", "Spin", "Date")

#单头里的**基础资料字段**（值会被金蝶解析成外键、界面显示名称而不是我们写的编码）。
#这批字段在 UiaSetHeaderField 里不能按"回读==写入值"判成败：
#  ①写后立刻回读可能是编码原样（解析发生在保存时，probe7 实测）；
#  ②也可能直接是名称，而新草稿默认值恰好等于目标值 → before==actual 不等于没写进去
#    （实测：申请组织默认=当前登录组织，Excel 也填同一个组织，被误判失败）。
#真正的把关在保存后：pipeline 用 UiaHeaderFilled 验编码有没有解析成名称（VerifyFK 节点）。
_BASEDATA_HEADER_FIELDS = frozenset((
    "apply_org", "source_org", "target_org", "applicant",
))

#field 参数 → 从 OrderItem 取值的函数。
#为什么不用 getattr(item, field)：field 是 pipeline 里手写的字符串，直接 getattr 能把
# "__class__"、"items" 这类也取出来 —— transfer_actions 的 _ORDER_FIELDS 白名单是同一个理由
_COLUMN_ATTR_BY_FIELD = {
    "material_code": lambda it: it.material_code,
    "quantity": lambda it: it.quantity,
    "note": lambda it: it.remark,
    "source_warehouse": lambda it: it.source_warehouse,
    "source_location": lambda it: it.source_location,
    "target_warehouse": lambda it: it.target_warehouse,
    "target_location": lambda it: it.target_location,
    "source_org": lambda it: it.source_org,
    "target_org": lambda it: it.target_org,
    #下面这些只有单头用得上（_header_value_from_order），明细列里没有对应字段，
    #所以不用进 _CELL_FIELDS —— UiaSetCell 拿到它们照样会报"未知field"
    "remark": lambda it: it.remark,
    "transfer_type": lambda it: it.transfer_type,
    "apply_org": lambda it: it.apply_org,
    "applicant": lambda it: it.applicant,
    "phone": lambda it: it.phone,
    "address": lambda it: it.address,
    "contact": lambda it: it.contact,
    "demand_date": lambda it: it.demand_date,
    "apply_type": lambda it: it.apply_type,
}

#金蝶主窗口的类名特征（标题是「金蝶云星空」，但类名里的GUID每次重启都变，所以只认前缀）
_MAIN_WINDOW_CLASS = "HwndWrapper[Kingdee.BOS.XPF.App"

#金蝶主窗口对象的本次动作内缓存（见 _main_window / _reset_cache）
_MAIN_CACHE = None

#找元素时最多往下钻几层。金蝶单据页的层级比这个深得多，但格子的 aid 在浅层就能命中；
#调大只是兜底，真正的防错是行号核对（见 _cell_node）
_MAX_DEPTH = 24

_DEFAULT_TIMEOUT_MS = 3000  #等列表弹窗的默认超时
_POLL_MS = 200  #轮询间隔

#uiautomation 是可选依赖：缺失时只在真正调用UIA动作时失败，不影响agent其它动作启动
try:
    import uiautomation as _uia
    _uia.SetGlobalSearchTimeout(_DEFAULT_TIMEOUT_MS / 1000.0)
    _UIA_AVAILABLE = True
except Exception:  #依赖没装/初始化失败都退化，动作内部再报错
    _uia = None
    _UIA_AVAILABLE = False


#动作开始前的前台窗口句柄（模块级）。装饰器写入，_fg_restore 读取。
#为什么不用参数往下传：偷焦点发生在 SetValue/Invoke 那一瞬间，而调用点散在
#UiaSetCell/UiaFillColumn/UiaSetHeaderField/UiaInvokeButton 各处，
#逐个加参数会把签名搅乱；MaaFramework 的任务队列是单线程串行执行动作，
#动作之间不并发，所以这一个全局槽位够用（下一个动作开头会覆盖它）。
_FG_ORIG = 0


def _fg_capture():
    #动作开始前记下当前前台窗口（用户此刻在看谁）。0=读不到，后面就不还原
    try:
        return int(ctypes.windll.user32.GetForegroundWindow())
    except Exception:
        return 0


def _fg_restore(fg0=None):
    #把被金蝶偷走的前台还给原窗口。§2.6 实测：ValuePattern.SetValue / Invoke 会让
    #金蝶 WPF 控件请求焦点、整窗被激活到前台（毫秒级、可逆）——2026-11-18 用户实测
    #「每一步都拉回前台」的真身（截图侧 FramePool 已零命中 ensure foreground）。
    #
    #调用时机有两处，配合起来才是"极致"：
    #  ①写入点即时归还（主路）：UiaSetCell/UiaFillColumn/UiaSetHeaderField/UiaInvokeButton
    #    在 SetValue/Invoke 之后**马上**调一次 —— 金蝶刚偷到手就还回去，动作内部那些
    #    sleep 等解析、轮询等页面（Navigate 最长 20s）全程用户窗口都拿着自己的前台。
    #  ②动作末尾兜底（装饰器）：万一某条路径没走到①（异常提前 return、或以后新增的
    #    写入点忘了加），这里再还一次。
    #只在"前台现在确实是金蝶主窗、且原来不是金蝶"时归还；用户本来就在看金蝶则不动。
    #SetForegroundWindow 可能被前台锁拒（返回 False）：只打一次日志、绝不重试 ——
    #被拒的调用会闪任务栏，重试连闪比"闪一下"扰民得多。
    if fg0 is None:
        fg0 = _FG_ORIG
    if not fg0:
        return
    kingdee = _main_hwnd()
    if not kingdee or fg0 == kingdee:
        return
    try:
        user32 = ctypes.windll.user32
        if int(user32.GetForegroundWindow()) == kingdee:
            if not user32.SetForegroundWindow(fg0):
                _log("fg", "前台归还被拒（Windows 前台锁），本次跳过")
    except Exception:
        pass


def _ensure_window_visible():
    #2026-11-18 用户实测：没人点最小化，只是**双击了任务栏上的金蝶图标**想切去别的应用
    #（第一击没反应、第二击缩后台——Windows 经典手势，真实用户天天这么干）。
    #金蝶一最小化，Win32 给主窗 237×39 的虚拟矩形：单头几何配对、Navigate 到位判据、
    #工具栏按钮可见性全灭（明细列按容器定界不看可见性，所以形态是"一路绿到单头突然红"）。
    #协作契约不能写成"用户别最小化"——agent 必须在每个动作开头自己把窗口捞回来。
    #SW_SHOWNOACTIVATE(4)：恢复显示但**不激活** —— 用户的前台窗口纹丝不动，
    #金蝶只是回到它底下的层（坐标恢复真实值），几何配对重新可用；
    #最小化前是最大化态的话，Windows 按最近尺寸还原（docs: most recent size and position）。
    #⚠️句柄必须走 _find_main_hwnd()（纯 Win32 EnumWindows），**不能用 _main_hwnd()**：
    #  后者要 ControlFromHandle 拿 COM 对象，而本函数在装饰器里跑于 COM 初始化之前 ——
    #  2026-11-18 GUI 日志实证：每个任务第一个动作在这里抛"尚未调用 CoInitialize"、
    #  静默返回 0，最小化根本没被捞回来（第二个动作起才有 COM）。三个 API 全是 Win32，
    #  本来就不需要 COM —— 别把它接回任何 UIA 调用链上。
    hwnd = _find_main_hwnd()
    if not hwnd:
        return
    try:
        user32 = ctypes.windll.user32
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, 4)
            _log("fg", "金蝶窗口被最小化了（多半是双击了任务栏图标），已恢复显示、不抢你当前的前台")
    except Exception:
        pass  #恢复失败不改变动作结果：下面的原逻辑会按"找不到标签"报错并自诊断


def _uia_inited(fn):
    #给每个 UIA 动作/识别套上"本线程的 COM 初始化"。
    #为什么必须在里面：agent 的 CustomAction/CustomRecognition 不在 main.py 那个线程跑，
    #而是跑在 MaaFramework 派下来的工作线程上，那个线程从没调过 CoInitialize。
    #症状（2026-11-16 09:35 实测）：stdout 打
    #  [WinError -2147221008] 尚未调用 CoInitialize。
    #  Can not load UIAutomationCore.dll.
    #  [UiaSetCell] 没找到金蝶主窗口（客户端是不是没开？）
    #—— 金蝶明明开着、窗口就在前台，但整棵 UIA 树都枚举不了，于是每个动作都 False。
    #别在主线程 import 时初始化一次就完事：COM 是按线程 apartment 算的，主线程初始化过
    #对工作线程毫无用处。实测（同一解释器、同一线程模型）：
    #  裸调用 → FAIL OSError(22, '尚未调用 CoInitialize。')
    #  with UIAutomationInitializerInThread() → OK 顶层窗口=18
    #退出时它自己 CoUninitialize（引用计数配平），所以嵌套调用也安全；
    #代价是"控件对象不能活过本次调用" —— 7 个入口开头都调了 _reset_cache()，
    #_MAIN_CACHE 不会把控件句柄带到下一个动作里，这条约束才成立。加新动作时必须照做。
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        if not _UIA_AVAILABLE:
            return fn(*args, **kwargs)  #依赖缺失时保持原样，由 _require_uia 报缺依赖
        global _FG_ORIG
        _FG_ORIG = _fg_capture()  #动作前记住用户正在看哪个窗口（写入点即时归还读它）
        _ensure_window_visible()  #金蝶被最小化就先捞回来（SW_SHOWNOACTIVATE，不动用户前台）
        with _uia.UIAutomationInitializerInThread():
            ret = fn(*args, **kwargs)
        _fg_restore()  #兜底：动作内部有 sleep 等解析，写入点已经即时还过一次的话这里无事可做
        _FG_ORIG = 0
        return ret

    return wrapper


def _log(tag, message):
    #安全的日志输出。agent 的 stdout 在中文 Windows 上常被当成 GBK 管道，
    #print 里出现编码不出来的字符（⚠、emoji 之类）会抛 UnicodeEncodeError，
    #而这个异常在动作里没人接 → 整个 CustomAction 崩掉，表现为"金蝶没动、任务失败、日志还没写全"。
    #先把整行按当前 stdout 编码试编一次，编不出来就转义再打 —— 不依赖 print 内部抛异常
    text = f"[{tag}] {message}"
    encoding = getattr(sys.stdout, "encoding", None) or "ascii"
    try:
        text.encode(encoding, "strict")
    except (UnicodeEncodeError, LookupError):
        text = text.encode("ascii", "backslashreplace").decode("ascii")
    try:
        print(text)
    except Exception:  #日志失败绝不能影响动作本身
        pass


def _require_uia(tag):
    #UIA 依赖缺失时统一在这里报错，别在每个动作里写一遍同样的 if
    if not _UIA_AVAILABLE:
        print(f"[{tag}] 缺少 uiautomation 依赖，请先在 agent 解释器里安装")
        return False
    return True


def _win32_class_name(hwnd):
    #取 Win32 真实类名。注意：UIA 的 ClassName 属性不是这个值 ——
    #金蝶主窗口 UIA 报 'Window'，Win32 报 'HwndWrapper[Kingdee.BOS.XPF.App.exe;;<GUID>]'。
    #两者对不上是这次整条 UIA 路线一次都没跑通的根因（见 _find_main_hwnd）。
    try:
        buf = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetClassNameW(int(hwnd), buf, 256)
        return buf.value or ""
    except Exception:
        return ""


def _find_main_hwnd():
    #用 Win32 按「类名前缀」找金蝶主窗口句柄，再交给 ControlFromHandle 换成 UIA 元素。
    #不能像旧写法那样遍历 UIA 顶层窗口比 ClassName：金蝶主窗口在 UIA 里 ClassName=='Window'，
    #永远不等于 _MAIN_WINDOW_CLASS，于是每个动作都打「没找到金蝶主窗口（客户端是不是没开？）」
    #—— 而客户端其实开着、还在前台（2026-11-16 09:55 实测：Win32 能枚举到 hwnd=2168456）。
    #类名里的 GUID 每次重启会变，所以只比前缀（这条沿用原设计的判断依据）。
    #同前缀的窗口可能有多个（主窗 + 它自己的弹窗），取面积最大的那个当主窗口。
    user32 = ctypes.windll.user32
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)  #回调必须返回 4 字节 BOOL；c_bool 是 1 字节，x64 下属于未定义行为（同 my_action._find_window）
    def enum_proc(hwnd, lparam):
        if not hwnd or not user32.IsWindowVisible(hwnd):
            return True
        if not _win32_class_name(hwnd).startswith(_MAIN_WINDOW_CLASS):
            return True
        r = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(r))
        area = max(0, r.right - r.left) * max(0, r.bottom - r.top)
        found.append((int(hwnd), area))
        return True

    try:
        user32.EnumWindows(enum_proc, 0)
    except Exception:
        return 0
    if not found:
        return 0
    return max(found, key=lambda item: item[1])[0]


def _main_window():
    #找金蝶主窗口：Win32 认句柄 → ControlFromHandle 换 UIA 元素（见 _find_main_hwnd 的说明）。
    #枚举和 COM 调用都可能抛，异常绝不能穿出 run()，所以整体兜住。
    #结果按"当前动作"缓存：一次动作里 _cell_node/_type_text/_main_pid 会各要一次，
    #每次都重新枚举会明显拖慢（审核提的性能问题）
    global _MAIN_CACHE
    if _MAIN_CACHE is not None:
        return _MAIN_CACHE
    if not _UIA_AVAILABLE:
        return None
    hwnd = _find_main_hwnd()
    if not hwnd:
        return None  #金蝶真没开（或进程完整性级别太高，EnumWindows 看不见）
    try:
        win = _uia.ControlFromHandle(hwnd)
    except Exception:
        return None
    if win is None:
        return None
    _MAIN_CACHE = win
    return win


def _reset_cache():
    #动作入口调一次：一个动作内复用同一次枚举结果，动作之间必须重查
    #（否则金蝶重启/切页后还拿着旧句柄，表现为所有动作静默失败）
    global _MAIN_CACHE
    _MAIN_CACHE = None


#页面级缓存（2026-11-20 性能优化）。_MAIN_CACHE 每个动作入口都被 _reset_cache 清掉，
#这是防"金蝶重启后拿旧句柄"的必要约束，但它逼每个字段操作重新全树遍历：
#_header_editor 一次遍历几千个跨进程 COM 调用、实测 5~15s，一张单 18 个字段操作
#就是 GUI 实测 5m1s 里的大头。这里缓存的是"同一张编辑页内不会变的东西"：
#单头扫描结果（标签+候选编辑框）、明细容器列表。
#生命周期 = 一次页面：四个"页面会变"的地方显式作废（UiaSetForm/UiaNavigate/UiaInvokeButton/
#UiaPickCombo提交后，见 _invalidate_page_cache 调用点），另加廉价存活探测兜底
#（用户手动切页签/改窗口尺寸时，缓存节点读属性会抛异常或主窗矩形变了 → 作废重扫）。
#缓存只省遍历，不改任何识别/配对/拒绝逻辑；配对和值读取仍然实时。
#动作队列单线程串行执行（同 _FG_ORIG 的成立理由），一个 dict 够用。
_PAGE_CACHE = {"hwnd": 0, "header": None, "panels": None}


def _invalidate_page_cache():
    _PAGE_CACHE["hwnd"] = 0
    _PAGE_CACHE["header"] = None
    _PAGE_CACHE["panels"] = None


def _page_cache_usable():
    #缓存必须属于当前主窗：句柄对得上且非 0（金蝶没开时 _main_hwnd()==0，绝不许命中缓存）
    return bool(_PAGE_CACHE["hwnd"]) and _PAGE_CACHE["hwnd"] == _main_hwnd()


def _main_pid():
    #金蝶进程号。弹窗归属校验用：别的 WPF 程序也有列表，只按"有 DataItem"认会认错爹
    main = _main_window()
    if main is None:
        return 0
    try:
        return main.ProcessId
    except Exception:
        return 0


def _walk(root, want_type=None, want_aid=None, max_depth=_MAX_DEPTH, stats=None):
    #栈式先序遍历：倒序压栈保证出栈顺序=界面顺序（第一行在第一）。
    #want_type 按控件类型收（如 DataItemControl），want_aid 按 AutomationId 收；
    #两者都给则同时满足才算。单节点读属性可能抛（已销毁/虚拟化），逐个跳过不中断遍历
    #stats 传一个 dict 进来可以拿到 truncated 标志：区分"这行确实不存在"和"树没走到那么深"
    out = []
    stack = [(root, 0)]
    while stack:
        node, depth = stack.pop()
        #want_type/want_aid 都不给 = 收集所有后代（_row_texts 的回退分支靠这个拿"任意带文本的节点"）。
        #root 自身不收集：调用方要的是它的子孙
        if node is not root:
            try:
                hit_type = want_type is None or node.ControlTypeName == want_type
                hit_aid = want_aid is None or node.AutomationId == want_aid
                if hit_type and hit_aid:
                    out.append(node)
            except Exception:
                pass
        if depth >= max_depth:
            if stats is not None:
                stats["truncated"] = True  #还有更深的内容没遍历到
            continue
        try:
            children = node.GetChildren()
        except Exception:
            continue
        for child in reversed(children):
            stack.append((child, depth + 1))
    return out


def _cell_prefix(name):
    #拆出金蝶格子 Name 里的行号/列号。格式是 "Row 0, Column 10: <该格当前值>"；
    #解析不出来返回 (None, None)，调用方按"无法核对"处理
    text = name or ""
    if not text.startswith("Row "):
        return None, None
    try:
        head, _sep, _value = text.partition(":")
        parts = head.replace("Row ", "").split(", Column ")
        return int(parts[0]), int(parts[1])
    except Exception:
        return None, None


def _read_cell_text(node):
    #读格子当前已提交的值。金蝶的 ValuePattern getter 恒返回空（实测），
    #所以优先读 Name（形如 "Row 0, Column 10: <该格当前值>"），再退化到 TextPattern
    name = ""
    try:
        name = node.Name or ""
    except Exception:
        pass
    if ":" in name:  #剥掉 "Row 0, Column 10:" 前缀，只留值
        name = name.split(":", 1)[1]
    name = name.strip()
    if name:
        return name
    try:
        text_pattern = node.GetPattern(_uia.PatternId.TextPattern)
        if text_pattern is not None:
            return (text_pattern.DocumentRange.GetText(-1) or "").strip()
    except Exception:
        pass
    return ""  #必须在这里再 strip 一次：格子内层 TextControl 的 Name 是单个空格 ' '，
    #旧写法只在带 ':' 的分支里 strip，空格原样返回 → 被当成"这格有值"，
    #于是每列都凭空多出一个"有值的行"（UiaRowCount 数出 2 行 / Excel 只有 1 行的根因）


def _detail_panels(tag):
    #全树找 aid=dataPresenter 的 PaneControl（明细容器候选），结果进页面级缓存。
    #存活探测：逐个读 IsOffscreen（每容器一个 COM 调用）；金蝶重建/切页后旧引用读属性会抛，
    #全部读不通就重扫。容器本身在一页里不变，变的是里面的行 —— 行不缓存（_grid_cells 现走）
    cached = _PAGE_CACHE["panels"]
    if cached is not None and _page_cache_usable():
        alive = []
        visible = 0
        for p in cached:
            try:
                off = bool(p.IsOffscreen)
            except Exception:
                continue  #节点已销毁
            alive.append(p)
            if not off:
                visible += 1
        #⚠️命中必须"至少一个容器仍可见"：用户中途切页签时，旧容器全部转 offscreen
        #但节点没销毁（读属性不抛）——只查存活会拿着"当前不可见的那张单"的容器继续配对，
        #_detail_grid 的可见性过滤就只剩旧页签一个候选 → 写进非活动单据（10:19 事故同型）。
        #全部不可见 → 作废重扫（最小化态也走重扫，行为与无缓存时一致）
        if alive and visible:
            return alive
    main = _main_window()
    if main is None:
        return None
    panels = [p for p in _walk(main, want_aid="dataPresenter")
              if p.ControlTypeName == "PaneControl"]
    if panels:  #空结果不缓存：可能只是页面还没加载完，下一动作要重扫
        _PAGE_CACHE["panels"] = panels
        _PAGE_CACHE["hwnd"] = _main_hwnd()
    return panels


def _detail_grid(aid, tag, _retry=True):
    #取"当前这张单据"的明细容器；分不清就返回 None（宁可不填也不能填进别的单据）。
    #失败先作废页面缓存重扫一次再下结论（_retry 防死循环）：缓存过期时的典型表现就是
    #"找不到字段/多张单据歧义"这类假失败，重扫一次既纠错又只在失败路径多花一次遍历
    #
    #为什么必须做这一步：金蝶每轮失败都会留下一个没关的「直接调拨单 - 新增」页签，
    #每个页签里都有一套**同名同 AutomationId** 的格子，Name 还都叫 "Row 0, Column 2"。
    #实测（2026-11-16 10:45，当时开着 3 个单据页）：
    #   aid=dataPresenter 的 Pane 共 6 个（每张单 2 个，只有 1 个装明细表格）
    #   容器[1] offscreen=True  FMATERIALID='01-002-01565'   ← 陈旧页签
    #   容器[3] offscreen=True  FMATERIALID=''               ← 陈旧页签
    #   容器[5] offscreen=False FMATERIALID=''               ← 当前页签（用户在看的那个）
    #旧写法 _walk(主窗口, want_aid=FMATERIALID) 把三个全捞出来按顺序取第一个
    #→ 10:19 那轮日志写着"写入成功、回读一致"，实际写进了一张看不见的陈旧单据，
    #  用户盯着当前单据看，结论是"根本没填进去"（他的观察完全正确）。
    #这不是识别精度问题，是**写错对象**：陈旧单据一旦被保存就是脏数据，事故级。
    #
    #判据：先只保留"装有该字段格子"的容器（每张单有 2 个 dataPresenter，
    #另一个装的是别的面板，不含明细格子），再看可见性。
    #窗口最小化时金蝶把所有元素都标成 offscreen（11-10 §2.5 的离屏虚拟坐标），
    #那时靠可见性分不出来 → 一律拒绝并提示关掉残留页签。
    #残留页签关掉后 holders==1，最小化态照样能写，不影响后台目标。
    panels = _detail_panels(tag)
    if panels is None:
        return None
    if not panels:
        if _retry:
            _invalidate_page_cache()
            return _detail_grid(aid, tag, _retry=False)
        print(f"[{tag}] 界面上没有明细表格（当前页不是单据编辑页？先点「新增」把单开出来）")
        return None
    holders = [p for p in panels if _grid_cells(p, aid)]
    if not holders:
        if _retry:
            _invalidate_page_cache()
            return _detail_grid(aid, tag, _retry=False)
        print(f"[{tag}] 明细表格里找不到字段 {aid}（列没渲染出来？横向滚动或列被隐藏）")
        return None
    if len(holders) == 1:
        return holders[0]
    visible = [p for p in holders if not _is_offscreen(p)]
    if len(visible) == 1:
        return visible[0]
    if _retry:
        #假歧义：缓存里可能还留着已被用户关掉的页签容器，重扫一次再判
        _invalidate_page_cache()
        return _detail_grid(aid, tag, _retry=False)
    print(f"[{tag}] 同时开着 {len(holders)} 张单据（可见 {len(visible)} 张），"
          f"分不清该往哪张单写，拒绝写入。"
          f"请把之前失败留下的「直接调拨单 - 新增」页签关掉再跑")
    return None


def _is_offscreen(node):
    #读 IsOffscreen，读不到时按"可见"处理（宁可多试一次，别把正常单挡在门外）
    try:
        return bool(node.IsOffscreen)
    except Exception:
        return False


def _grid_cells(grid, aid):
    #取该容器里指定字段的所有格子，排除格子内层的文本节点。
    #同一个 AutomationId 会同时挂在"格子本身(Custom)"和"格子里的文本(Text)"上，
    #旧写法两者都收 → 每行被数两次（10:19 那轮 UiaRowCount 数出 2 行 / Excel 只有 1 行的另一半根因）
    out = []
    for node in _walk(grid, want_aid=aid):
        if node.ControlTypeName == "TextControl":
            continue
        out.append(node)
    return out


def _in_main_window(rect, win_rect):
    #矩形是否落在主窗口内。IsOffscreen==False **不代表可见** ——
    #金蝶把单据区做成可横向虚拟滚动的面板，滚出视野的列、别的页签里算过布局的元素，
    #IsOffscreen 都还是 False，只有矩形坐标把它们露出来：
    #2026-11-16 实测主窗口右边缘 x=1931，而「备注」标签有两个 offscreen=False 的实例，
    #一个在 (1046,461)（真单头），一个在 (2902,904)（明细表格里 FNOTEENTRY 那一列的列头）。
    #单头定位如果不过这道闸，就会把单头的值写到明细列头上去。
    if not win_rect:
        return True
    left, top, right, bottom = win_rect
    return (rect[0] >= left - 2 and rect[1] >= top - 2
            and rect[2] <= right + 2 and rect[3] <= bottom + 2)


def _header_rect(node):
    try:
        r = node.BoundingRectangle
    except Exception:
        return None
    return (r.left, r.top, r.right, r.bottom)


def _read_header_value(node, pattern):
    #读单头编辑框当前值。明细格子的 ValuePattern getter 恒为空（见 _read_cell_text 的注释），
    #但单头这几个控件不是 —— 实测 KDPopupBaseEdit 能直接读出解析后的完整名称。
    #所以这里 ValuePattern 优先，读不到再退 Name / TextPattern。
    try:
        if pattern is not None:
            raw = pattern.Value
            if raw is not None and str(raw).strip():
                return str(raw)
    except Exception:
        pass
    try:
        name = (node.Name or "").strip()
    except Exception:
        name = ""
    #金蝶自绘控件的 Name 常常就是类名本身（'KDPopupBaseEdit'），这种当没值，
    #否则"回读到一个类名"会被误判成"填上了"
    try:
        cls_name = node.ClassName or ""
    except Exception:
        cls_name = ""
    if name and name != cls_name:
        return name
    try:
        tp = node.GetPattern(_uia.PatternId.TextPattern)
        if tp is not None:
            return (tp.DocumentRange.GetText(-1) or "").strip()
    except Exception:
        pass
    return ""


def _header_scan(tag):
    #全树扫一遍：_HEADER_FIELDS 里所有候选标签文本 + 带 ValuePattern 的候选编辑框。
    #返回 (labels_by_text, edits, win_rect)；labels_by_text: 文本→[(节点,矩形)]，
    #edits: [(节点, pattern, 矩形, 类名)]。一次遍历服务所有字段（旧写法每个字段各扫一遍全树，
    #单头 8 个字段就是 8 次几千 COM 调用的遍历 —— 实测 5m1s 里的最大单项开销）
    main = _main_window()
    if main is None:
        print(f"[{tag}] 没找到金蝶主窗口（客户端是不是没开/停在登录页？）")
        return None
    win_rect = _header_rect(main)
    wanted = set(_HEADER_FIELDS.values())
    labels_by_text = {}
    edits = []
    for node in _walk(main):
        rect = _header_rect(node)
        if rect is None or _is_offscreen(node) or not _in_main_window(rect, win_rect):
            continue
        try:
            ct = node.ControlTypeName
            cls = node.ClassName or ""
            aid = node.AutomationId or ""
            name = (node.Name or "").strip()
        except Exception:
            continue
        if ct == "TextControl" and name in wanted:
            labels_by_text.setdefault(name, []).append((node, rect))
            continue
        if aid == "PART_Editor":  #内层部件，和父控件矩形重合，留着会配出两个候选
            continue
        if not any(hint in cls for hint in _HEADER_CLASS_HINTS):
            continue
        try:
            pattern = node.GetPattern(_uia.PatternId.ValuePattern)
        except Exception:
            pattern = None
        if pattern is None:
            continue
        edits.append((node, pattern, rect, cls))
    return labels_by_text, edits, win_rect


def _header_scan_cached(tag):
    cached = _PAGE_CACHE["header"]
    if cached is not None and _page_cache_usable():
        labels_by_text, edits, win_rect = cached
        probe = None
        for bucket in labels_by_text.values():
            if bucket:
                probe = bucket[0][0]
                break
        if probe is None and edits:
            probe = edits[0][0]
        alive = True
        if probe is not None:
            try:
                #同 _detail_panels：切页签后旧标签节点还活着但已 offscreen，
                #只查存活会把"上一页的编辑框"当这页的结果交出去 → 必须连可见性一起复核
                if probe.IsOffscreen:
                    alive = False
            except Exception:
                alive = False
        if alive:
            main = _main_window()
            cur = _header_rect(main) if main is not None else None
            #主窗矩形变了（最大化/还原/挪动）→ 缓存里所有坐标作废
            if cur is not None and cur == win_rect:
                return cached
    fresh = _header_scan(tag)
    if fresh is not None and (fresh[0] or fresh[1]):  #空结果不缓存：页面可能还没加载完
        _PAGE_CACHE["header"] = fresh
        _PAGE_CACHE["hwnd"] = _main_hwnd()
    return fresh


def _header_editor(field, tag, _retry=True):
    #按界面标签定位单头编辑框，返回 (node, value_pattern, 标签rect, 控件rect, 类名)；失败返回 None。
    #失败先作废页面缓存重扫一次再下结论（_retry 防死循环），理由同 _detail_grid
    #
    #为什么不能像明细那样用 AutomationId：单头控件的 aid 是 WPF 模板部件名 'PART_Editor'，
    #一张单里几十个控件共用同一个名字，根本不唯一（实测单据编辑页没有任何含 ORG 的 aid）。
    #唯一可靠的锚是**标签文本**：金蝶把「调出库存组织」这类标签渲染成可见 TextControl，
    #Name 就是这几个字，编辑框就在它右边。
    #
    #配对判据用"垂直重叠 > 4px"而不是"中心线差 <= 14px"：
    #「备注」是多行高控件（实测 rect 高 70px），中心线差达 20px，用中心线判会把它判成"没有编辑框"
    #—— 我第一版探针就是这么漏掉备注的，横向还漏了整类自绘控件，两个错叠一起得出
    #"单头 0 个可写框、UIA 直写这条路不成立"的错误结论。
    label = _HEADER_FIELDS.get(field)
    if not label:
        print(f"[{tag}] 未知单头字段 {field!r}，可选: {list(_HEADER_FIELDS)}")
        return None
    scanned = _header_scan_cached(tag)
    if scanned is None:
        return None
    labels_by_text, edits, _win_rect = scanned
    labels = labels_by_text.get(label, [])
    if not labels and _retry:
        _invalidate_page_cache()
        return _header_editor(field, tag, _retry=False)
    if not labels:
        #2026-11-18 用户实测：明细列全绿、到单头突然"找不到标签"——真身是**窗口最小化**。
        #最小化时 Win32 给主窗 237×39 的虚拟矩形，所有元素坐标都落在窗外，
        #而明细靠容器定界不看可见性，所以只有单头几何配对这条路会断。
        #报错必须自证这一点，别让下一个人再去猜"标签文本不对"。
        iconic = False
        try:
            iconic = bool(ctypes.windll.user32.IsIconic(_main_hwnd()))
        except Exception:
            pass
        if iconic:
            print(f"[{tag}] 找不到「{label}」标签——**金蝶窗口当前是最小化的**：最小化时"
                  f"全树坐标变虚拟值，单头几何配对必失败（明细列不受影响，所以会出现"
                  f"「前面几步全绿、单头突然全红」的形态）。把窗口还原/最大化再跑；"
                  f"pipeline 里 UiaNavigate 的 wait_label 把关也有同样的可见性依赖")
        else:
            print(f"[{tag}] 界面上没有可见的「{label}」标签（当前页不是单据编辑页？"
                  f"还是这张单的标签文本和 _HEADER_FIELDS 里记的不一样？）")
        return None
    if not edits and _retry:
        _invalidate_page_cache()
        return _header_editor(field, tag, _retry=False)
    if not edits:
        print(f"[{tag}] 找到了「{label}」标签，但附近没有任何带 ValuePattern 的编辑框")
        return None

    def pair(label_rect):
        #标签右边、垂直有重叠、距离最近的那个编辑框
        lx, ly, lr, lb = label_rect
        best = None
        for node, pattern, (ex, ey, er, eb), cls in edits:
            overlap = min(lb, eb) - max(ly, ey)
            dx = ex - lr
            if overlap > 4 and -10 <= dx <= 300:
                key = (abs(dx), -overlap)
                if best is None or key < best[0]:
                    best = (key, node, pattern, (ex, ey, er, eb), cls)
        return best

    #歧义处理：同名标签出现多次（同时开着几张单据页时就会这样）——
    #和 _detail_grid 同一个原则：分不清往哪张单写就拒绝，绝不猜（陈旧单据被保存=脏数据，事故级）
    hits = []
    for _node, rect in labels:
        b = pair(rect)
        if b is not None:
            hits.append((b[1], b[2], rect, b[3], b[4]))
    if not hits and _retry:
        _invalidate_page_cache()
        return _header_editor(field, tag, _retry=False)
    if not hits:
        print(f"[{tag}] 「{label}」标签右侧 300px 内没有编辑框，拒绝写入。标签位置: {labels[0][1]}")
        return None
    if len(labels) > 1 and _retry:
        #假歧义：缓存里可能算进了已关闭页签的标签，重扫一次再判
        _invalidate_page_cache()
        return _header_editor(field, tag, _retry=False)
    if len(labels) > 1:
        print(f"[{tag}] 「{label}」标签在窗口内出现 {len(labels)} 次（可能同时开着多张单据），"
              f"分不清该往哪张单写，拒绝写入。请把失败留下的单据页签关掉再跑")
        return None
    #hits 里每项本来就是 (node, pattern, 标签rect, 控件rect, 类名)，和返回值约定一致
    return hits[0]


def _header_write_target(node):
    #找单头字段真正的可写元素，返回 (元素, ValuePattern)。
    #2026-11-16 probe4/probe7 实测：外层复合控件（KDPopupBaseEdit 等）的 ValuePattern **只读**，
    #SetValue 抛 COMError 0x80040200（UIA_E_INVALIDOPERATION），但同一个 pattern 读值正常；
    #可写的是它内层的 EditControl（probe5 解剖：IsReadOnly=False，写后外层读回就是新值）。
    #优先内层 EditControl（矩形必须在主窗口内 —— 别的页签算过布局的内层也在树里），
    #没有内层（KDComboBoxEdit 这类）就退外层本身：外层 SetValue 大概率抛，
    #抛了由调用方判失败 —— 不静默，也不拿 SetFocus/键盘注入顶替（那是前台路线）。
    for c in _walk(node, want_type="EditControl", max_depth=6):
        if _is_offscreen(c):
            continue
        try:
            pattern = c.GetPattern(_uia.PatternId.ValuePattern)
        except Exception:
            pattern = None
        if pattern is None:
            continue
        try:
            if pattern.IsReadOnly:
                continue  #只读的 Edit 写了也白写（金蝶的显示框 KDNumEditorDisplayer 一类）
        except Exception:
            pass
        return c, pattern
    try:
        return node, node.GetPattern(_uia.PatternId.ValuePattern)
    except Exception:
        return node, None


def _header_value_from_order(field, param):
    #没给 value 时从当前单据队列取单头字段（和 _cell_value 同一套 Excel 契约）。
    #单头字段整单同值，由 transfer_actions._current_order() 的第一行提供。
    getter = _COLUMN_ATTR_BY_FIELD.get(field)
    if getter is None:
        return None
    try:
        import transfer_actions
        order = transfer_actions._current_order()
        if order is None or not order.items:
            return None
        return getter(order.items[0])
    except Exception:
        return None


def _as_index(value, default=0):
    #行号参数的宽松解析：JSON 里写成字符串 "0" 很常见，布尔值不算数字。
    #与 common._as_ms / _parse 的宽松策略一致，别让一个笔误把动作搞成静默跳过
    if isinstance(value, bool):
        return default
    if isinstance(value, int):
        return value if value >= 0 else default
    try:
        parsed = int(str(value).strip())
        return parsed if parsed >= 0 else default
    except Exception:
        return default


def _cell_node(field, tag, row=0, want_lookup=None):
    #按字段名+行号取明细格子。返回 (元素, 字段标识)；任一步失败返回 (None, None)
    #want_lookup=None 不校验；True/False 时校验该字段是否属于查找类（防止动作配错）
    aid = _CELL_FIELDS.get(field)
    if not aid:
        print(f"[{tag}] 未知field: {field!r}，可选: {list(_CELL_FIELDS)}")
        return None, None
    if want_lookup is True and field not in _LOOKUP_FIELDS:
        print(f"[{tag}] {field} 不是查找类字段，请用 UiaSetCell")
        return None, None
    if want_lookup is False and field in _LOOKUP_FIELDS:
        print(f"[{tag}] {field} 是基础资料查找字段，请用 UiaPickList（直写会得到未解析的显示文本）")
        return None, None
    main = _main_window()
    if main is None:
        print(f"[{tag}] 没找到金蝶主窗口（客户端是不是没开？还是停在登录页？）")
        return None, None

    row = _as_index(row)
    #只在"当前这张单据"的容器里找格子：残留的旧单据页签里有同名同 AutomationId 的格子，
    #整窗去捞会写错单据（详见 _detail_grid 的说明）
    grid = _detail_grid(aid, tag)
    if grid is None:
        return None, None
    nodes = _grid_cells(grid, aid)
    if not nodes:
        print(f"[{tag}] 界面上找不到字段 {field}({aid})，当前单据页可能没打开")
        return None, None

    #金蝶每个已填格子的 Name 自带真实行号，优先用它核对 —— 只看节点顺序会被虚拟化/滚动骗到
    #（WPF DataGrid 未实现的行不在树里，第 N 个节点未必是第 N 行）
    #但真实单据是**混合**的：已填行有 "Row N" 前缀、刚新增的空行没有，而且空行可能出现在中间。
    #所以按界面顺序走一遍游标：遇到有前缀就以它为准，遇到无前缀就当"占下一个未知位置"。
    #注意这是启发式 —— 无前缀节点无法自证属于第几行，所以降级时一定打日志
    cursor = -1
    for node in nodes:
        row_no, _col_no = _cell_prefix(_safe_name(node))
        if row_no is None:
            cursor += 1  #无前缀：占下一个还没被行号声明的位置
            if cursor == row:
                print(f"[{tag}] {field} 第 {row} 行没有行号前缀（新增行），按无前缀节点序命中")
                return node, aid
        else:
            cursor = row_no
            if row_no == row:
                return node, aid
    print(f"[{tag}] {field} 找不到第 {row} 行（本单据 {len(nodes)} 个该字段格子）")
    return None, None


def _type_text(text, key_delay_ms=40, hwnd=None):
    #逐字符输入到金蝶窗口（默认主窗；弹窗/下拉这类独立顶层传 hwnd）。
    #必须走字符消息而不是Ctrl+V：
    #WPF 组合键读 GetKeyState，后台 PostMessage 不更新修饰键状态，Ctrl+V 在后台必然失效
    #（这是后台粘贴失效的根因，实测确认）
    #目标只可能是金蝶自己的窗口：绝不退化到 GetForegroundWindow —— 后台跑时前台是用户的程序，
    #把字符投过去就是"打字打进别人窗口"的事故
    hwnd = hwnd or _main_hwnd()
    if not hwnd:
        print("[_type_text] 拿不到金蝶窗口句柄，拒绝投递按键（不能打到用户窗口）")
        return False
    user32 = ctypes.windll.user32
    user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
    user32.PostMessageW.restype = ctypes.c_bool
    for ch in text:
        #WM_CHAR=0x0102，wParam=字符码点，lParam=1 表示真实按键（部分控件靠它区分重复字符）
        if not user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0102, ord(ch), ctypes.c_void_p(1)):
            return False
        if key_delay_ms:
            time.sleep(key_delay_ms / 1000.0)
    return True


def _select_row_and_confirm(row_node):
    #用 UIA 选中列表里的目标行，再点弹窗自己的「确定」。
    #全程零输入注入：不发方向键、不发回车 —— 那些走 keybd_event/SendInput 会打到用户前台程序，
    #走 PostMessage 又要赌 WPF 列表控件是否处理 WM_KEYDOWN，都不如直接调 UIA 稳
    #**关键**：必须用 SelectionItemPattern.Select() 真正改变选中项。
    #不能拿 SetFocus 顶替 —— 焦点不等于选中，列表选中项没变就点「确定」会选到**别的行**，
    #那是静默写错仓库。所以 Select 不可用/失败时直接返回 False，宁可失败也不瞎点
    try:
        selection = row_node.GetPattern(_uia.PatternId.SelectionItemPattern)
    except Exception as exc:
        print(f"[_select_row_and_confirm] 读取 SelectionItem 失败: {exc!r}")
        return False
    if selection is None:
        print("[_select_row_and_confirm] 该列表行不支持 SelectionItemPattern，"
              "无法保证选中目标行（不用 SetFocus 顶替，避免选中项没变就点确定）")
        return False
    try:
        if not selection.Select():
            print("[_select_row_and_confirm] Select() 返回失败")
            return False
    except Exception as exc:
        print(f"[_select_row_and_confirm] Select() 抛异常: {exc!r}")
        return False
    time.sleep(0.4)

    #弹窗底部常有「确定」按钮；有就点，没有就靠"选中即生效"的列表行为
    for button in _walk(_popup_root(row_node), want_type="ButtonControl"):
        if _safe_name(button) not in ("确定", "选择"):
            continue
        try:
            invoke = button.GetPattern(_uia.PatternId.InvokePattern)
            if invoke is not None:
                invoke.Invoke()
                return True
        except Exception:
            continue
    #没有确定按钮：靠"选中即生效"的列表行为。这条路径的安全性依赖两件事——
    #Select 已真正改变选中项（强约束，见上），且调用方会用归一化等值回读校验，
    #真选错了会在校验那步返回False，不会静默写错
    return True


def _popup_root(node):
    #从行节点往上找到弹窗根。GetTopLevelControl 内部走父链，元素销毁时会抛 COMError，
    #绝不能让它穿出 run()，所以这里整体兜住并退化到节点自身
    try:
        root = node.GetTopLevelControl()
        return root if root is not None else node
    except Exception:
        return node


def _main_hwnd():
    #金蝶主窗口句柄。**必须用句柄比较而不是对象比较**：
    #uiautomation 的 Control 没有实现 __eq__，`win == _main_window()` 是对象身份比较，
    #两次 GetChildren() 拿到的即使指向同一窗口也是不同对象 → 恒为 False → 排除条件失效。
    #这个坑正是"清理残留弹窗时把主窗口也关了"的成因
    main = _main_window()
    if main is None:
        return 0
    try:
        return main.NativeWindowHandle or 0
    except Exception:
        return 0


def _close_stale_popups():
    #关掉上一次失败残留的查找列表。
    #为什么必须有：_wait_popup_rows 用"调用前的窗口基线"排除旧弹窗，如果旧的没关，
    #它就会一直躺在基线里 → 本次永远看不到新弹窗 → 失败退出 → 下次基线还是它 → 永久死锁。
    #删掉按键发送后我们失去了"发ESC关弹窗"的手段，所以改用 UIA 点「取消」/Close。
    #**安全红线**：这里绝不碰金蝶主窗口、也绝不点标题栏关闭按钮 —— 关错窗口是灾难级事故
    pid = _main_pid()
    main_hwnd = _main_hwnd()
    if not pid or not main_hwnd:
        return 0
    closed = 0
    try:
        tops = _uia.GetRootControl().GetChildren()
    except Exception:
        return 0
    for win in tops:
        try:
            if win.ProcessId != pid:
                continue
            #排除主窗口只认句柄等值：类名帮不上忙 —— 金蝶自己的弹窗和主窗口同属一个进程，
            #Win32 类名是同一个 HwndWrapper[Kingdee.BOS.XPF.App.exe;;<同一个GUID>]（GUID 按
            #AppDomain 生成，不是按窗口），UIA ClassName 又统一是 'Window'，两种写法都区分不了。
            #旧代码这里写的是「UIA ClassName 比 _MAIN_WINDOW_CLASS」，那是恒 False 的死条件，
            #看着像第二重保险其实一直空转 —— 别再往这条路上加"保险"，真正的判据是下面的形状校验。
            if win.NativeWindowHandle == main_hwnd:
                continue
            if win.NativeWindowHandle == 0:
                continue  #没有独立句柄的不是普通顶层窗，别动
        except Exception:
            continue
        if not _looks_like_selector(win):
            continue
        if _close_window(win):
            closed += 1
    if closed:
        print(f"[_close_stale_popups] 关掉了 {closed} 个残留弹窗")
    return closed


def _looks_like_selector(win):
    #判断一个窗口是不是"金蝶的查找列表弹窗"。
    #只按"有 DataItem"认太宽（主窗口明细也是 DataItem），所以要求同时满足：
    #  ① 有带文本的 DataItem 行
    #  ② 有成对的「确定」+「取消」按钮（查找列表的典型结构；普通单据页没有这对按钮）
    #宁可漏关（本次失败，下次重试）也不误关
    if not [c for c in _walk(win, want_type="DataItemControl", max_depth=10) if _safe_name(c)]:
        return False
    names = set()
    for button in _walk(win, want_type="ButtonControl", max_depth=10):
        names.add(_safe_name(button))
    return "确定" in names and "取消" in names


def _close_window(win):
    #关一个查找列表弹窗：只点「取消」。绝不点「关闭」——
    #金蝶主窗口标题栏的按钮 Name 就是"关闭"，认错窗口+认错按钮会直接关掉客户端
    for button in _walk(win, want_type="ButtonControl", max_depth=10):
        if _safe_name(button) != "取消":
            continue
        try:
            invoke = button.GetPattern(_uia.PatternId.InvokePattern)
            if invoke is not None:
                invoke.Invoke()
                time.sleep(0.4)
                return True
        except Exception:
            continue
    return False


def _safe_name(node):
    #读节点 Name，异常时返回空串（遍历期间元素被销毁是常态）
    try:
        return (node.Name or "").strip()
    except Exception:
        return ""


def _wait_popup_rows(timeout_ms=_DEFAULT_TIMEOUT_MS, exclude=None, strict=False):
    #等金蝶的查找列表弹出来，返回 (弹窗元素, 行元素列表)。
    #三个必须同时满足，否则会认错窗口/认错列表：
    #  ① 属于金蝶进程（别的 WPF 程序也有 DataItem 列表）
    #  ② 不是调用前就存在的旧窗口（上一次失败残留的列表会被当成本次的，然后按旧内容数行）
    #  ③ 真有带文本的 DataItem 行
    #strict=True 时额外要求 _looks_like_selector（有 DataItem + 成对确定/取消）。
    #**为什么默认不 strict**：等待侧有 exclude 基线兜底，宽松能少一截误杀；
    #但清理侧用的是 strict 判据，两边不一致会导致"能等到但清不掉"的窗口跨重试残留死锁。
    #所以调用方在这轮失败后，会用 strict=True 再试一次（见 UiaPickList），两侧标准就对齐了
    pid = _main_pid()
    if not pid:
        return None, []
    exclude = set(exclude or ())
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        try:
            tops = _uia.GetRootControl().GetChildren()
        except Exception:  #COM 树被重绘时可能抛，重试即可
            time.sleep(_POLL_MS / 1000.0)
            continue
        for win in tops:
            try:
                if win.ProcessId != pid:
                    continue  #不是金蝶的窗口，直接跳过
                if win.NativeWindowHandle in exclude:
                    continue  #调用前就存在的窗口，不是这次弹出的列表
                if strict and not _looks_like_selector(win):
                    continue  #严格模式：必须长得像查找列表
                rows = [c for c in _walk(win, want_type="DataItemControl") if _safe_name(c)]
                if rows:
                    return win, rows
            except Exception:
                continue
        time.sleep(_POLL_MS / 1000.0)
    return None, []


def _top_handles():
    #记录当前所有顶层窗口句柄，用作"哪些弹窗是调用前就有的"基线
    handles = set()
    if not _UIA_AVAILABLE:
        return handles
    try:
        for win in _uia.GetRootControl().GetChildren():
            try:
                handles.add(win.NativeWindowHandle)
            except Exception:
                continue
    except Exception:
        return handles
    return handles


def _row_texts(row):
    #把一个 DataItem 行的所有文本片段收集起来，用于包含匹配与最终等值校验。
    #先按 TextControl 收（金蝶列表行里的名称/编码通常是文本控件）；
    #但表格单元格本身是 CustomControl，列表行也可能用 CustomControl 包文本，
    #所以收不到时**回退到任意后代里非空的 Name** —— 只认 TextControl 会让命中率归零
    texts = []
    for node in _walk(row, want_type="TextControl", max_depth=8):
        try:
            name = (node.Name or "").strip()
        except Exception:
            continue
        if name:
            texts.append(name)
    if texts:
        return texts
    for node in _walk(row, max_depth=8):
        try:
            name = (node.Name or "").strip()
            cls_name = (node.ClassName or "").strip()
        except Exception:
            continue
        if not name or node is row:
            continue
        #跳过"Name 就是控件类名"的噪声节点（UIA 拿不到内容时会拿类名充当 Name）。
        #注意只在类名非空时比 —— CustomControl 的 ClassName 常是空串，
        #无条件比会把真正带值的节点也过滤掉（这个坑被单测抓到过）
        if cls_name and name == cls_name:
            continue
        texts.append(name)
    return texts


def _cell_value(param):
    #取要写入的值：优先 value 参数（固定值）；没给就用 field 从当前单据队列里取，
    #复用 PasteOrderColumn 的数据源，pipeline 侧换动作时不用改 Excel 契约
    #取法必须走显式映射表，不能 getattr(field)
    value = param.get("value")
    if value is not None:
        return str(value)
    field = param.get("field", "")
    getter = _COLUMN_ATTR_BY_FIELD.get(field)
    if getter is None:
        return None
    try:
        import transfer_actions
        order = transfer_actions._current_order()
        if order is None or not order.items:
            return None
        row = _as_index(param.get("row", 0))
        if not 0 <= row < len(order.items):
            return None
        return getter(order.items[row])
    except Exception:
        return None


def _anchor_row_nos(anchor="material_code"):
    #用"锚点列"（默认物料编码）判定哪些行是**真实数据行**，返回有值的行号集合。
    #为什么不能拿目标列自己判：UiaFillColumn 的用途就是把一列**从空填满**，
    #此时目标列必然全是空串，用它判行会得出"一行都没有"，动作在主场景下 100% 失败。
    #锚点列是粘进来的，必有值，所以用它定行才可靠
    aid = _CELL_FIELDS.get(anchor)
    if not aid:
        return None
    grid = _detail_grid(aid, "UiaRowCount")
    if grid is None:
        return None
    row_nos = set()
    cursor = -1
    for node in _grid_cells(grid, aid):
        row_no, _col = _cell_prefix(_safe_name(node))
        if row_no is None:
            cursor += 1
            row_no = cursor
        else:
            cursor = row_no
        if _read_cell_text(node):
            row_nos.add(row_no)
    return row_nos


def _rows_by_no(field):
    #取该列所有格子，返回 {行号: 元素}。行号取 Name 里的真实 Row N，拿不到时按顺序补号。
    #**这里不按"有没有值"过滤** —— 空格子也是真实存在的可写目标，
    #"哪些行该填"由 _anchor_row_nos 决定，两件事分开
    aid = _CELL_FIELDS.get(field)
    if not aid:
        return None
    grid = _detail_grid(aid, "UiaFillColumn")
    if grid is None:
        return None
    mapping = {}
    cursor = -1
    for node in _grid_cells(grid, aid):
        row_no, _col = _cell_prefix(_safe_name(node))
        if row_no is None:
            cursor += 1
            row_no = cursor
        else:
            cursor = row_no
        mapping[row_no] = node
    return mapping


def _visible_rows(field, anchor="material_code"):
    #返回 [(行号, 元素)]（按行号升序），行号来自锚点列判定的真实数据行。
    #该行在目标列上没有格子节点时也保留行号、元素为 None ——
    #让调用方能报"这一行没渲染出来"，而不是静默少写一行（静默少录比报错糟）
    row_nos = _anchor_row_nos(anchor)
    if row_nos is None:
        return None
    by_no = _rows_by_no(field)
    if by_no is None:
        return None
    return [(row_no, by_no.get(row_no)) for row_no in sorted(row_nos)]


def _fill_rows(field, values, tag, anchor="material_code"):
    #把 values 按**Excel 行号**逐行写进 field 这一列，返回 (成功行数, 失败原因)
    aid = _CELL_FIELDS.get(field)
    if not aid:
        return 0, f"未知field: {field!r}"
    if field in _LOOKUP_FIELDS:
        return 0, f"{field} 是查找字段，本动作不支持（请用 UiaPickList）"
    rows = _visible_rows(field, anchor)
    if rows is None:
        return 0, "没找到金蝶主窗口"
    if not rows:
        return 0, f"锚点列 {anchor} 没有已填的数据行（先把物料粘进来再填这一列）"
    return _write_each_row(field, values, rows, tag)


def _write_each_row(field, values, rows, tag):
    #逐行写值。**按真实行号配对**，不用列表下标——
    #中间有空行时下标会整体顺移，把"Excel第2行"写进"界面第1行"（静默错位）
    written = 0
    expected = len(values)
    by_no = {row_no: node for row_no, node in rows}
    missing = []
    for index, value in enumerate(values):
        node = by_no.get(index)  #Excel 第 index 行 → 界面 Row index
        if node is None:
            if value not in (None, ""):
                missing.append(index)  #有值但界面没这一行，属于"物料没粘全"
            continue
        if value is None or value == "":
            continue  #空值跳过：不写，也不清空原有内容
        try:
            pattern = node.GetPattern(_uia.PatternId.ValuePattern)
        except Exception:
            pattern = None
        if pattern is None:
            _log(tag, f"第 {index} 行的 {field} 格子不支持写值")
            continue
        try:
            pattern.SetValue(str(value))
            written += 1
        except Exception as exc:
            _log(tag, f"第 {index} 行写值失败: {exc!r}")
        time.sleep(0.12)  #逐行写要留一点点间隙，太快金蝶来不及刷新格子
    if expected > len(rows):
        _log(tag, f"注意：Excel 有 {expected} 个值，界面只有 {len(rows)} 行（物料可能没粘全）")
    if missing:
        _log(tag, f"注意：第 {missing[:8]} 行有值但界面没有对应格子（没粘全或未渲染）")
    #逐行写完整列归还一次（不在循环里每行还）：每行 SetValue 都会偷一次前台，
    #连还 66 次是 66 下抖动；写完统一还一次，金蝶占用时长=整列写入时长，可接受
    _fg_restore()
    return written, None


@AgentServer.custom_action("UiaSetForm")  #切换明细字段白名单（transfer|requisition），任务第一个节点调
class UiaSetForm(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)  #拉取参数
        if param is None:
            return False
        form = str(param.get("form", "") or "").strip()
        if form not in _FORM_FIELDS:
            print(f"[UiaSetForm] 未知 form {form!r}，可选: {list(_FORM_FIELDS)}")
            return False
        #原地改内容不重新绑定：_cell_node/_rows_by_no 等一堆函数读的都是这些全局容器
        _CELL_FIELDS.clear()
        _CELL_FIELDS.update(_FORM_FIELDS[form])
        #查找类闸跟着切：申请单仓库/仓位按编码直写就是正路（probe25/26），闸清空；
        #直接调拨单维持"查找字段必须走 UiaPickList"的老约束
        _LOOKUP_FIELDS.clear()
        _LOOKUP_FIELDS.update(
            _LOOKUP_FIELDS_REQUISITION if form == "requisition" else _LOOKUP_FIELDS_TRANSFER)
        _invalidate_page_cache()  #任务入口：页面级缓存从这一张新单开始攒
        print(f"[UiaSetForm] 明细字段白名单已切到 {form}（{sorted(_CELL_FIELDS)}），"
              f"查找类闸={sorted(_LOOKUP_FIELDS) or '无'}")
        return True  #白名单切换本身不碰界面，成败只取决于 form 名字对不对


#整列广播的字段（申请单的组织两列）：界面上**每行必填**，但 Excel 契约里它们是 order 级
#——整单同值、只有首行有单元格内容。逐行取值会得到「首行有值、其余全空」，
#UiaFillColumn 默认语义（空值跳过）会把其余行留空 → 保存时金蝶弹「每行组织必录」。
#broadcast:true 时改从 TransferOrder 整单级取一个值，铺满所有数据行。
#白名单只列这两个，别把 quantity/note 也广播（那是真·逐行数据）
_BROADCAST_ATTR_BY_ORDER = {
    "source_org": lambda order: order.source_org,
    "target_org": lambda order: order.target_org,
}


@AgentServer.custom_action("UiaFillColumn")  #按Excel顺序逐行填满一列（跳过金蝶末尾的空行）
class UiaFillColumn(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaFillColumn"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        field = param.get("field", "")
        if param.get("broadcast"):
            #整列同值广播（申请单组织两列专用，见 _BROADCAST_ATTR_BY_ORDER 的说明）：
            #从当前单据的整单级取一个值，按 Excel 行数铺满；取不到值=整单该列为空
            values = _broadcast_column_values(field)
            if values is None:
                print(f"[UiaFillColumn] broadcast 取不到 {field} 的整单值（field 不在广播白名单/队列没就绪）")
                return False
            if not values:
                print(f"[UiaFillColumn] {field} 整列为空，无需填写")
                return True
            #广播列**不能走 _fill_rows 的锚点判行**（2026-11-17 probe18 实测）：
            #组织列要在物料还没写之前填，此时锚点列 material_code 整列是空的，
            #_anchor_row_nos 会判"没有数据行"直接失败 —— 而金蝶草稿默认带当前登录组织，
            #失败被默认值掩盖，真实业务组织≠默认组织时就是每张单静默填错组织（事故级）。
            #广播的行集合由 Excel 行数定：第 i 个值 → 界面 Row i，格子节点从目标列自己拿
            rows = _rows_by_no(field)
            if rows is None:
                _log("UiaFillColumn", f"{field} 广播取格子失败（主窗口/容器没找到）")
                return False
            indexed = [(i, rows.get(i)) for i in range(len(values))]
            written, error = _write_each_row(field, values, indexed, "UiaFillColumn")
        else:
            values = _column_values(param)
            if values is None:
                print(f"[UiaFillColumn] 取不到 {field} 这一列的值（队列没就绪或field不在取值表）")
                return False
            if not values:
                print(f"[UiaFillColumn] {field} 整列为空，无需填写")
                return True  #空列不是错误，跟 UiaSetCell 的空值语义保持一致
            written, error = _fill_rows(field, values, "UiaFillColumn")
        if error:
            _log("UiaFillColumn", error)
            return False

        #判据必须"写满所有非空值"才算过，不能"写了一行就算过"：
        #只判 written>0 时，66 行里成功 3 行也返回 True，而下游校验通常只看第 0 行，
        #第一行对、后面全错 → 校验通过 → 静默少录（比报错严重得多）
        expected = len([v for v in values if v not in (None, "")])
        if written < expected:
            _log("UiaFillColumn", f"{field} 只写入 {written} 行，预期 {expected} 行（有值行数），判失败")
            return False
        if written == 0:
            _log("UiaFillColumn", "一行都没写进去")
            return False
        _log("UiaFillColumn", f"{field} 写入 {written} 行（队列 {len(values)} 个值）")
        time.sleep(0.6)  #给金蝶把这一列渲染完，后面校验节点读到的是最终值
        return True  #写满全部非空值才算成功


@AgentServer.custom_recognition("UiaColumnFilled")  #整列每一行都填了（对账用，比只看首行强）
class UiaColumnFilled(CustomRecognition):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def analyze(
            self,
            context: Context,
            argv: CustomRecognition.AnalyzeArg,
    ) -> "CustomRecognition.AnalyzeResult":

        if not _require_uia("UiaColumnFilled"):
            return None
        _reset_cache()  #每个识别重新认窗口
        param = _params(argv, attr="custom_recognition_param", tag="UiaColumnFilled")
        if param is None:
            return None
        field = param.get("field", "")
        nonzero = bool(param.get("nonzero"))
        aid = _CELL_FIELDS.get(field)
        if not aid:
            _log("UiaColumnFilled", f"未知field: {field!r}")
            return None

        #行集合来自锚点列（物料），保证"该有几行"和"物料真粘了几行"一致，
        #不会被金蝶末尾的空行带偏；再逐行去目标列读值
        row_nos = _anchor_row_nos(param.get("anchor", "material_code"))
        if row_nos is None:
            return None
        by_no = _rows_by_no(field)
        if by_no is None:
            return None

        empty_rows = []
        zero_rows = []
        for row_no in sorted(row_nos):
            node = by_no.get(row_no)
            value = _read_cell_text(node) if node is not None else ""
            if not value:
                empty_rows.append(row_no)
            elif nonzero and not _is_nonzero(value):
                zero_rows.append(row_no)

        total = len(row_nos)
        _log("UiaColumnFilled",
             f"{field} 共 {total} 行：空 {len(empty_rows)} 行、为0 {len(zero_rows)} 行")
        if not row_nos:
            _log("UiaColumnFilled", f"锚点列没有数据行，无法校验 {field}")
            return None  #一行都没有说明物料没粘上，不该把这里判成"填满了"
        if empty_rows or zero_rows:
            return None  #有空格子/0值就不命中，交给调用方重填或报错
        return CustomRecognition.AnalyzeResult(
            box=(0, 0, 1, 1),
            detail={"field": field, "rows": total})


@AgentServer.custom_recognition("UiaRowCount")  #界面实际可填行数 vs Excel 队列行数（对账用）
class UiaRowCount(CustomRecognition):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def analyze(
            self,
            context: Context,
            argv: CustomRecognition.AnalyzeArg,
    ) -> "CustomRecognition.AnalyzeResult":

        if not _require_uia("UiaRowCount"):
            return None
        _reset_cache()  #每个识别重新认窗口
        param = _params(argv, attr="custom_recognition_param", tag="UiaRowCount")
        if param is None:
            return None
        field = param.get("field", "material_code")
        values = _column_values({"field": field})
        if values is None:
            _log("UiaRowCount", "取不到队列（ReadTransferExcel/NextOrder 跑过了吗）")
            return None
        #行数一律用锚点列（物料）判定 —— 目标列在被填之前是空的，拿它数会永远是 0
        row_nos = _anchor_row_nos(field if field in _CELL_FIELDS else "material_code")
        if row_nos is None:
            return None
        queue_rows = len(values)
        ui_rows = len(row_nos)
        _log("UiaRowCount", f"界面已填 {ui_rows} 行 / Excel 队列 {queue_rows} 行"
                            f"（金蝶末尾的空行已排除）")
        #只做正向判断（相等才命中）。要"不一致才命中"就在 pipeline 节点上写 inverse:true，
        #由框架取反 —— 项目里 OrderRemaining 就是这个惯例。
        #不在这里读 inverse：两层都取反会变回原语义，是隐形的双重否定 bug
        if ui_rows != queue_rows:
            return None  #数量对不上就不命中，交给 on_error / 兜底分支
        return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1),
                                               detail={"ui": ui_rows, "queue": queue_rows})


def _broadcast_column_values(field):
    #broadcast 版取列：整单一个值复制 N 份（N=Excel 行数），给 UiaFillColumn 的 broadcast 分支。
    #值取不到（队列没就绪/field 不在广播白名单）返回 None 让调用方报错，
    #整单该列本来就是空 → 返回空串列表，走"整列为空，无需填写"那条现成分支
    getter = _BROADCAST_ATTR_BY_ORDER.get(field)
    if getter is None:
        return None
    try:
        import transfer_actions
        order = transfer_actions._current_order()
        if order is None:
            return None
        value = str(getter(order) or "")
        return [value] * len(order.items)
    except Exception:
        return None


def _column_values(param):
    #从当前单据队列取某一列的**全部**行值（不去空行、不截首行），给 UiaFillColumn/UiaRowCount 用
    #取值走显式 lambda 表，不用 getattr
    field = param.get("field", "")
    getter = _COLUMN_ATTR_BY_FIELD.get(field)
    if getter is None:
        return None
    try:
        import transfer_actions
        order = transfer_actions._current_order()
        if order is None:
            return None
        return [getter(item) for item in order.items]
    except Exception:
        return None


@AgentServer.custom_action("UiaSetCell")  #按字段标识给明细格子写值（UIA，零鼠标零焦点）
class UiaSetCell(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaSetCell"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        field = param.get("field", "")
        node, aid = _cell_node(field, "UiaSetCell", param.get("row", 0), want_lookup=False)
        if node is None:
            return False  #字段名写错/界面没就绪/把查找字段配到了这个动作

        value = _cell_value(param)
        if value is None or value == "":
            #空值不写（写了会把已有数据清空），但也**不能算失败**：
            #Excel 里备注/数量这类列本来就允许留空，判失败会让 on_error 白转一圈。
            #所以按"本行无需填这一列"处理，返回 True 让流程继续
            print(f"[UiaSetCell] {field} 没有值（Excel空值），按无需填写处理")
            return True

        try:
            pattern = node.GetPattern(_uia.PatternId.ValuePattern)
        except Exception:
            pattern = None
        if pattern is None:
            print(f"[UiaSetCell] {field} 格子不支持写值（{aid}）")
            return False
        try:
            pattern.SetValue(value)
        except Exception as exc:
            print(f"[UiaSetCell] 写值失败: {exc!r}")
            return False
        _fg_restore()  #金蝶在 SetValue 那一瞬间偷走前台，立刻还回去再等解析（2026-11-18 用户实测"归还慢"）
        time.sleep(1.5)  #给金蝶解析基础资料留时间（物料编码会连带带出名称规格）

        actual = _read_cell_text(node)
        if not actual:
            #回读空有两种可能：①金蝶把这个值吞了（名称形态的基础资料字段必吞，
            #  实测基础资料名称无论简称还是全称都被吞，还会把格子原值清掉）；
            #  ②写入触发了自动增行/重绘，手里这个 node 引用过期了（值其实在新节点上）。
            #所以先重取格子再判 —— 但**重取后还是空就必须失败**：
            #旧版这里静默返回 True，等于"名称形态填进 Excel 就带病往下走"，
            #保存时要么弹必录要么存张缺仓库的错单（probe 实测回读='' 仍 True 放行）
            _reset_cache()
            retry_node, _retry_aid = _cell_node(field, "UiaSetCell", param.get("row", 0), want_lookup=False)
            if retry_node is not None:
                actual = _read_cell_text(retry_node)
                node = retry_node
        print(f"[UiaSetCell] {field} ← {value!r}，界面回读 {actual!r}")
        if not actual:
            print(f"[UiaSetCell] {field} 写入后回读仍是空 —— 金蝶把这个值吞了。"
                  f"基础资料列（仓库/仓位）只认**编码**，名称形态一律被吞；"
                  f"Excel 该列请改填编码，或留空（非必填列留空合法）")
            return False  #吞值=配置/数据错误，重跑一遍也是吞，直接失败让 pipeline 停下
        expect = param.get("expect", "")
        if expect and expect not in actual:
            print(f"[UiaSetCell] 回读校验不通过：期望含 {expect!r}")
            return False  #回读不符，让pipeline走on_error
        return True  #写入并校验通过


def _pick_text(param):
    #取查找列表的过滤词。两种来源，优先级：text_field（从当前单据取）> text（固定值）
    #为什么优先从单据取：仓库/仓位每张单都不一样，写死 text 只能跑一种单。
    #text_field 的值走和 UiaSetCell 同一套字段白名单，Excel 模板里就有这四列
    text_field = str(param.get("text_field", "") or "").strip()
    if text_field:
        getter = _COLUMN_ATTR_BY_FIELD.get(text_field)
        if getter is None:
            _log("UiaPickList", f"text_field 未知: {text_field!r}，可选: {list(_COLUMN_ATTR_BY_FIELD)}")
            return ""
        try:
            import transfer_actions
            order = transfer_actions._current_order()
            if order is None or not order.items:
                _log("UiaPickList", f"text_field={text_field} 取不到当前单据（队列没就绪）")
                return ""
            row = _as_index(param.get("row", 0))
            if not 0 <= row < len(order.items):
                _log("UiaPickList", f"当前单只有 {len(order.items)} 行，取不到第 {row} 行")
                return ""
            value = getter(order.items[row])
            _log("UiaPickList", f"text_field={text_field} 第 {row} 行 → 过滤词 {value!r}")
            return str(value or "").strip()
        except Exception as exc:
            _log("UiaPickList", f"text_field={text_field} 取值异常: {exc!r}")
            return ""
    return str(param.get("text", "") or "").strip()


@AgentServer.custom_action("UiaPickList")  #查找字段：给过滤词→选中列表命中行（整段后台完成）
class UiaPickList(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaPickList"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)
        if param is None:
            return False
        field = param.get("field", "")
        text = _pick_text(param)
        if not text:
            _log("UiaPickList", "取不到过滤词：text 和 text_field 都没给有效值"
                                "（没有过滤词会列出全表，选中哪行全看运气）")
            return False

        node, aid = _cell_node(field, "UiaPickList", param.get("row", 0), want_lookup=True)
        if node is None:
            return False

        #先清掉上一次失败残留的列表：否则它永远躺在下面的句柄基线里，
        #_wait_popup_rows 会把它排除掉，本次永远等不到"新"弹窗 → 跨重试死锁
        _close_stale_popups()
        baseline = _top_handles()  #调用前就存在的窗口，不算这次弹出的列表
        try:
            if not node.SetFocus():
                print(f"[UiaPickList] {field} SetFocus 失败，不敢继续投递字符")
                return False  #焦点没进去就投字符，等于往主窗口乱打字
        except Exception as exc:
            print(f"[UiaPickList] SetFocus 异常: {exc!r}")
            return False
        time.sleep(0.5)  #等金蝶把编辑框挂出来

        if not _type_text(text):
            print("[UiaPickList] 输入过滤词失败")
            return False
        time.sleep(2.0)  #等列表加载（后台可能比前台慢，实测约1.5~2.2秒）

        popup, rows = _wait_popup_rows(exclude=baseline)
        if popup is None:
            #宽判据没等到，可能是"宽判据认到的窗口其实是残留的旧弹窗、已被基线排除"。
            #清一次残留再按严格判据重试，避免"能认到但清不掉"的窗口跨重试死锁
            _close_stale_popups()
            popup, rows = _wait_popup_rows(timeout_ms=1500, exclude=baseline, strict=True)
        if popup is None:
            print(f"[UiaPickList] {field} 的查找列表没出来")
            return False  #偶发失效，交给on_error重试

        hit_index = -1
        hit_texts = []
        for index, row_node in enumerate(rows):
            texts = _row_texts(row_node)
            if any(text in t for t in texts):
                hit_index = index
                hit_texts = texts
                break
        print(f"[UiaPickList] {field} 列表 {len(rows)} 行，命中第 {hit_index} 行 {hit_texts[:2]!r}")
        if hit_index < 0:
            print(f"[UiaPickList] 列表里没有含 {text!r} 的行")
            return False  #没命中，不瞎选第一行

        if not _select_row_and_confirm(rows[hit_index]):
            print(f"[UiaPickList] {field} 行选中失败")
            return False
        time.sleep(2.0)

        actual = _read_cell_text(node)
        #校验必须是**归一化等值**，两头都不能松：
        #  判"包含"会让选到隔壁行也通过（"仓库A" in "仓库A临时仓"）；
        #  要求完全等值又会被列表行格式差异搞成假失败（"名称(编码)" vs "名称"）
        #所以先剥掉行文本里的括号后缀再比等值 —— 宁可失败也不放过选错行
        matched = any(_normalize_row_text(t) == actual for t in hit_texts if t)
        print(f"[UiaPickList] {field} 选择后界面值 {actual!r}，命中行文本 {hit_texts!r}")
        if not matched:
            print(f"[UiaPickList] 选择没生效或选错行（命中行 {hit_texts!r}，实际 {actual!r}）")
            return False
        return True  #选择并校验通过


def _normalize_row_text(text):
    #把列表行文本归一成"纯名称"，用于和格子回读值做等值比对。
    #只剥"末尾像编码"的括号后缀（金蝶列表常显示 "名称(编码)"），
    #不能见括号就剥 —— 名字本身可能带括号（"仓库(临时)" 与 "仓库(长期)" 会归一成同一个，放行选错行）
    #所以要求括号后缀短且像编码（无中文、无空格）
    value = (text or "").strip()
    for left, right in (("(", ")"), ("（", "）")):
        index = value.rfind(left)
        if index <= 0 or not value.rstrip().endswith(right):
            continue
        suffix = value[index + 1:value.rstrip().rfind(right)]
        if not suffix or len(suffix) > 24:
            continue  #太长的不像编码，当名字的一部分保留
        if any("\u4e00" <= ch <= "\u9fff" for ch in suffix):
            continue  #含中文的多半是名字本身，不剥
        return value[:index].strip()
    return value


@AgentServer.custom_action("UiaReadCell")  #回读明细格子值（代替OCR做校验）
class UiaReadCell(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaReadCell"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)
        if param is None:
            return False
        field = param.get("field", "")
        node, aid = _cell_node(field, "UiaReadCell", param.get("row", 0))
        if node is None:
            return False

        actual = _read_cell_text(node)
        _log("UiaReadCell", f"{field}({aid}) = {actual!r}")

        #not_expect：值**不等于**它才算通过。数量列这类"默认值"场景必须用它 ——
        #数量列默认是 0，粘上物料后肯定不是 0，所以判"非0"就是在判"粘贴生效了"。
        #用 expect 做不到这件事：expect 是包含匹配，"0" 会命中 "10"、"20"、"100"
        #比较要走数值：金蝶把 0 写成 "0.0000000000"，纯字符串比 "0.0000000000" != "0" 会误判通过
        not_expect = param.get("not_expect")
        if not_expect is not None and not_expect != "":
            return not _same_number(actual, not_expect)

        #nonzero：值不是 0（含 "0"、"0.0"、"0.0000000000" 这类金蝶的补零写法）
        if param.get("nonzero"):
            return _is_nonzero(actual)

        expect = param.get("expect", "")
        if not expect:
            return True  #只读不校验
        if param.get("match_any"):
            keys = param.get("keys")
            if not isinstance(keys, list) or not keys:
                _log("UiaReadCell", "match_any 需要 keys 数组")
                return False  #类型不对直接失败，别按字符迭代字符串
            return any(str(k) in actual for k in keys)
        return expect in actual


def _is_nonzero(text):
    #判断格子值"不是 0"。金蝶数量列会把 5 显示成 "5.0000000000"，
    #所以不能直接比字符串，要按数值判；读不出数字（空串/文字）一律算"非0"，
    #因为空值意味着粘贴根本没生效，交给调用方失败更安全
    value = (text or "").strip()
    if not value:
        return False  #空值 = 没有数量 = 判失败，让pipeline走on_error
    try:
        return float(value) != 0.0
    except (TypeError, ValueError):
        return True  #不是数字（比如读到提示文字），按"有内容"处理，不误判成0


def _same_number(left, right):
    #两个值是否"数值相等"。都解析成数字就比数值（"0.0000000000" == "0"），
    #任一边不是数字就退回字符串比较（仓库名这类）
    left_text = str(left or "").strip()
    right_text = str(right or "").strip()
    try:
        return float(left_text) == float(right_text)
    except (TypeError, ValueError):
        return left_text == right_text


@AgentServer.custom_recognition("UiaCellHasValue")  #当前行该字段有值时命中（分流用）
class UiaCellHasValue(CustomRecognition):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def analyze(
            self,
            context: Context,
            argv: CustomRecognition.AnalyzeArg,
    ) -> "CustomRecognition.AnalyzeResult":

        if not _require_uia("UiaCellHasValue"):
            return None
        _reset_cache()  #每个识别重新认窗口
        param = _params(argv, attr="custom_recognition_param", tag="UiaCellHasValue")
        if param is None:
            return None
        field = param.get("field", "")
        node, aid = _cell_node(field, "UiaCellHasValue", param.get("row", 0))
        if node is None:
            return None
        actual = _read_cell_text(node)
        print(f"[UiaCellHasValue] {field}({aid}) = {actual!r}")
        if not actual:
            return None  #空着就不命中，让pipeline走"需要填写"那条分支
        return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1), detail={"field": field, "value": actual})

@AgentServer.custom_action("UiaSetHeaderField")  #给单头字段写值（UIA，零鼠标零键盘注入，不抢前台）
class UiaSetHeaderField(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaSetHeaderField"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)
        if param is None:
            return False
        field = param.get("field", "")
        if field in _CELL_FIELDS:
            print(f"[UiaSetHeaderField] {field} 是明细列，请用 UiaSetCell，别配到这个动作")
            return False

        #label 参数：同名字段在别的单据上标签不一样（申请单的「调出组织」），允许覆盖默认标签
        if param.get("label"):
            _HEADER_FIELDS[field] = str(param["label"])
        found = _header_editor(field, "UiaSetHeaderField")
        if found is None:
            return False
        node, pattern, label_rect, edit_rect, cls = found

        value = param.get("value")
        value = str(value) if value is not None else _header_value_from_order(field, param)
        if value is None:
            print(f"[UiaSetHeaderField] 没给 value，Excel 队列里也取不到 {field}"
                  f"（字段名不在取值白名单里？）")
            return False
        if value == "":
            #和 UiaSetCell 一致：空值不算失败（备注这类允许留空），但也绝不能写空串把默认值清掉
            print(f"[UiaSetHeaderField] {field} 没有值（Excel空值），按无需填写处理")
            return True

        before = _read_header_value(node, pattern)
        #写通道（2026-11-16 probe4/7 实测）：外层控件的 ValuePattern 只读 —— SetValue 直接抛
        #COMError 0x80040200（UIA_E_INVALIDOPERATION），但同一个 pattern 读值正常（before 就是这么来的）。
        #能写的是内层 EditControl（KDPopupBaseEdit/ButtonEdit/KDDatePicker 实测都有、IsReadOnly=False，
        #probe5 解剖确认）。先内层、找不到退外层 —— 外层那条对 KDComboBoxEdit 这类没有内层 Edit 的
        #控件仍是唯一路径（虽然实测它 SetValue 会抛，抛了就走下面的失败分支，不会静默）
        target, target_pattern = _header_write_target(node)
        if target_pattern is None:
            print(f"[UiaSetHeaderField] {field}「{_HEADER_FIELDS.get(field)}」找不到可写的 ValuePattern"
                  f"（外层 {cls} @ {edit_rect}，内层也没有带 ValuePattern 的 Edit）")
            return False
        try:
            target_pattern.SetValue(value)
        except Exception as exc:
            print(f"[UiaSetHeaderField] {field}「{_HEADER_FIELDS.get(field)}」写值失败: {exc!r}"
                  f"（写入目标 {getattr(target, 'ClassName', '?')} @ {getattr(target, 'BoundingRectangle', '?')}，外层 {cls} @ {edit_rect}）")
            return False
        _fg_restore()  #SetValue 偷走的前台立刻归还，再等金蝶刷新
        time.sleep(1.2)  #给金蝶把编辑框刷新出来（外键解析要等保存，见文件头说明）

        actual = _read_header_value(node, pattern)
        print(f"[UiaSetHeaderField] {field}「{_HEADER_FIELDS.get(field)}」"
              f"原值={before!r} ← {value!r}，回读={actual!r}")
        expect = str(param.get("expect", "") or "")
        if expect:
            if expect not in actual:
                print(f"[UiaSetHeaderField] 回读校验不通过：期望含 {expect!r}，实际 {actual!r}。"
                      f"编码没被解析成基础资料的可能性最大（金蝶只把 {value!r} 当文本显示了）")
                return False
            return True
        #没写 expect 时的"没写进去"判据，只对**纯文本字段**成立（电话/地址/联系人/日期：
        #写什么回读什么，值没变就是没写进去）。
        #⚠️基础资料字段（组织/申请人）必须跳过这条 —— 2026-11-18 GUI 实测误判：
        #  新草稿的申请组织默认值=当前登录组织，Excel 也填同一个组织 →
        #  写入成功、回读当然和原值一样；而写入值是编码、回读是名称，
        #  "认不出写入值"也成立 → 两条撞上把一张**正确的单**判成失败。
        #  这类字段写前根本区分不了"我写的"和"默认值恰好相同"（结果本来也一样），
        #  真正的把关在保存之后：VerifyFK 用 UiaHeaderFilled 验编码有没有解析成名称。
        is_basedata = field in _BASEDATA_HEADER_FIELDS
        #需求日期也跳过"值没变=没写进去"判据：金蝶把 2026/11/18 重排成 2026-11-18 00:00:00，
        #"认不出写入值"对日期**必然成立**（分隔符/补零/时间尾巴全不同）；新草稿默认恰好是当天时
        #before==actual 也成立 → 两条撞上会把一张正确的单误杀（和申请组织同一类坑，只是还没轮到它报）。
        if field == "demand_date":
            if not actual:
                print(f"[UiaSetHeaderField] demand_date 写入后回读是空，判失败")
                return False
            print(f"[UiaSetHeaderField] demand_date 已写入（回读={actual!r}）；"
                  f"日期格式金蝶会重排，不做等值校验，真错单由保存后的单据本身把关")
            return True
        if not is_basedata and before == actual and value not in actual:
            print(f"[UiaSetHeaderField] 回读和原值一模一样、也认不出写入值 {value!r}（{actual!r}），判为没写进去")
            return False
        if is_basedata:
            if not actual:
                print(f"[UiaSetHeaderField] {field}「{_HEADER_FIELDS.get(field)}」写入后回读是空，判失败")
                return False
            #11-16 probe7 实测：编码→名称的外键解析发生在**保存提交时**，写后立刻回读还是编码原样。
            #所以这里不判失败、也不催 expect（expect 在这类字段上必误判），只提醒校验位置。
            print(f"[UiaSetHeaderField] {field} 已写入（回读={actual!r}）。"
                  f"组织/申请人是基础资料字段：显示编码≠没绑上外键，解析在保存时才发生，"
                  f"外键校验请放到保存节点之后用 UiaHeaderFilled（expect 填名称片段）")
        return True


@AgentServer.custom_recognition("UiaHeaderFilled")  #单头字段已有值时命中（分流/收尾校验用）
class UiaHeaderFilled(CustomRecognition):

    @_uia_inited
    def analyze(
            self,
            context: Context,
            argv: CustomRecognition.AnalyzeArg,
    ) -> "CustomRecognition.AnalyzeResult":

        if not _require_uia("UiaHeaderFilled"):
            return None
        _reset_cache()
        param = _params(argv, attr="custom_recognition_param", tag="UiaHeaderFilled")
        if param is None:
            return None
        field = param.get("field", "")
        if param.get("label"):
            _HEADER_FIELDS[field] = str(param["label"])
        found = _header_editor(field, "UiaHeaderFilled")
        if found is None:
            return None
        node, pattern, _lr, _er, _cls = found
        actual = _read_header_value(node, pattern)
        print(f"[UiaHeaderFilled] {field}「{_HEADER_FIELDS.get(field)}」= {actual!r}")
        expect = str(param.get("expect", "") or "")
        if expect:
            if expect not in actual:
                return None
            return CustomRecognition.AnalyzeResult(
                box=(0, 0, 1, 1), detail={"field": field, "value": actual})
        if not actual:
            return None  #空着不命中，让 pipeline 走"需要填写"那条分支
        return CustomRecognition.AnalyzeResult(
            box=(0, 0, 1, 1), detail={"field": field, "value": actual})


def _nav_page_ready(wait_label, tag):
    #目标编辑页是否已就位。判据不是"wait_label 文本可见"——列表页的列头/筛选区
    #同样印着「申请组织」这几个字（2026-11-17 用户实测误判：Invoke 列表回列表页后
    #仍被判成"已在目标页"）。真判据 = 该标签右侧配到了带 ValuePattern 的编辑框，
    #也就是 _header_editor 能成功配对（列表页的列头文本右侧没有编辑框）。
    for field, label in _HEADER_FIELDS.items():
        if label != wait_label:
            continue
        return _header_editor(field, tag) is not None
    #wait_label 不在单头标签表里：退化成"文本可见+主窗内"（调用方自己保证标签唯一）
    main = _main_window()
    if main is None:
        return False
    win_rect = _header_rect(main)
    for node in _walk(main, want_type="TextControl", max_depth=16):
        try:
            nm = (node.Name or "").strip()
            off = _is_offscreen(node)
        except Exception:
            continue
        if nm == wait_label and not off and _in_main_window(_header_rect(node), win_rect):
            return True
    return False


def _invoke_toolbar_item(name, tag):
    #按名字 Invoke 工具栏项（MenuItem 名→内层带 InvokePattern 的 Button，退 ButtonControl 按名）。
    #UiaInvokeButton 和 UiaNavigate 的「新增」路线共用这份实现。
    #只认非 offscreen：编辑页预置着大量隐藏的 是/否/确定/取消 按钮（§2.9-2）
    main = _main_window()
    if main is None:
        return False
    for node in _walk(main, want_type="MenuItemControl"):
        try:
            nm = (node.Name or "").strip()
            off = _is_offscreen(node)
        except Exception:
            continue
        if nm != name or off:
            continue
        for b in _walk(node, want_type="ButtonControl", max_depth=4):
            try:
                if _is_offscreen(b):
                    continue
                p = b.GetPattern(_uia.PatternId.InvokePattern)
            except Exception:
                continue
            if p is not None:
                try:
                    p.Invoke()
                    _fg_restore()  #Invoke 偷走的前台立刻归还
                    print(f"[{tag}] 已 Invoke 工具栏 {name!r}（MenuItem 内层 Button）")
                    return True
                except Exception as exc:
                    print(f"[{tag}] MenuItem {name!r} 内层 Invoke 异常: {exc!r}")
                    return False
    for b in _walk(main, want_type="ButtonControl"):
        try:
            nm = (b.Name or "").strip()
            off = _is_offscreen(b)
        except Exception:
            continue
        if nm != name or off:
            continue
        try:
            p = b.GetPattern(_uia.PatternId.InvokePattern)
        except Exception:
            p = None
        if p is not None:
            try:
                p.Invoke()
                _fg_restore()  #Invoke 偷走的前台立刻归还
                print(f"[{tag}] 已 Invoke 按钮 {name!r}")
                return True
            except Exception as exc:
                print(f"[{tag}] 按钮 {name!r} Invoke 异常: {exc!r}")
                return False
    return False


def _find_nav_target(text):
    #首页磁贴：Name **等于**单据名的 ButtonControl（实测「调拨申请单」磁贴 rect≈(340,917,610,965)
    #带 InvokePattern，Invoke 直达新增编辑页 —— 2026-11-17 实测）。
    #用等值不用包含：下拉项「DBSQD 调拨申请单」那种后台拿不到（下拉只在金蝶前台时渲染，
    #主窗树/顶层窗口双零增量，probe50），包含匹配反而会被页签标题等杂项干扰。
    main = _main_window()
    if main is None:
        return None
    win_rect = _header_rect(main)
    for node in _walk(main, want_type="ButtonControl"):
        try:
            nm = (node.Name or "").strip()
            if nm != text or _is_offscreen(node):
                continue
            if not _in_main_window(_header_rect(node), win_rect):
                continue
            p = node.GetPattern(_uia.PatternId.InvokePattern)
        except Exception:
            continue
        if p is not None:
            return p.Invoke, nm
    return None


@AgentServer.custom_action("UiaNavigate")  #直达单据编辑页（全后台三路线：磁贴→列表页新增→搜索回车兜底）
class UiaNavigate(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaNavigate"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        _invalidate_page_cache()  #本动作就是要换页面：幂等检查必须用新鲜树，不许命中上一页缓存
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        text = str(param.get("text", "") or "").strip()
        if not text:
            print("[UiaNavigate] 没给 text（要导航的单据名，如「调拨申请单」）")
            return False
        wait_label = str(param.get("wait_label", "") or "").strip()
        timeout_ms = _as_index(param.get("timeout_ms", 15000), 15000)

        if _main_window() is None:
            print("[UiaNavigate] 没找到金蝶主窗口（客户端是不是没开/停在登录页？）")
            return False

        def ready():
            return (not wait_label) or _nav_page_ready(wait_label, "UiaNavigate")

        #幂等：已经在目标编辑页就直接过（上一张单保存后可能留在编辑页）
        if ready():
            print(f"[UiaNavigate] 已在目标页（「{wait_label}」右侧编辑框配对成功），无需导航")
            return True

        budget = max(1000, timeout_ms) / 1000.0

        def wait_ready():
            deadline = time.time() + budget
            while time.time() < deadline:
                time.sleep(1.0)
                #每轮先还前台再认树：导航 Invoke 之后金蝶抱着前台等页面加载，
                #这个轮询最长 budget（pipeline 里配到 60s）——不在这儿还，
                #用户窗口就要替金蝶白占整整一段加载时间。_fg_restore 只在
                #前台确实是金蝶主窗时动手，页面加载完/本来没偷时是空操作。
                #_ensure_window_visible 同款理由：60s 窗口足够用户双击一次任务栏图标，
                #中途被最小化的话后面每轮配对都会瞎，每轮捞一次（IsIconic 查询很便宜）。
                _ensure_window_visible()
                _fg_restore()
                _reset_cache()  #页面在变，每轮重新认树
                _invalidate_page_cache()  #…以及重扫页面级缓存（否则上一轮缓存会假报"已就位"）
                if ready():
                    return True
            return False

        #路线1：首页磁贴 Invoke（probe50 前一轮实测：直达「调拨申请单 - 新增」编辑页，全后台）
        tgt = _find_nav_target(text)
        if tgt is not None:
            invoke_fn, nm = tgt
            try:
                invoke_fn()
                print(f"[UiaNavigate] 路线1：已 Invoke 首页磁贴 {nm!r}")
                if wait_ready():
                    print("[UiaNavigate] 导航到位（磁贴路线）")
                    return True
            except Exception as exc:
                print(f"[UiaNavigate] 磁贴 Invoke 异常: {exc!r}")

        #路线2：列表页工具栏「新增」（probe20 实测 Invoke 成立）。
        #不在列表页时这条找不到按钮，静默过。
        _reset_cache()
        if _invoke_toolbar_item("新增", "UiaNavigate"):
            if wait_ready():
                print("[UiaNavigate] 导航到位（列表页新增路线）")
                return True
            print("[UiaNavigate] 新增点了但页面没到位")

        #路线3：顶部搜索框 SetValue + PostMessage 回车 —— 只在金蝶恰好是前台时有效
        #（2026-11-17 用户实测定案：非前台时 WPF 不触发搜索、下拉不渲染，此路仅兜底）
        box = None
        win_rect = _header_rect(_main_window())
        for node in _walk(_main_window(), want_type="EditControl"):
            rect = _header_rect(node)
            if rect is None or _is_offscreen(node) or not _in_main_window(rect, win_rect):
                continue
            try:
                if (node.AutomationId or "") == "PART_Editor":
                    continue
            except Exception:
                continue
            if rect[1] < 80:
                box = node
                break
        if box is not None:
            try:
                box.SetFocus()
                time.sleep(0.5)
                vp = box.GetPattern(_uia.PatternId.ValuePattern)
                if vp is not None:
                    vp.SetValue(text)
                    hwnd = _main_hwnd()
                    user32 = ctypes.windll.user32
                    user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
                    user32.PostMessageW.restype = ctypes.c_bool
                    user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0100, 0x0D, ctypes.c_void_p(1))
                    user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0101, 0x0D, ctypes.c_void_p(0xC0000001))
                    print(f"[UiaNavigate] 路线3：搜索框已写 {text!r} 并发回车")
                    if wait_ready():
                        print("[UiaNavigate] 导航到位（搜索回车路线）")
                        return True
            except Exception as exc:
                print(f"[UiaNavigate] 搜索框路线异常: {exc!r}")
        else:
            print("[UiaNavigate] 顶部没有可见搜索框（登录页？窗口最小化？）")

        print(f"[UiaNavigate] 三条路线都没到位（text={text!r} wait_label={wait_label!r}）"
              f"—— 窗口是不是被最小化了？金蝶最小化时全树坐标变虚拟值，磁贴/搜索框都判不可见")
        return False  #让 pipeline 走 on_error 重导航


@AgentServer.custom_action("UiaInvokeButton")  #按名字 Invoke 工具栏按钮（保存/提交/列表/新增…），零坐标
class UiaInvokeButton(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaInvokeButton"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        name = str(param.get("name", "") or "").strip()
        if not name:
            print("[UiaInvokeButton] 没给 name（工具栏按钮的可见文字，如「保存」）")
            return False
        post_wait_ms = _as_index(param.get("post_wait_ms", 3000), 3000)

        main = _main_window()
        if main is None:
            print("[UiaInvokeButton] 没找到金蝶主窗口（客户端是不是没开？）")
            return False

        #金蝶工具栏的 保存/提交/审核 是 MenuItemControl，带 InvokePattern 的是它内层
        #Button（aid=PART_DefaultButt，**Name 是空的**）—— 直接按 ButtonControl 找名会漏
        #（11-16 probe2b 就是这么"找不到保存"的，probe3 全类型扫才发现）。
        #先 MenuItem 按名 → 内层第一个非 offscreen 的 Invoke；再退 ButtonControl 按名（列表/退出这类）
        invoke_target = None
        for node in _walk(main, want_type="MenuItemControl"):
            try:
                nm = (node.Name or "").strip()
                off = _is_offscreen(node)
            except Exception:
                continue
            if nm != name or off:
                continue
            for b in _walk(node, want_type="ButtonControl", max_depth=4):
                try:
                    if _is_offscreen(b):
                        continue
                    p = b.GetPattern(_uia.PatternId.InvokePattern)
                except Exception:
                    continue
                if p is not None:
                    invoke_target = p.Invoke
                    break
            break
        if invoke_target is None:
            for b in _walk(main, want_type="ButtonControl"):
                try:
                    nm = (b.Name or "").strip()
                    off = _is_offscreen(b)
                except Exception:
                    continue
                if nm != name or off:
                    continue  #offscreen 闸必须留着：编辑页预置着大量隐藏的是/否/确定/取消按钮
                try:
                    p = b.GetPattern(_uia.PatternId.InvokePattern)
                except Exception:
                    p = None
                if p is not None:
                    invoke_target = p.Invoke
                    break
        if invoke_target is None:
            print(f"[UiaInvokeButton] 界面上没有可见的「{name}」按钮（当前页不是单据页？还是被弹窗挡住了？）")
            return False

        try:
            invoke_target()
        except Exception as exc:
            print(f"[UiaInvokeButton] Invoke「{name}」异常: {exc!r}")
            return False
        _fg_restore()  #Invoke 偷走的前台立刻归还，界面跳转的等待期用户窗口拿回前台
        time.sleep(max(0, post_wait_ms) / 1000.0)  #给界面跳转留时间（保存→列表约2-3s）
        _invalidate_page_cache()  #保存/列表/新增 之后页面可能整个换掉，缓存作废
        print(f"[UiaInvokeButton] 已 Invoke「{name}」，等 {post_wait_ms}ms")
        return True  #Invoke 不抛即算发出；生效与否由下一个节点（识别/回读）把关


@AgentServer.custom_action("UiaPickCombo")  #组合框下拉选择（Expand+DoDefaultAction，全程后台）
class UiaPickCombo(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaPickCombo"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        field = param.get("field", "")
        found = _header_editor(field, "UiaPickCombo")
        if found is None:
            return False  #标签配对失败（不是编辑页/标签不对），_header_editor 已经报过原因
        node, pattern, label_rect, edit_rect, cls = found

        value = param.get("value")
        value = str(value) if value is not None else _header_value_from_order(field, param)
        if value is None:
            print(f"[UiaPickCombo] 没给 value，Excel 队列里也取不到 {field}（字段名不在取值白名单里？）")
            return False
        if value == "":
            #空值=保留金蝶默认选项、什么都不动（和 UiaSetCell 空值语义一致）
            print(f"[UiaPickCombo] {field} 没有值（Excel空值），保留金蝶默认选项")
            return True

        try:
            current = str(pattern.Value or "")
        except Exception:
            current = ""
        if current == value:
            #现值已经等于目标值（金蝶默认值恰好正确，或重跑上一轮已选过）→ 不展开不点击
            print(f"[UiaPickCombo] {field}「{_HEADER_FIELDS.get(field)}」现值已是 {value!r}，无需选择")
            return True

        try:
            ecp = node.GetPattern(_uia.PatternId.ExpandCollapsePattern)
        except Exception:
            ecp = None
        if ecp is None:
            print(f"[UiaPickCombo] {field}「{_HEADER_FIELDS.get(field)}」不是组合框（{cls} @ {edit_rect}），"
                  f"纯文本走 UiaSetHeaderField、基础资料走 UiaSetCell")
            return False

        #匹配必须**等值**：下拉项每个公司/账套不一样，用包含匹配会把「选项甲」错选成「选项甲临时」这类
        #相近项静默填错。找不到 → 日志列出全部可选项，让改 Excel，绝不「差不多选一个」
        def scan_options():
            #下拉项可能挂在主窗树，也可能挂在同进程的独立 Popup 顶层窗口（probe58 两处都出现过，
            #随渲染时机二选一），两个来源都扫；同名项在两处各出现一次，按名字去重、目标取首个活引用。
            tgt, opts = None, []
            roots = [_main_window()]
            try:
                pid = _main_pid()
                base = _main_hwnd()
                for w in _uia.GetRootControl().GetChildren():
                    try:
                        if w.ProcessId == pid and w.NativeWindowHandle != base:
                            roots.append(w)
                    except Exception:
                        continue
            except Exception:
                pass
            for root in roots:
                if root is None:
                    continue
                for it in _walk(root, want_type="ListItemControl", max_depth=18):
                    try:
                        nm = (it.Name or "").strip()
                        if not nm or _is_offscreen(it):
                            continue
                    except Exception:
                        continue
                    if nm in opts:
                        continue
                    opts.append(nm)
                    if nm == value and tgt is None:
                        tgt = it
            return tgt, opts

        #展开前**强制归位**：ExpandCollapse 是开关不是"打开"指令。上一轮 Enter 提交后下拉状态未知，
        #直接 Expand() 有可能正好把它关掉（2026-11-18 T5 实测：连续调用后 可选项=[]）。
        #每轮先 Collapse 再 Expand 把状态钉成"开着"，扫不到目标就再来一轮（渲染慢的兜底）。
        target, options = None, []
        for attempt in range(2):
            try:
                ecp.Collapse()
            except Exception:
                pass
            time.sleep(0.4)
            try:
                ecp.Expand()
            except Exception as exc:
                print(f"[UiaPickCombo] {field} 第 {attempt + 1} 轮 Expand 异常: {exc!r}")
                break
            time.sleep(1.2)  #等下拉项渲染进树（probe58：Expand 后 ListItem 直接挂主窗树/Popup）
            _reset_cache()
            target, options = scan_options()
            if target is not None:
                break
        if target is None:
            try:
                ecp.Collapse()
            except Exception:
                pass
            print(f"[UiaPickCombo] {field} 下拉里没有 {value!r}。可选项={options}。"
                  f"Excel 该格改成上面某个可选项再跑（每个账套选项不同，这里不猜）")
            return False

        #提交序列（2026-11-18 probe60 对照实测，S3 路线成立且最干净）：
        #  Expand() → SelectionItem.Select() 高亮 → PostMessage Enter 到主窗 = 提交。
        #三条失败路都钉过：只 Select 不高亮提交（probe59 路A）；Select+DoDefaultAction
        #对同一旧对象（T2/T4 实测回读不变）；双展开 DoDefaultAction 虽成立（S2）但要收起
        #再展开、还依赖鼠标语义。S3 全程 PostMessage，不碰鼠标、DoDefaultAction 那种
        #偷焦点也没有，Enter 走金蝶自己的"确认当前高亮项"逻辑。
        try:
            sip = target.GetPattern(_uia.PatternId.SelectionItemPattern)
            if sip is None:
                raise RuntimeError("目标项没有 SelectionItemPattern")
            sip.Select()
            time.sleep(0.4)
        except Exception as exc:
            print(f"[UiaPickCombo] {field} 高亮 {value!r} 失败: {exc!r}")
            try:
                ecp.Collapse()
            except Exception:
                pass
            return False
        hwnd = _main_hwnd()
        if not hwnd:
            print(f"[UiaPickCombo] {field} 拿不到主窗口句柄，发不了确认回车")
            return False
        user32 = ctypes.windll.user32
        scan = user32.MapVirtualKeyW(0x0D, 0)
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0100, 0x0D, ctypes.c_void_p(1 | (scan << 16)))
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0101, 0x0D, ctypes.c_void_p(1 | (scan << 16) | 0xC0000000))
        time.sleep(1.2)

        #回读必须**重新配对**再读：点击提交后金蝶可能重建控件，旧引用读到的是缓存值
        #（页面级缓存同理必须作废，否则"重新配对"拿到的还是提交前那棵树）
        _invalidate_page_cache()
        actual = ""
        again = _header_editor(field, "UiaPickCombo")
        if again is not None:
            try:
                actual = str(again[1].Value or "")
            except Exception:
                actual = ""
        print(f"[UiaPickCombo] {field}「{_HEADER_FIELDS.get(field)}」"
              f"{current!r} → 选 {value!r}，回读={actual!r}")
        if actual != value:
            print(f"[UiaPickCombo] 回读和所选不一致：金蝶没接住这次点击（或下拉被别的东西吃了），判失败")
            return False
        return True  #组合框是枚举值不是基础资料，等值回读就是最终校验，没有"保存才解析"的后账


#── UiaPickFlex：弹性域（仓位这类多维值集组合字段）的维度录入弹窗驱动 ──
#
#为什么不能像仓库那样直写：仓库是单一基础资料（编码→外键，金蝶能解析），
#仓位是**值集组合值** —— 一个格子要接收 N 个维度的组合，没有哪个字符串能唯一确定它，
#所以名称/编码/编码+名称/仓库+编码 全被吞（实测）。正路只有驱动金蝶自己的弹窗。
#
#为什么必须"点"而不能 Expand/F7：放大镜是金蝶自绘的，
#在 UIA 树里**没有任何对应元素**（格子内层只有 ButtonEdit + PART_Editor 文本块，
#Expand()/F7/Legacy DoDefaultAction/右键/批量填充 全部实测开不出弹窗）。
#但"点"是 PostMessage 到主窗口，不碰用户鼠标键盘，所以仍然是无人值守的。
#
#坐标怎么做到不怕用户改列宽/列顺序/DPI：
#  ①所有落点都从**格子自己的 UIA 矩形**按比例算，列宽变了 rect 跟着变，落点自动跟着变；
#  ②放大镜位置给一组候选（右缘内 3%~12% 宽 × 纵向 40%/50%/60% 高），
#    命中判据是"弹窗真的出现了"而不是"我点完了"，所以猜错只多试一次，绝不会写错数据；
#  ③列被横向滚出视口时先用明细网格的 ScrollPattern 滚回来，
#    滚不回来就直接失败并说明原因 —— 不拿窗外坐标瞎点（窗外坐标等于点到别的控件，事故级）。
_FLEX_MAGNIFIER_X_FRACS = (0.05, 0.03, 0.08, 0.12, 0.02, 0.00)  #距右缘多少个"列宽"
_FLEX_MAGNIFIER_Y_FRACS = (0.50, 0.40, 0.60)  #距顶边多少个"行高"
_FLEX_DIALOG_WAIT_MS = 3000  #一次候选点击后等录入窗的时间。
#实测所有成功路径里录入窗都在第 1~2 个候选位、3s 内出现。
#真正的坑不是"等太短"，而是**等不到就在候选之间关窗** ——
#迟到的合法录入窗会被当成"点歪弹出的杂物"关掉，下一个候选重开的窗带着坏掉的查询状态
#（值集列表 0 行）。所以循环里绝不关窗，见下面的候选循环注释
_FLEX_LIST_WAIT_MS = 6000  #等值集列表弹出的时间
_FLEX_DROPDOWN_WAIT_MS = 4000  #打字后等「模糊匹配下拉」的时间（实测 2.5s 内出现）
_FLEX_MAX_PAGES = 12  #值集列表最多翻几页（首页只渲染十几行，目标常在后面）


#金蝶这两类弹窗的固定标题（产品自带文案，不是账套数据）。
#为什么还要按标题认一遍、不全靠形状判据：值集列表窗的工具栏按钮是 返回/过滤/刷新/退出 + closeButton，
#**没有「取消」**，也没有 KDPopupBaseEdit —— 两个形状判据都认不出它（实测 _looks_like_selector=False、
#_flex_dialog_shape=False）。于是上一轮失败留下的列表窗永远关不掉，
#而金蝶下一次开列表会**复用同一个句柄**，那个句柄在"调用前基线"里 →
#_wait_flex_list 把它当旧窗滤掉 → 报"列表没出来"，其实是清理盲区（2026-11-19 GUI 实测整轮失败的真因）
_FLEX_DIALOG_TITLE = "维度数据录入"
_FLEX_LIST_TITLE = "值集资料选择列表"

#仓位列 → 必须先填好的仓库列。仓库为空时仓位格是禁用态（里面只有 TextBlock，没有编辑控件），
#点了也不会有任何反应 —— 历史上一批"仓位写不进/弹窗开不出"的结论全是在这个禁用态下测的。
_FLEX_PREREQ_WAREHOUSE = {
    "source_location": "source_warehouse",
    "target_location": "target_warehouse",
}


def _post_click(hwnd, x, y, double=False):
    #PostMessage 合成一次左键点击（可选双击）。全程不移动用户光标、不动用户键盘。
    #WM_MOUSEACTIVATE 的参数按实测生效的那组写（wParam=窗口句柄、lParam=0x0101）——
    #PostMessage 不取返回值，所以这组"不规范"的实参和金蝶自己的响应互不相干，
    #换成规范值没有额外收益，别改（改了就是拿未验证的写法替换已验证的写法）
    #⚠️lparam 必须按 c_void_p 传：WM_LBUTTONUP 那类消息的 lParam 高位是坐标，
    #  默认 argtypes 会按有符号 32 位截 → 坐标超过 2^31 时直接 OverflowError
    user32 = ctypes.windll.user32
    user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
    user32.PostMessageW.restype = ctypes.c_bool
    pt = wintypes.POINT(int(x), int(y))
    user32.ScreenToClient(hwnd, ctypes.byref(pt))
    lparam = (pt.y << 16) | (pt.x & 0xFFFF)
    try:
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0021, hwnd, ctypes.c_void_p(0x0101))  # WM_MOUSEACTIVATE
        time.sleep(0.07)
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0200, 0, ctypes.c_void_p(lparam))  # WM_MOUSEMOVE
        time.sleep(0.07)
        if double:
            user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0203, 1, ctypes.c_void_p(lparam))  # WM_LBUTTONDBLCLK
            time.sleep(0.06)
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0201, 1, ctypes.c_void_p(lparam))  # WM_LBUTTONDOWN
        time.sleep(0.06)
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0202, 0, ctypes.c_void_p(lparam))  # WM_LBUTTONUP
    except Exception as exc:
        print(f"[_post_click] 投递点击失败: {exc!r}")
        return False
    time.sleep(0.2)
    return True


def _post_key(hwnd, vk):
    #PostMessage 一个虚拟键（按下+抬起）。只投给金蝶自己的窗口，绝不全局注入（踩坑清单#40）
    #argtypes 同 _post_click：释放键的 lParam 带 0xC0000000 那一位，必须走 c_void_p
    user32 = ctypes.windll.user32
    user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_void_p]
    user32.PostMessageW.restype = ctypes.c_bool
    scan = user32.MapVirtualKeyW(vk, 0)
    try:
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0100, vk, ctypes.c_void_p(1 | (scan << 16)))
        time.sleep(0.05)
        user32.PostMessageW(ctypes.c_void_p(hwnd), 0x0101, vk, ctypes.c_void_p(0xC0000001 | (scan << 16)))
    except Exception as exc:
        print(f"[_post_key] 投递按键 vk={vk:#04x} 失败: {exc!r}")
        return False
    time.sleep(0.05)
    return True


def _force_foreground(hwnd):
    #把前台**强制**交给指定窗口，绕开 Windows 前台锁。
    #为什么必须有：WPF 的命中测试要求窗口处于激活态 —— 实测（2026-11-19）
    #  列已滚进视口 + 前台是别的应用 → PostMessage 点击**完全没反应**，弹窗开不出来；
    #  同一坐标 + 前台是金蝶 → 弹窗立刻出现。
    #裸 SetForegroundWindow 会被前台锁拒（返回 False、前台纹丝不动，实测踩过），
    #所以走 AttachThreadInput：把自己的输入线程临时挂到当前前台的线程上，
    #这段时间内 SetForegroundWindow 就被系统放行，摘掉附件再恢复原状。
    #⚠️调用方必须在动作结尾用 _fg_restore 把前台还给用户，否则等于长期霸屏
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    try:
        tid_self = kernel32.GetCurrentThreadId()
        current = int(user32.GetForegroundWindow() or 0)
        tid_cur = user32.GetWindowThreadProcessId(current, None) if current else 0
        for _ in range(3):
            attached = False
            if tid_cur and tid_cur != tid_self:
                attached = bool(user32.AttachThreadInput(tid_self, tid_cur, True))
            try:
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
                user32.SetActiveWindow(hwnd)
            finally:
                if attached:
                    user32.AttachThreadInput(tid_self, tid_cur, False)
            time.sleep(0.35)
            if int(user32.GetForegroundWindow() or 0) == hwnd:
                return True
        return False
    except Exception as exc:
        print(f"[_force_foreground] 异常: {exc!r}")
        return False


def _hold_foreground(hwnd, tag, doing):
    #每次要投递点击/按键之前，确认前台还在目标窗口上；被抢走就当场抢回来。
    #
    #为什么不能"动作开头抢一次就够了"（2026-11-19 GUI 实测的整轮失败根因）：
    #本动作从抢前台到点完放大镜，中间夹着十几秒横滚 + 最多 18 个候选位约 60 秒，
    #这期间企业微信 / UU远程 / 各类弹窗会把前台抢走（日志抓到的原话：
    #  候选 1 前台=196706(企业微信) 目标窗=132442 → 候选 2 前台=1773176 → 18 个全空）。
    #WPF 的命中测试要求窗口处于激活态，前台不在金蝶时 PostMessage 点击**完全没反应**，
    #于是 18 次点击全部白打，现象却是"放大镜开不出弹窗"，极易误判成坐标问题。
    #
    #抢不到就返回 False 让调用方立刻止损：继续试只会再白烧几十秒，
    #而且日志里会留下"试了 18 个候选位"这种把环境问题伪装成坐标问题的误导结论
    user32 = ctypes.windll.user32
    try:
        if int(user32.GetForegroundWindow() or 0) == int(hwnd):
            return True
    except Exception:
        return False
    if _force_foreground(hwnd):
        return True
    try:
        stolen = int(user32.GetForegroundWindow() or 0)
    except Exception:
        stolen = 0
    print(f"[{tag}] {doing} 之前前台不在目标窗口（现在={stolen}，目标={int(hwnd)}），"
          f"且抢不回来 —— 停止尝试。继续投递点击只会在非激活窗口上白打，"
          f"现象和「放大镜位置不对」一模一样，别往坐标上查")
    return False


def _flex_dialog_shape(win):
    #判断一个顶层窗口是不是「维度数据录入」这类弹性域录入窗。
    #判据用结构不用标题（标题是金蝶界面文案，会随版本/语言变）：
    #  ① 有可见的 KDPopupBaseEdit（维度值输入框）
    #  ② 有可见的「确定」和「取消」（录入窗的固定按钮对）
    #两条都要，少一条就可能是别的弹窗（比如查找列表也有确定/取消，但它没有 KDPopupBaseEdit）
    has_edit = False
    names = set()
    for node in _walk(win, want_type="EditControl", max_depth=14):
        try:
            if node.ClassName == "KDPopupBaseEdit" and not _is_offscreen(node):
                has_edit = True
                break
        except Exception:
            continue
    if not has_edit:
        return False
    for button in _walk(win, want_type="ButtonControl", max_depth=14):
        try:
            if not _is_offscreen(button):
                names.add(_safe_name(button))
        except Exception:
            continue
    return "确定" in names and "取消" in names


def _kingdee_top_controls(exclude=None):
    #枚举金蝶进程当前**可见**的顶层窗口，返回 [(句柄, UIA控件)]。
    #
    #⚠️为什么用 Win32 EnumWindows 而不是 _uia.GetRootControl().GetChildren()：
    #「维度数据录入」和它的值集列表都是带 owner 的 WPF 弹窗，实测（2026-11-19）
    #它们出现在 EnumWindows 里、用 ControlFromHandle 拿到的控件形状判据也成立
    #（pid 对、不是主窗、不在基线），但**不在 UIA 根节点的子列表中**。
    #按根子节点轮询就永远等不到，表现成"点了 18 个候选位都没弹窗"，
    #其实是检测方式看不见它 —— UiaPickList 用的 _wait_popup_rows 走的正是那条路，
    #本动作不复用它，原因同上。
    user32 = ctypes.windll.user32
    pid = _main_pid()
    if not pid:
        return []
    #主窗句柄两条路都取，任一命中就跳过：
    #_main_hwnd() 依赖 COM（本函数可能在 COM 就绪前被调），_find_main_hwnd() 是纯 Win32。
    #⚠️这里绝不能只信一个 —— _looks_like_selector 对主窗本身返回 True
    #（主窗有 DataItem 行 + 隐藏预置的确定/取消按钮，实测成立），
    #主窗一旦被当成"残留弹窗"，清理动作就会在单据页上点「取消」= 关掉用户正在填的单
    protected = {h for h in (_main_hwnd(), _find_main_hwnd()) if h}
    exclude = set(exclude or ())
    found = []
    EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

    def callback(hwnd, _lparam):
        try:
            proc_id = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
            if proc_id.value != pid or int(hwnd) in protected or int(hwnd) in exclude:
                return True
            if not user32.IsWindowVisible(hwnd):
                return True
            rect = wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            if rect.right - rect.left < 100 or rect.bottom - rect.top < 60:
                return True  #气泡/工具条一类的小窗，不可能是我们要的弹窗
            found.append(int(hwnd))
        except Exception:
            return True
        return True

    try:
        user32.EnumWindows(EnumProc(callback), 0)
    except Exception:
        return []
    out = []
    for handle in found:
        try:
            out.append((handle, _uia.ControlFromHandle(handle)))
        except Exception:
            continue
    return out


def _kingdee_top_handles():
    #本轮"已存在"的顶层句柄基线。
    #⚠️不能用 _top_handles()：它走 UIA 根子节点，而录入窗/值集列表恰恰不在那里（见
    #_kingdee_top_controls 的注释）—— 那样基线里永远没有这些弹窗，
    #上一轮残留的录入窗就会被当成"本轮新弹出来的"，对着它里头的旧维度值提交 = 静默写错仓位
    return {handle for handle, _win in _kingdee_top_controls()}


def _wait_flex_dialog(exclude, timeout_ms=_FLEX_DIALOG_WAIT_MS):
    #等一个"新的、属于金蝶的、形状像弹性域录入窗"的顶层窗口。
    #exclude 是调用前的顶层句柄基线 —— 上一轮残留的录入窗必须排除掉，
    #否则会把旧窗当成本次的，然后对着旧窗里别的维度值提交，静默写错仓位
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        for _handle, win in _kingdee_top_controls(exclude):
            try:
                if _flex_dialog_shape(win):
                    return win
            except Exception:
                continue
        time.sleep(_POLL_MS / 1000.0)
    return None


def _wait_flex_list(exclude, timeout_ms=_FLEX_LIST_WAIT_MS):
    #等值集列表弹出，返回 (窗口控件, 带文本的 DataItem 行)。
    #枚举走 _kingdee_top_controls（这些弹窗不在 UIA 根子节点里，见那里的注释）。
    #
    #⚠️必须把"窗口出现了"和"窗口里有行了"分开返回：
    #金蝶查询偶发慢一拍 —— 窗先到、行后到，甚至这一轮根本不给行。
    #旧写法只在有行时才返回窗口，"有窗无行"也回 (None, [])，
    #调用方就分不清该重试还是该放弃（重试分支变成死代码，实测整轮白等 6s×3）
    #返回 (None, []) = 窗口始终没出现；(窗口, []) = 窗口在但没数据
    deadline = time.time() + timeout_ms / 1000.0
    seen = None
    while time.time() < deadline:
        for _handle, win in _kingdee_top_controls(exclude):
            try:
                if not _looks_like_flex_list(win):
                    continue
                seen = win
                rows = [c for c in _walk(win, want_type="DataItemControl") if _safe_name(c)]
                if rows:
                    return win, rows
            except Exception:
                continue
        time.sleep(_POLL_MS / 1000.0)
    return seen, []


def _looks_like_flex_list(win):
    #形状判据：有 DataItem 行，或有列表专属的翻页按钮（bNext/bPrev 之一）。
    #加第二条是因为"窗到了、行还没到"的窗口里 DataItem 数为 0，
    #只按行判就认不出它，等于回到旧写法的盲区
    try:
        if [c for c in _walk(win, want_type="DataItemControl", max_depth=12) if _safe_name(c)]:
            return True
    except Exception:
        return False
    for button in _walk(win, want_type="ButtonControl", max_depth=12):
        try:
            if (getattr(button, "AutomationId", "") or "") in ("bNext", "bPrev"):
                return True
        except Exception:
            continue
    return False


def _dump_flex_windows(tag):
    #把金蝶当前所有可见顶层窗口打出来。只在失败路径调用。
    #为什么要：这一串环节的失败原因高度同形（"弹窗没出来/列表没出来"），
    #光看这句日志下一步只能瞎猜。窗口清单能一眼区分三种情况：
    #  什么都没有 → 点击没生效（坐标/前台问题）
    #  有录入窗没有列表窗 → Invoke 没开列表
    #  两个都有 → 是检测判据没认出来，不是没弹
    try:
        handles = _kingdee_top_controls()
    except Exception as exc:
        print(f"[{tag}] 窗口清单读取异常: {exc!r}")
        return
    if not handles:
        print(f"[{tag}] 金蝶当前没有任何非主窗顶层窗口")
        return
    for handle, win in handles:
        try:
            title = (win.Name or "")[:24]
            flex = _flex_dialog_shape(win)
            #max_depth 必须和 _wait_flex_list 用的同一个值：
            #这里写小了会数到 0 行，把"判据没认出来"误报成"列表是空的"，自诊断反而骗下一个人
            rows = len([c for c in _walk(win, want_type="DataItemControl") if _safe_name(c)])
            print(f"[{tag}]   hwnd={handle} title={title!r} 录入窗形状={flex} 数据行={rows}")
        except Exception:
            print(f"[{tag}]   hwnd={handle} 读取异常")


def _flex_edit_of(dialog):
    #取录入窗里的维度值输入框（KDPopupBaseEdit）
    for node in _walk(dialog, want_type="EditControl", max_depth=16):
        try:
            if node.ClassName == "KDPopupBaseEdit" and not _is_offscreen(node):
                return node
        except Exception:
            continue
    return None


def _flex_magnifier_of(edit):
    #取输入框自己的放大镜按钮（aid=PART_Item，带 Invoke）。
    #⚠️这是**弹窗里**的放大镜 —— 它和明细格子那个不一样，这个是正经 WPF Button，
    #UIA 里能拿到、Invoke 有效（实测能开出值集列表）。明细格子里那个是自绘的，拿不到。
    for node in _walk(edit, want_type="ButtonControl", max_depth=8):
        try:
            if (getattr(node, "AutomationId", "") or "") != "PART_Item":
                continue
            if node.GetPattern(_uia.PatternId.InvokePattern) is not None:
                return node
        except Exception:
            continue
    return None


def _flex_inner_textbox(edit):
    #找 KDPopupBaseEdit 内层真正的 TextBox。
    #⚠️WM_CHAR 要打进"有键盘焦点的输入控件"才会走金蝶的文本变更事件 —— 实测 SetValue 赋值
    #不触发模糊下拉，只有逐字符输入触发；而逐字符前必须先把焦点放进这个内层 TextBox
    for node in _walk(edit, want_type="EditControl", max_depth=8):
        try:
            if node.ClassName == "TextBox":
                return node
        except Exception:
            continue
    return None


def _wait_flex_dropdown(exclude, keyword, timeout_ms=_FLEX_DROPDOWN_WAIT_MS):
    #等「全模糊匹配」下拉：打字后新出现的可见顶层窗（实测无标题、带 DataItem 行）。
    #返回 (下拉窗, 命中 keyword 的行)。窗没出现 → (None, [])；出现了但没命中 → (窗, [])
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        for _handle, win in _kingdee_top_controls(exclude):
            try:
                rows = [c for c in _walk(win, want_type="DataItemControl", max_depth=16)
                        if _safe_name(c)]
                if not rows:
                    continue
                matched = [c for c in rows if any(keyword in t for t in _row_texts(c))]
                return win, matched
            except Exception:
                continue
        time.sleep(_POLL_MS / 1000.0)
    return None, []


def _flex_dropdown_pick(dialog, dialog_handle, edit, keyword, tag):
    #「模糊匹配下拉」支路（用户人工路径的脚本化，2026-11-19 实测打通）：
    #  清空输入框 → 焦点进内层 TextBox → 逐字 WM_CHAR 输入 keyword →
    #  金蝶自动弹「全模糊匹配」下拉（它按相似度只给最相关的几行，天然比全量列表准）→
    #  双击唯一命中行 → 输入框回显编码 = 选中成功，交回调用方点「确定」。
    #为什么优先走这条：PART_Item 的「值集资料选择列表」是另一条查询，实测会返回空
    #（账套/会话状态相关，见 UiaPickFlex 失败记录）；而下拉支路不依赖它。
    #返回 "ok"=已选中 / "fallback"=下拉没出现（交回旧支路）/ "abort"=出现但歧义或无命中（弹窗已取消）
    try:
        edit.GetPattern(_uia.PatternId.ValuePattern).SetValue("")
    except Exception:
        return "fallback"
    time.sleep(0.4)
    inner = _flex_inner_textbox(edit)
    if inner is not None:
        try:
            inner.SetFocus()
            time.sleep(0.4)
        except Exception:
            pass
    base = _kingdee_top_handles()
    if not _type_text(keyword, key_delay_ms=150, hwnd=dialog_handle):
        print(f"[{tag}] 往录入窗投字符失败，退回值集列表支路")
        return "fallback"
    win, matched = _wait_flex_dropdown(base, keyword)
    if win is None:
        print(f"[{tag}] 打字后没出模糊下拉，退回值集列表支路")
        return "fallback"
    if not matched:
        print(f"[{tag}] 模糊下拉出现了但没有含 {keyword!r} 的行 —— 值集里没有这个值，"
              f"Excel 该列改成下拉里真实出现的名称，或留空（仓位非必填）")
        _invoke_named(dialog, "取消")
        return "abort"
    if len(matched) > 1:
        print(f"[{tag}] 模糊下拉有 {len(matched)} 行都含 {keyword!r}，歧义不选。"
              f"Excel 该列写得更具体一点（用全称）")
        _invoke_named(dialog, "取消")
        return "abort"
    rect = _header_rect(matched[0])
    if not rect:
        return "fallback"
    try:
        dd_handle = int(win.NativeWindowHandle) or dialog_handle
    except Exception:
        dd_handle = dialog_handle
    _post_click(dd_handle, (rect[0] + rect[2]) / 2.0, (rect[1] + rect[3]) / 2.0, double=True)
    time.sleep(1.5)
    try:
        echo = str(edit.GetPattern(_uia.PatternId.ValuePattern).Value or "")
    except Exception:
        echo = ""
    #判据：输入框从 keyword 变成编码（金蝶选中行后回显的是代码不是名称）。
    #echo 恰好等于 keyword 时（用户本来就填的编码）无法用"变了没"判定 → 交回旧支路重选，
    #方向是安全的：最多多走一条支路，不会静默选错行
    if echo and echo != keyword:
        print(f"[{tag}]   模糊下拉命中，输入框回显编码={echo!r}")
        return "ok"
    print(f"[{tag}]   双击后输入框={echo!r}，金蝶没接住这次选中，退回值集列表支路")
    return "fallback"


def _invoke_named(root, name):
    #按 Name 点一个可见按钮，返回是否点到
    for button in _walk(root, want_type="ButtonControl", max_depth=20):
        try:
            if _safe_name(button) != name or _is_offscreen(button):
                continue
            invoke = button.GetPattern(_uia.PatternId.InvokePattern)
            if invoke is not None:
                invoke.Invoke()
                time.sleep(0.5)
                return True
        except Exception:
            continue
    return False


def _flex_page_next(root):
    #值集列表底部翻页按钮 aid=bNext。翻到下一页，返回是否翻页成功。
    #为什么必须翻页：金蝶列表**只渲染当前页**（实测首页 13 行），
    #目标仓位排在后面时"没命中"不等于"不存在" —— 不翻页就报"找不到"会误杀正常单据
    for button in _walk(root, want_type="ButtonControl", max_depth=20):
        try:
            if (getattr(button, "AutomationId", "") or "") != "bNext":
                continue
            invoke = button.GetPattern(_uia.PatternId.InvokePattern)
            if invoke is None:
                continue
            invoke.Invoke()
            time.sleep(1.2)
            return True
        except Exception:
            continue
    return False


def _grid_scroll_pattern(tag, node=None):
    #取明细网格的横向 ScrollPattern（实测只有可见的那个 DataGridControl HorizontallyScrollable=True）。
    #优先从格子往上爬祖先链找 DataGridControl —— 格子就在网格里，爬十几层、十几个 COM 调用；
    #全树遍历要几千个调用、实测 15~25s，能免则免。爬到顶没找到才回退全树遍历
    if node is not None:
        cur = node
        for _ in range(30):
            try:
                cur = cur.GetParentControl()
            except Exception:
                break
            if cur is None:
                break
            try:
                if cur.ControlTypeName == "DataGridControl":
                    scroll = cur.GetPattern(_uia.PatternId.ScrollPattern)
                    if scroll is not None and scroll.HorizontallyScrollable:
                        return scroll
            except Exception:
                continue
    win = _main_window()
    if win is None:
        return None
    for grid in _walk(win, want_type="DataGridControl", max_depth=30):
        try:
            if _is_offscreen(grid):
                continue  #陈旧页签里的网格，滚它没用
            scroll = grid.GetPattern(_uia.PatternId.ScrollPattern)
            if scroll is not None and scroll.HorizontallyScrollable:
                return scroll
        except Exception:
            continue
    return None


def _grid_hscroll_to_cell(field, tag, row, start_node=None):
    #把目标格子滚进视口，返回 (格子元素, 新鲜矩形)；弄不进来返回 (None, None)。
    #
    #为什么必须滚：明细列多，用户改过列宽/列顺序、窗口非最大化、或分辨率变小时，
    #仓位列会整列滚在窗口右缘之外（实测 rect x 到 2200 而窗右缘只有 1288）。
    #这时格子在树里"存在且 offscreen=False"，屏幕上却点不到 ——
    #拿它的坐标去点等于点到别的列上（改错数据），比失败更危险，所以弄不进来就直接失败。
    #
    #为什么滚动后要重读矩形：WPF 虚拟化在中间滚动位置会给过期的布局坐标，
    #只有滚动事件之后的读取才和屏幕渲染一致（实测踩过：拿旧 rect 点 → 18 个候选位全空）。
    #列滚进视口时旧格子元素会被回收，对它读矩形返回 None → 那时才全树重找一次。
    #
    #策略（2026-11-19 重写；旧版"归零+24步盲滚、每步两次全树遍历"实测烧 2~4 分钟/列，
    #GUI 表现就是"卡了好久没日志"，用户手动拖横滚条才解卡）：
    #  ①优先复用调用方已经找到的格子（start_node），省一次遍历
    #  ②已可见 → 微滚 +2% 强制重排 + 重找一次拿新鲜矩形
    #  ③不可见 → 滚 8% 校准"每 1% 多少像素"（实测线性）→ 一步跳到目标位 →
    #    最多 3 次微调 → 交付前保证点击坐标来自滚动后的新鲜元素
    #  ④滚不动/滚到底仍看不见 → 立刻失败并说清"窗口太窄，最大化后重跑"
    #⚠️SetScrollPercent 的第二个参数必须直接写 -1.0：这个版本的 uiautomation
    #  没有 BasePattern.NoScroll 常量，用它会抛 AttributeError（实测踩过）
    def _visible(rect):
        if rect is None:
            return False
        win = _main_window()
        wr = _header_rect(win) if win is not None else None
        if not wr:
            return False
        margin = 40  #右缘留白：放大镜在格子内部右侧，贴边点击容易飘到相邻列
        return wr[0] < rect[0] and rect[2] < wr[2] - margin

    def _refind():
        #一次全树遍历（实测 15~25s，本函数唯一的贵操作）：只在廉价重读失败时才用
        fresh, _aid = _cell_node(field, tag, row, want_lookup=None)
        if fresh is None:
            return None, None
        return fresh, _header_rect(fresh)

    def _cheap_rect():
        #廉价重读：滚动事件已经强制 WPF 重新布局，同一个节点再读一次矩形就是新鲜值，
        #一次 COM 调用（对比 _refind 的几千次）。只有节点被虚拟化回收（读不到矩形）
        #才值得付一次全树遍历的代价
        nonlocal node
        try:
            r = _header_rect(node)
        except Exception:
            r = None
        if r is not None:
            return node, r
        node, r = _refind()
        return node, r

    node = start_node
    rect = _header_rect(node) if node is not None else None
    if node is None or rect is None:
        node, rect = _refind()
    if node is None or rect is None:
        return None, None

    scroll = _grid_scroll_pattern(tag, node)  #从格子往上爬，别再全树遍历
    if scroll is None:
        print(f"[{tag}] {field} 找不到可横向滚动的明细网格（rect={rect}）："
              f"当前界面不是单据编辑页？")
        return None, None

    def _percent():
        try:
            p = scroll.HorizontalScrollPercent
            return p if p is not None and p >= 0 else 0.0
        except Exception:
            return 0.0

    if _visible(rect):
        #微滚 +2% 强制重排；再全树重找一次 —— 点击坐标必须来自滚动后的新鲜元素
        try:
            scroll.SetScrollPercent(min(100.0, _percent() + 2.0), -1.0)
            time.sleep(0.5)
        except Exception:
            pass
        node, r2 = _cheap_rect()
        if r2 is not None and _visible(r2):
            return node, r2
        return node, rect  #重读失败：退回原矩形，调用方按现值处理

    #校准：滚 8% 量出线性比例（列左缘 x 随百分比线性左移，实测 ≈25px/1%）
    p0 = _percent()
    x0 = rect[0]
    try:
        scroll.SetScrollPercent(min(100.0, p0 + 8.0), -1.0)
    except Exception:
        return None, None
    time.sleep(0.6)
    node, r1 = _cheap_rect()
    if node is None or r1 is None:
        return None, None
    moved = x0 - r1[0]
    rect = r1
    if moved <= 10:
        node, r2 = _cheap_rect()
        if r2 is not None and _visible(r2):
            return node, r2
        win = _main_window()
        wr = _header_rect(win) if win is not None else None
        print(f"[{tag}] {field} 第 {row} 行滚不动（rect={rect} 窗右缘="
              f"{wr[2] if wr else '?'}）：窗口太窄装不下这列，"
              f"请把金蝶最大化/提高分辨率后重跑")
        return None, None
    px_per_pct = moved / 8.0

    win = _main_window()
    wr = _header_rect(win)
    cell_w = rect[2] - rect[0]
    x_target = max(wr[0] + 80, wr[2] - 40 - cell_w)  #目标：整列进视口且右缘留白

    for _ in range(3):
        gap = rect[0] - x_target
        if _visible(rect) and gap <= 0:
            return node, rect
        pct = _percent() + max(2.0, gap / px_per_pct)
        try:
            scroll.SetScrollPercent(min(100.0, pct), -1.0)
        except Exception:
            break
        time.sleep(0.6)
        node, r2 = _cheap_rect()
        if node is None or r2 is None:
            return None, None
        rect = r2

    if _visible(rect):
        node, r2 = _cheap_rect()  #交付前保证：点击坐标来自滚动后的新鲜矩形
        if r2 is not None and _visible(r2):
            return node, r2
        return node, rect
    win = _main_window()
    wr = _header_rect(win) if win is not None else None
    print(f"[{tag}] {field} 第 {row} 行滚到底仍在视口外（rect={rect} 窗右缘="
          f"{wr[2] if wr else '?'}）：窗口太窄装不下这列，请把金蝶最大化后重跑")
    return None, None

@AgentServer.custom_action("UiaPickFlex")  #弹性域字段：驱动金蝶「维度数据录入」弹窗选值
class UiaPickFlex(CustomAction):

    @_uia_inited  #本线程要自己初始化 COM，否则整棵 UIA 树枚举不了
    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        if not _require_uia("UiaPickFlex"):
            return False
        _reset_cache()  #每个动作重新认窗口，不跨动作复用旧句柄
        param = _params(argv)  #拉取参数
        if param is None:
            return False
        field = param.get("field", "")
        if field not in _LOOKUP_FIELDS:
            print(f"[UiaPickFlex] {field!r} 不在当前单据的查找类字段里"
                  f"（可选: {sorted(_LOOKUP_FIELDS)}）。是不是没先调 UiaSetForm？")
            return False
        row = _as_index(param.get("row", 0))
        _t0 = time.time()  #从取参数就计时：旧写法在滚动之后才起表，"全程17.2s"漏掉了前面 2 分钟的滚动

        keyword = param.get("keyword") or param.get("value")
        keyword = str(keyword).strip() if keyword is not None else ""
        if not keyword:
            getter = _COLUMN_ATTR_BY_FIELD.get(field)
            keyword = ""
            if getter is not None:
                try:
                    import transfer_actions
                    order = transfer_actions._current_order()
                    if order is not None and row < len(order.items):
                        keyword = str(getter(order.items[row]) or "").strip()
                except Exception as exc:
                    print(f"[UiaPickFlex] 从队列取 {field} 异常: {exc!r}")
        if not keyword:
            print(f"[UiaPickFlex] {field} 第 {row} 行没有值（Excel 空），按无需填写处理")
            return True  #空值=这行本来就不填仓位，和 UiaSetCell 空值语义一致

        expect = str(param.get("expect", "") or "")

        #── 前置：同行的仓库列必须已填 ──
        #仓库空 → 仓位格是禁用态（没有编辑控件），点它不会有任何反应。
        #这条必须**先报清楚**，否则现象是"弹窗开不出来"，下一个人会去查弹窗，其实卡在这
        prereq = _FLEX_PREREQ_WAREHOUSE.get(field)
        if prereq:
            wh_node, _wh_aid = _cell_node(prereq, "UiaPickFlex", row, want_lookup=None)
            wh_value = _read_cell_text(wh_node) if wh_node is not None else ""
            if not wh_value:
                print(f"[UiaPickFlex] {field} 的前置没满足：同行 {prereq} 是空的。"
                      f"仓库为空时仓位格处于禁用态（金蝶不给它编辑控件），弹窗开不出来 —— "
                      f"请先用 UiaSetCell 填仓库（编码形态），再填仓位")
                return False

        node, aid = _cell_node(field, "UiaPickFlex", row, want_lookup=None)
        if node is None:
            return False
        current = _read_cell_text(node)
        if current and keyword in current:
            print(f"[UiaPickFlex] {field} 第 {row} 行已经是 {current!r}，跳过填写")
            return True  #重试幂等：上一轮其实填成功了，别重复开弹窗再选一遍

        hwnd = _main_hwnd()
        if not hwnd:
            print(f"[UiaPickFlex] 拿不到金蝶主窗口句柄")
            return False
        #⚠️抢前台必须在滚动之前：滚动要十几秒，抢不到前台的话那十几秒全是白等
        #（实测出现过"抢不到前台"却已经烧掉 189s 的日志，就是顺序反了）
        if not _hold_foreground(hwnd, "UiaPickFlex", "开始定位仓位列"):
            return False
        #start_node=node：把前面 _cell_node 已经找到的格子传进去，滚动函数不再重复全树遍历
        node, rect = _grid_hscroll_to_cell(field, "UiaPickFlex", row, start_node=node)
        if node is None or rect is None:
            return False
        left, top, right, bottom = rect
        width = max(1, right - left)
        height = max(1, bottom - top)
        print(f"[UiaPickFlex] {field}({aid}) 第 {row} 行 rect={rect} 现值={current!r}"
              f"（定位+滚动 {time.time() - _t0:.1f}s）")

        #── ①触发录入弹窗：先点格子中心进编辑态，再点放大镜候选位 ──
        #⚠️中途绝不 _fg_restore：从"抢前台"到"格子填完"整段都要金蝶拿着前台，
        #  归还统一交给动作结尾装饰器那次兜底。
        #先清掉上一轮残留的弹窗，再取基线：残留不排掉的话，_wait_flex_dialog 会把它当"本轮新窗"，
        #对着旧窗里别的维度值提交 = 静默写错仓位（基线判据只有在"旧窗已被关掉"时才成立）
        _close_flex_windows()
        baseline = _kingdee_top_handles()
        center_x = left + width / 2.0
        if not _hold_foreground(hwnd, "UiaPickFlex", "进编辑态"):
            return False
        _post_click(hwnd, center_x, top + height / 2.0)  #进编辑态（放大镜只在编辑态渲染）
        time.sleep(0.8)
        dialog = None
        tried = 0
        lost_foreground = False
        for x_frac in _FLEX_MAGNIFIER_X_FRACS:
            for y_frac in _FLEX_MAGNIFIER_Y_FRACS:
                px = right - x_frac * width
                py = top + y_frac * height
                tried += 1
                if not _hold_foreground(hwnd, "UiaPickFlex", f"第 {tried} 个候选位"):
                    lost_foreground = True
                    break
                if tried > 1:
                    #上一轮没弹窗时，金蝶可能已经把格子退出了编辑态（被抢前台、
                    #焦点被别的控件带走都会），这时重抢前台后直接点放大镜还是打空，
                    #现象和「坐标不对」一样。所以每轮先重新点一次格子中心进编辑态，再点放大镜。
                    _post_click(hwnd, left + width / 2.0, top + height / 2.0)
                    time.sleep(0.8)
                clicked = _post_click(hwnd, px, py)
                dialog = _wait_flex_dialog(baseline)
                if dialog is not None:
                    print(f"[UiaPickFlex] 弹窗已出现（第 {tried} 个候选位 ({px:.0f},{py:.0f}），"
                          f"距开始 {time.time() - _t0:.1f}s）")
                    break
                if tried <= 2 or not clicked:
                    #前两轮打明细：这一串环节的失败全靠"坐标对不对 + 屏幕上有没有东西"定位
                    print(f"[UiaPickFlex]   候选 {tried} ({px:.1f},{py:.1f}) 投递={clicked} "
                          f"前台={int(ctypes.windll.user32.GetForegroundWindow() or 0)} 目标窗={hwnd}")
                    _dump_flex_windows("UiaPickFlex")
                #这里**不**关窗（实测踩过的坑）：等不到经常只是窗还在路上，
                #关掉它反而制造坏状态；迟到的窗由下一个候选的等待收下来照样能用
            if dialog is not None or lost_foreground:
                break  #前台丢了整轮停：再投十几次点击只会在非激活窗口上白打
        if dialog is None and lost_foreground:
            return False  #_hold_foreground 已经把原因说清楚了，别再报"坐标不对"误导
        if dialog is None:
            print(f"[UiaPickFlex] {field} 试了 {tried} 个放大镜候选位都没弹出录入窗"
                  f"（全程前台都在金蝶，所以这次可以排除前台因素）。"
                  f"可能原因：①金蝶这个版本把放大镜画在别处 ②该列不是弹性域 "
                  f"③格子矩形与实际渲染不符")
            _dump_flex_windows("UiaPickFlex")
            _close_flex_windows()  #确认全败之后才清，别在候选之间清
            return False

        #── ②填维度关键词，开值集列表 ──
        edit = _flex_edit_of(dialog)
        if edit is None:
            print(f"[UiaPickFlex] 录入窗里找不到维度输入框（KDPopupBaseEdit）")
            _invoke_named(dialog, "取消")
            return False
        dim_label = _flex_dimension_label(dialog)

        #── ②支路一：「模糊匹配下拉」（打字触发；实测最快，且不依赖值集列表那条查询）──
        #2026-11-19 实测：SetValue 不触发下拉，只有逐字 WM_CHAR 触发；双击唯一命中行后
        #输入框回显编码，直接跳到 ⑤ 点「确定」。旧支路（PART_Item→值集列表）保留为兜底。
        dropdown_ok = False
        try:
            dialog_handle = int(dialog.NativeWindowHandle)
        except Exception:
            dialog_handle = 0
        if dialog_handle:
            outcome = _flex_dropdown_pick(dialog, dialog_handle, edit, keyword, "UiaPickFlex")
            if outcome == "abort":
                return False
            dropdown_ok = outcome == "ok"

        if not dropdown_ok:
            try:
                edit.GetPattern(_uia.PatternId.ValuePattern).SetValue(keyword)
            except Exception as exc:
                print(f"[UiaPickFlex] 往维度输入框写 {keyword!r} 失败: {exc!r}")
                _invoke_named(dialog, "取消")
                return False
            time.sleep(1.0)  #给金蝶按关键词预筛列表留时间（0.8s 实测偶发来不及）

            magnifier = _flex_magnifier_of(edit)
            if magnifier is None:
                print(f"[UiaPickFlex] 维度输入框里没有 PART_Item 放大镜，开不出值集列表")
                _invoke_named(dialog, "取消")
                return False
            list_baseline = _kingdee_top_handles()
            try:
                magnifier.GetPattern(_uia.PatternId.InvokePattern).Invoke()
            except Exception as exc:
                print(f"[UiaPickFlex] Invoke 弹窗放大镜失败: {exc!r}")
                _invoke_named(dialog, "取消")
                return False
            popup, rows = _wait_flex_list(list_baseline)
            print(f"[UiaPickFlex]   列表窗={popup is not None} 数据行={len(rows)} "
                  f"距开始 {time.time() - _t0:.1f}s")
            #这里不做"关掉空列表窗再重开"的重试：实测关掉之后同一个放大镜再也开不出列表窗，
            #重试逻辑自己把现场毁掉，把一次本可以靠 pipeline on_error 救回的偶发失败变成必然失败。
            #偶发的"窗到了行没到"交给上层重跑整节点（动作开头会彻底清场，重跑是干净的）
            if popup is None:
                print(f"[UiaPickFlex] {field} 的值集列表窗根本没出现（维度={dim_label}）")
                _dump_flex_windows("UiaPickFlex")
                _invoke_named(dialog, "取消")
                return False
            if not rows:
                print(f"[UiaPickFlex] {field} 的值集列表窗出现了但没有一行数据（维度={dim_label}）"
                      f"—— 两种可能：①该账套在这个维度下确实没有可选值 ②金蝶这次查询没返回（偶发）。"
                      f"分不清就先让 pipeline 重试本节点；仍失败就把 Excel 该列留空（仓位非必填）")
                _dump_flex_windows("UiaPickFlex")
                _invoke_named(dialog, "取消")
                return False
            #这里不归还前台：下一步要往列表窗口 PostMessage Enter，同样吃"激活态"这道门槛

            #── ③在列表里找唯一命中行（含翻页）──
            hit = None
            hit_texts = []
            page = 0
            while page < _FLEX_MAX_PAGES:
                candidates = []
                for row_node in rows:
                    texts = _row_texts(row_node)
                    if any(keyword in t for t in texts):
                        candidates.append((row_node, texts))
                if len(candidates) == 1:
                    hit, hit_texts = candidates[0]
                    break
                if len(candidates) > 1:
                    #多行都含关键词 = 有歧义，宁可不填也不能猜。
                    #（"仓位A" 和 "仓位A临时" 同时存在时，选错就是发错货）
                    print(f"[UiaPickFlex] {field} 列表第 {page + 1} 页有 {len(candidates)} 行都含 {keyword!r}，"
                          f"歧义不选。各行文本={[t[:2] for t, _ in candidates]}")
                    _close_flex_windows()
                    _invoke_named(dialog, "取消")
                    return False
                if not _flex_page_next(popup):
                    break  #没有下一页按钮，或已经翻到最后一页
                page += 1
                time.sleep(0.6)
                rows = [c for c in _walk(popup, want_type="DataItemControl") if _safe_name(c)]
            if hit is None:
                print(f"[UiaPickFlex] {field} 值集列表翻了 {page + 1} 页都没有含 {keyword!r} 的行"
                      f"（维度={dim_label}）。Excel 该列改成列表里真实存在的名称，或留空（仓位非必填）")
                _close_flex_windows()
                _invoke_named(dialog, "取消")
                return False

            #── ④选中并提交 ──
            #必须用 SelectionItemPattern.Select() 真改选中项，不能用 SetFocus 顶替（焦点≠选中）
            try:
                selection = hit.GetPattern(_uia.PatternId.SelectionItemPattern)
                if selection is None or not selection.Select():
                    raise RuntimeError("该行没有 SelectionItemPattern 或 Select 返回失败")
            except Exception as exc:
                print(f"[UiaPickFlex] {field} 选中目标行失败: {exc!r}")
                _close_flex_windows()
                _invoke_named(dialog, "取消")
                return False
            time.sleep(0.5)

            list_handle = None
            try:
                list_handle = popup.NativeWindowHandle
            except Exception:
                pass
            committed = _flex_confirm_selection(list_handle, popup, edit)
            if not committed:
                print(f"[UiaPickFlex] {field} 选完行没能提交（列表关不掉/输入框没回显编码）")
                _close_flex_windows()
                _invoke_named(dialog, "取消")
                return False

        #── ⑤点录入窗「确定」，回读格子 ──
        #「确定」点不到不立刻判死：实测（2026-11-19 p171/p172）模糊下拉双击后弹窗可能
        #自行提交并关闭 —— 最终成败以格子回读为准，回读非空就是成功
        if not _invoke_named(dialog, "确定"):
            print(f"[UiaPickFlex] 录入窗里没有可见的「确定」（可能已随下拉选中自行提交），"
                  f"改用格子回读判定")
        time.sleep(2.0)
        _fg_restore()

        again, _again_aid = _cell_node(field, "UiaPickFlex", row, want_lookup=None)
        actual = _read_cell_text(again) if again is not None else ""
        print(f"[UiaPickFlex] {field} 第 {row} 行 {current!r} → 选 {keyword!r}，回读={actual!r}"
              f"（全程 {time.time() - _t0:.1f}s）")
        if not actual:
            print(f"[UiaPickFlex] {field} 回读是空：金蝶没接住这次选择（弹窗流程中途被别的窗吃掉了？）")
            return False
        if expect and expect not in actual:
            print(f"[UiaPickFlex] 回读校验不通过：期望含 {expect!r}，实际 {actual!r}")
            return False
        return True  #格子回读到值 = 弹性域已绑定，保存时不会再被"必录"拦下


def _flex_dimension_label(dialog):
    #读录入窗左侧的维度名（值集里那个维度的显示文本），只为日志好看，读不到就返回空。
    #判据必须是"在输入框左边"——录入窗里同时有「确定」「取消」等文本，
    #不加这条会把按钮名当维度名打出来
    edit = _flex_edit_of(dialog)
    try:
        rect = _header_rect(edit) if edit is not None else None
    except Exception:
        rect = None
    if not rect:
        return ""
    best = ""
    best_x = rect[0]
    for node in _walk(dialog, want_type="TextControl", max_depth=16):
        try:
            name = _safe_name(node)
            if not name or len(name) > 12 or _is_offscreen(node):
                continue
            own = _header_rect(node)
            if not own or own[0] >= rect[0]:
                continue
            if own[1] < rect[1] - 40 or own[1] > rect[3] + 40:  #不在输入框那一行附近
                continue
            if own[0] < best_x:
                best_x, best = own[0], name
        except Exception:
            continue
    return best


def _flex_echo_code(edit):
    #读维度输入框当前回显的值。选中一行后金蝶把它换成**编码**（不是名称），
    #这个回显是"列表选择真的生效了"的判据 —— 光看列表窗口关了没并不可靠
    try:
        return str(edit.GetPattern(_uia.PatternId.ValuePattern).Value or "").strip()
    except Exception:
        return ""


def _flex_confirm_selection(list_handle, popup, edit):
    #把"已选中的列表行"提交回录入窗。两条路按优先级试，判据都是输入框回显变成编码：
    #  ①PostMessage Enter 到列表窗口（实测最干净，全程零鼠标）
    #  ②PostMessage 双击目标行（有些金蝶版本列表只认双击）
    #⚠️绝不用 keybd_event/SendInput —— 那是全局注入，前台是用户程序时等于往用户窗口里敲键
    before = _flex_echo_code(edit)
    if list_handle:
        #⚠️Enter 投的是列表窗，所以确认的前台目标也是列表窗自己，不是主窗
        if not _hold_foreground(list_handle, "UiaPickFlex", "往列表窗投 Enter"):
            return False
        _post_key(list_handle, 0x0D)  # VK_RETURN
        time.sleep(2.0)
        if _flex_list_closed(popup) and _flex_echo_code(edit) != before:
            return True
        if _flex_echo_code(edit) not in ("", before):
            return True
    #兜底：双击列表里那一行
    try:
        rows = [c for c in _walk(popup, want_type="DataItemControl") if _safe_name(c)]
    except Exception:
        rows = []
    for row_node in rows:
        rect = _header_rect(row_node)
        if not rect:
            continue
        if list_handle:
            if not _hold_foreground(list_handle, "UiaPickFlex", "往列表窗投双击"):
                return False
            _post_click(list_handle, (rect[0] + rect[2]) / 2.0, (rect[1] + rect[3]) / 2.0, double=True)
            time.sleep(2.0)
            if _flex_echo_code(edit) not in ("", before):
                return True
    return False


def _flex_list_closed(popup):
    #列表窗口是否已经关掉（句柄失效或不再可见）
    try:
        handle = popup.NativeWindowHandle
    except Exception:
        return True
    if not handle:
        return True
    try:
        user32 = ctypes.windll.user32
        return not (bool(user32.IsWindow(handle)) and bool(user32.IsWindowVisible(handle)))
    except Exception:
        return True


def _invoke_close_button(win):
    #点窗口自己的 closeButton（标题栏那个 X）。
    #⚠️只在调用方已经确认"这不是主窗"之后用 —— _kingdee_top_controls 已按句柄排除主窗，
    #两层保护都别拆：主窗的 closeButton 会把整个客户端关掉
    for button in _walk(win, want_type="ButtonControl", max_depth=14):
        try:
            if "close" not in (getattr(button, "AutomationId", "") or "").lower():
                continue
            invoke = button.GetPattern(_uia.PatternId.InvokePattern)
            if invoke is not None:
                invoke.Invoke()
                time.sleep(0.8)
                return True
        except Exception:
            continue
    return False


def _close_flex_windows():
    #关掉金蝶在场的「维度数据录入」窗和它的值集列表，只点「取消」。
    #
    #为什么 _close_stale_popups 不够用：它的形状判据 _looks_like_selector 要求"有 DataItem 行
    #+ 成对确定/取消"，而录入窗是 KDPopupBaseEdit + 确定/取消、**没有 DataItem 行** → 认不出来。
    #后果是灾难级的：录入窗是模态的，它挂在那里，主窗口收不到任何点击，
    #于是本动作后面每个候选位都白试（实测一轮烧掉 68s 全失败），
    #而且每次重试都还是被同一个旧窗挡着 —— 跨重试永久死锁。
    #
    #为什么只点「取消」不点「确定」：残留窗属于上一轮，窗里那个维度值不是这一轮的答案，
    #点确定等于把上一轮的仓位静默写进这一张单
    closed = 0
    for _handle, win in _kingdee_top_controls():
        title = ""
        try:
            title = win.Name or ""
        except Exception:
            pass
        if _flex_dialog_shape(win) or _FLEX_DIALOG_TITLE in title:
            #录入窗：只点「取消」，点「确定」会把窗里上一轮的旧维度值提交进单据
            if _invoke_named(win, "取消"):
                closed += 1
        elif _FLEX_LIST_TITLE in title or _looks_like_selector(win):
            #列表窗没有「取消」，走它自己的「退出」；再兜底标题栏 closeButton
            #（能走到这里说明 _kingdee_top_controls 已经排除了主窗，点不到用户的单据页）
            if _invoke_named(win, "退出") or _invoke_close_button(win):
                closed += 1
    if closed:
        print(f"[_close_flex_windows] 关掉了 {closed} 个残留的维度录入窗/值集列表")
    return closed

