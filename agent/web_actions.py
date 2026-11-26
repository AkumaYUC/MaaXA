#网页后台字段填写自定义动作（CDP 驱动已登录的金蝶云 HTML5 网页端，不抢前台、不动用户键鼠）
#pipeline侧接口约定（节点action都填Custom，custom_action填下面的动作名；param 与 uia_actions 同名同义，Excel 契约不变）：
#  前置：pretask start_web.ps1 拉起带 --remote-debugging-port 的专用 Edge（独立 profile，登录一次长期复用），
#        页面停在 K3Cloud html5 主控台；本模块只连这个已打开的页面，绝不 navigate（裸导航会销毁金蝶会话）
#  依赖：websocket-client（requirements.txt）；Edge 没开 / 页面不在 K3Cloud / 登录态失效 → 动作 False 走 on_error，
#        日志会写明是哪种（排查用 WebPing）
#  WebSetForm         param: {"form": "requisition"}  切换单据白名单；v1 只支持 requisition（直接调拨单为前台基线）
#  WebNavigate        param: {"text": "调拨申请单列表", "wait_label": "申请组织", "timeout_ms": 60000}
#                     主控台点 text → 等列表 → 点「新增」→ 等编辑页就位（wait_label 对应单头字段出现）
#  WebSetCell         param: {"field": "material_code", "row": 0, "value": "05-002-03064"}
#                     明细格子写值：双击进编辑态 → 清空 → 键入 → F7 列 Enter / 普通列 Tab
#                     F7 列值含中文时改走逐字符键入 + Enter：insertText 不触发 typeahead，
#                     逐字符键入才会弹模糊下拉，唯一命中时 Enter 直接选中；多命中时金蝶会留空或选错，
#                     由校验报错，不假成功
#                     写入必须走真实输入链：model.set 只改显示层、不进 billModel，保存时被判空；
#                     只有真实键入触发的金蝶 UpdateValue 才进模型，F7 列 Enter 后由金蝶解析编码带出名称
#                     写完读显示层校验：物料列看名称列被带出、其他 F7 看非空非 select、普通列看含写入值
#                     空值=本行无需填这列，记日志返回True（与 UiaSetCell 一致）
#                     仓位弹性域（source_location/target_location）走「维度数据录入」弹窗链路：
#                       双击格子 → td 内维度图标 → 维度弹窗。填值两条路：
#                       路A（主路）：弹窗内 F7 输入框逐字符键入（先清空，防重跑追加）→ 金蝶「全模糊匹配」
#                         下拉 → 点命中行 → 弹窗「确定」。编码和名称片段都命中，能查到值集列表首页之外的编码
#                       路B（兜底）：F7 放大镜 → 值集列表窗 → dataSource 匹配 → 「返回」→ 确定（只覆盖首页15项）
#                       两条路都查不到就清场返回 False 并列出样例，绝不选第一行；格子最终显示值集项名称
#  WebFillColumn      param: {"field": "quantity"} / {"field": "source_org", "broadcast": true}
#                     整列逐行走真实输入链；broadcast=整单一个值铺每行（申请单调出/调入组织是明细列，每行都填）
#                     界面行数不足 Excel 行数时点「新增行」补齐（真实输入链不像 model.set 能凭空扩行）
#  WebSetHeaderField  param: {"field": "contact", "value": "张三"}
#                     文本/日期 → CDP 真实输入链（focus→insertText→Tab）；纯 JS 写 value 只改显示层，保存校验照样判空
#                     基础资料（apply_org/applicant）→ 放大镜 → 金蝶自己的选择弹窗 → 快捷过滤/可见行按名称匹配 → 「返回数据」
#                       匹配不到就关弹窗返回 False，绝不选第一行
#  WebPickCombo       param: {"field": "transfer_type", "value": "组织内调拨"}
#                     下拉：遍历候选 value 逐个试选、读显示文本与目标等值匹配后提交再回读
#                     账套选项不同也通用（零硬编码业务词）；找不到等值项还原原值并 False，日志列出候选数
#  WebBlockPaste      param: {"field": "quantity"} / {"field": "source_warehouse", "broadcast": true} / {"field": "material_code", "values": ["a","b"]}
#                     整列块粘贴（几百物料的正解）：右键首格 → 点「块粘贴」→ 弹窗 textarea 灌整列换行文本 →
#                     点「确定」→ 校验整列非空
#                     提交必须点「确定」：textarea 里 Enter 只是换行，按回车整列不落地
#                     粘贴前不补行：金蝶块粘贴自动扩行；新草稿 0 行时兜底点一次「新增行」给粘贴一个落点
#                     物料列校验读名称列被金蝶带出（编码解析成功的标志）；弹性域（仓位）不支持块粘贴
#  WebPatchFill       param: {"field": "source_warehouse"}  整单同值列一次铺开（批量填充）：
#                     源格填值 → 单击该格选中 → 点明细工具栏「批量填充」→ 一次铺满整列
#                     前提：物料列必须已有值，金蝶对空行拒绝填充；按钮只认真实鼠标事件，JS .click() 无效
#                     源格填值失败自动重试一次；铺开段（单击选中→点填充→校验）整体重试 2 轮，校验 8s 轮询，
#                     current 列名漂移则重单击
#  WebReadCell        param: {"field": "source_warehouse", "row": 0, "expect": "关键词"}
#                     回读明细格子（编码+名称拼接后判断），判据三选一优先级同 UiaReadCell：not_expect > nonzero > expect
#  WebRowCount        (custom_recognition) param: {"field": "material_code"}
#                     界面已填物料的行数 == Excel 队列行数才命中；inverse:true 反向（配自环重试）
#  WebPing            param: {}  探测 CDP/登录态/当前页形态，排查连接问题用（正常流程不调）
#  WebSubmit          param: {}  顶栏「提交」真实鼠标点击 + 确认弹窗点确定。落地硬判据=
#                     「单据编号」由占位「保存时自动生成」变真编号（或页签标题脱离「新增」）
#                     真单进 ERP 工作流：由任务开关「是否提交并下推单据」接入链路（Yes 才走），默认 No=只填不提交
#  WebAudit           param: {}  顶栏「审核」，同 WebSubmit 机制；前置=已提交（按钮灰则拒点）
#                     落地硬判据=「下推」按钮由灰变亮（不赌状态文案）
#  WebPushDown        param: {}  顶栏「下推」；前置=已审核（灰则拒点）。落地判据=弹窗关+新目标单据页签；
#                     随后自动切到新页签对目标单（如 分步式调出单）点一次「提交」——整单才算完成

import json
import os
import re
import subprocess
import threading
import time
import urllib.parse
import urllib.request

from maa.agent.agent_server import AgentServer
from maa.context import Context
from maa.custom_action import CustomAction
from maa.custom_recognition import CustomRecognition

import uia_actions  #复用它的 _COLUMN_ATTR_BY_FIELD：两套后端共用同一张 Excel 取值契约，改一处两边都生效
from common import _params

#websocket-client 是可选依赖：缺失时只有 Web* 动作报错，不影响 agent 其它模块启动（同 uiautomation 的处理）
try:
    import websocket
    _WS_AVAILABLE = True
except Exception:
    websocket = None
    _WS_AVAILABLE = False

_CDP_PORT = int(os.environ.get("MAAXA_CDP_PORT", "9333"))  #start_web.ps1 的约定端口，环境变量可覆盖
_CJK_RE = re.compile(r"[\u4e00-\u9fff]")  #含中文 = 走名称模糊匹配（typeahead 只认逐字符键入）
_WEB_URL = "https://erp.xair.cn/K3Cloud/html5/"  #与 start_web.ps1 的默认 -Url 一致；页签被关时自愈重开用
_JS_TIMEOUT = 20.0
_POLL_S = 0.5

#明细列：Excel字段 → 网页 grid 列标识 + 是否基础资料（True=F7 写 __value.number；"flex"=弹性域 v1 不支持）
#列标识与 UIA 的 AutomationId、WebAPI 字段名同源（见 uia_actions 文件头），不受列顺序/横向滚动影响
_ENTRY_FIELDS = {
    "material_code": ("FMATERIALID", True),
    "source_org": ("FSTOCKORGID", True),
    "target_org": ("FSTOCKORGINID", True),
    "source_warehouse": ("FSTOCKID", True),
    "target_warehouse": ("FSTOCKINID", True),
    "quantity": ("FQTY", False),
    "note": ("FNOTE", False),
    "source_location": ("FSTOCKLOCID", "flex"),
    "target_location": ("FSTOCKLOCINID", "flex"),
}

#单头文本/日期：Excel字段 → (网页key, 类型)；写入走 CDP 真实输入链
_HEADER_TEXT_FIELDS = {
    "contact": ("FCONTACTS", "text"),
    "phone": ("FCONTACTNUMBER", "text"),
    "address": ("FMAILINGADDRESS", "text"),
    "remark": ("FREMARKS", "textarea"),
    "demand_date": ("F_JF_XQDATE", "date"),
}

#单头基础资料：Excel字段 → 网页key；写入走放大镜→选择弹窗→返回数据
_HEADER_F7_FIELDS = {
    "apply_org": "FAPPORGID",
    "applicant": "FAPPLICANT",
}

#单头下拉：Excel字段 → 网页key；写入走"试值读文本"等值匹配
_HEADER_COMBO_FIELDS = {
    "transfer_type": "FTRANSTYPE",
    "apply_type": "F_JF_REQUIRETYPE",
}

#wait_label（中文标签）→ 单头网页key：WebNavigate 用它判定编辑页就位
_LABEL_TO_HEADER = {
    "申请组织": "FAPPORGID", "申请人": "FAPPLICANT", "联系人": "FCONTACTS",
    "联系电话": "FCONTACTNUMBER", "通讯地址": "FMAILINGADDRESS", "需求日期": "F_JF_XQDATE",
    "调拨类型": "FTRANSTYPE", "调拨申请类型": "F_JF_REQUIRETYPE", "备注": "FREMARKS",
}

_FORMS = ("requisition",)  #v1 只做调拨申请单；直接调拨单维持前台基线
_ACTIVE_FORM = None

#====================== CDP 管道 ======================

def _targets():
    #返回 (targets列表, None) 或 (None, 错误说明)
    try:
        return json.load(urllib.request.urlopen(
            f"http://127.0.0.1:{_CDP_PORT}/json", timeout=5)), None
    except Exception as e:
        return None, f"CDP 端口 {_CDP_PORT} 无响应: {str(e)[:60]}"


def _ws_from_targets():
    ts, err = _targets()
    if ts is None:
        return None, err
    for t in ts:
        if t.get("type") == "page" and "/k3cloud/" in (t.get("url") or "").lower():
            return t.get("webSocketDebuggerUrl"), None
    return None, "CDP 在，但没有 K3Cloud 页签（页面被关了/导航走了）"


_EDGE_LAUNCH_LOCK = threading.Lock()
_EDGE_LAUNCHED = False


def _ensure_edge():
    #拉起专用调试 Edge（独立 profile、带 CDP 端口）；普通 Edge 无调试端口，CDP 连不上
    global _EDGE_LAUNCHED
    with _EDGE_LAUNCH_LOCK:
        if _EDGE_LAUNCHED:
            return
        here = os.path.dirname(os.path.abspath(__file__))
        script = os.path.join(here, "start_web.ps1")
        if not os.path.exists(script):
            print(f"[web] 找不到 {script}，无法自动拉起 Edge")
            return
        print(f"[web] CDP {_CDP_PORT} 未响应，自动执行 start_web.ps1 拉起 Edge…")
        try:
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-File", script, "-Port", str(_CDP_PORT), "-WaitSeconds", "40"],
                timeout=70, capture_output=True, text=True, encoding="utf-8", errors="replace")
            _EDGE_LAUNCHED = True  #脚本幂等：端口已活/拉起成功都算跑过，不反复重试
        except Exception as e:
            print(f"[web] start_web.ps1 执行失败: {e}")


def _open_k3cloud_tab():
    #Edge 活着但 K3Cloud 页签被关掉/导航走 → 走 CDP HTTP 接口（PUT /json/new）重开一个页签，免重启 GUI
    try:
        req = urllib.request.Request(
            f"http://127.0.0.1:{_CDP_PORT}/json/new?{urllib.parse.quote(_WEB_URL, safe=':/')}", method="PUT")
        urllib.request.urlopen(req, timeout=5)
        return True
    except Exception:
        return False


def _page_ws_url():
    #从 CDP 目标列表里找 K3Cloud 那个页签；找不到就自愈：
    #  ① 端口没活 → 拉起专用 Edge（start_web.ps1，幂等）
    #  ② 端口活但没 K3Cloud 页签 → PUT /json/new 开一个，等它挂上
    #返回 (ws地址, None) 或 (None, 错误说明)
    url, err = _ws_from_targets()
    if url:
        return url, None
    _ensure_edge()
    url, err2 = _ws_from_targets()
    if url:
        return url, None
    if err2 and "没有 K3Cloud 页签" in err2 and _open_k3cloud_tab():
        deadline = time.time() + 20.0
        while time.time() < deadline:
            time.sleep(1.5)
            url, err2 = _ws_from_targets()
            if url:
                print("[web] K3Cloud 页签缺失，已自动新开页签接回")
                return url, None
    return None, (err2 or err or "未知错误")


class _Cdp:
    #一次动作一条连接：动作串行、建连毫秒级，换来"永不复用脏会话"
    def __init__(self, ws_url):
        self.ws = websocket.create_connection(ws_url, timeout=_JS_TIMEOUT,
                                              max_size=None, enable_multithread=True)
        self._id = 0
        self._cmd("Runtime.enable")
        #窗口被遮挡时浏览器会节流 rAF/定时器，金蝶渲染随之变慢；声明页面恒为聚焦态以解除节流（只改页面可见性，不抢真实焦点）
        self._cmd("Emulation.setFocusEmulationEnabled", {"enabled": True})

    def _cmd(self, method, params=None):
        self._id += 1
        try:
            self.ws.send(json.dumps({"id": self._id, "method": method, "params": params or {}}))
        except Exception as e:
            return {"err": "CDP 发送失败(连接断开): " + str(e)[:40]}
        return self._wait(self._id)

    def _wait(self, mid):
        deadline = time.time() + _JS_TIMEOUT
        while time.time() < deadline:
            try:
                msg = json.loads(self.ws.recv())
            except Exception:
                return {"err": "CDP 应答超时/断连"}
            if msg.get("id") == mid:
                if msg.get("error"):
                    return {"err": "CDP: " + str(msg["error"].get("message"))[:80]}
                return {"raw": msg.get("result", {})}
        return {"err": "CDP 应答超时"}

    def js(self, expression):
        self._id += 1
        mid = self._id
        try:
            self.ws.send(json.dumps({"id": mid, "method": "Runtime.evaluate", "params": {
                "expression": expression, "awaitPromise": True, "returnByValue": True}}))
        except Exception as e:
            return {"err": "CDP 发送失败(连接断开): " + str(e)[:40]}
        deadline = time.time() + _JS_TIMEOUT
        while time.time() < deadline:
            try:
                msg = json.loads(self.ws.recv())
            except Exception:
                return {"err": "CDP 应答超时/断连"}
            if msg.get("id") == mid:
                if msg.get("exceptionDetails"):
                    return {"err": "JS异常: " + str(msg["exceptionDetails"].get("text"))[:80]}
                return {"value": msg.get("result", {}).get("result", {}).get("value")}
        return {"err": "CDP 应答超时"}

    def click(self, x, y):
        for typ in ("mousePressed", "mouseReleased"):
            self._cmd("Input.dispatchMouseEvent",
                      {"type": typ, "x": x, "y": y, "button": "left", "clickCount": 1})

    def type_text(self, text):
        self._cmd("Input.insertText", {"text": str(text)})

    def key_tab(self):
        for typ in ("keyDown", "keyUp"):
            self._cmd("Input.dispatchKeyEvent", {"type": typ, "key": "Tab", "code": "Tab",
                                                 "windowsVirtualKeyCode": 9, "nativeVirtualKeyCode": 9})

    def key_enter(self):
        for typ in ("keyDown", "keyUp"):
            self._cmd("Input.dispatchKeyEvent", {"type": typ, "key": "Enter", "code": "Enter",
                                                 "windowsVirtualKeyCode": 13, "nativeVirtualKeyCode": 13})

    # ---- 真实输入链原语（明细格子/单元格提交必须走金蝶自己的键盘监听，
    #      只有真实键入触发的 UpdateValue 才是正道；insertText 不算真实键入，但填编辑器值可用） ----

    def dbl_click(self, x, y):
        for typ, cc in (("mousePressed", 1), ("mouseReleased", 1),
                        ("mousePressed", 2), ("mouseReleased", 2)):
            self._cmd("Input.dispatchMouseEvent", {"type": typ, "x": x, "y": y,
                                                   "button": "left", "clickCount": cc})

    def triple_click(self, x, y):
        #三击全选：金蝶 F7 输入框里 Ctrl+A 不生效（格子有值时键入会追加而非替换），
        #只有三击建立选区 + Delete 才能清空重填
        for typ, cc in (("mousePressed", 1), ("mouseReleased", 1),
                        ("mousePressed", 2), ("mouseReleased", 2),
                        ("mousePressed", 3), ("mouseReleased", 3)):
            self._cmd("Input.dispatchMouseEvent", {"type": typ, "x": x, "y": y,
                                                   "button": "left", "clickCount": cc})

    def select_all(self):
        for typ in ("rawKeyDown", "keyUp"):
            self._cmd("Input.dispatchKeyEvent", {"type": typ, "key": "a", "code": "KeyA",
                                                 "modifiers": 2,
                                                 "windowsVirtualKeyCode": 65, "nativeVirtualKeyCode": 65})

    def key_delete(self):
        for typ in ("rawKeyDown", "keyUp"):
            self._cmd("Input.dispatchKeyEvent", {"type": typ, "key": "Delete", "code": "Delete",
                                                 "windowsVirtualKeyCode": 46, "nativeVirtualKeyCode": 46})

    def key(self, name, code, vk, text=None):
        #单键（Enter/Tab 等）：rawKeyDown(+char)+keyUp，浏览器监听器收得到
        self._cmd("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": name, "code": code,
                                             "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk})
        if text:
            self._cmd("Input.dispatchKeyEvent", {"type": "char", "key": name, "text": text})
        self._cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": name, "code": code,
                                             "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk})

    def key_enter2(self):
        self.key("Enter", "Enter", 13)

    def key_tab2(self):
        self.key("Tab", "Tab", 9)

    def type_chars(self, text, gap=0.35):
        #逐字符真实按键（rawKeyDown+char+keyUp）：金蝶 F7 的全模糊匹配下拉只认这种输入，
        #insertText 只设 DOM value、不触发 typeahead
        for ch in str(text):
            self._cmd("Input.dispatchKeyEvent", {"type": "rawKeyDown", "key": ch, "code": "Key",
                                                 "windowsVirtualKeyCode": 0, "nativeVirtualKeyCode": 0})
            self._cmd("Input.dispatchKeyEvent", {"type": "char", "key": ch, "text": ch})
            self._cmd("Input.dispatchKeyEvent", {"type": "keyUp", "key": ch, "code": "Key",
                                                 "windowsVirtualKeyCode": 0, "nativeVirtualKeyCode": 0})
            time.sleep(gap)

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass


def _open(tag):
    if not _WS_AVAILABLE:
        print(f"[{tag}] websocket-client 没装（agent/requirements.txt 里的 websocket-client），网页端动作不可用")
        return None
    url, err = _page_ws_url()
    if url is None:
        print(f"[{tag}] {err}")
        return None
    try:
        return _Cdp(url)
    except Exception as e:
        print(f"[{tag}] CDP 连不上: {e}")
        return None

#====================== JS 模板 ======================

_JS_FIND_DRAFT = r"""
(() => {
  if (/index\.aspx/.test(location.pathname)) return {err: 'login-page'};
  const seen = new Set();
  for (const el of document.querySelectorAll('[id]')) {
    const m = el.id.match(/^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})-FENTITY-c$/i);
    if (!m) continue;
    const r = el.getBoundingClientRect();
    if (r.width > 50 && r.height > 50) seen.add(m[1]);
  }
  const pids = [...seen];
  if (!pids.length) return {err: 'no-draft'};
  if (pids.length > 1) return {err: 'multi-draft', pids};
  return {pid: pids[0]};
})()
"""

_JS_CLICK_RECT = r"""
(() => {
  const want = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 10 && r.height > 8; };
  const els = [...document.querySelectorAll('a,div,span,button,td,li')].filter(
    e => (e.textContent || '').trim() === want && vis(e));
  if (!els.length) return {err: 'not-found'};
  els.sort((a, b) => (a.textContent || '').length - (b.textContent || '').length);
  const el = els[0];
  el.scrollIntoView({block: 'center'});
  const r = el.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_TOOLBAR_NEW = r"""
(() => {
  //编辑页顶栏（y 60~130）的「新增」。⚠️不能用 _JS_CLICK_RECT：页面里"新增"文字有 4 处
  //（顶栏按钮、明细工具栏「新增行」、页内标签），等值匹配会选错。
  //限定顶栏区域 + 取最内层元素，实测命中 a.kdmenuitem（2026-11-23 开新单实验）
  const els = [...document.querySelectorAll('.kd-menu a, a.kdmenuitem, .kdmenuitem-content, a.k-link')].filter(e => {
    const tx = (e.textContent || '').trim();
    if (tx !== '新增') return false;
    const r = e.getBoundingClientRect();
    return r.width > 10 && r.height > 8 && r.top > 60 && r.top < 130;
  });
  if (!els.length) return {err: 'not-found'};
  els.sort((a, b) => (a.textContent || '').length - (b.textContent || '').length);
  const r = els[0].getBoundingClientRect();
  return {x: Math.round(r.x + 12), y: Math.round(r.y + r.height / 2)};  //左侧主体，避开下拉箭头
})()
"""

_JS_ACTIVE_TAB = r"""
(() => {
  //当前活动页签标题：提交后金蝶把「调拨申请单 - 新增」换成带单据编号/「- 修改」的标题
  //（提交落地信号之一，也是 WebNewBill 判「当前单已进 ERP 不必再放弃」的依据）
  //⚠️限定 top<100 顶栏：弹窗内部 tabstrip 也带 k-state-active，不过滤会读错（2026-11-25 dump 实证）
  const li = [...document.querySelectorAll('.k-tabstrip-items .k-item')]
    .find(e => /k-state-active/.test(e.className || '')
             && e.getBoundingClientRect().width > 10 && e.getBoundingClientRect().top < 100);
  return li ? (li.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 30) : '';
})()
"""

_JS_LABEL_INPUT = r"""
(() => {
  //按中文标签找单头输入框，读 value + placeholder（「单据编号」占位=保存时自动生成，
  //变成真实编号即提交落库的铁证）。找不到就返回 err，调用方不赌文案
  const label = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 10 && r.height > 8; };
  const lb = [...document.querySelectorAll('label,span,div,a')].filter(e =>
    vis(e) && (e.textContent || '').trim() === label)
    .sort((a, b) => (a.textContent || '').length - (b.textContent || '').length)[0];
  if (!lb) return {err: 'no-label'};
  let host = lb.parentElement;
  for (let up = 0; up < 5 && host; up++) {
    const inp = host.querySelector('input,textarea');
    if (inp) return {value: inp.value || '', placeholder: inp.placeholder || ''};
    host = host.parentElement;
  }
  return {err: 'no-input'};
})()
"""

_JS_TAB_BY_TEXT = r"""
(() => {
  //顶栏页签里【标题精确等于 want】的那个坐标（下推后切回来源单页签用，want 传完整标题）。
  //⚠️精确等值不能前缀/includes：「调拨申请单」会命中「调拨申请单列表」和空白「- 新增」，
  //  切错页签再点新增就开错单（2026-11-25 实测踩全两种）。限定 top<100 排除弹窗内部 tab
  const want = %s;
  const el = [...document.querySelectorAll('.k-tabstrip-items .k-item')].find(e => {
    const r = e.getBoundingClientRect();
    return r.width > 10 && r.top < 100 && (e.innerText || '').replace(/\\s+/g, ' ').trim() === want;
  });
  if (!el) return {err: 'not-found', want};
  const r = el.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), text: want};
})()
"""

_JS_TAB_TITLES = r"""
(() => [...document.querySelectorAll('.k-tabstrip-items .k-item')]
  //⚠️必须限定顶栏区域（top<100）：下推「选择单据」弹窗内部也有 k-tabstrip（参数选择），
  //  不过滤会把弹窗内部 tab 误当新页签（2026-11-25 实测假阳性）
  .filter(e => { const r = e.getBoundingClientRect(); return r.width > 10 && r.top < 100; })
  .map(e => (e.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 24)))()
"""

_JS_PUSH_OK_BTN = r"""
(() => {
  //下推「选择单据」弹窗（分步式调出单/直接调拨单 单选 + 转换规则/单据类型，异步加载）。
  //只认含「选择单据」的窗；返回其【确定】坐标 + 弹窗文本。窗没弹/已关 → noPop
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 100 && r.height > 80; };
  const w = [...document.querySelectorAll('.k-window,.kd-dynamicform-active,[class*=dialog]')]
    .filter(vis).find(e => /选择单据/.test(e.innerText || ''));
  if (!w) return {noPop: true};
  const b = [...w.querySelectorAll('a,button,span,div')].find(e => {
    const r = e.getBoundingClientRect();
    return (e.textContent || '').trim() === '确定' && r.width > 10 && r.height > 6
      && !/disabled/i.test((e.className || '') + (e.parentElement && e.parentElement.className || ''));
  });
  if (!b) return {noOk: true, text: (w.innerText || '').replace(/\\s+/g, ' ').slice(0, 60)};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          text: (w.innerText || '').replace(/\\s+/g, ' ').slice(0, 60)};
})()
"""

_JS_CONFIRM_POPUP_OK = r"""
(() => {
  //最上层可见弹窗（确认窗/子表单窗）里的【确定】按钮坐标 + 窗文本（校准日志用）。
  //⚠️只认弹窗内按钮，不碰 kd-msg 横幅（横幅按钮由调用方按类型分别处理）
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window,.kd-dynamicform-active')].filter(vis);
  if (!ws.length) return {none: true};
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  const b = [...w.querySelectorAll('a,button,span,div')].find(e =>
    (e.textContent || '').trim() === '确定' && e.getBoundingClientRect().width > 8);
  if (!b) return {noOk: true, text: (w.innerText || '').replace(/\\s+/g, ' ').slice(0, 60)};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          text: (w.innerText || '').replace(/\\s+/g, ' ').slice(0, 60)};
})()
"""

_JS_TOOLBAR_BTN = r"""
(() => {
  //顶栏（y 60~130）文字恰为 label 的按钮 + 灰态检测（下推/审核未满足前置时是灰的）。
  //与 _JS_TOOLBAR_NEW 同构；坐标取左侧主体避开下拉箭头（金蝶按钮只认真实鼠标事件）。
  const label = %s;
  const els = [...document.querySelectorAll('.kd-menu a, a.kdmenuitem, .kdmenuitem-content, a.k-link')].filter(e => {
    const tx = (e.textContent || '').trim();
    if (tx !== label) return false;
    const r = e.getBoundingClientRect();
    return r.width > 10 && r.height > 8 && r.top > 60 && r.top < 130;
  });
  if (!els.length) return {err: 'not-found'};
  els.sort((a, b) => (a.textContent || '').length - (b.textContent || '').length);
  const el = els[0];
  let grey = false;
  for (let n = el; n && n !== document.body; n = n.parentElement) {
    const c = (n.className || '').toString();
    if (/disabled/i.test(c) || /state-disabled/.test(c)) { grey = true; break; }
    if (n.getAttribute && n.getAttribute('aria-disabled') === 'true') { grey = true; break; }
    if (parseFloat(getComputedStyle(n).opacity) < 0.5) { grey = true; break; }
  }
  const r = el.getBoundingClientRect();
  return {x: Math.round(r.x + 12), y: Math.round(r.y + r.height / 2), grey};
})()
"""

_JS_DISCARD_NO = r"""
(() => {
  //顶部黄色横幅「内容已经修改，是否保存？」里的【否】。
  //⚠️它不是 .k-window（清场/弹窗检测都扫不到），容器特征 kd-msg-btn-container（2026-11-23 实测）
  const btns = [...document.querySelectorAll('a,button,span,div')].filter(e => {
    const tx = (e.textContent || '').trim();
    if (tx !== '否') return false;
    const r = e.getBoundingClientRect();
    return r.width > 8 && r.height > 6 && r.top > 0 && r.top < 400;
  });
  if (!btns.length) return {none: true};
  btns.sort((a, b) => (a.textContent || '').length - (b.textContent || '').length);
  const r = btns[0].getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_EDIT_ACTIVE = r"""
(() => {
  //目标格子是否真进了编辑态：activeElement 必须是【本格子内】的输入控件。
  //⚠️必须限定到目标 td：上一列批量填充后网格异步收尾（setBusy），双击被吞时
  //  activeElement 还停在上一列的旧编辑器里，只查"grid 内有输入框"会假阳性
  //  （2026-11-24 复现实测：target_org 键进 source_org 的旧编辑器，校验 0/1）
  const pid = %s, field = %s, row = %s;
  const $ = window.jQuery;
  const ae = document.activeElement;
  if (!ae) return {ok: false};
  if (!(["INPUT", "TEXTAREA"].includes(ae.tagName) || ae.isContentEditable)) return {ok: false};
  const el = document.getElementById(pid + '-FENTITY-c');
  if (!el) return {ok: false, err: 'grid-gone'};
  const g = $(el).data().kendoKDGrid;
  if (!g) return {ok: false, err: 'no-grid'};
  const cols = g.columns.map(c => c.field || c.dataField);
  const idx = cols.indexOf(field);
  if (idx < 0) return {ok: false, err: 'no-col'};
  const data = g.dataSource.data();
  if (row >= data.length) return {ok: false, err: 'row-out'};
  let tr = g.tbody ? g.tbody.find('tr[data-uid="' + data[row].uid + '"]') : $();
  if (!tr.length) tr = g.tbody ? g.tbody.find('tr').eq(row) : $();
  const td = tr.find('td').eq(idx);
  return {ok: !!(td.length && td[0].contains(ae))};
})()
"""

_JS_CELL_DISABLED = r"""
(() => {
  //目标格是否被金蝶联动锁定（celldiv 带 k-gridCellstate-disabled）。
  //实测（2026-11-24）：组织内调拨时调入组织=调出组织的镜像，格子永久 disabled、
  //双击进不去 —— 这列的值由金蝶自己铺，不该也填不了
  const pid = %s, field = %s, row = %s;
  const $ = window.jQuery;
  const g = $(document.getElementById(pid + '-FENTITY-c')).data().kendoKDGrid;
  if (!g) return {err: 'no-grid'};
  const cols = g.columns.map(c => c.field || c.dataField);
  const idx = cols.indexOf(field);
  if (idx < 0) return {err: 'no-col'};
  const data = g.dataSource.data();
  if (row >= data.length) return {err: 'row-out'};
  let tr = g.tbody.find('tr[data-uid="' + data[row].uid + '"]');
  if (!tr.length) tr = g.tbody.find('tr').eq(row);
  const cd = tr.find('td').eq(idx).find('[class*=celldiv]');
  if (!cd.length) return {disabled: false};
  return {disabled: /k-gridCellstate-disabled/.test(cd.attr('class') || '')};
})()
"""

_JS_TOP_BANNER_BTN = r"""
(() => {
  //顶部黄色横幅（kd-msg-btn-container 一族，血泪结论#10：不是 .k-window）上 label 按钮的坐标，
  //pattern 限定横幅文本。三种横幅共用此通道：「内容已经修改，是否保存」「主业务组织已改变」
  //「刚刚执行的操作不能完成（网络）」。调用方负责 pattern 与 label 的安全配对：
  //保存横幅绝不点「是/确定」（会真保存进 ERP），组织变更横幅只在最前面点「确定」。
  const label = %s, pattern = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 6 && r.height > 6; };
  const msg = [...document.querySelectorAll('[class*=kd-msg],[class*=msg],[class*=banner],[class*=confirm]')].find(e =>
    pattern.test(e.innerText || '') && e.getBoundingClientRect().width > 30);
  if (!msg) return {none: true};
  let host = msg;
  for (let up = 0; up < 4 && host; up++) {
    const b = [...host.querySelectorAll('a,button,span,div')].find(e =>
      (e.textContent || '').trim() === label && vis(e));
    if (b) {
      const r = b.getBoundingClientRect();
      return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
    }
    host = host.parentElement;
  }
  return {noBtn: true, text: (msg.innerText || '').replace(/\\s+/g, ' ').slice(0, 60)};
})()
"""

_JS_HEADER_VISIBLE = r"""(() => {
  const root = document.getElementById(%s + '-' + %s + '_FSITEM');
  if (!root) return {ok: false};
  const r = root.getBoundingClientRect();
  return {ok: r.width > 10 && r.height > 8};
})()
"""

#====================== 明细真实输入链（model.set 只改显示层，保存必须走真实键入） ======================

_JS_CELL_XY = r"""
(() => {
  const pid = %s, field = %s, row = %s;
  const $ = window.jQuery;
  const el = document.getElementById(pid + '-FENTITY-c');
  if (!el) return {err: 'grid-gone'};
  const d = $(el).data(); const g = d && d.kendoKDGrid;
  if (!g || !g.columns) return {err: 'no-grid'};
  const cols = g.columns.map(c => c.field || c.dataField);
  const idx = cols.indexOf(field);
  if (idx < 0) return {err: 'no-col'};
  const data = g.dataSource.data();
  if (row >= data.length) return {err: 'row-out-of-range', rows: data.length};
  const uid = data[row].uid;
  let tr = g.tbody ? g.tbody.find('tr[data-uid="' + uid + '"]') : $();
  if (!tr.length) tr = g.tbody ? g.tbody.find('tr').eq(row) : $();
  if (!tr.length) return {err: 'no-tr'};
  const td = tr.find('td').eq(idx);
  if (!td.length) return {err: 'no-td'};
  td[0].scrollIntoView({block: 'center', inline: 'center'});
  const r = td[0].getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

#====================== 仓位弹性域（维度数据录入弹窗链路）======================
#仓位格编辑器 = k-textbox + .ui-cloumn-editor-icon 图标；点图标弹「维度数据录入」窗，窗内是一个 bosf7field
#（维度名由账套配置，代码不写死，只认弹窗里第一个 F7 编辑器）；F7 走放大镜→列表→返回数据，最后弹窗「确定」提交。

_JS_FLEX_ICON = r"""
(() => {
  //金蝶 kdgrid 的编辑器直接渲染在 td 里（td 自带 kd-edit-cell），不是 kendo 标准的 _editContainer
  //（2026-11-23 诊断：绕 _editContainer 找图标永远 no-edit-container；改为直接进 td 找，editCell 兜底）
  const pid = %s, field = %s, row = %s;
  const $ = window.jQuery;
  const g = $(document.getElementById(pid + '-FENTITY-c')).data().kendoKDGrid;
  if (!g) return {err: 'no-grid'};
  const cols = g.columns.map(c => c.field || c.dataField);
  const idx = cols.indexOf(field);
  if (idx < 0) return {err: 'no-col'};
  const data = g.dataSource.data();
  if (row >= data.length) return {err: 'row-out-of-range'};
  let tr = g.tbody.find('tr[data-uid="' + data[row].uid + '"]');
  if (!tr.length) tr = g.tbody.find('tr').eq(row);
  const td = tr.find('td').eq(idx);
  if (!td.length) return {err: 'no-td'};
  function findIcon() {
    const b = td.find('.ui-cloumn-editor-icon').get(0);
    if (!b) return null;
    const r = b.getBoundingClientRect();
    return (r.width > 2 && r.height > 2) ? {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)} : null;
  }
  let ic = findIcon();
  if (!ic) {
    try { g.editCell(td); } catch (e) {}
    ic = findIcon();
    if (!ic) return {err: 'no-flex-icon', tdCls: td.attr('class') || ''};
  }
  return ic;
})()
"""

_JS_TOPWIN_TEXT = r"""
(() => {
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window')].filter(vis);
  if (!ws.length) return {none: true};
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  return {text: (w.innerText || '').replace(/\s+/g, ' ').slice(0, 80)};
})()
"""

_JS_FLEX_MAGNET = r"""
(() => {
  //最上层弹窗（维度数据录入）里第一个 F7 编辑器的放大镜
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window')].filter(vis);
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  if (!w) return {err: 'no-window'};
  const ed = w.querySelector('[id$="-EDITOR"], [data-role=kdbosf7]');
  if (!ed) return {err: 'no-f7-editor'};
  const btn = ed.querySelector('.k-select.kd-editor-btn, .kd-editor-btn, .k-select');
  if (!btn) return {err: 'no-magnifier'};
  const r = btn.getBoundingClientRect();
  if (r.width < 2) return {err: 'magnifier-hidden'};
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_FLEX_INPUT = r"""
(() => {
  //维度弹窗内第一个 F7（本账套叫「供应商」，代码不写死维度名）的文本输入框坐标。
  //这是全模糊匹配主路的入口：往这里逐字符键入，金蝶自己弹模糊匹配下拉。
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const w = [...document.querySelectorAll('.k-window')].find(e => vis(e) && /维度数据录入/.test(e.innerText || ''));
  if (!w) return {err: 'no-flex-window'};
  const inp = [...w.querySelectorAll('input')].find(e => {
    const r = e.getBoundingClientRect();
    return r.width > 40 && r.height > 8 && !e.readOnly && (e.type === 'text' || !e.type);
  });
  if (!inp) return {err: 'no-input'};
  const r = inp.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_FUZZY_HIT_ROW = r"""
(() => {
  //金蝶 F7 全模糊匹配下拉：挂在 body 的 .k-animation-container（不是 k-window），内含「代码/名称」表头 + 命中行。
  //want 可以是编码(120)或名称片段(领益)，命中行文本「120 珠海领益通信技术有限公司」两者都含，includes 即可命中。
  //2026-11-23 实测：这是金蝶查值集分页外编码（如120）的唯一通道，值集列表窗只加载首页15项查不到。
  const want = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 20; };
  const pops = [...document.querySelectorAll('.k-animation-container')]
    .filter(vis).filter(e => /代码|名称/.test(e.innerText || ''));
  if (!pops.length) return {noPop: true};
  pops.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);  //最小的含表头浮层
  const p = pops[0];
  const dataRows = [...p.querySelectorAll('tr')].filter(e => {
    const r = e.getBoundingClientRect();
    if (r.width < 30 || r.height < 8) return false;
    if ((e.className || '').includes('filter-row')) return false;
    return !/^\s*代码\s*名称\s*$/.test((e.innerText || '').trim());  //去表头
  });
  const hit = dataRows.find(e => (e.innerText || '').includes(want));
  if (!hit) return {noHit: true, sample: dataRows.slice(0, 3).map(e => (e.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 40))};
  const r = hit.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          text: (hit.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 40)};
})()
"""

_JS_LIST_HIT_ROW = r"""
(() => {
  //值集资料选择列表：数据全在 dataSource（2026-11-23 实测 pageSize=200、15 项一页放完，
  //关键字框回车/查询按钮对这个弹窗都不生效，DOM tr 抓到的多是表头/空行）。
  //所以匹配走 dataSource：编码全等 或 名称包含目标；命中后用 tr[data-uid] 定位真实行坐标。
  //匹配不到返回 noItem —— 调用方绝不选第一行；值集超过一页且当前页没有则报 need-paging。
  const want = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window')].filter(vis);
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  if (!w) return {err: 'no-window'};
  const $ = window.jQuery;
  let g = null;
  $(w).find('*').addBack().each(function () {
    if (g) return;
    const d = $(this).data && $(this).data(); if (!d) return;
    for (const k of Object.keys(d)) {
      const v = d[k];
      if (v && typeof v === 'object' && v.dataSource && typeof v.columns !== 'undefined'
          && typeof v.select === 'function') { g = v; return; }
    }
  });
  if (!g) return {err: 'no-grid', title: (w.innerText || '').replace(/\s+/g, ' ').slice(0, 30)};
  const G = (it, f) => { try { const v = it.get ? it.get(f) : it[f]; return String(v || ''); } catch (e) { return ''; } };
  const data = g.dataSource.data();
  if (!data.length) return {loading: true};  //值集异步拉取中，别误判成"没有这个值"
  let hit = data.find(it => G(it, 'FNUMBER') === want)
         || data.find(it => G(it, 'FNAME').includes(want));
  if (!hit) {
    let total = null, pageLen = data.length;
    try { total = typeof g.dataSource.total === 'function' ? g.dataSource.total() : null; } catch (e) {}
    return {noItem: true, pageLen, total,
            sample: data.slice(0, 3).map(it => G(it, 'FNUMBER') + '/' + G(it, 'FNAME').slice(0, 16))};
  }
  const tr = w.querySelector('tr[data-uid="' + hit.uid + '"]');
  if (!tr) return {noTr: true, uid: hit.uid, name: G(hit, 'FNAME')};
  tr.scrollIntoView({block: 'center'});
  const r = tr.getBoundingClientRect();
  if (r.width < 2) return {noTr: true, reason: 'row-hidden'};
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          name: G(hit, 'FNAME')};
})()
"""

_JS_LIST_BACK_BTN = r"""
(() => {
  //值集列表窗的确认按钮文字是「返回」（员工/物料列表是「返回数据」），两种都认
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window')].filter(vis);
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  if (!w) return {err: 'no-window'};
  const b = [...w.querySelectorAll('a,button,span,div')].find(e =>
    ['返回', '返回数据'].includes((e.textContent || '').trim()) && e.getBoundingClientRect().width > 8);
  if (!b) return {err: 'no-back-btn'};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_FLEX_OK = r"""
(() => {
  //「维度数据录入」窗（列表窗关掉后它在上）里的确定。
  //⚠️按钮判定必须用独立小阈值：误用窗口的 vis(>30×30) 会永远找不到按钮（2026-11-23 实测 no-ok-btn）
  const visWin = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const w = [...document.querySelectorAll('.k-window')].find(e => visWin(e) && /维度数据录入/.test(e.innerText || ''));
  if (!w) return {err: 'no-flex-window'};
  const b = [...w.querySelectorAll('a,button,span,div')].find(e => {
    const r = e.getBoundingClientRect();
    return (e.textContent || '').trim() === '确定' && r.width > 8 && r.height > 6;
  });
  if (!b) return {err: 'no-ok-btn'};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_FLEX_CANCEL = r"""
(() => {
  //清场用：返回【最上层】可见窗的取消按钮坐标（取消/退出/×）。
  //⚠️绝不返回 确定/返回 的坐标（那会把当前选中的值提交进字段）。
  //一次只报一个窗（关掉最上面那个，下面的随之成为最上层），由调用方 CDP 点击后轮询。
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const ws = [...document.querySelectorAll('.k-window')].filter(vis);
  if (!ws.length) return {none: true};
  ws.sort((a, b) => (parseInt(getComputedStyle(a).zIndex) || 0) - (parseInt(getComputedStyle(b).zIndex) || 0));
  const w = ws[ws.length - 1];
  const b = [...w.querySelectorAll('a,button,span,div')].find(e => {
    const r = e.getBoundingClientRect();
    return ['取消', '退出'].includes((e.textContent || '').trim()) && r.width > 8 && r.height > 6;
  }) || w.querySelector('.k-i-close, .k-window-action');
  if (!b) return {noClose: true, title: (w.innerText || '').replace(/\s+/g, ' ').slice(0, 24)};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_BLOCKPASTE_MENU = r"""
(() => {
  //右键菜单里的「块粘贴」项坐标（菜单异步出现，调用方轮询）
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 5 && r.height > 5; };
  const el = [...document.querySelectorAll('li,.k-item,[role=menuitem],a,span,div')]
    .find(e => (e.textContent || '').trim() === '块粘贴' && vis(e) && e.children.length <= 1);
  if (!el) return {noItem: true};
  const r = el.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_BLOCKPASTE_DIALOG = r"""
(() => {
  //块粘贴弹窗：文本含「ctrl+v/粘贴」，内有 textarea + 「确定」按钮。
  //⚠️提交必须点「确定」——textarea 里 Enter 只是换行（2026-11-24 实测按 Enter 整列不落地）
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 20 && r.height > 10; };
  const ws = [...document.querySelectorAll('.k-window,.k-dialog,[role=dialog]')].filter(vis);
  const w = ws.find(e => /ctrl\+v|粘贴/.test(e.innerText || '')) || ws[ws.length - 1];
  if (!w) return {noWin: true};
  const ta = [...w.querySelectorAll('textarea')].find(vis)
    || [...w.querySelectorAll('input')].find(e => vis(e) && !e.readOnly);
  if (!ta) return {noTextarea: true, text: (w.innerText || '').replace(/\s+/g, ' ').slice(0, 60)};
  const r = ta.getBoundingClientRect();
  const ok = [...w.querySelectorAll('a,button,span,div')].find(e =>
    (e.textContent || '').trim() === '确定' && vis(e));
  if (!ok) return {noOk: true};
  const o = ok.getBoundingClientRect();
  return {tx: Math.round(r.x + r.width / 2), ty: Math.round(r.y + r.height / 2),
          okx: Math.round(o.x + o.width / 2), oky: Math.round(o.y + o.height / 2)};
})()
"""

_JS_PATCH_FILL_BTN = r"""
(() => {
  //明细工具栏「批量填充」(tbPatchFill) 按钮坐标。
  //金蝶按钮只认真实鼠标事件（JS .click() 无效，2026-11-23 实测），调用方必须用 cdp.click。
  const pid = %s;
  const bar = document.getElementById(pid + '-FENTITY_ENTRYBAR');
  const el = (bar && bar.querySelector('[id$="-tbPatchFill"]'))
    || [...(bar || document).querySelectorAll('a,span,div,li')].find(e =>
         (e.textContent || '').trim() === '批量填充' && e.getBoundingClientRect().width > 10);
  if (!el) return {err: 'no-patchfill-btn'};
  const r = el.getBoundingClientRect();
  if (r.width < 8 || r.height < 6) return {err: 'patchfill-hidden'};
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_GRID_CURRENT = r"""
(() => {
  //读明细格 current 单元格的列 field。批量填充的源值由服务器从 current 行该列读取
  //（血泪结论#16），current 漂到别的列就会铺空。API 取不到时返回 unknown，调用方跳过检查。
  const pid = %s;
  const $ = window.jQuery;
  try {
    const g = $(document.getElementById(pid + '-FENTITY-c')).data().kendoKDGrid;
    if (!g) return {unknown: true};
    let td = null;
    if (typeof g.current === 'function') {
      const c = g.current();
      if (c && c.length) td = c[0];
    }
    if (!td) td = (g.tbody && g.tbody.find('td.k-current-data').get(0)) || null;
    if (!td) return {unknown: true};
    const tr = td.closest('tr');
    const idx = [...tr.children].indexOf(td);
    const cols = g.columns.map(c => c.field || c.dataField);
    return {field: cols[idx] || ''};
  } catch (e) { return {unknown: true, err: String(e).slice(0, 60)}; }
})()
"""

_JS_COL_TEXTS = r"""
(() => {
  //读整列各行显示文本（DOM 层，抗半销毁 grid；批量填充前后对比用）
  const pid = %s, field = %s;
  try {
    const g = $(document.getElementById(pid + '-FENTITY-c')).data().kendoKDGrid;
    if (!g) return {err: 'no-grid'};
    const cols = g.columns.map(c => c.field || c.dataField);
    const idx = cols.indexOf(field);
    if (idx < 0) return {err: 'no-col'};
    const trs = g.tbody.find('tr');
    const out = [];
    for (let i = 0; i < trs.length; i++) out.push((trs.eq(i).find('td').eq(idx).text() || '').trim());
    return {texts: out};
  } catch (e) { return {err: String(e).slice(0, 80)}; }
})()
"""

_JS_CELL_TEXT = r"""
(() => {
  const pid = %s, field = %s, row = %s;
  const $ = window.jQuery;
  const el = document.getElementById(pid + '-FENTITY-c');
  if (!el) return {err: 'grid-gone'};
  const g = $(el).data().kendoKDGrid;
  const cols = g.columns.map(c => c.field || c.dataField);
  const idx = cols.indexOf(field);
  if (idx < 0) return {err: 'no-col'};
  const data = g.dataSource.data();
  if (row >= data.length) return {err: 'row-out-of-range'};
  let tr = g.tbody ? g.tbody.find('tr[data-uid="' + data[row].uid + '"]') : $();
  if (!tr.length) tr = g.tbody ? g.tbody.find('tr').eq(row) : $();
  if (!tr.length) return {err: 'no-tr'};
  return {text: (tr.find('td').eq(idx).text() || '').trim()};
})()
"""

_JS_FOCUS_INPUT = r"""
(() => {
  const pid = %s, key = %s;
  const root = document.getElementById(pid + '-' + key + '_FSITEM');
  if (!root) return {ok: false, err: 'no-field'};
  const inp = root.querySelector('input,textarea');
  if (!inp) return {ok: false, err: 'no-input'};
  inp.focus();
  const ok = document.activeElement === inp;
  if (ok) { try { inp.select && inp.select(); } catch (e) {} }
  const r = inp.getBoundingClientRect();
  return {ok, err: ok ? null : 'focus-fail',
          x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_READ_INPUT = r"""
(() => {
  const root = document.getElementById(%s + '-' + %s + '_FSITEM');
  const inp = root && root.querySelector('input,textarea');
  return {value: inp ? String(inp.value) : null};
})()
"""

_JS_F7_READ = r"""
(() => {
  const pid = %s, key = %s;
  const $ = window.jQuery;
  const el = document.getElementById(pid + '-' + key + '-EDITOR');
  if (!el) return {err: 'no-editor'};
  const d = $(el).data(); const w = d.kendoKDBosF7 || d.handler;
  if (!w) return {err: 'no-widget'};
  let vd = null;
  try { vd = w.valData(); } catch (e) {}
  return {name: vd ? String(vd.name || '') : '', number: vd ? String(vd.number || '') : ''};
})()
"""

_JS_F7_MAGNET = r"""
(() => {
  const root = document.getElementById(%s + '-' + %s + '_FSITEM');
  if (!root) return {err: 'no-field'};
  const btn = root.querySelector('.k-select.kd-editor-btn, .kd-editor-btn, .k-select');
  if (!btn) return {err: 'no-magnifier'};
  const r = btn.getBoundingClientRect();
  if (r.width < 2 || r.height < 2) return {err: 'magnifier-hidden'};
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_POPUP_FILTER = r"""
(() => {
  //弹窗 = 可见且含「返回数据」按钮的 .k-window（员工/物料列表都长这样，类名不稳定）
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const w = [...document.querySelectorAll('.k-window')].find(e => vis(e) && /返回数据/.test(e.innerText || ''));
  if (!w) return {err: 'no-popup'};
  //优先「请输入关键字后回车查询」框，其次任意可见 text input
  const inputs = [...w.querySelectorAll('input')].filter(e => {
    const r = e.getBoundingClientRect();
    return r.width > 30 && r.height > 8 && !e.readOnly && (e.type === 'text' || !e.type);
  });
  const inp = inputs.find(e => /关键字/.test(e.placeholder || '')) || inputs[0];
  if (!inp) return {noFilter: true};
  const r = inp.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), ph: inp.placeholder || ''};
})()
"""

_JS_POPUP_ROW = r"""
(() => {
  const target = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 50 && r.height > 10 && r.top > 0; };
  const w = [...document.querySelectorAll('.kd-dynamicform-active')].filter(vis)[0];
  if (!w) return {err: 'no-popup'};
  const rows = [...w.querySelectorAll('tr')].filter(e =>
    vis(e) && !/设置/.test(e.innerText || '') && (e.innerText || '').includes(target));
  if (!rows.length) return {noRow: true};
  rows.sort((a, b) => (a.innerText || '').length - (b.innerText || '').length);
  const r = rows[0].getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          text: (rows[0].innerText || '').replace(/\s+/g, ' ').trim().slice(0, 50)};
})()
"""

_JS_POPUP_BACK = r"""
(() => {
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 10 && r.height > 8; };
  const w = [...document.querySelectorAll('.kd-dynamicform-active')].filter(vis)[0];
  if (!w) return {err: 'no-popup'};
  const b = [...w.querySelectorAll('a,button,span,div')].find(e => (e.textContent || '').trim() === '返回数据' && vis(e));
  if (!b) return {err: 'no-back-btn'};
  const r = b.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_POPUP_CLOSE = r"""
(() => {
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 30 && r.height > 30; };
  const w = [...document.querySelectorAll('.kd-dynamicform-active')].filter(vis)[0];
  if (!w) return {gone: true};
  const x = w.querySelector('.k-i-close, .k-window-action');
  if (x) { x.click(); return {clicked: true}; }
  return {stuck: true};
})()
"""

_JS_COMBO_TRIGGER = r"""
(() => {
  const root = document.getElementById(%s + '-' + %s + '_FSITEM');
  if (!root) return {err: 'no-field'};
  const e = root.querySelector('.k-select');
  const t = e || root;
  const r = t.getBoundingClientRect();
  if (!(r.width > 4 && r.height > 4)) return {err: 'no-trigger'};
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_COMBO_OPTIONS = r"""
(() => {
  const target = %s;
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 20 && r.height > 8 && r.top > 0; };
  const lis = [...document.querySelectorAll('li')].filter(e =>
    vis(e) && (e.innerText || '').replace(/\s+/g, '').includes(target.replace(/\s+/g, '')));
  if (!lis.length) return {noOption: true};
  const r = lis[0].getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
          n: lis.length, text: (lis[0].innerText || '').replace(/\s+/g, ' ').trim().slice(0, 20)};
})()
"""

_JS_COMBO_READ = r"""
(() => {
  const $ = window.jQuery;
  const root = document.getElementById(%s + '-' + %s + '_FSITEM');
  if (!root) return {err: 'no-field'};
  let w = null;
  $(root).find('*').addBack().each(function () {
    if (w) return;
    const d = $(this).data && $(this).data(); if (!d) return;
    for (const k of Object.keys(d)) {
      const o = d[k];
      if (o && typeof o.value === 'function' && typeof o.text === 'function' && /dropdown|combo/i.test(k)) { w = o; return; }
    }
  });
  const inp = root.querySelector('input');
  return {text: w ? String(w.text()).trim() : '', code: w ? String(w.value()) : '',
          inputVal: inp ? inp.value : ''};
})()
"""

_JS_MATERIAL_COUNT = r"""
(() => {
  const pid = %s;
  const $ = window.jQuery;
  const el = document.getElementById(pid + '-FENTITY-c');
  if (!el) return {err: 'grid-gone'};
  const g = $(el).data().kendoKDGrid;
  const cols = g.columns.map(c => c.field || c.dataField);
  const mIdx = cols.indexOf('FMATERIALID');
  if (mIdx < 0) return {err: 'no-col'};
  //按显示层数行：真实输入链提交后显示文本才有值（model.set 的假值显示层是空的）
  let n = 0;
  g.dataSource.data().forEach((it, r) => {
    let tr = g.tbody ? g.tbody.find('tr[data-uid="' + it.uid + '"]') : $();
    if (!tr.length) tr = g.tbody ? g.tbody.find('tr').eq(r) : $();
    if (!tr.length) return;
    const tx = (tr.find('td').eq(mIdx).text() || '').trim();
    if (tx && tx !== 'select') n++;
  });
  return {count: n, rows: g.dataSource.data().length};
})()
"""

_JS_ROW_COUNT = r"""
(() => {
  const pid = %s;
  const $ = window.jQuery;
  const el = document.getElementById(pid + '-FENTITY-c');
  if (!el) return {err: 'grid-gone'};
  const g = $(el).data().kendoKDGrid;
  return {rows: g.dataSource.data().length};
})()
"""

_JS_ADD_ROW = r"""
(() => {
  //明细页签工具栏「新增行」。它是 k-split-button：A 主体直接加行、右侧箭头弹菜单。
  //⚠️只返回坐标由调用方 CDP 真实点击——金蝶工具栏按钮的 JS .click() 无效
  //（2026-11-23 实测：ok:true 但 rowCount 纹丝不动，和取消按钮同一个坑）
  const pid = %s;
  const bar = document.getElementById(pid + '-FENTITY_ENTRYBAR');
  const scope = bar || document;
  const el = [...scope.querySelectorAll('a.kdmenuitem, a')].find(e => {
    const tx = (e.textContent || '').trim();
    const r = e.getBoundingClientRect();
    return tx === '新增行' && r.width > 10 && r.height > 8;
  });
  if (!el) return {err: 'no-add-row'};
  const r = el.getBoundingClientRect();
  return {x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2)};
})()
"""

_JS_PING = r"""
(() => ({
  href: location.href.split('?')[0],
  login: /index\.aspx/.test(location.pathname),
  console: (document.body.innerText || '').includes('常用功能'),
  jq: typeof window.jQuery,
  drafts: [...new Set([...document.querySelectorAll('[id]')]
    .map(e => (e.id.match(/^([0-9a-f-]{36})-FENTITY-c$/i) || [])[1]).filter(Boolean))],
})())
"""

#====================== 公共小件 ======================

def _norm(text):
    #归一化比对：去空白、剥掉「名称(编码)」尾巴（金蝶显示层常带）
    t = re.sub(r"\s+", "", str(text or ""))
    t = re.sub(r"[（(][^（）()]*[)）]$", "", t)
    return t


def _draft(cdp, tag):
    #定位当前【唯一可见】的草稿编辑页；返回 (pid, None) 或 (None, 原因)
    r = cdp.js(_JS_FIND_DRAFT)
    v = r.get("value") if "value" in r else r
    if not isinstance(v, dict):
        return None, f"探测异常: {str(v)[:60]}"
    if v.get("err") == "login-page":
        return None, "登录态失效，请在 Edge 里重新登录"
    if v.get("err") == "no-draft":
        return None, "没有打开的调拨申请单编辑页（WebNavigate 跑了吗）"
    if v.get("err") == "multi-draft":
        return None, f"同时可见多个草稿页签 {v.get('pids')}，分不清写哪张 → 拒写（同 _detail_grid 原则）"
    return v.get("pid"), None


def _row_value(field, row):
    #没给 value 时从当前单据队列按 field 取（Excel 契约不变，同 UiaSetCell）
    getter = uia_actions._COLUMN_ATTR_BY_FIELD.get(field)
    if getter is None:
        return None
    try:
        import transfer_actions
        order = transfer_actions._current_order()
        if order is None or row >= len(order.items):
            return None
        v = getter(order.items[row])
    except Exception:
        return None
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")  #日期对象 → 网页输入框吃的写法
    return str(v)


def _order_value(field):
    #单头字段整单同值，取首行（同 uia_actions._header_value_from_order）
    return _row_value(field, 0)


def _ensure_rows(cdp, pid, need, tag):
    #界面行数不够 Excel 行数时点「新增行」补齐（真实输入链不像 model.set 能凭空扩行；
    #点击必须走 CDP 真实鼠标：金蝶工具栏按钮 JS .click() 无效（ok:true 但行数不动）
    for _ in range(60):
        r = cdp.js(_JS_ROW_COUNT % json.dumps(pid))
        v = r.get("value") if "value" in r else r
        rows = (v or {}).get("rows")
        if rows is None:
            print(f"[{tag}] 数行失败: {v}")
            return False
        if rows >= need:
            return True
        a = cdp.js(_JS_ADD_ROW % json.dumps(pid))
        av = a.get("value") if "value" in a else a
        if not (isinstance(av, dict) and av.get("x")):
            print(f"[{tag}] 界面 {rows} 行 < 需要 {need} 行，且点不到「新增行」: {av}")
            return False
        cdp.click(av["x"], av["y"])
        time.sleep(0.9)
    return False


def _cell_text(cdp, pid, key, row, tag):
    #读明细格子的【显示文本】（显示层=billModel 层，model.set 的假值读不出来）——
    #校验口径：金蝶没解析成功时格子里是空或 select
    r = cdp.js(_JS_CELL_TEXT % (json.dumps(pid), json.dumps(key), row))
    v = r.get("value") if "value" in r else r
    if not isinstance(v, dict) or "text" not in v:
        return None, f"读取失败: {v}"
    return str(v.get("text") or ""), None


def _type_into_cell(cdp, pid, key, row, value, f7, tag):
    #真实输入链写明细格（唯一被保存确认进 billModel 的通道）：
    #双击进编辑态 → 三击全选编辑器内容 → Delete → 键入 → F7 列 Enter 触发编码解析 / 普通列 Tab 失焦提交
    #⚠️清空必须三击不能用 Ctrl+A：双击偶发没进编辑态时 Ctrl+A 会把整页字段名选中，
    #  Delete 触发「内容已修改」横幅、键入落空；三击只在格子内建选区，安全
    xy = cdp.js(_JS_CELL_XY % (json.dumps(pid), json.dumps(key), row))
    v = xy.get("value") if "value" in xy else xy
    if not isinstance(v, dict) or "x" not in v:
        return False, f"定位格子失败: {v}"
    cdp.dbl_click(v["x"], v["y"])
    time.sleep(0.8)
    #编辑态门：上一列批量填充后网格异步收尾，这一下双击会被吞；没进编辑态就 Ctrl+A
    #会把整页字段名选中（Delete 弹「内容已修改」、键入落空）。
    #⚠️清空仍用 Ctrl+A 不用三击：三击会把刚建立的格子编辑态点退出，三击只适用于弹窗过滤框；
    #  有了这道门，Ctrl+A 只在格子编辑器内生效，安全
    for _ in range(5):
        a = cdp.js(_JS_EDIT_ACTIVE % (json.dumps(pid), json.dumps(key), str(row))).get("value") or {}
        if a.get("ok"):
            break
        cdp.dbl_click(v["x"], v["y"])
        time.sleep(1.0)
    else:
        print(f"[{tag}] {key} 双击 5 次仍未进编辑态，放弃")
        return False, "双击没进编辑态"
    cdp.select_all()
    time.sleep(0.2)
    cdp.key_delete()
    time.sleep(0.2)
    if f7 and _CJK_RE.search(value):
        #F7 列填中文名称（如 电子仓/委外仓）：insertText 不触发 typeahead，必须逐字符键入
        #让金蝶弹模糊匹配下拉，唯一命中时 Enter 直接选中。多命中时 Enter 行为不确定
        #（业务前提是组织内名称唯一）；不唯一时格子会被金蝶清空 → 校验报错，不会假成功
        cdp.type_chars(value)
        time.sleep(1.5)  #等模糊下拉过滤稳定再 Enter（留 1.5s 余量）
    else:
        cdp.type_text(value)
    time.sleep(0.3)
    if f7:
        cdp.key_enter2()  #F7：Enter 让金蝶解析编码/唯一模糊命中并联动带出名称/单位
    else:
        cdp.key_tab2()    #普通列：Tab 失焦提交（金蝶会发 UpdateValue）
    time.sleep(1.2)
    return True, None


def _verify_cell(cdp, pid, key, row, value, f7, tag):
    #写后校验（显示层判据）：
    #  物料列 → 名称列被金蝶带出非空才算解析成功（编码本身在名称解析前就显示，读它会假阳）
    #  其他 F7 → 格子文本非空且不是 select 占位
    #  普通列 → 文本含写入值（数量显示 3.0000000000 这类补零写法）
    #⚠️F7 列 Enter 后金蝶发 GetLookupList2More 异步查询，回来才 SetItemValueByIDFromClient 提交显示；
    #  读一次就下结论会把"查询还没回来"当成空，必须轮询等异步回填（仓库列比物料列慢）
    read_key = "FMATERIALNAME" if (f7 and key == "FMATERIALID") else key
    #⚠️普通列也要给足等待：前面逐行填完 F7 列后金蝶还在异步联动（单位/库存量刷新），窗口太短会把「联动还没完成」当成空。
    deadline = time.time() + (8.0 if f7 else 6.0)
    actual = ""
    while True:
        actual, err = _cell_text(cdp, pid, read_key, row, tag)
        if err:
            return False, err
        if f7:
            if actual and "select" not in actual:
                return True, actual
        elif _norm(value) in _norm(actual) or _norm(actual) == _norm(value):
            return True, actual
        if time.time() >= deadline:
            return False, actual
        time.sleep(0.5)


def _wait_js(cdp, template, arg=None, timeout=8.0, ok_key="x", abort_keys=()):
    #轮询一个返回 {x,y} 的 JS 模板直到拿到 ok_key 或超时；返回最后一个结果 dict
    #arg：已 json.dumps 好的 JS 字面量（单值或元组）；None=模板无占位符
    #abort_keys：明确的业务否定（如 noRow 过滤后真没有），命中立即返回；
    #  ⚠️no-window/no-edit-container 这类"结构还没长出来"不能当否定 —— 弹窗是异步加载的，点开后要等一会儿才查得到。
    if arg is None:
        expr = template
    else:
        expr = template % (arg if isinstance(arg, tuple) else (arg,))
    end = time.time() + timeout
    last = {}
    while time.time() < end:
        r = cdp.js(expr)
        last = r.get("value") if "value" in r else r
        if not isinstance(last, dict):
            last = {}
            time.sleep(0.5)
            continue
        if last.get(ok_key) or any(last.get(k) is not None for k in abort_keys):
            return last
        time.sleep(0.5)
    return last


def _clear_windows(cdp, to=6.0):
    #清场：一次关一个最上层窗（模板返回其取消按钮坐标），用 CDP 真实鼠标点击
    #⚠️必须 CDP 点击不能 JS .click()：金蝶窗的取消按钮绑的是真实鼠标事件，JS 合成 click 点了窗不关
    end = time.time() + to
    while time.time() < end:
        r = cdp.js(_JS_FLEX_CANCEL)
        v = r.get("value") if "value" in r else r
        if not isinstance(v, dict) or v.get("none"):
            #窗清完再扫 kd-msg 顶部横幅（残留横幅会挡住放大镜/网格）。两种横幅都只点【取消】：
            #「主业务组织已改变」取消=放弃待生效的组织变更；「内容已经修改，是否保存」
            #取消=留在当前单不保存（点「是/确定」会真保存单据进 ERP，测试期绝对禁止）
            clicked = False
            for label, pattern in (("取消", r"/主业务组织/"), ("取消", r"/内容已经修改|是否保存/")):
                b = _banner_btn(cdp, label, pattern)
                if b.get("x"):
                    cdp.click(b["x"], b["y"])
                    time.sleep(0.6)
                    clicked = True
            if _dismiss_net_banner(cdp, to=0.5):  #网络横幅也顺手清（上一步提交未生效时代码会重跑）
                clicked = True
            if clicked:
                continue
            return True
        if v.get("noClose") or not v.get("x"):
            return False  #有窗但找不到取消按钮 / 连接断开拿不到坐标，直接判失败
        cdp.click(v["x"], v["y"])
        time.sleep(0.6)
    return False


def _banner_btn(cdp, label, pattern):
    #顶部黄色横幅（kd-msg-btn-container）上 label 按钮的坐标；pattern 是 JS 正则源码，
    #限定只认哪条横幅（保存横幅绝不点「是/确定」，组织变更横幅绝不点「确定」重置单）
    r = cdp.js(_JS_TOP_BANNER_BTN % (json.dumps(label, ensure_ascii=False), pattern))
    v = r.get("value") if "value" in r else r
    return v if isinstance(v, dict) else {}


def _orgchange_btn(cdp, label):
    #「主业务组织已改变，系统将创建新单，是否继续？」横幅上的按钮坐标
    return _banner_btn(cdp, label, r"/主业务组织/")


def _confirm_orgchange(cdp, tag, to=6.0):
    #apply_org 键入回车后金蝶弹「主业务组织已改变」横幅，不点确定值不生效且挡住后续操作
    #点确定后金蝶把同页签重置成新单（明细/单头全清）→ 申请组织必须排在一切填写之前。
    #返回 True=已确认且草稿重新就位（调用方要重取 pid）；False=无横幅（正常）或确认后就位超时
    end = time.time() + to
    r = {}
    while time.time() < end:
        r = _orgchange_btn(cdp, "确定")
        if r.get("x"):
            break
        if r.get("noBtn"):
            print(f"[{tag}] 主业务组织横幅找不到「确定」: {r}")
            return False
        time.sleep(0.5)
    if not r.get("x"):
        return False
    cdp.click(r["x"], r["y"])
    print(f"[{tag}] 已点主业务组织横幅「确定」（金蝶重置同页签为新单，等就位）")
    deadline = time.time() + 30.0
    while time.time() < deadline:
        time.sleep(1.0)
        pid, _ = _draft(cdp, tag)
        if pid and _orgchange_btn(cdp, "确定").get("none"):
            return True
    print(f"[{tag}] 主业务组织确认后 30s 新单未就位")
    return False


def _dismiss_net_banner(cdp, to=2.0):
    #网络横幅「刚刚执行的操作不能完成，请尝试以下办法…联系系统管理员」：
    #它一出现就说明上一步对服务器的提交被吞了（值停在显示层，billModel 没进），
    #且会挡住后续所有操作。点【确定】清掉后调用方要重跑这一步（这类横幅只有 确定 一个按钮，点它不保存单据）。
    #重跑策略=「谁被吞谁重跑」：每个写动作自带回读校验，校验没过就是自己被吞了，
    #清横幅后重跑本步即可（各动作的 2 轮重试循环开头都会 _clear_windows，那里已接横幅清除）。
    #不做跨动作台账重放：重放 WebNewBill 会弃单、重放块粘贴会二次扩行，风险远大于收益。
    end = time.time() + to
    while time.time() < end:
        b = _banner_btn(cdp, "确定", r"/刚刚执行|不能完成|网络状况/")
        if b.get("x"):
            cdp.click(b["x"], b["y"])
            time.sleep(0.8)
            print("[net-banner] 检测到网络横幅，已点「确定」清场，调用方重跑这一步")
            return True
        if b.get("none") or b.get("noBtn"):
            return False  #没横幅（正常）或横幅在但按钮找不到：不空转，交给调用方报错
        time.sleep(0.4)
    return False


def _fill_flex(cdp, pid, key, row, value, tag):
    #仓位弹性域：双击格子→维度图标→维度弹窗。填值两条路：
    #  路A（主路）：弹窗内 F7 输入框【逐字符键入】→ 金蝶「全模糊匹配」下拉 → 点命中行 → 弹窗确定。
    #    这是查值集分页外编码（如120）的唯一通道——值集列表窗只加载首页15项，120 不在其中。
    #  路B（兜底）：F7放大镜 → 值集列表窗 → dataSource 匹配选行 → 返回 → 弹窗确定。
    #两条路都匹配不到就清场返回 False，绝不选第一行。value 可以是编码(120)或名称片段(领益)。
    def step(js, arg=None, to=8.0, abort=("noRow",)):
        #abort 只放"窗口已存在但缺目标元素"的硬否定；no-window 类继续轮询（弹窗异步）
        return _wait_js(cdp, js, arg, to, abort_keys=abort)

    def read_cell():
        #路A/B 提交后统一读仓位格（维度确定后金蝶异步回填，轮询到非空非 select）
        actual, err = _cell_text(cdp, pid, key, row, tag)
        deadline = time.time() + 8.0
        while not err and (not actual or "select" in actual) and time.time() < deadline:
            time.sleep(0.5)
            actual, err = _cell_text(cdp, pid, key, row, tag)
        return actual, err

    def confirm_flex():
        #维度弹窗点「确定」把选中的维度值提交进仓位格
        ok = step(_JS_FLEX_OK, None, to=6.0)
        if not ok.get("x"):
            return False
        cdp.click(ok["x"], ok["y"])
        time.sleep(1.0)
        return True

    #前置清场：残留的维度/列表窗会挡住后续，先全关
    _clear_windows(cdp)

    xy = step(_JS_CELL_XY, (json.dumps(pid), json.dumps(key), str(row)))
    if not xy.get("x"):
        return False, f"定位格子失败: {xy}"
    cdp.dbl_click(xy["x"], xy["y"])
    time.sleep(1.0)

    icon = step(_JS_FLEX_ICON, (json.dumps(pid), json.dumps(key), str(row)), to=6.0,
                abort=("no-flex-icon", "no-col", "row-out-of-range", "no-grid"))
    if not icon.get("x"):
        return False, f"维度图标不可用: {icon}"
    cdp.click(icon["x"], icon["y"])

    inp = step(_JS_FLEX_INPUT, None, to=10.0)
    if not inp.get("x"):
        _clear_windows(cdp)
        return False, f"维度弹窗/输入框没出现: {inp}"

    #⚠️先清空输入框：格子已有值时（重跑同一单）直接键入会追加在旧值后面，模糊匹配必失败
    #清空用三击+Delete：金蝶 F7 框里 Ctrl+A 不生效，只有三击建立选区才删得干净
    cdp.triple_click(inp["x"], inp["y"])
    time.sleep(0.3)
    cdp.key_delete()
    time.sleep(0.3)
    cdp.type_chars(value)  #insertText 不触发 typeahead，必须逐字符键盘事件
    hit = {}
    end = time.time() + 10.0
    while time.time() < end:
        r = cdp.js(_JS_FUZZY_HIT_ROW % json.dumps(value, ensure_ascii=False))
        hit = r.get("value") if "value" in r else r
        if not isinstance(hit, dict):
            hit = {}
        if hit.get("x") or hit.get("noHit"):
            break  #命中 / 下拉出来了但没这个值 —— 确定结果，不再空等
        time.sleep(0.6)
    if hit.get("x"):
        cdp.click(hit["x"], hit["y"])
        time.sleep(1.5)
        if confirm_flex():
            actual, err = read_cell()
            if not err and actual and "select" not in actual:
                return True, actual
        #路A点了行却没提交成功：清残留落到路B
    elif hit.get("noHit"):
        #下拉弹出来了但没这个值 = 值集确认没有，不必再开值集列表（它只是首页15项的子集）。
        #⚠️先 Escape 关掉模糊匹配浮层：它盖着维度弹窗的取消按钮，直接清场会把点击落到浮层上，浮层关了窗还在。
        cdp.key("Escape", "Escape", 27)
        time.sleep(0.5)
        _clear_windows(cdp)
        return False, (f"值集里没有「{value}」（模糊匹配下拉无此项，样例 {hit.get('sample')}）"
                       f"——检查 Excel 仓位填的是编码还是名称")

    #⑤ 路B（仅当路A 下拉根本没弹出时才试）：放大镜 → 值集列表 → dataSource 匹配 → 返回 → 确定
    #先 Escape 关掉可能挂着的模糊匹配浮层（它会盖住放大镜，导致点了不开列表窗），
    #再三击+Delete 清空输入框避免脏值干扰（Ctrl+A 在金蝶 F7 框不生效）
    cdp.key("Escape", "Escape", 27)
    time.sleep(0.4)
    cdp.triple_click(inp["x"], inp["y"])
    time.sleep(0.3)
    cdp.key_delete()
    time.sleep(0.3)
    mag = step(_JS_FLEX_MAGNET, None, to=8.0)
    if not mag.get("x"):
        _clear_windows(cdp)
        return False, f"路A未命中({hit})，且放大镜不可用: {mag}"
    cdp.click(mag["x"], mag["y"])
    hit2 = {}
    end = time.time() + 12.0
    while time.time() < end:
        r = cdp.js(_JS_LIST_HIT_ROW % json.dumps(value, ensure_ascii=False))
        hit2 = r.get("value") if "value" in r else r
        if not isinstance(hit2, dict):
            hit2 = {}
        if hit2.get("x") or hit2.get("noItem") or hit2.get("noTr"):
            break
        time.sleep(0.8)
    if not hit2.get("x"):
        _clear_windows(cdp)
        if hit2.get("noItem"):
            return False, (f"两条路都没查到「{value}」：模糊匹配={hit}；值集首页 "
                           f"{hit2.get('pageLen')}/{hit2.get('total')} 项 样例 {hit2.get('sample')}"
                           f"——检查 Excel 仓位填的是编码还是名称")
        return False, f"值集列表选行失败: {hit2}"
    cdp.click(hit2["x"], hit2["y"])
    time.sleep(1.0)
    back = step(_JS_LIST_BACK_BTN, None, to=6.0)
    if not back.get("x"):
        _clear_windows(cdp)
        return False, f"列表窗没有「返回」按钮: {back}"
    cdp.click(back["x"], back["y"])
    time.sleep(1.5)
    if not confirm_flex():
        _clear_windows(cdp)
        return False, "维度弹窗没有「确定」按钮"
    actual, err = read_cell()
    if err:
        return False, err
    passed = bool(actual) and "select" not in actual
    return passed, actual

#====================== 动作 ======================

@AgentServer.custom_action("WebSetForm")  #切换单据白名单，任务第一个节点调
class WebSetForm(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        form = str(param.get("form", "") or "").strip()
        if form not in _FORMS:
            print(f"[WebSetForm] 未知 form {form!r}，v1 可选: {list(_FORMS)}（直接调拨单冻结为前台基线）")
            return False
        global _ACTIVE_FORM
        _ACTIVE_FORM = form
        print(f"[WebSetForm] 网页端单据已切到 {form}（明细 {len(_ENTRY_FIELDS)} 列 / "
              f"单头 文本{len(_HEADER_TEXT_FIELDS)} F7{len(_HEADER_F7_FIELDS)} 下拉{len(_HEADER_COMBO_FIELDS)}）")
        return True


@AgentServer.custom_action("WebNavigate")  #主控台→列表→新增，等编辑页就位
class WebNavigate(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        text = str(param.get("text", "") or "").strip()
        wait_label = str(param.get("wait_label", "") or "").strip()
        timeout_s = float(param.get("timeout_ms", 60000) or 60000) / 1000.0
        header_key = _LABEL_TO_HEADER.get(wait_label)
        if not text:
            print("[WebNavigate] 缺 text（主控台入口文字，如 调拨申请单列表）")
            return False
        if not header_key:
            print(f"[WebNavigate] wait_label {wait_label!r} 不在 _LABEL_TO_HEADER，编辑页没法判就位")
            return False

        cdp = _open("WebNavigate")
        if cdp is None:
            return False
        try:
            #入口：主控台点文字（若已在列表页，入口可能不可见 → 直接找「新增」也算就位）
            rect = cdp.js(_JS_CLICK_RECT % json.dumps(text, ensure_ascii=False)).get("value") or {}
            if rect.get("x"):
                cdp.click(rect["x"], rect["y"])
                print(f"[WebNavigate] 已点「{text}」")
            else:
                print(f"[WebNavigate] 「{text}」不可见，可能已在列表页，直接找新增")

            deadline = time.time() + min(timeout_s, 30.0)
            add = {}
            while time.time() < deadline:
                time.sleep(_POLL_S)
                add = cdp.js(_JS_CLICK_RECT % json.dumps("新增", ensure_ascii=False)).get("value") or {}
                if add.get("x"):
                    break
            if not add.get("x"):
                print("[WebNavigate] 等不到「新增」按钮（列表没打开/登录失效）")
                return False
            cdp.click(add["x"], add["y"])
            print("[WebNavigate] 已点「新增」，等编辑页…")

            #等草稿页 + wait_label 字段出现。
            #⚠️必须带补点循环：列表页刚点开时「新增」可能还在异步渲染，第一下会点到旧 DOM 空位
            #每 8s 没草稿就重新定位「新增」再点一次（坐标要重取，列表可能已重排）
            deadline = time.time() + timeout_s
            next_click = time.time() + 8.0
            clicks = 1
            while time.time() < deadline:
                time.sleep(1.0)
                pid, err = _draft(cdp, "WebNavigate")
                if pid:
                    hv = cdp.js(_JS_HEADER_VISIBLE % (json.dumps(pid), json.dumps(header_key))).get("value") or {}
                    if hv.get("ok"):
                        print(f"[WebNavigate] 编辑页就位 pid={pid[:8]} 「{wait_label}」可见（共点新增 {clicks} 次）")
                        return True
                if time.time() >= next_click:
                    add = cdp.js(_JS_CLICK_RECT % json.dumps("新增", ensure_ascii=False)).get("value") or {}
                    if add.get("x"):
                        cdp.click(add["x"], add["y"])
                        clicks += 1
                        print(f"[WebNavigate] 无动静，补点「新增」第 {clicks} 次")
                    next_click = time.time() + 8.0
            print(f"[WebNavigate] {timeout_s:.0f}s 内编辑页没就位（点了 {clicks} 次新增，最后一次: {err if pid is None else '字段未出现'}）")
            return False
        finally:
            cdp.close()


@AgentServer.custom_action("WebNewBill")  #当前单填完后开下一张：点顶栏「新增」→ 顶部横幅点「否」放弃当前单
class WebNewBill(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        #最后一张单不开新单：留在当前单（提交链路下当前单已进 ERP）或直接收尾，
        #省得每张跑完多开一张空白新增单（单张 Excel 同样走这条）
        try:
            import transfer_actions
            if transfer_actions.is_last_order():
                print("[WebNewBill] 已是最后一张 → 不点新增，留在当前单收尾")
                return True
        except Exception:
            pass  #读不到队列就按常规开新单，不报错
        wait_label = str(param.get("wait_label", "") or "").strip()
        timeout_s = float(param.get("timeout_ms", 60000) or 60000) / 1000.0
        header_key = _LABEL_TO_HEADER.get(wait_label)
        if not header_key:
            print(f"[WebNewBill] wait_label {wait_label!r} 不在 _LABEL_TO_HEADER，无法判就位")
            return False

        cdp = _open("WebNewBill")
        if cdp is None:
            return False
        try:
            #① 点顶栏「新增」
            add = cdp.js(_JS_TOOLBAR_NEW).get("value") or {}
            if not add.get("x"):
                print(f"[WebNewBill] 顶栏「新增」不可点: {add}")
                return False
            cdp.click(add["x"], add["y"])
            print("[WebNewBill] 已点「新增」")

            #② 等顶部横幅「内容已修改」出现（当前单没填过内容时金蝶不弹，跳过点否）
            no = {}
            deadline = time.time() + min(timeout_s, 12.0)
            while time.time() < deadline:
                time.sleep(0.6)
                no = cdp.js(_JS_DISCARD_NO).get("value") or {}
                if no.get("x"):
                    break
                if no.get("none") and time.time() > deadline - 6.0:
                    break
            if no.get("x"):
                cdp.click(no["x"], no["y"])
                print("[WebNewBill] 已点横幅「否」（放弃当前未保存单）")
            else:
                print("[WebNewBill] 无「内容已修改」横幅（当前单未填过），视为已直接开新单")

            #③ 等重置完成：横幅消失 + 唯一草稿 + wait_label 字段可见
            #（点否后 pid 复用不变，金蝶把同一页签重置成新单，申请人等字段清空）
            deadline = time.time() + timeout_s
            last = ""
            while time.time() < deadline:
                time.sleep(1.0)
                gone = (cdp.js(_JS_DISCARD_NO).get("value") or {}).get("none")
                pid, err = _draft(cdp, "WebNewBill")
                if not pid:
                    last = err or "无草稿"
                    continue
                hv = cdp.js(_JS_HEADER_VISIBLE % (json.dumps(pid), json.dumps(header_key))).get("value") or {}
                if gone and hv.get("ok"):
                    print(f"[WebNewBill] 新单就位 pid={pid[:8]}")
                    return True
                last = "横幅未消失" if not gone else "字段未出现"
            print(f"[WebNewBill] {timeout_s:.0f}s 内新单没就位（最后: {last}）")
            return False
        finally:
            cdp.close()


@AgentServer.custom_action("WebSetCell")  #明细格子写值（F7 编码 / 普通列直写）
class WebSetCell(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        spec = _ENTRY_FIELDS.get(field)
        if not spec:
            print(f"[WebSetCell] 未知field: {field!r}，可选: {sorted(_ENTRY_FIELDS)}")
            return False
        key, kind = spec
        row = int(param.get("row", 0) or 0)
        value = param.get("value")
        if value is None or str(value).strip() == "":
            value = _row_value(field, row)
        value = str(value or "").strip()
        if value == "":
            print(f"[WebSetCell] {field} 第{row}行为空 → 本行无需填这列，跳过")
            return True  #空值语义与 UiaSetCell 一致
        if kind == "flex":
            cdp = _open("WebSetCell")
            if cdp is None:
                return False
            try:
                pid, err = _draft(cdp, "WebSetCell")
                if not pid:
                    print(f"[WebSetCell] {err}")
                    return False
                passed, actual = _fill_flex(cdp, pid, key, row, value, "WebSetCell")
                if not passed:
                    print(f"[WebSetCell] 仓位弹性域填写失败 {field}(row {row}) ← {value!r}: {actual}")
                    return False
                print(f"[WebSetCell] {field}(row {row}) ← {value!r} 显示层 {actual!r} 已提交（维度弹窗链路）")
                return True
            finally:
                cdp.close()
        f7 = (kind is True)

        cdp = _open("WebSetCell")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebSetCell")
            if not pid:
                print(f"[WebSetCell] {err}")
                return False
            ok, e = _type_into_cell(cdp, pid, key, row, value, f7, "WebSetCell")
            if not ok:
                print(f"[WebSetCell] {field}(row {row}) 输入失败: {e}")
                return False
            passed, actual = _verify_cell(cdp, pid, key, row, value, f7, "WebSetCell")
            if not passed:
                hint = ("该编码可能不属于本行的库存组织（金蝶仓库/仓位按组织过滤，"
                        "跨组织调拨时调出仓库必须属于调出组织）"
                        if f7 and key in ("FSTOCKID", "FSTOCKINID") else
                        "金蝶没解析这个编码/名称，或格子没提交")
                print(f"[WebSetCell] 校验不过 {field}(row {row}) ← {value!r} 显示层读到 {actual!r}（{hint}）")
                return False
            print(f"[WebSetCell] {field}(row {row}) ← {value!r} 显示层 {actual!r} 已提交")
            return True
        finally:
            cdp.close()


@AgentServer.custom_action("WebFillColumn")  #整列逐行写（broadcast=整单一个值铺满每行）
class WebFillColumn(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        spec = _ENTRY_FIELDS.get(field)
        if not spec:
            print(f"[WebFillColumn] 未知field: {field!r}，可选: {sorted(_ENTRY_FIELDS)}")
            return False
        key, kind = spec
        f7 = (kind is True)
        flex = (kind == "flex")  #仓位弹性域：逐行走维度弹窗链路

        try:
            import transfer_actions
            order = transfer_actions._current_order()
        except Exception:
            order = None
        if order is None or not order.items:
            print("[WebFillColumn] 队列没就绪（ReadTransferExcel/NextOrder 跑过了吗）")
            return False
        getter = uia_actions._COLUMN_ATTR_BY_FIELD.get(field)
        if getter is None:
            print(f"[WebFillColumn] field {field!r} 不在 Excel 契约表")
            return False
        if param.get("broadcast"):
            head = str(getter(order.items[0]) or "").strip()
            values = [head] * len(order.items)  #整单同值铺每行（申请单组织列就是这个用法）
        else:
            values = [str(getter(it) or "").strip() for it in order.items]

        cdp = _open("WebFillColumn")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebFillColumn")
            if not pid:
                print(f"[WebFillColumn] {err}")
                return False
            need = sum(1 for x in values if x != "")
            if not _ensure_rows(cdp, pid, need, "WebFillColumn"):
                return False
            written = 0
            for i, value in enumerate(values):
                if value == "":
                    continue  #Excel 该行为空 → 不动这格（同 UiaFillColumn）
                if flex:
                    passed, actual = _fill_flex(cdp, pid, key, i, value, "WebFillColumn")
                    if not passed:
                        print(f"[WebFillColumn] {field} 第{i}行仓位填写失败 ← {value!r}: {actual}")
                        return False
                    written += 1
                    continue
                ok, e = _type_into_cell(cdp, pid, key, i, value, f7, "WebFillColumn")
                if not ok:
                    print(f"[WebFillColumn] {field} 第{i}行输入失败: {e}")
                    return False
                passed, actual = _verify_cell(cdp, pid, key, i, value, f7, "WebFillColumn")
                if not passed:
                    print(f"[WebFillColumn] {field} 第{i}行校验不过 ← {value!r} 显示层 {actual!r}")
                    return False
                written += 1
            print(f"[WebFillColumn] {field} 已铺 {written} 行（{'维度弹窗链路' if flex else '真实输入链'}提交）")
            return True
        finally:
            cdp.close()


@AgentServer.custom_action("WebSetHeaderField")  #单头写值：文本/日期真实输入链，基础资料走选择弹窗
class WebSetHeaderField(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        value = param.get("value")
        if value is None or str(value).strip() == "":
            value = _order_value(field)
        value = str(value or "").strip()
        if value == "":
            print(f"[WebSetHeaderField] {field} 值为空 → 不动，保留金蝶默认（同 UIA 空值语义）")
            return True

        cdp = _open("WebSetHeaderField")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebSetHeaderField")
            if not pid:
                print(f"[WebSetHeaderField] {err}")
                return False

            if field in _HEADER_TEXT_FIELDS:
                key, _kind = _HEADER_TEXT_FIELDS[field]
                #2 轮重试：网络横幅会吞掉 Tab 失焦提交（回读没过就是自己被吞了），清横幅后重跑本步
                for rnd in range(2):
                    if self._set_text_once(cdp, pid, key, field, value):
                        return True
                    if rnd == 0:
                        print(f"[WebSetHeaderField] {field} 第1次没落地，清横幅 2s 后重跑")
                        _clear_windows(cdp)
                        time.sleep(2.0)
                return False

            if field in _HEADER_F7_FIELDS:
                key = _HEADER_F7_FIELDS[field]
                for rnd in range(2):
                    if self._pick_f7(cdp, pid, field, key, value):
                        return True
                    if rnd == 0:
                        print(f"[WebSetHeaderField] {field} 路A/B 第1次没落地，清横幅 2s 后重跑")
                        _clear_windows(cdp)
                        time.sleep(2.0)
                return False

            if field in _HEADER_COMBO_FIELDS:
                print(f"[WebSetHeaderField] {field} 是下拉，请配 WebPickCombo")
                return False
            print(f"[WebSetHeaderField] 未知field: {field!r}，可选: "
                  f"{sorted(set(_HEADER_TEXT_FIELDS) | set(_HEADER_F7_FIELDS) | set(_HEADER_COMBO_FIELDS))}")
            return False
        finally:
            cdp.close()

    def _set_text_once(self, cdp, pid, key, field, value):
        #文本单头字段真实输入链写一次：聚焦→键入→Tab 失焦提交→回读校验
        f = cdp.js(_JS_FOCUS_INPUT % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if not f.get("ok"):
            print(f"[WebSetHeaderField] {field}({key}) 聚焦失败: {f.get('err')}")
            return False
        cdp.type_text(value)
        cdp.key_tab()  #Tab 失焦提交 —— 只有真实输入链会进金蝶模型
        time.sleep(0.4)
        back = (cdp.js(_JS_READ_INPUT % (json.dumps(pid), json.dumps(key))).get("value") or {}).get("value") or ""
        digits = lambda x: re.sub(r"\D", "", str(x))
        same = (_norm(back) == _norm(value)) or (digits(back) and digits(back) == digits(value))
        print(f"[WebSetHeaderField] {field} ← {value!r} 回读 {back!r} {'OK' if same else '不符'}")
        return bool(same)

    def _pick_f7(self, cdp, pid, field, key, target):
        #基础资料写值，两条路按顺序试：
        #  路A（优先）：聚焦编辑器→清空→键入→Enter，金蝶自己精确匹配解析编码/名称
        #  路B（兜底）：点放大镜开金蝶选择弹窗→关键字框过滤→点命中行→「返回数据」
        #判据统一用 valData（金蝶控件的已解析值 {name,number}），绝不选第一行
        def parsed_ok(cur):
            return (_norm(cur.get("name")) == _norm(target)
                    or str(cur.get("number")) == str(target)
                    or (_norm(target) and _norm(target) in _norm(cur.get("name"))))

        cur = cdp.js(_JS_F7_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if parsed_ok(cur):
            print(f"[WebSetHeaderField] {field} 现值已是 {target!r}，跳过")
            return True

        # ---- 路A：键入 + Enter ----
        f = cdp.js(_JS_FOCUS_INPUT % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if f.get("ok"):
            #清空：F7 框里 Ctrl+A 不生效，用三击建立选区再 Delete
            cdp.triple_click(f["x"], f["y"])
            time.sleep(0.2)
            cdp.key_delete()
            time.sleep(0.2)
            if _CJK_RE.search(target):
                #中文名称走逐字符键入→typeahead 模糊下拉→唯一命中 Enter 直接选中
                #（insertText 不触发 typeahead，申请人/仓库中文名都靠这条）
                cdp.type_chars(target)
                time.sleep(1.5)  #等模糊下拉过滤稳定
            else:
                cdp.type_text(target)
                time.sleep(0.3)
            cdp.key_enter2()
            time.sleep(2.0)
            #申请组织=主业务组织：变更后金蝶弹横幅要确认，确认后同页签被重置成新单
            #（明细全清）→ 本字段必须排在一切填写之前
            if _confirm_orgchange(cdp, "WebSetHeaderField"):
                pid2, _ = _draft(cdp, "WebSetHeaderField")
                cur2 = cdp.js(_JS_F7_READ % (json.dumps(pid2 or pid), json.dumps(key))).get("value") or {}
                if parsed_ok(cur2):
                    print(f"[WebSetHeaderField] {field} 路A+组织变更确认成功 valData={cur2}（页签已重置，后续动作重取 pid）")
                    return True
                print(f"[WebSetHeaderField] {field} 组织已确认变更但回读不符 valData={cur2}")
                return False
            cur2 = cdp.js(_JS_F7_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
            if parsed_ok(cur2):
                print(f"[WebSetHeaderField] {field} 路A 键入解析成功 valData={cur2}")
                return True
            print(f"[WebSetHeaderField] {field} 路A 未解析（valData={cur2}），转路B 弹窗")

        # ---- 路B：放大镜弹窗 ----
        mag = cdp.js(_JS_F7_MAGNET % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if not mag.get("x"):
            print(f"[WebSetHeaderField] {field} 放大镜不可用: {mag.get('err')}")
            return False
        cdp.click(mag["x"], mag["y"])
        print(f"[WebSetHeaderField] {field} 已点放大镜，轮询等选择弹窗…")

        #弹窗是子表单异步加载，固定 sleep 不够会误判 no-popup → 轮询最多 10s
        filt = {}
        for _ in range(10):
            time.sleep(1.0)
            filt = cdp.js(_JS_POPUP_FILTER).get("value") or {}
            if filt.get("x") or filt.get("noFilter"):
                break
        if filt.get("x"):
            cdp.click(filt["x"], filt["y"])
            time.sleep(0.3)
            cdp.type_text(target)
            cdp.key_enter2()  #快捷过滤回车触发服务端查询（弹窗内真实输入，不碰用户键鼠）
            for _ in range(6):
                time.sleep(1.0)
                row = cdp.js(_JS_POPUP_ROW % json.dumps(target, ensure_ascii=False)).get("value") or {}
                if row.get("x"):
                    break
            else:
                row = cdp.js(_JS_POPUP_ROW % json.dumps(target, ensure_ascii=False)).get("value") or {}
        else:
            print(f"[WebSetHeaderField] {field} 弹窗无过滤框（{filt}），直接扫可见行")
            row = cdp.js(_JS_POPUP_ROW % json.dumps(target, ensure_ascii=False)).get("value") or {}

        if not row.get("x"):
            cdp.js(_JS_POPUP_CLOSE)
            print(f"[WebSetHeaderField] {field} 弹窗里找不到「{target}」行（{row}）→ 关窗返回 False，绝不选第一行")
            return False
        cdp.click(row["x"], row["y"])
        time.sleep(1.0)

        back = cdp.js(_JS_POPUP_BACK).get("value") or {}
        if not back.get("x"):
            cdp.js(_JS_POPUP_CLOSE)
            print(f"[WebSetHeaderField] {field} 弹窗没有「返回数据」按钮: {back}")
            return False
        cdp.click(back["x"], back["y"])
        time.sleep(2.0)

        cur3 = cdp.js(_JS_F7_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
        ok = parsed_ok(cur3)
        print(f"[WebSetHeaderField] {field} 路B 选择后 valData={cur3} {'OK' if ok else '校验不符'}")
        return bool(ok)


@AgentServer.custom_action("WebPickCombo")  #下拉：试值读文本等值匹配（零硬编码业务词）
class WebPickCombo(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        key = _HEADER_COMBO_FIELDS.get(field)
        if not key:
            print(f"[WebPickCombo] 未知field: {field!r}，可选: {sorted(_HEADER_COMBO_FIELDS)}")
            return False
        value = param.get("value")
        if value is None or str(value).strip() == "":
            value = _order_value(field)
        value = str(value or "").strip()
        if value == "":
            print(f"[WebPickCombo] {field} 值为空 → 保留金蝶默认，返回True")
            return True  #同 UiaPickCombo：留空=不动

        cdp = _open("WebPickCombo")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebPickCombo")
            if not pid:
                print(f"[WebPickCombo] {err}")
                return False
            #⚠️下拉必须走【真实鼠标选值】：w.value()+fireChange 只改显示层，
            #  网络横幅一吞 billModel 就不进（报成功却没生效）。
            #点 .k-select 弹层 → 轮询目标 li → 真点 → 回读 widget.text 校验，整段重试 2 轮，
            #轮间清网络横幅（点确定）再重跑，超出两轮才报错停下。
            ok = False
            detail = ""
            for rnd in range(2):
                ok, detail = self._pick_combo_real(cdp, pid, key, value, "WebPickCombo")
                if ok:
                    break
                if rnd == 0:
                    print(f"[WebPickCombo] {field} 第1次失败({detail})，清横幅 2s 后重跑")
                    _dismiss_net_banner(cdp)
                    _clear_windows(cdp)
                    time.sleep(2.0)
            if not ok:
                print(f"[WebPickCombo] {field} ← {value!r} 两轮都没落地: {detail}"
                      f"（选项文本必须与账套一字不差，宁可停下）")
                return False
            print(f"[WebPickCombo] {field} = {value!r}（{detail}）")
            return True
        finally:
            cdp.close()

    def _pick_combo_real(self, cdp, pid, key, target, tag):
        #真实鼠标下拉选值：点 .k-select → 轮询含 target 的可见 li → 真点 → 回读校验。
        #返回 (ok, detail)。回读用 widget.text（显示层）；只有真实鼠标操作才会进模型
        cur = cdp.js(_JS_COMBO_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if cur.get("text", "").strip() == target:
            return True, "编码已是"

        trig = cdp.js(_JS_COMBO_TRIGGER % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if not trig.get("x"):
            return False, f"找不到下拉箭头 {trig}"
        cdp.click(trig["x"], trig["y"])

        opt = {}
        deadline = time.time() + 3.0
        while time.time() < deadline:
            time.sleep(0.4)
            opt = cdp.js(_JS_COMBO_OPTIONS % json.dumps(target, ensure_ascii=False)).get("value") or {}
            if opt.get("x"):
                break
        if not opt.get("x"):
            cdp.key("Escape", "Escape", 27)
            return False, f"弹层里没找到选项 {target!r}: {opt}"
        cdp.click(opt["x"], opt["y"])
        time.sleep(1.5)

        #选完立刻查网络横幅：服务器提交是异步的，横幅一出现就说明本次提交被吞，
        #widget.text 可能已先改成目标值（假阳性），不能只信回读 → 清横幅并让外层重跑本步
        if _dismiss_net_banner(cdp, to=1.0):
            back0 = cdp.js(_JS_COMBO_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
            return False, f"选值提交被网络横幅吞（回读 {back0!r} 不可信），重跑"

        back = cdp.js(_JS_COMBO_READ % (json.dumps(pid), json.dumps(key))).get("value") or {}
        if back.get("text", "").strip() == target:
            return True, f"编码 {back.get('code')} 已选"
        return False, f"选后回读 {back!r} 不是 {target!r}"


@AgentServer.custom_action("WebReadCell")  #回读明细格子（判据同 UiaReadCell）
class WebReadCell(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        spec = _ENTRY_FIELDS.get(field)
        if not spec:
            print(f"[WebReadCell] 未知field: {field!r}")
            return False
        key, kind = spec
        row = int(param.get("row", 0) or 0)

        cdp = _open("WebReadCell")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebReadCell")
            if not pid:
                print(f"[WebReadCell] {err}")
                return False
            #读显示层（billModel 口径）：F7 列金蝶没解析成功时是空或 select
            actual, err = _cell_text(cdp, pid, key, row, "WebReadCell")
            if err:
                print(f"[WebReadCell] {err}")
                return False
            print(f"[WebReadCell] {field}(row {row}) = {actual!r}")
            not_expect = str(param.get("not_expect", "") or "")
            if not_expect:
                return actual != not_expect
            if param.get("nonzero"):
                digits = re.sub(r"\D", "", actual)
                return bool(digits) and set(digits) != {"0"}
            expect = str(param.get("expect", "") or "")
            if expect:
                return expect in actual
            return bool(actual) and "select" not in actual  #没给判据 = 非空即命中
        finally:
            cdp.close()


@AgentServer.custom_recognition("WebRowCount")  #界面已填物料行数 vs Excel 队列行数（对账）
class WebRowCount(CustomRecognition):

    def analyze(
            self,
            context: Context,
            argv: CustomRecognition.AnalyzeArg,
    ) -> "CustomRecognition.AnalyzeResult":

        param = _params(argv, attr="custom_recognition_param", tag="WebRowCount")
        if param is None:
            return None
        try:
            import transfer_actions
            order = transfer_actions._current_order()
            queue_rows = len(order.items) if order else None
        except Exception:
            queue_rows = None
        if queue_rows is None:
            print("[WebRowCount] 队列没就绪")
            return None

        cdp = _open("WebRowCount")
        if cdp is None:
            return None
        try:
            pid, err = _draft(cdp, "WebRowCount")
            if not pid:
                print(f"[WebRowCount] {err}")
                return None
            r = cdp.js(_JS_MATERIAL_COUNT % json.dumps(pid))
            v = r.get("value") if "value" in r else r
            ui_rows = (v or {}).get("count")
            if ui_rows is None:
                print(f"[WebRowCount] 数行失败: {v}")
                return None
            equal = ui_rows == queue_rows
            if bool(param.get("inverse")):
                equal = not equal
            if not equal:
                print(f"[WebRowCount] 界面 {ui_rows} 行 ≠ 队列 {queue_rows} 行，不命中")
                return None
            return CustomRecognition.AnalyzeResult(box=(0, 0, 1, 1),
                                                   detail={"ui": ui_rows, "queue": queue_rows})
        finally:
            cdp.close()


@AgentServer.custom_action("WebPatchFill")  #整单同值列一次铺开：源格填值→单击选中→批量填充按钮铺满整列
class WebPatchFill(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        spec = _ENTRY_FIELDS.get(field)
        if not spec:
            print(f"[WebPatchFill] 未知field: {field!r}，可选: {sorted(_ENTRY_FIELDS)}")
            return False
        key, kind = spec
        srow = int(param.get("row", 0) or 0)
        value = param.get("value")
        if value is None or str(value).strip() == "":
            value = _row_value(field, srow)
        value = str(value or "").strip()
        if value == "":
            print(f"[WebPatchFill] {field} 值为空 → 整列留空，跳过")
            return True  #空值语义与 UiaSetCell 一致

        cdp = _open("WebPatchFill")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebPatchFill")
            if not pid:
                print(f"[WebPatchFill] {err}")
                return False

            #① 源格填值前先判镜像锁定列：组织内调拨时调入组织=调出组织镜像，格子永久 disabled。
            #硬填会双击失败、焦点残留旧编辑器 → Ctrl+A 误选整页 → 弹保存横幅 → 键入落空。
            #金蝶自己已铺，这里只校验不填
            dis = cdp.js(_JS_CELL_DISABLED % (json.dumps(pid), json.dumps(key), str(srow)))
            dis = dis.get("value") if "value" in dis else dis
            if isinstance(dis, dict) and dis.get("disabled"):
                col = cdp.js(_JS_COL_TEXTS % (json.dumps(pid), json.dumps(key)))
                col = col.get("value") if "value" in col else col
                texts = (col or {}).get("texts") or []
                try:
                    import transfer_actions
                    order = transfer_actions._current_order()
                    expect = len(order.items) if order else len(texts)
                except Exception:
                    expect = len(texts)
                filled = sum(1 for x in texts[:expect] if x and "select" not in x)
                if expect and filled >= expect:
                    print(f"[WebPatchFill] {field} 是金蝶镜像锁定列（组织内调拨自动同步），已铺 {filled}/{expect}，跳过填充")
                    return True
                print(f"[WebPatchFill] {field} 镜像锁定列但未铺满 {filled}/{expect}（调拨类型与组织不符？）")
                return False

            #① 源格填值：F7/普通列走真实输入链，弹性域（仓位）走维度弹窗
            #重试一次：紧跟批量填充之后时，网格可能还在异步收尾（setBusy），第一次输入会落空。
            detail = None
            for attempt in range(2):
                if kind == "flex":
                    ok, detail = _fill_flex(cdp, pid, key, srow, value, "WebPatchFill")
                else:
                    ok, detail = _type_into_cell(cdp, pid, key, srow, value, (kind is True), "WebPatchFill")
                    if ok:
                        #⚠️_type_into_cell 只管敲键不验落地；源格必须真有值才配当批量填充的源，
                        #  否则填充拿到空格铺 0/N
                        ok, detail = _verify_cell(cdp, pid, key, srow, value, (kind is True), "WebPatchFill")
                if ok:
                    break
                if attempt == 0:
                    print(f"[WebPatchFill] {field} 源格第1次失败({detail})，等 2s 重试")
                    _clear_windows(cdp)
                    time.sleep(2.0)
            if not ok:
                print(f"[WebPatchFill] 源格填值失败 {field} ← {value!r}: {detail}")
                return False

            #②③④ 单击选中 → 点「批量填充」 → 轮询校验，整段连做 2 轮。
            #（金蝶按钮只认真实鼠标事件，JS .click() 无效；填充源值由服务器从 current 行该列读，
            #  所以必须先单击让 current 落在目标列）
            #整段重试的原因：抖动可能出在单击被吞 / current 漂移 / 填充回包慢，只重试源格填值盖不住
            try:
                import transfer_actions
                order = transfer_actions._current_order()
                expect = len(order.items) if order else 0
            except Exception:
                expect = 0
            filled = 0
            texts = []
            for rnd in range(2):
                xy = cdp.js(_JS_CELL_XY % (json.dumps(pid), json.dumps(key), str(srow)))
                xy = xy.get("value") if "value" in xy else xy
                if not (isinstance(xy, dict) and xy.get("x")):
                    print(f"[WebPatchFill] {field} 第{rnd + 1}轮源格定位失败: {xy}")
                    if rnd == 0:
                        _clear_windows(cdp)
                        time.sleep(2.0)
                    continue
                cdp.click(xy["x"], xy["y"])
                time.sleep(0.6)
                cdp.key("Escape", "Escape", 27)  #退编辑态但保留选中
                time.sleep(0.4)

                #current 落在别的列就再单击一次源格（漂了硬点填充会铺空/铺错列）
                cur = cdp.js(_JS_GRID_CURRENT % json.dumps(pid))
                cur = cur.get("value") if "value" in cur else cur
                if isinstance(cur, dict) and cur.get("field") and cur["field"] != key:
                    print(f"[WebPatchFill] {field} current 漂到 {cur['field']}，重单击源格")
                    cdp.click(xy["x"], xy["y"])
                    time.sleep(0.6)
                    cdp.key("Escape", "Escape", 27)
                    time.sleep(0.4)

                bf = cdp.js(_JS_PATCH_FILL_BTN % json.dumps(pid))
                bf = bf.get("value") if "value" in bf else bf
                if not (isinstance(bf, dict) and bf.get("x")):
                    print(f"[WebPatchFill] {field} 第{rnd + 1}轮批量填充按钮不可点: {bf}")
                    if rnd == 0:
                        _clear_windows(cdp)
                        time.sleep(2.0)
                    continue
                cdp.click(bf["x"], bf["y"])

                #校验轮询 8s：批量填充是服务端回包，读一次就下结论会把"还没回来"当成空
                #判据用「≥队列行数」而非等于 DOM 行数：粘完末尾可能多一行空行
                #⚠️比对基准是【源格显示文本】而不是"非空"：金蝶组织内模式会把调入组织镜像成调出组织，
                #  铺进去的是错值但非空，"非空计数"会假阳性报 7/7
                src, serr = _cell_text(cdp, pid, key, srow, "WebPatchFill")
                src = (src or "").strip()
                comparable = bool(src) and "select" not in src
                filled = 0
                texts = []
                deadline = time.time() + 8.0
                while time.time() < deadline:
                    col = cdp.js(_JS_COL_TEXTS % (json.dumps(pid), json.dumps(key)))
                    col = col.get("value") if "value" in col else col
                    texts = (col or {}).get("texts") or []
                    if comparable:
                        filled = sum(1 for x in texts[:expect] if x.strip() == src)
                    elif expect:
                        filled = sum(1 for x in texts[:expect] if x and "select" not in x)
                    else:  #队列行数未知（transfer_actions 不可用）：只判「列里至少有一个非空值」
                        filled = sum(1 for x in texts if x and "select" not in x)
                    if expect and filled >= expect:
                        break
                    if not expect and filled >= 1:
                        break
                    time.sleep(0.8)
                if expect and filled >= expect:
                    print(f"[WebPatchFill] {field} 已铺开 {filled}/{expect} 行"
                          f"（批量填充，DOM {len(texts)} 行，第{rnd + 1}轮）")
                    return True
                if not expect and filled >= 1:  #拿不到队列行数：非空即算铺开，不做两轮
                    print(f"[WebPatchFill] {field} 已铺开（队列行数未知，非空 {filled} 行，DOM {len(texts)} 行）")
                    return True
                if comparable:
                    #源格文本对不上 = 金蝶把值改写了（镜像/外键拒绝），打印明细便于定位
                    print(f"[WebPatchFill] {field} 第{rnd + 1}轮铺开 {filled}/{expect}，"
                          f"源格={src!r} 各行={texts[:expect]}（与源格不一致，疑似金蝶改值）")
                else:
                    print(f"[WebPatchFill] {field} 第{rnd + 1}轮铺开 {filled}/{expect}"
                          f"（源格文本不可比: {serr or repr(src)}），清场重走选中→填充")
                _clear_windows(cdp)
                time.sleep(2.0)
            print(f"[WebPatchFill] {field} 写入 {filled}/{expect} 行，两轮批量填充未铺开，"
                  f"界面行数 {len(texts)}（行数不对？先跑 WebRowCount 分流）")
            return False
        finally:
            cdp.close()


@AgentServer.custom_action("WebBlockPaste")  #整列块粘贴：右键首格→块粘贴→弹窗textarea灌整列TSV→点确定（几百物料一次填完）
class WebBlockPaste(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        param = _params(argv)
        if param is None:
            return False
        field = param.get("field")
        spec = _ENTRY_FIELDS.get(field)
        if not spec:
            print(f"[WebBlockPaste] 未知field: {field!r}，可选: {sorted(_ENTRY_FIELDS)}")
            return False
        key, kind = spec
        if kind == "flex":
            print(f"[WebBlockPaste] {field} 是弹性域（仓位），块粘贴不适用，请用 WebPatchFill")
            return False

        #整列值：优先 param.values，否则按 Excel 队列取（broadcast=整单同值铺每行）
        try:
            import transfer_actions
            order = transfer_actions._current_order()
        except Exception:
            order = None
        if order is None or not order.items:
            print("[WebBlockPaste] 队列没就绪（ReadTransferExcel/NextOrder 跑过了吗）")
            return False
        values = param.get("values")
        if values is None:
            getter = uia_actions._COLUMN_ATTR_BY_FIELD.get(field)
            if getter is None:
                print(f"[WebBlockPaste] field {field!r} 不在 Excel 契约表")
                return False
            if param.get("broadcast"):
                head = str(getter(order.items[0]) or "").strip()
                values = [head] * len(order.items)
            else:
                values = [str(getter(it) or "").strip() for it in order.items]
        values = [str(v or "").strip() for v in values]
        nonempty = [v for v in values if v]
        if not nonempty:
            print(f"[WebBlockPaste] {field} 整列为空 → 跳过")
            return True

        cdp = _open("WebBlockPaste")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebBlockPaste")
            if not pid:
                print(f"[WebBlockPaste] {err}")
                return False
            #新草稿打开瞬间 dataSource 可能是 0 行（物料空 → 仓库/仓位被金蝶拒绝批量填充，连锁整单失败）。
            #等初始行出现，6s 不出点一次「新增行」兜底；只补 1 行，块粘贴自动扩到 N 行（勿预补 N 行）
            for _ in range(12):
                r = cdp.js(_JS_ROW_COUNT % json.dumps(pid))
                v = r.get("value") if "value" in r else r
                if ((v or {}).get("rows") or 0) >= 1:
                    break
                time.sleep(0.5)
            else:
                a = cdp.js(_JS_ADD_ROW % json.dumps(pid))
                av = a.get("value") if "value" in a else a
                if isinstance(av, dict) and av.get("x"):
                    cdp.click(av["x"], av["y"])
                    time.sleep(1.0)
                    print("[WebBlockPaste] 草稿 0 行，已点一次「新增行」再粘贴")
            xy = cdp.js(_JS_CELL_XY % (json.dumps(pid), json.dumps(key), "0"))
            xy = xy.get("value") if "value" in xy else xy
            if not (isinstance(xy, dict) and xy.get("x")):
                print(f"[WebBlockPaste] 首格定位失败: {xy}")
                return False
            #不补行：金蝶块粘贴会自动扩行（1 行草稿粘 7 值 → 8 行），
            #几百物料一次粘贴即可；若先 _ensure_rows 反而退化成几百次点「新增行」

            #① 右键首格 → 点「块粘贴」（菜单异步，重试两轮）
            dlg = {}
            for _ in range(2):
                for typ in ("mousePressed", "mouseReleased"):
                    cdp._cmd("Input.dispatchMouseEvent", {"type": typ, "x": xy["x"], "y": xy["y"],
                                                          "button": "right", "clickCount": 1})
                item = {}
                for _ in range(8):
                    time.sleep(0.5)
                    r = cdp.js(_JS_BLOCKPASTE_MENU)
                    item = r.get("value") if "value" in r else r
                    if isinstance(item, dict) and item.get("x"):
                        break
                if isinstance(item, dict) and item.get("x"):
                    cdp.click(item["x"], item["y"])
                    for _ in range(8):
                        time.sleep(0.5)
                        r = cdp.js(_JS_BLOCKPASTE_DIALOG)
                        dlg = r.get("value") if "value" in r else r
                        if isinstance(dlg, dict) and dlg.get("tx") and dlg.get("okx"):
                            break
                    if isinstance(dlg, dict) and dlg.get("tx"):
                        break
                cdp.key("Escape", "Escape", 27)
                time.sleep(0.5)
            if not (isinstance(dlg, dict) and dlg.get("tx") and dlg.get("okx")):
                print(f"[WebBlockPaste] 块粘贴弹窗未就绪: {dlg}")
                _clear_windows(cdp)
                return False

            #② 灌整列 TSV（textarea 是普通输入框，insertText 有效；Enter 只换行，必须点确定）
            cdp.click(dlg["tx"], dlg["ty"])
            time.sleep(0.4)
            cdp.type_text("\n".join(values))
            time.sleep(0.5)
            cdp.click(dlg["okx"], dlg["oky"])
            print(f"[WebBlockPaste] {field} 已粘贴 {len(nonempty)} 行并提交，等金蝶解析…")

            #③ 校验：物料列看名称列被带出，其他列看整列非空非 select。
            #⚠️判据用「文本含任一期望值」而不是「非空非 select」：F7 首行渲染成 select+编码 混排，按非 select 判会把已填上的首行当成空。
            read_key = "FMATERIALNAME" if key == "FMATERIALID" else key
            deadline = time.time() + 15.0
            texts = []
            while time.time() < deadline:
                r = cdp.js(_JS_COL_TEXTS % (json.dumps(pid), json.dumps(read_key)))
                col = r.get("value") if "value" in r else r
                texts = (col or {}).get("texts") or []
                if key == "FMATERIALID":
                    filled = sum(1 for x in texts if x and "select" not in x)
                else:
                    filled = sum(1 for x in texts if x and any(v in x for v in nonempty))
                if filled >= len(nonempty):
                    print(f"[WebBlockPaste] {field} 块粘贴成功 {filled} 行（一次提交，金蝶自动扩行）")
                    return True
                time.sleep(1.0)
            print(f"[WebBlockPaste] {field} 块粘贴校验不足: 期望{len(nonempty)} 实到{sum(1 for x in texts if x)} texts={texts[:8]}")
            return False
        finally:
            cdp.close()


def _toolbar_step(cdp, pid, label, tag, probe_wait=2.0):
    #顶栏按钮一步走：点按钮 → 有确认弹窗就点确定 → 等 probe_wait 秒读落地信号。
    #⚠️本函数不做任何"测试性"点击：调用它的节点默认不在跑单链路里
    btn = cdp.js(_JS_TOOLBAR_BTN % json.dumps(label, ensure_ascii=False)).get("value") or {}
    if not btn.get("x"):
        return False, f"顶栏「{label}」找不到: {btn}"
    if btn.get("grey"):
        return False, f"顶栏「{label}」是灰的（前置条件没满足，绝不硬点）"
    cdp.click(btn["x"], btn["y"])
    print(f"[{tag}] 已点顶栏「{label}」，等确认弹窗…")
    time.sleep(0.8)
    pop = cdp.js(_JS_CONFIRM_POPUP_OK).get("value") or {}
    if pop.get("x"):
        print(f"[{tag}] 确认弹窗「{pop.get('text', '')[:40]}」→ 点确定")
        cdp.click(pop["x"], pop["y"])
    elif not pop.get("none"):
        print(f"[{tag}] 弹窗里没有确定按钮（{pop.get('text', '')}），继续等落地信号")
    time.sleep(probe_wait)
    return True, None


def _submit_landed(cdp, tag, before_tab):
    #提交落地判据（申请单/下推目标单共用，DRY）：轮询 8s 等异步出号，
    #硬判据=「单据编号」出真值；次判据=页签标题脱离「新增」（金蝶保存/提交后改「- 修改」）。
    #都读不到 → False 停下，绝不假装成功
    end = time.time() + 8.0
    while True:
        num, tab, st = _landing_signals(cdp)
        val = (num.get("value") or "").strip()
        if val and val != (num.get("placeholder") or "").strip():
            print(f"[{tag}] 提交落地：单据编号={val!r}")
            return True
        if tab and tab != before_tab and "新增" not in tab:
            print(f"[{tag}] 提交疑似落地（页签 {before_tab!r}→{tab!r}，编号未读到，请校准判据）")
            return True
        if time.time() >= end:
            return False
        time.sleep(1.0)


def _landing_signals(cdp):
    #提交/审核的落地信号（不赌状态文案，全部打日志）：
    #  ①「单据编号」从占位变成真实编号 ②活动页签标题变化 ③单据状态字段文本
    num = cdp.js(_JS_LABEL_INPUT % json.dumps("单据编号", ensure_ascii=False)).get("value") or {}
    tab = cdp.js(_JS_ACTIVE_TAB).get("value") or ""
    st = cdp.js(_JS_LABEL_INPUT % json.dumps("单据状态", ensure_ascii=False)).get("value") or {}
    print(f"[signals] 单据编号={num.get('value', '')!r} 占位={num.get('placeholder', '')!r} "
          f"页签={tab!r} 状态={st.get('value', '')!r}")
    return num, tab, st


@AgentServer.custom_action("WebSubmit")  #顶栏「提交」：真实鼠标点击 + 确认弹窗 + 落地信号记录
class WebSubmit(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        cdp = _open("WebSubmit")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebSubmit")
            if not pid:
                print(f"[WebSubmit] {err}")
                return False
            tab0 = cdp.js(_JS_ACTIVE_TAB).get("value") or ""
            ok, detail = _toolbar_step(cdp, pid, "提交", "WebSubmit")
            if not ok:
                print(f"[WebSubmit] {detail}")
                return False
            return _submit_landed(cdp, "WebSubmit", tab0)
        finally:
            cdp.close()


@AgentServer.custom_action("WebAudit")  #顶栏「审核」：同 WebSubmit；前置=已提交，灰态绝不硬点
class WebAudit(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        cdp = _open("WebAudit")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebAudit")
            if not pid:
                print(f"[WebAudit] {err}")
                return False
            ok, detail = _toolbar_step(cdp, pid, "审核", "WebAudit")
            if not ok:
                print(f"[WebAudit] {detail}")
                return False
            _landing_signals(cdp)
            #审核落地的判据=「下推」由灰变亮；读不到就停下报错
            dp = cdp.js(_JS_TOOLBAR_BTN % json.dumps("下推", ensure_ascii=False)).get("value") or {}
            if dp.get("x") and not dp.get("grey"):
                print("[WebAudit] 审核落地：「下推」已点亮")
                return True
            print(f"[WebAudit] 点了审核但「下推」仍是灰/找不到（{dp}）→ 停下（首轮真单请校准判据）")
            return False
        finally:
            cdp.close()


@AgentServer.custom_action("WebPushDown")  #顶栏「下推」：前置=已审核（按钮灰则拒点）；落地=新目标单据页签出现
class WebPushDown(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        cdp = _open("WebPushDown")
        if cdp is None:
            return False
        try:
            pid, err = _draft(cdp, "WebPushDown")
            if not pid:
                print(f"[WebPushDown] {err}")
                return False
            before = cdp.js(_JS_TAB_TITLES).get("value") or []
            #先记住来源单页签的完整标题（此刻活动页签就是它），下推后按精确标题切回——
            #不能按「调拨申请单」前缀找：会命中「...列表」或遗留空白「- 新增」
            src_tab = cdp.js(_JS_ACTIVE_TAB).get("value") or ""
            btn = cdp.js(_JS_TOOLBAR_BTN % json.dumps("下推", ensure_ascii=False)).get("value") or {}
            if not btn.get("x"):
                print(f"[WebPushDown] 顶栏「下推」找不到: {btn}")
                return False
            if btn.get("grey"):
                print("[WebPushDown] 「下推」是灰的（审核没成功，绝不硬点）")
                return False
            cdp.click(btn["x"], btn["y"])
            print("[WebPushDown] 已点顶栏「下推」，等「选择单据」弹窗…", flush=True)

            #「选择单据」弹窗里转换规则/单据类型是异步加载的，早于就绪点确定=点了没生效，
            #必须轮询到「确定」可点再点
            ok = {}
            end = time.time() + 12.0
            while time.time() < end:
                time.sleep(0.8)
                ok = cdp.js(_JS_PUSH_OK_BTN).get("value") or {}
                if ok.get("x"):
                    break
            if not ok.get("x"):
                print(f"[WebPushDown] 「选择单据」弹窗没就绪（{ok}）→ 停下")
                return False
            cdp.click(ok["x"], ok["y"])
            print(f"[WebPushDown] 已点弹窗「确定」（{ok.get('text','')[:40]}）", flush=True)

            #下推落地=弹窗关闭 + 冒出一个新顶栏页签（目标单据，如「直接调拨单 - 新增」）。
            #⚠️新页签判据必须等弹窗关掉后再读：弹窗内部「参数选择」tab 会被误当页签
            landed = False
            new = []
            end2 = time.time() + 12.0
            while time.time() < end2:
                time.sleep(1.0)
                still = cdp.js(_JS_PUSH_OK_BTN).get("value") or {}
                after = cdp.js(_JS_TAB_TITLES).get("value") or []
                new = [t for t in after if t not in before]
                if still.get("noPop") and new:
                    landed = True
                    break
            if not landed:
                print(f"[WebPushDown] 点了确定但没落地（弹窗还在={not still.get('noPop')} 新页签={new}）→ 停下")
                return False
            print(f"[WebPushDown] 下推落地：新页签 {new}", flush=True)

            #下推出的目标单（如「分步式调出单 - 新增」）还要点一次「提交」整单才算完成。
            tgt = new[0]
            tsel = cdp.js(_JS_TAB_BY_TEXT % json.dumps(tgt, ensure_ascii=False)).get("value") or {}
            if not tsel.get("x"):
                print(f"[WebPushDown] 找不到下推结果页签 {tgt!r} → 停下")
                return False
            cdp.click(tsel["x"], tsel["y"])
            time.sleep(2.0)
            ok, detail = _toolbar_step(cdp, None, "提交", "WebPushDown", probe_wait=2.5)
            if not ok:
                print(f"[WebPushDown] 目标单提交失败: {detail}")
                return False
            if not _submit_landed(cdp, "WebPushDown·目标单", tgt):
                print("[WebPushDown] 目标单点了提交但没读到落地信号 → 停下")
                return False

            #切回来源单页签再收工：NewBill 点的是【当前页签】顶栏「新增」
            #停在目标单页签上会开出错单据类型。
            #⚠️目标单提交后标题会从「- 新增」变「- 修改」，所以按"提交后仍在活动位"
            #  的实时标题找来源页签，不能拿落地前的旧标题（会 not-found）
            now_tab = cdp.js(_JS_ACTIVE_TAB).get("value") or ""
            if src_tab and now_tab != src_tab:
                #活动位还停在目标单（提交后改名了）：来源页签用记住的 src_tab 精确找
                back = cdp.js(_JS_TAB_BY_TEXT % json.dumps(src_tab, ensure_ascii=False)).get("value") or {}
                if back.get("x"):
                    cdp.click(back["x"], back["y"])
                    time.sleep(1.5)
                    print(f"[WebPushDown] 已切回来源页签「{back.get('text')}」")
                    return True
                print(f"[WebPushDown] ⚠️下推+目标单提交都成功，但切不回来源页签（{back}）→ 停下，防 NewBill 开错单")
                return False
            print("[WebPushDown] 目标单已提交，活动页签即来源单，无需切换")
            return True
        finally:
            cdp.close()


@AgentServer.custom_action("WebStop")  #异常停止：必须让任务判失败（Stop 走 DoNothing 会把失败伪装成"任务完成"）
class WebStop(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        print("[WebStop] 流程走到异常停止节点 → 任务判失败。往前翻日志找第一条 [Web*] 失败原因")
        return False  #节点失败且本节点无 on_error → Tasker 报 Task.Failed，GUI 显示红色失败


@AgentServer.custom_action("WebPing")  #探测 CDP/登录态/页面形态，排查用
class WebPing(CustomAction):

    def run(
            self,
            context: Context,
            argv: CustomAction.RunArg,
    ) -> bool:

        cdp = _open("WebPing")
        if cdp is None:
            return False
        try:
            v = cdp.js(_JS_PING).get("value")
            print(f"[WebPing] {json.dumps(v, ensure_ascii=False)}")
            return bool(v) and not v.get("login")
        finally:
            cdp.close()
