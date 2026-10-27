# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (C) 2026 三扣科技 (3Q Tech)
# 本项目基于 PyMuPDF（Copyright (C) Artifex Software, Inc.，AGPL-3.0 或商业许可），
# 依 AGPL-3.0 第 5 条整体以 AGPL-3.0 授权。完整文本见同目录 LICENSE，
# 第三方声明见 NOTICE。
"""
3Q PDF开源免费工具 · MCP 服务器（pdf_mcp）

把 pdf_cli 的命令表原样暴露成 MCP 工具，零新增第三方依赖：
只用标准库 + pdf_cli（pdf_cli 再复用 pdf_core，处理逻辑仍然只有一份）。

两种传输：
  stdio   —— 子进程方式，AI 客户端在本机拉起。属于「分发」，不触发 AGPL §13。
  http    —— Streamable HTTP，对外提供网络服务。**触发 AGPL-3.0 第 13 条**，
             因此本文件把 §13 义务做进协议结构本身，而不是写在文档里：
             1) initialize 响应的 instructions 字段（标准字段，客户端必然收到）
             2) 每个 tools/call 结果的 structuredContent 里带 _source
             3) 专用工具 pdf_source（拿源码地址 + 依赖组件许可清单）
             4) HTTP 模式内置 GET /source 与 GET /source/<file> 自持源码端点，
                即使不配外部仓库地址也能满足「免费提供对应源码」

用法：
  python pdf_mcp.py                                   # stdio
  python pdf_mcp.py --transport http --port 8765
  python pdf_mcp.py --transport http --port 8765 --source-url https://github.com/you/pdf-tools
  python pdf_mcp.py --transport http --port 8765 --allow-dir D:\\pdf_in --allow-dir D:\\pdf_out
  python pdf_mcp.py --transport http --host 0.0.0.0 --code-only
      # 远程代码服务：只暴露 pdf-code / pdf_source 两个**不碰任何文件**的工具，
      # 远程智能体连上来自己发现工具、按需取可执行代码，到它自己机器上跑。
      # 因为不读写本机任何文件，绑定非回环地址也不需要 --allow-dir。
"""
import argparse
import base64
import hashlib
import io
import json
import os
import socket
import sys
import threading
import time
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pdf_cli
import pdf_core as core

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROTOCOL_VERSION = "2025-06-18"
SERVER_NAME = "pdf-tools-mcp"
SERVER_VERSION = "1.0.0"
LICENSE_ID = "AGPL-3.0-only"
LICENSE_TEXT_URL = "https://www.gnu.org/licenses/agpl-3.0.txt"

# 对外源码地址。为空时 HTTP 模式用自带的 /source 端点自持。
SOURCE_URL = ""
PUBLIC_BASE = ""
ALLOW_DIRS = []          # 路径沙箱；为空 = 不限制（与本地 CLI 行为一致）
SOURCE_TOOL_NAME = "pdf_source"

# 远程代码服务模式：只暴露 pdf-code / pdf_source。
# 这两个工具都是纯代码/信息生成，不读写服务器上的任何文件，
# 所以绑定非回环地址也不会把文件系统对网络开放——
# 「远程智能体接入」档正是靠这一点才能安全地监听局域网地址。
CODE_ONLY = False

# code_only 模式下允许暴露的工具（其余一律拒绝）
CODE_ONLY_TOOLS = frozenset((SOURCE_TOOL_NAME, "pdf-code"))

# 生成脚本对远程端的依赖清单（make_instructions 与 gen_pdf_code 共用）
DEPS = ["pymupdf", "pillow", "openpyxl", "python-docx"]

SECTION_13_NOTICE = (
    "本服务使用 PyMuPDF（GNU AFFERO GPL 3.0）。依 AGPL-3.0 第 13 条"
    "「Remote Network Interaction」，通过网络使用本服务的用户有权免费获取"
    "本服务所运行版本的对应源码。"
)


# ---------- 源码/许可取证（运行时实测，不写死）----------

def _source_files():
    """可以被对外提供下载的源码文件清单（不含 __pycache__/dist/tests 中间物）。"""
    names = []
    for n in sorted(os.listdir(BASE_DIR)):
        if not n.endswith(".py"):
            continue
        if n.startswith("_"):
            continue
        names.append(n)
    for n in ("LICENSE", "NOTICE"):
        p = os.path.join(BASE_DIR, n)
        if os.path.exists(p):
            names.append(n)
    return names


def _sha256_of(rel):
    p = os.path.join(BASE_DIR, rel)
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest(), os.path.getsize(p)


def _probe_components():
    """扫本机已安装的 PyMuPDF 相关包，读出真实许可字段。

    这是取证行为：许可信息从安装的包元数据现读，不是从本文件的注释里抄的。
    """
    rows = []
    try:
        import site
        dirs = []
        for p in site.getsitepackages() + [site.getusersitepackages()]:
            if os.path.isdir(p):
                dirs.append(p)
        # PyInstaller 冻结后 site.getsitepackages() 指向包内并不存在的
        # Lib\site-packages，而 dist-info 实际解包在 _MEIPASS 根目录，
        # 不补这一条，pdf_source 的组件许可清单在成品 exe 里会是空的。
        meipass = getattr(sys, "_MEIPASS", "")
        if meipass and os.path.isdir(meipass) and meipass not in dirs:
            dirs.append(meipass)
        for d in dirs:
            for entry in sorted(os.listdir(d)):
                if not entry.endswith(".dist-info"):
                    continue
                if not entry.lower().startswith(("pymupdf", "mupdf")):
                    continue
                meta = os.path.join(d, entry, "METADATA")
                if not os.path.exists(meta):
                    continue
                name = entry.rsplit(".dist-info", 1)[0]
                version, lic, vendor = "", "", ""
                with open(meta, encoding="utf-8", errors="replace") as f:
                    for line in f:
                        if line.startswith("Name: "):
                            name = line[6:].strip()
                        elif line.startswith("Version: "):
                            version = line[9:].strip()
                        elif line.startswith("License: "):
                            lic = line[9:].strip()
                        elif line.startswith("Author: "):
                            vendor = line[8:].strip()
                        if version and lic and vendor:
                            break
                rows.append({"name": name, "version": version,
                             "license": lic or "(未声明)",
                             "vendor": vendor or "(未声明)"})
    except Exception as e:
        rows.append({"name": "(扫描失败)", "version": "",
                     "license": type(e).__name__, "vendor": str(e)})
    return rows


def local_source_url():
    if PUBLIC_BASE:
        return PUBLIC_BASE.rstrip("/") + "/source"
    if SOURCE_URL:
        return SOURCE_URL
    return ""


def source_ref():
    """每个 tools/call 结果都会带上这个，满足 §13 的「显著提供」。"""
    ref = {
        "license": LICENSE_ID,
        "license_text": LICENSE_TEXT_URL,
        "obligation": SECTION_13_NOTICE,
    }
    u = local_source_url()
    if u:
        ref["source_url"] = u
    return ref


def source_info():
    files = []
    for n in _source_files():
        try:
            sha, size = _sha256_of(n)
            files.append({"file": n, "bytes": size, "sha256": sha})
        except Exception:
            pass
    info = {
        "server": SERVER_NAME,
        "version": SERVER_VERSION,
        "protocol_version": PROTOCOL_VERSION,
        "license": LICENSE_ID,
        "license_text": LICENSE_TEXT_URL,
        "license_file": os.path.join(BASE_DIR, "LICENSE"),
        "notice_file": os.path.join(BASE_DIR, "NOTICE"),
        "obligation": SECTION_13_NOTICE,
        "components": _probe_components(),
        "source_files": files,
    }
    u = local_source_url()
    if u:
        info["source_url"] = u
    return info


# ---------- 工具集（从 pdf_cli.SPECS 数据驱动生成）----------

def _schema_for(o):
    kw = o.get("kw") or {}
    if kw.get("action") == "store_true":
        s = {"type": "boolean"}
    elif kw.get("nargs") == "+":
        s = {"type": "array", "items": {"type": "string"}}
    else:
        t = kw.get("type")
        s = {"type": "integer" if t is int else ("number" if t is float
                                                 else "string")}
    if o.get("enum"):
        s["enum"] = o["enum"]
    s["description"] = o["help"] or o["dest"]
    if o.get("default") is not None:
        s["default"] = o["default"]
    return s


def build_tools():
    tools = [{
        "name": SOURCE_TOOL_NAME,
        "title": "获取源码与许可信息",
        "description": "返回本服务运行版本的源码获取方式、开源许可与第三方组件"
                       "许可清单（AGPL-3.0 第 13 条要求对网络用户显著提供）。",
        "inputSchema": {"type": "object", "properties": {}},
    }, {
        "name": "pdf-code",
        "title": "生成可执行代码，交给远程智能体自己执行",
        "description": "本工具不读写任何文件，也不做实际处理。它按操作意图"
                       "生成一段自包含的 Python 脚本：调用方把它保存到自己"
                       "机器上执行即可，不需要 Windows、不需要安装本程序，"
                       "文件全程不出它的机器。跨机器 / 跨平台场景用这个，"
                       "不要直接调用那些会读写本地文件的同名工具。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "command": {
                    "type": "string",
                    "description": "要执行的操作名，取值见同名工具列表",
                    "enum": sorted(n for n, _h, _o, _f in pdf_cli.SPECS),
                },
                "params": {
                    "type": "object",
                    "description": "该操作的参数，键名与同名工具的参数名一致。"
                                   "例：compress 传 "
                                   "{\"input\": \"a.pdf\", \"target_mb\": 5}",
                },
            },
            "required": ["command"],
        },
    }]
    for name, help_, opts, _fn in pdf_cli.SPECS:
        props = {}
        required = []
        for o in opts:
            props[o["dest"]] = _schema_for(o)
            if o["required"]:
                required.append(o["dest"])
        tools.append({
            "name": name,
            "title": help_,
            "description": help_,
            "inputSchema": {"type": "object", "properties": props,
                            "required": required},
        })
    return tools


TOOLS = build_tools()
TOOL_NAMES = {t["name"] for t in TOOLS}


def active_tools():
    """当前模式对外可见的工具表。

    全量模式 = 全部工具；code_only（远程代码服务）= 只有 pdf-code 与
    pdf_source——这两个工具不读写服务器上的任何文件。
    """
    if not CODE_ONLY:
        return TOOLS
    return [t for t in TOOLS if t["name"] in CODE_ONLY_TOOLS]


def make_instructions(code_only=None):
    if code_only is None:
        code_only = CODE_ONLY
    if code_only:
        return (
            "PDF 文档处理代码服务（远程模式）。本服务不提供任何读写"
            "服务器文件的工具——你的 PDF 在你自己机器上，处理流程是："
            "1) 从 tools/list 里挑操作，调用 pdf-code（参数：command=操作名，"
            "params=该操作参数）；2) 它返回一段自包含 Python 脚本；"
            "3) 把脚本保存到你自己的机器上执行即可。不需要 Windows、"
            "不用装任何额外程序，只要 Python 和 pip install "
            + " ".join(DEPS)
            + "，文件全程不出你的机器。"
            + SECTION_13_NOTICE
            + "源代码：" + (local_source_url() or "见工具 pdf_source 返回值")
            + "；许可证：" + LICENSE_ID
            + "。处理产出的 PDF / 图片等文件不受 AGPL 约束。"
        )
    return (
        "PDF 文档处理。两种用法："
        "一、直接调用与操作同名的工具（merge / compress / crop ...），"
        "会在本机读写文件，适用于你与本服务同机；"
        "二、跨机器或跨平台时改用 pdf-code：它返回一段自包含 Python 脚本，"
        "你保存到自己机器上执行即可，不需要 Windows、不用装本程序，"
        "文件全程不出你的机器。"
        "三、print 是高危操作（真的打印），未显式放行时需要人工确认，"
        "被拒绝时会收到包含「已取消」的错误。"
        + SECTION_13_NOTICE
        + "源代码：" + (local_source_url() or "见工具 pdf_source 返回值")
        + "；许可证：" + LICENSE_ID
        + "。处理产出的 PDF / 图片等文件不受 AGPL 约束。"
    )


INSTRUCTIONS = make_instructions()


# ---------- JSON-RPC 2.0 ----------

class RpcError(Exception):
    def __init__(self, code, message):
        Exception.__init__(self, message)
        self.code = code
        self.message = message


def _res(mid, result):
    return {"jsonrpc": "2.0", "id": mid, "result": result}


def _err(mid, code, message):
    return {"jsonrpc": "2.0", "id": mid,
            "error": {"code": code, "message": message}}


def _allowed(p):
    if not ALLOW_DIRS:
        return True
    ap = os.path.abspath(p)
    for d in ALLOW_DIRS:
        if ap == d or ap.startswith(d + os.sep) or ap.startswith(d + "/"):
            return True
    return False


def _check_sandbox(args):
    for key in ("input", "output", "image"):
        v = args.get(key)
        vals = v if isinstance(v, list) else ([v] if v else [])
        for p in vals:
            if not isinstance(p, str):
                continue
            if not _allowed(p):
                raise RpcError(-32602, "路径不在允许目录内：%s" % p)


# ---------- 高危操作确认闸 ----------

# 会对外界产生不可撤销后果的操作。print 会真的占用纸张与墨水，
# 智能体误调用的代价不是一条错误消息，而是一沓纸。
# 注意 batch 的 op 枚举（BATCH_OPS）里没有 print，所以不用动态判断。
HIGH_RISK_COMMANDS = {"print"}

# 没有确认途径时（stdio / 命令行模式）是否放行。默认 False：
# 宁可让智能体拿到一条能看懂的拒绝，也不要默默执行。
# GUI 的「仅本机」档不设它，改注册 CONFIRM_HOOK 弹人工确认框。
ALLOW_HIGH_RISK = False

# 确认钩子：callable(command, args, reason) -> bool。
# GUI 注册它弹模态框；返回 False、抛异常、超时，一律按拒绝处理。
CONFIRM_HOOK = None

_HIGH_RISK_TEXT = {
    "print": "向物理打印机发送打印任务（纸张与墨水不可撤销）",
}


def high_risk_reason(name, args):
    """高危原因；不高危返回 None。"""
    if name in HIGH_RISK_COMMANDS:
        return _HIGH_RISK_TEXT.get(name, "高危操作：%s" % name)
    return None


def confirm_high_risk(name, args, reason):
    """执行高危命令前的统一闸口。放行返回 True，否则 False。"""
    if ALLOW_HIGH_RISK:
        return True
    hook = CONFIRM_HOOK
    if hook is None:
        return False
    try:
        return bool(hook(name, args, reason))
    except Exception:
        return False        # 钩子坏了倒向拒绝，不倒向放行


def dispatch(req):
    """处理单条 JSON-RPC 请求；通知类返回 None（不回响应）。"""
    if not isinstance(req, dict):
        return _err(None, -32600, "请求不是 JSON 对象")
    mid = req.get("id")
    method = req.get("method")
    params = req.get("params") or {}

    if method == "initialize":
        return _res(mid, {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            "instructions": INSTRUCTIONS,
        })
    if method == "ping":
        return _res(mid, {})
    if method == "tools/list":
        return _res(mid, {"tools": active_tools()})
    if method == "tools/call":
        return _call_tool(mid, params)
    if isinstance(method, str) and method.startswith("notifications/"):
        return None
    if isinstance(method, str) and method.startswith("$/setTrace"):
        return None
    return _err(mid, -32601, "不支持的方法：%s" % method)


def _call_tool(mid, params):
    name = params.get("name")
    args = params.get("arguments") or {}
    if name == SOURCE_TOOL_NAME:
        info = source_info()
        return _res(mid, {
            "content": [{"type": "text",
                         "text": json.dumps(info, ensure_ascii=False)}],
            "structuredContent": info,
            "isError": False,
        })
    if name == "pdf-code":
        r = gen_pdf_code(args.get("command"), args.get("params"))
        if not r.get("ok"):
            return _err(mid, -32602, r.get("error", "生成失败"))
        return _res(mid, {
            "content": [{"type": "text", "text": r["code"]}],
            "structuredContent": r,
            "isError": False,
        })
    if name not in TOOL_NAMES:
        return _err(mid, -32602, "未知工具：%s" % name)
    if CODE_ONLY:
        return _err(
            mid, -32602,
            "本服务运行在远程代码服务模式，只提供 pdf-code（生成可执行代码）"
            "与 pdf_source（源码与许可信息），不提供任何读写服务器文件的"
            "工具。请改调 pdf-code：传 command=%r 与它的参数，返回的脚本"
            "保存到你自己的机器上执行。" % name)
    try:
        _check_sandbox(args)
    except RpcError as e:
        return _err(mid, e.code, e.message)
    reason = high_risk_reason(name, args)
    if reason and not confirm_high_risk(name, args, reason):
        return _err(mid, -32602,
                    "已取消：%s。高危操作未获确认——请人工确认后重试，"
                    "或以 --allow-print 启动本服务显式放行。" % reason)
    res = pdf_cli.execute(name, args)
    res = dict(res)
    res.setdefault("elapsed_ms", 0)
    res["_source"] = source_ref()
    text = json.dumps(res, ensure_ascii=False)
    return _res(mid, {
        "content": [{"type": "text", "text": text}],
        "structuredContent": res,
        "isError": not res.get("ok", False),
    })


def handle_payload(payload):
    """入：解析后的 JSON（对象或数组）。出：响应对象 / 数组 / None。"""
    if isinstance(payload, list):
        outs = [dispatch(x) for x in payload]
        outs = [o for o in outs if o is not None]
        return outs or None
    return dispatch(payload)


# ---------- HTTP 传输 ----------

PAGE_CSS = (
    "body{font:14px/1.7 -apple-system,'Segoe UI','Microsoft YaHei',sans-serif;"
    "margin:0;padding:32px;color:#1c1c1c;background:#fff}"
    "h1{font-size:18px;font-weight:600;margin:0 0 4px}"
    "h2{font-size:14px;font-weight:600;margin:24px 0 8px}"
    "code{background:#f2f2f2;padding:1px 5px;border-radius:3px}"
    "a{color:#0a5aa8}table{border-collapse:collapse;font-size:13px}"
    "td,th{border:1px solid #dcdcdc;padding:5px 10px;text-align:left}"
    "th{background:#f6f6f6}"
)


def _html_source_page():
    rows = []
    for n in _source_files():
        try:
            sha, size = _sha256_of(n)
            rows.append("<tr><td><a href='/source/%s'>%s</a></td>"
                        "<td>%d</td><td><code>%s</code></td></tr>"
                        % (n, n, size, sha[:16]))
        except Exception:
            rows.append("<tr><td>%s</td><td>-</td><td>-</td></tr>" % n)
    comp = []
    for c in _probe_components():
        comp.append("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>"
                    % (c["name"], c["version"], c["license"], c["vendor"]))
    u = local_source_url()
    return (
        "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
        "<title>源代码 · PDF文档处理服务</title><style>%s</style></head><body>"
        "<h1>源代码</h1>"
        "<p>本服务使用 PyMuPDF（AGPL-3.0）。依 <b>AGPL-3.0 第 13 条</b>"
        "「Remote Network Interaction」，下列文件必须<b>免费</b>提供给"
        "所有通过网络使用本服务的用户。</p>"
        "<p>许可证：<b>%s</b> — <a href='%s'>%s</a></p>"
        "%s"
        "<h2>本服务运行版本的对应源码</h2>"
        "<table><tr><th>文件</th><th>字节</th><th>SHA-256（前16位）</th></tr>"
        "%s</table>"
        "<h2>第三方组件</h2>"
        "<table><tr><th>组件</th><th>版本</th><th>许可</th><th>权利方</th></tr>"
        "%s</table>"
        "<h2>说明</h2><p>处理产出的 PDF / 图片等文件不属于本许可证约束范围。</p>"
        "<p>MCP 端点：<code>POST /mcp</code>（JSON-RPC 2.0）</p>"
        "</body></html>"
    ) % (PAGE_CSS, LICENSE_ID, LICENSE_TEXT_URL, LICENSE_TEXT_URL,
         ("<p>仓库地址：<a href='%s'>%s</a></p>" % (u, u)) if u else "",
         "".join(rows), "".join(comp))


def _html_root():
    u = local_source_url()
    return (
        "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>"
        "<title>PDF文档处理服务（MCP）</title><style>%s</style></head><body>"
        "<h1>PDF文档处理服务</h1>"
        "<p>MCP 端点：<code>POST /mcp</code>，协议版本 %s，共 %d 个工具。</p>"
        + ("本服务运行在<b>远程代码服务</b>模式：只提供 pdf-code / "
           "pdf_source，不读写服务器上的任何文件。"
           if CODE_ONLY else "")
        + "<p><b>源代码</b>：<a href='/source'>%s</a></p>"
        "<p>许可证：%s（<a href='%s'>全文</a>）</p>"
        "</body></html>"
    ) % (PAGE_CSS, PROTOCOL_VERSION, len(active_tools()), u or "/source",
         LICENSE_ID, LICENSE_TEXT_URL)


class Handler(BaseHTTPRequestHandler):
    server_version = "pdf-tools-mcp"

    def log_message(self, fmt, *a):
        sys.stderr.write("[mcp] %s %s\n" % (self.command, self.path))

    def _send(self, code, body, ctype, extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _send_sse(self, obj):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.end_headers()
        data = json.dumps(obj, ensure_ascii=False)
        self.wfile.write(("event: message\ndata: %s\n\n" % data)
                         .encode("utf-8"))
        self.wfile.flush()

    def do_HEAD(self):
        self._send(200, "", "text/plain; charset=utf-8")

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in ("/", "/index.html"):
            self._send(200, _html_root(), "text/html; charset=utf-8")
        elif path == "/source":
            self._send(200, _html_source_page(), "text/html; charset=utf-8")
        elif path.startswith("/source/"):
            rel = path[len("/source/"):]
            self._send_file(rel)
        elif path == "/health":
            self._send(200, json.dumps({"ok": True, "code_only": CODE_ONLY,
                                        "tools": len(active_tools())},
                                       ensure_ascii=False),
                       "application/json")
        else:
            self._send(404, json.dumps({"error": "not found"},
                                       ensure_ascii=False),
                       "application/json")

    def _send_file(self, rel):
        if rel not in _source_files():
            self._send(404, json.dumps({"error": "不允许的文件：%s" % rel},
                                       ensure_ascii=False),
                       "application/json")
            return
        p = os.path.join(BASE_DIR, rel)
        data = open(p, "rb").read()
        if rel.endswith(".html"):
            ct = "text/html; charset=utf-8"
        elif rel.endswith(".py"):
            ct = "text/x-python; charset=utf-8"
        else:
            ct = "text/plain; charset=utf-8"
        self._send(200, data, ct,
                   {"Content-Disposition":
                    'attachment; filename="%s"' % os.path.basename(rel)})

    def do_POST(self):
        if self.path.split("?", 1)[0] != "/mcp":
            self._send(404, json.dumps({"error": "not found"},
                                       ensure_ascii=False),
                       "application/json")
            return
        n = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(n) if n else b""
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception as e:
            self._send(400, json.dumps(_err(None, -32700, "JSON 解析失败：%s" % e),
                                       ensure_ascii=False),
                       "application/json")
            return
        out = handle_payload(payload)
        if out is None:
            self.send_response(202)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        # MCP 客户端按规范会同时发 application/json 与 text/event-stream。
        # 只有它明确只收 event-stream 时才走 SSE，否则一律回普通 JSON，
        # 免得简单客户端被迫去解 SSE 帧。
        accept = self.headers.get("Accept") or ""
        want_sse = ("text/event-stream" in accept
                    and "application/json" not in accept)
        if want_sse:
            if isinstance(out, list):
                for o in out:
                    self._send_sse(o)
            else:
                self._send_sse(out)
            return
        self._send(200, json.dumps(out, ensure_ascii=False),
                   "application/json")


LOOPBACK_HOSTS = ("127.0.0.1", "localhost", "::1", "")


def _is_loopback(host):
    return (host or "").strip().lower() in LOOPBACK_HOSTS


def serve_http(host, port):
    srv = ThreadingHTTPServer((host, port), Handler)
    srv.daemon_threads = True
    u = local_source_url()
    lines = [
        "[mcp] 监听 http://%s:%d/mcp" % (host, port),
        "[mcp] 工具数 %d，协议 %s%s"
        % (len(active_tools()), PROTOCOL_VERSION,
           "（远程代码服务：仅 pdf-code / pdf_source）" if CODE_ONLY else ""),
        "[mcp] AGPL-3.0 第 13 条义务：必须向网络用户免费提供对应源码",
        "[mcp]   自持源码页 http://%s:%d/source" % (host, port),
    ]
    if u:
        lines.append("[mcp]   仓库地址 %s" % u)
    else:
        lines.append("[mcp]   警告：未配 --source-url / --public-base，"
                     "目前仅靠自持 /source 端点满足义务")
    sys.stderr.write("\n".join(lines) + "\n")
    sys.stderr.flush()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


def _stdio_streams():
    """stdio 的输入输出流。

    PyInstaller --windowed 打包后，没被重定向时 sys.stdout/sys.stdin 是
    None（pythonw 行为）。MCP 客户端拉起进程时一定带管道，但兜底直接用
    fd 0/1，保证任何拉起方式下协议帧都不丢。
    """
    try:
        rin = sys.stdin.buffer
    except AttributeError:
        rin = os.fdopen(0, "rb", closefd=False)
    try:
        rout = sys.stdout.buffer
    except AttributeError:
        rout = os.fdopen(1, "wb", closefd=False)
    return rin, rout


def serve_stdio():
    rin, rout = _stdio_streams()
    reader = io.TextIOWrapper(rin, encoding="utf-8",
                              errors="replace", newline=None)
    out = io.TextIOWrapper(rout, encoding="utf-8",
                           errors="replace", newline="\n",
                           line_buffering=True)
    for line in reader:
        line = line.strip()
        if not line:
            continue
        try:
            payload = json.loads(line)
        except Exception as e:
            out.write(json.dumps(_err(None, -32700, "JSON 解析失败：%s" % e),
                                 ensure_ascii=False) + "\n")
            out.flush()
            continue
        r = handle_payload(payload)
        if r is None:
            continue
        out.write(json.dumps(r, ensure_ascii=False) + "\n")
        out.flush()


# ---------- 给 GUI 内嵌用：后台起停 ----------

def start_http_background(host="127.0.0.1", port=0, allow_dirs=None,
                          source_url="", code_only=False):
    """在后台线程起 MCP HTTP 服务，供 GUI 调用。

    port=0 交给系统分配端口，返回 (server, 真实端口)。
    code_only=True：远程代码服务模式，只暴露 pdf-code / pdf_source——
    这两个工具不读写本机任何文件，所以绑 0.0.0.0 也不需要 allow_dirs。
    非 code_only 时契约不变：非回环地址必须已配 allow_dirs，
    否则等于把文件系统对网络敞开。
    """
    global SOURCE_URL, ALLOW_DIRS, CODE_ONLY, INSTRUCTIONS
    if source_url:
        SOURCE_URL = source_url.rstrip("/")
    if allow_dirs:
        ALLOW_DIRS = [os.path.abspath(d) for d in allow_dirs]
    CODE_ONLY = bool(code_only)
    INSTRUCTIONS = make_instructions()
    srv = ThreadingHTTPServer((host, port), Handler)
    srv.daemon_threads = True
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    return srv, srv.server_address[1]


def stop_http(srv):
    """非阻塞停止。shutdown() 会一直阻塞到 serve_forever 返回，
    在 GUI 主线程直接调会卡死界面，所以丢到后台线程。"""
    def _go():
        try:
            srv.shutdown()
        finally:
            srv.server_close()
    threading.Thread(target=_go, daemon=True).start()


def _frozen_exe():
    """PyInstaller 打包后返回 exe 路径，否则空串。"""
    if getattr(sys, "frozen", False):
        return sys.executable or ""
    return ""


def lan_ip():
    """本机在局域网里的 IP（UDP connect 探测，不发任何数据包）。

    远程代码服务要告诉用户「把哪个地址交给远程智能体」，127.0.0.1
    对别的机器没用。拿不到（无网络/多网卡异常）就退回 127.0.0.1。
    """
    s = None
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0)
        s.connect(("10.255.255.255", 1))
        return s.getsockname()[0] or "127.0.0.1"
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"
    finally:
        try:
            if s is not None:
                s.close()
        except Exception:
            pass


def remote_config_snippet(port, host=""):
    """远程智能体接入配置：告诉对方把 MCP 端点填进它的客户端。"""
    ip = host or lan_ip()
    return {
        "说明": "远程智能体接入：把下面的 url 填进它的 MCP 客户端配置。"
                "它连上后用 tools/list 能看到 pdf-code 与 pdf_source；"
                "需要处理 PDF 时调 pdf-code，把返回的 code 保存为 .py "
                "在它自己的机器上执行，文件不经过本机。",
        "url": "http://%s:%d/mcp" % (ip, port),
        "协议": "MCP Streamable HTTP（JSON-RPC 2.0），POST /mcp",
        "工具": ["pdf-code", "pdf_source"],
        "调用示例": {"tool": "pdf-code",
                     "arguments": {"command": "compress",
                                   "params": {"input": "a.pdf",
                                              "target_mb": 5}}},
        "源码义务": "AGPL-3.0 第 13 条：GET /source 免费提供对应源码",
    }


def client_config_snippets(port, script_dir=None):
    """本机接入的两种接法配置片段，供 GUI 展示与复制。

    frozen（打包成 exe）时没有 pdf_mcp.py 文件可拉，stdio 接法改为
    直接拉起 exe 自身 + --mcp-stdio 隐藏入口（pdf_tool._mcp_cli 实现）。
    """
    exe = _frozen_exe()
    if exe:
        stdio = {
            "说明": "打包版专用：客户端直接拉起本 exe 的 MCP 模式，"
                    "不用预先启动服务，也不需要本机装 Python",
            "command": exe,
            "args": ["--mcp-stdio"],
        }
        claude = {"command": exe, "args": ["--mcp-stdio"]}
    else:
        d = (script_dir or BASE_DIR).replace("\\", "/")
        stdio = {
            "说明": "客户端直接拉起进程，不用预先启动服务",
            "command": "python",
            "args": [d + "/pdf_mcp.py"],
        }
        claude = {"command": "python", "args": [d + "/pdf_mcp.py"]}
    return {
        "回环 HTTP": {
            "说明": "先在本面板点「启动」；服务只听 127.0.0.1，外网进不来",
            "url": "http://127.0.0.1:%d/mcp" % port,
        },
        "stdio 子进程": stdio,
        "Claude Desktop 示例": {
            "mcpServers": {
                "pdf-tools": claude
            }
        },
    }


# ---------- 方案 B：把程序交给远程方，让它在自己机器上跑 ----------

_BAT = '''@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo MCP 服务（仅本机 127.0.0.1:8765）...
where python >nul 2>nul && (set PY=python) || (set PY=py)
%PY% pdf_mcp.py --transport http --host 127.0.0.1 --port 8765
pause
'''

_SH = '''#!/usr/bin/env bash
cd "$(dirname "$0")"
echo "MCP 服务（仅本机 127.0.0.1:8765）..."
exec python3 pdf_mcp.py --transport http --host 127.0.0.1 --port 8765
'''

_TXT = """3Q PDF开源免费工具 · MCP 智能体接入（分发版）

这一份是给你本机用的。你的 PDF 文件不会离开你的电脑。

【怎么跑起来】
  Windows：双击 启动MCP服务.bat
  Mac/Linux：终端执行 bash 启动MCP服务.sh
  依赖：pip install pymupdf pillow openpyxl python-docx

【怎么接进你的 AI 客户端】
  见 MCP客户端配置示例.json，两种接法二选一：
  - 回环 HTTP：先启动上面的服务，再把 url 填进客户端
  - stdio：客户端自己拉起 pdf_mcp.py，不用预先启动
  两种都是本机通信，监听 127.0.0.1，网络上其他机器连不进来。

【许可】
  本软件整体以 GNU Affero General Public License v3.0 授权，
  完整文本见 LICENSE，第三方组件声明见 NOTICE。
  本软件基于 PyMuPDF（Copyright (C) Artifex Software, Inc.），
  PyMuPDF 采用 AGPL-3.0 或 Artifex 商业许可双轨。
  你处理产生的 PDF / 图片等产出物不受本许可约束。

【为什么不用上传文件】
  程序跑在你自己机器上，直接读写你本机的 PDF。
  没有服务器、没有文件上传、没有第三方经手你的文档。
"""


def export_distribution(dest_zip):
    """打包一份可交给远程方的自包含分发包（方案 B）。

    对方拿到后在本机运行，文件不出他的机器。对提供方而言这是
    「分发（convey）」而非「提供网络服务」，AGPL 第 13 条不触发。
    """
    dest_zip = os.path.abspath(dest_zip)
    root = "pdf-tools-mcp"
    with zipfile.ZipFile(dest_zip, "w", zipfile.ZIP_DEFLATED) as z:
        for n in _source_files():
            z.write(os.path.join(BASE_DIR, n), root + "/" + n)
        z.writestr(root + "/启动MCP服务.bat", _BAT)
        z.writestr(root + "/启动MCP服务.sh", _SH)
        z.writestr(root + "/MCP客户端配置示例.json",
                   json.dumps(client_config_snippets(8765, "<本目录>"),
                              ensure_ascii=False, indent=2))
        z.writestr(root + "/给远程方看这里.txt", _TXT)
    return dest_zip


# ---------- 方案 C：生成定制代码，交给远程智能体自己执行 ----------

GEN_HEADER = '''# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (C) 2026 三扣科技 (3Q Tech)
# 本文件由 pdf_mcp 的 pdf-code 工具生成，内含 3Q PDF开源免费工具 未修改的源码副本，
# 整体以 GNU Affero General Public License v3.0 授权。
# 完整许可文本：https://www.gnu.org/licenses/agpl-3.0.txt
# 第三方组件（PyMuPDF, Copyright (C) Artifex Software, Inc.）声明见 NOTICE。
# 本文件处理产出的 PDF / 图片等文件不受本许可约束。
'''


def _read_src(name):
    with open(os.path.join(BASE_DIR, name), encoding="utf-8") as f:
        return f.read()


def gen_pdf_code(command, params=None):
    """生成一段自包含的 Python 脚本，交给调用方在它自己机器上跑。

    与「分发整个软件」的区别：这里发出去的是**按需定制的代码**，不是安装包。
    调用方只要有 Python 和依赖就能执行，不需要 Windows，也不需要我们的程序；
    文件全程在它自己机器上，我们不经手。

    法律上这仍是分发（convey）——脚本内含 AGPL 源码副本——所以 §4 要求的
    版权声明与许可指引必须长在生成代码的头部，不能只在文档里说。
    """
    if command not in pdf_cli.RESULT:
        return {"ok": False, "command": command,
                "error": "未知命令：%s" % command}
    core_src = _read_src("pdf_core.py")
    cli_src = _read_src("pdf_cli.py")
    fn = pdf_cli.RESULT[command].__name__
    merged = {}
    for o in pdf_cli.OPTS[command]:
        merged[o["dest"]] = o["default"]
    for k, v in (params or {}).items():
        merged[k] = v
    code = "\n".join([
        GEN_HEADER,
        "# ---- 用法：把本文件保存为任意 .py，然后 python <文件> ----",
        "import base64 as _b64",
        "import json as _json",
        "import sys",
        "import types",
        "",
        "_CORE_B64 = %r" % base64.b64encode(
            core_src.encode("utf-8")).decode("ascii"),
        "_CLI_B64 = %r" % base64.b64encode(
            cli_src.encode("utf-8")).decode("ascii"),
        # 注意这里必须用 repr 而不是 json.dumps：JSON 的 null/true/false
        # 不是 Python 字面量，拼进源码会 NameError。
        "_PARAMS = %r" % merged,
        "",
        "",
        "def _load():",
        "    # 先把 core 注册进 sys.modules，cli 的 import pdf_core 才接得上",
        "    m = types.ModuleType('pdf_core')",
        "    m.__file__ = '<generated>/pdf_core.py'",
        "    exec(_b64.b64decode(_CORE_B64).decode('utf-8'), m.__dict__)",
        "    sys.modules['pdf_core'] = m",
        "    c = types.ModuleType('pdf_cli')",
        "    c.__file__ = '<generated>/pdf_cli.py'",
        "    exec(_b64.b64decode(_CLI_B64).decode('utf-8'), c.__dict__)",
        "    return c",
        "",
        "",
        "if __name__ == '__main__':",
        "    _cli = _load()",
        "    from argparse import Namespace",
        "    _r = getattr(_cli, %r)(Namespace(**_PARAMS))" % fn,
        "    print(_json.dumps(_r, ensure_ascii=False, default=str))",
        "",
    ])
    info = {
        "ok": True,
        "command": command,
        "entrypoint": fn,
        "params": merged,
        "code": code,
        "code_bytes": len(code.encode("utf-8")),
        "code_sha256": hashlib.sha256(code.encode("utf-8")).hexdigest(),
        "run": "把本段代码保存为 .py 文件后执行 python <文件>；"
               "它只会读写你在 params 里给出的路径",
        "dependencies": DEPS,
        "pip": "pip install " + " ".join(DEPS),
        "license": LICENSE_ID,
        "license_text": LICENSE_TEXT_URL,
        "note": "本段代码内含 3Q PDF开源免费工具 完整源码副本，整体以 AGPL-3.0 授权；"
                "你处理产出的 PDF / 图片等文件不受该许可约束。",
    }
    u = local_source_url()
    if u:
        info["source_url"] = u
    return info


def main(argv=None):
    global SOURCE_URL, PUBLIC_BASE, ALLOW_DIRS, INSTRUCTIONS, ALLOW_HIGH_RISK
    global CODE_ONLY
    ap = argparse.ArgumentParser(
        prog="pdf_mcp.py",
        description="3Q PDF开源免费工具 · MCP 服务器（stdio 或 Streamable HTTP）。")
    ap.add_argument("--transport", choices=["stdio", "http"], default="stdio")
    ap.add_argument("--host", default="127.0.0.1",
                    help="监听地址。默认只听本机回环；全量模式下改绑 "
                         "0.0.0.0/局域网 IP 必须同时用 --allow-dir 收紧路径。"
                         "--code-only 模式不受此限（不暴露文件工具）")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--source-url", default="",
                    help="对外公开的源码仓库地址（满足 AGPL §13）")
    ap.add_argument("--public-base", default="",
                    help="对外 URL 前缀，如 https://pdf.example.com")
    ap.add_argument("--allow-dir", action="append", default=[],
                    help="路径沙箱，可重复；未设则不限制")
    ap.add_argument("--allow-print", action="store_true",
                    help="放行 print 高危操作（跳过人工确认）。"
                         "默认拒绝，防止智能体误调用真的打印")
    ap.add_argument("--code-only", action="store_true",
                    help="远程代码服务：只暴露 pdf-code / pdf_source，"
                         "不暴露任何读写本机文件的工具，因此允许绑定 "
                         "0.0.0.0 等非回环地址而不需要 --allow-dir")
    ap.add_argument("--self-test", action="store_true",
                    help="只跑一遍自检（tools/list 与 pdf_source），不开服务")
    a = ap.parse_args(argv)

    SOURCE_URL = (a.source_url or "").rstrip("/")
    PUBLIC_BASE = (a.public_base or "").rstrip("/")
    ALLOW_DIRS = [os.path.abspath(d) for d in (a.allow_dir or [])]
    ALLOW_HIGH_RISK = bool(a.allow_print)
    CODE_ONLY = bool(a.code_only)
    INSTRUCTIONS = make_instructions()

    if a.self_test:
        r = _call_tool(1, {"name": SOURCE_TOOL_NAME, "arguments": {}})
        info = r["result"]["structuredContent"]
        print(json.dumps({
            "tools": len(active_tools()),
            "code_only": CODE_ONLY,
            "protocol": PROTOCOL_VERSION,
            "license": info["license"],
            "source_files": [f["file"] for f in info["source_files"]],
            "components": info["components"],
        }, ensure_ascii=False, indent=2))
        return 0

    if a.transport == "http":
        if not _is_loopback(a.host) and not ALLOW_DIRS and not CODE_ONLY:
            # 响亮失败：非回环 + 无沙箱 + 暴露全量文件工具
            # = 把整个文件系统对网络开放，绝不能让它静悄悄跑起来。
            # code_only 模式例外：它只出代码，不碰任何文件。
            sys.stderr.write(
                "拒绝启动：监听地址 %s 不是本机回环，且未设 --allow-dir。\n"
                "这等于把本机整个文件系统对网络开放——任何能访问该地址的人都能\n"
                "读走你机器上的任意文件。请三选一：\n"
                "  1) 只给本机智能体用：去掉 --host（默认 127.0.0.1）\n"
                "  2) 确要给局域网用：补上 --allow-dir <共享目录>\n"
                "  3) 只想让远程智能体拿代码：加 --code-only（不暴露文件工具）\n"
                % a.host)
            return 2
        if not _is_loopback(a.host):
            if CODE_ONLY:
                sys.stderr.write(
                    "[mcp] 注意：监听 %s，网络内其他设备可以访问；"
                    "远程代码服务模式，不暴露任何读写文件的工具\n" % a.host)
            else:
                sys.stderr.write(
                    "[mcp] 注意：监听 %s，网络内其他设备可以访问；"
                    "路径已限制在 %s\n" % (a.host, "、".join(ALLOW_DIRS)))
        serve_http(a.host, a.port)
    else:
        serve_stdio()
    return 0


if __name__ == "__main__":
    sys.exit(main())
