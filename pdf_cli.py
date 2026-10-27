# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (C) 2026 三扣科技 (3Q Tech)
# 本作品基于 PyMuPDF（Copyright (C) Artifex Software, Inc.，AGPL-3.0 或商业许可），
# 依 AGPL-3.0 第 5 条整体以 AGPL-3.0 授权。完整文本见同目录 LICENSE，第三方声明见 NOTICE。
"""
3Q PDF开源免费工具 · Agent 入口（pdf_cli.py）

给 AI 调用的三条硬约定（也是本文件存在的理由）：
1. **输出永远是一行 JSON**。成功/失败都走同一个信封：
   成功 {"ok": true, "command": "...", ...结果字段..., "elapsed_ms": 123}
   失败 {"ok": false, "command": "...", "error": "中文说明"} 且 exit code = 1
   —— agent 不需要解析 traceback。
2. **先问 `capabilities` 再干活**。它返回全部命令 + 参数 + 枚举值的机器可读清单，
   agent 读一遍就知道自己能干什么、参数怎么填。
3. **路径一律给绝对路径**，输出路径缺省时按输入文件派生，结果里必定回传真实落盘路径。

人类也能用：直接 `python pdf_cli.py 命令 --help`，或在命令行里串着跑。

依赖：只依赖 pdf_core.py（不 import tkinter），与 GUI 版 pdf_tool.py 共用同一份核心。
"""
import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_core as core  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


class ToolError(Exception):
    """业务错误：只回中文说明，不回 traceback。"""


# ---------- 通用参数 ----------

def _o(flags, dest, help, default=None, kw=None, enum=None, required=False):
    return {"flags": flags, "dest": dest, "help": help, "default": default,
            "kw": kw or {}, "enum": enum, "required": required}


IN_PDF = _o(["--input", "-i"], "input", "输入 PDF 路径", required=True)
OUT = _o(["--output", "-o"], "output",
         "输出路径（缺省：按输入文件同名派生，结果里回传真路径）")
OUTDIR = _o(["--output", "-o"], "output", "输出文件夹（缺省：输入同目录下新建）")
PAGES = _o(["--pages"], "pages", '页码表达式，如 "1,3-5,10-"（缺省全部页）')
PW = _o(["--password"], "password", "源 PDF 的打开密码（若已加密）")
MANY_PDF = _o(["--input", "-i"], "input", "输入 PDF，可多个", kw={"nargs": "+"},
              required=True)


# ---------- 工具函数 ----------

def _in(a):
    p = getattr(a, "input", None)
    if isinstance(p, list):
        return [os.path.abspath(x) for x in p]
    if not p:
        raise ToolError("缺少 --input 参数")
    p = os.path.abspath(p)
    if not os.path.exists(p):
        raise ToolError("文件不存在：%s" % p)
    return p


def _ins(a, exts=None):
    ps = getattr(a, "input", None) or []
    if not ps:
        raise ToolError("缺少 --input 参数")
    out = []
    for x in ps:
        p = os.path.abspath(x)
        if not os.path.exists(p):
            raise ToolError("文件不存在：%s" % p)
        if exts and os.path.splitext(p)[1].lower() not in exts:
            raise ToolError("不支持的文件类型：%s" % p)
        out.append(p)
    return out


def _out(a, suffix):
    """输出文件路径：给了就用给的，没给就按输入派生。"""
    p = getattr(a, "output", None)
    if not p:
        src = getattr(a, "input", None)
        src = src[0] if isinstance(src, list) and src else src
        if isinstance(src, str):
            d = os.path.dirname(os.path.abspath(src))
            stem = os.path.splitext(os.path.basename(src))[0]
            p = os.path.join(d, "%s_%s.pdf" % (stem, suffix))
        else:
            p = os.path.join(os.getcwd(), "output_%s.pdf" % suffix)
    p = os.path.abspath(p)
    src = getattr(a, "input", None)
    src = src[0] if isinstance(src, list) and src else src
    if isinstance(src, str) and os.path.abspath(src) == p:
        raise ToolError("输出路径不能等于输入文件（本工具从不覆盖原文件）")
    d = os.path.dirname(p)
    if d:
        os.makedirs(d, exist_ok=True)
    return p


def _outdir(a, name):
    p = getattr(a, "output", None)
    if not p:
        src = getattr(a, "input", None)
        src = src[0] if isinstance(src, list) and src else src
        d = os.path.dirname(os.path.abspath(src)) if isinstance(src, str) else os.getcwd()
        p = os.path.join(d, name)
    p = os.path.abspath(p)
    os.makedirs(p, exist_ok=True)
    return p


def _pages(a, total):
    e = getattr(a, "pages", None)
    if not e:
        return list(range(total))
    return core.parse_page_expr(e, total)


def _color(s):
    """颜色：支持 黑/红/蓝/绿/浅灰 或 "r,g,b"（0~1）或 "0-255,0-255,0-255"。"""
    if not s:
        return (0, 0, 0)
    if s in core.COLOR_PRESETS:
        return core.COLOR_PRESETS[s]
    parts = [x.strip() for x in s.replace("，", ",").split(",") if x.strip()]
    if len(parts) != 3:
        raise ToolError('颜色格式应为 "r,g,b"，如 "0.8,0.1,0.1"，或直接写 黑/红/蓝/绿/浅灰')
    try:
        vals = [float(x) for x in parts]
    except ValueError:
        raise ToolError("颜色分量必须是数字")
    if max(vals) > 1.0:
        vals = [v / 255.0 for v in vals]
    return tuple(max(0.0, min(1.0, v)) for v in vals)


def _jsonarg(s, what):
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        raise ToolError("%s 不是合法 JSON：%s" % (what, s[:120]))


def _doc(a):
    return core.open_pdf_doc(_in(a), getattr(a, "password", None))


def _size(p):
    try:
        return os.path.getsize(p)
    except OSError:
        return 0


# ---------- 各命令实现 ----------

def c_ink(a):
    p = _in(a)
    out = _out(a, "画笔")
    strokes = _jsonarg(getattr(a, "strokes", None), "--strokes") or []
    if not isinstance(strokes, list):
        raise ToolError("--strokes 应为 JSON 数组")
    n = core.do_ink(p, out, strokes, pw=getattr(a, "password", None))
    return {"output": out, "strokes": n, "size_bytes": _size(out)}


def c_save_as(a):
    """另存为新文件（原文件不动），可顺带写画笔笔迹。"""
    p = _in(a)
    out = _out(a, "副本")
    if os.path.abspath(out) == os.path.abspath(p):
        raise ToolError("输出路径不能等于输入文件（另存请换个名字）")
    strokes = _jsonarg(getattr(a, "strokes", None), "--strokes") or None
    path, n = core.do_save_as(p, out, pw=getattr(a, "password", None),
                              strokes=strokes)
    return {"output": path, "strokes": n, "size_bytes": _size(path)}


def c_save(a):
    """覆盖保存：直接写回 --input 自身（可顺带写画笔笔迹）。"""
    p = _in(a)
    if not os.path.exists(p):
        raise ToolError("输入文件不存在：%s" % p)
    strokes = _jsonarg(getattr(a, "strokes", None), "--strokes") or None
    path, n = core.do_save_inplace(p, pw=getattr(a, "password", None),
                                  strokes=strokes)
    return {"output": path, "strokes": n, "size_bytes": _size(path)}


def c_capabilities(a):
    cmds = []
    for name, help_, opts, _fn in SPECS:
        if name == "run":
            continue
        cmds.append({
            "name": name,
            "help": help_,
            "params": [{"flag": o["flags"][0], "dest": o["dest"],
                        "required": o["required"], "help": o["help"],
                        "default": o["default"], "enum": o["enum"]}
                       for o in opts],
        })
    return {"tool": "3Q PDF开源免费工具", "version": core.VERSION, "commands": cmds}


def c_info(a):
    p = _in(a)
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    sizes = []
    has_text = False
    for i in range(doc.page_count):
        pg = doc[i]
        dw, dh = core.page_geometry(pg)[:2]
        sizes.append([round(dw, 1), round(dh, 1)])
        if not has_text and pg.get_text().strip():
            has_text = True
    md = {k: v for k, v in (doc.metadata or {}).items() if v}
    toc = doc.get_toc()
    doc.close()
    return {"input": p, "pages": len(sizes), "page_sizes_pt": sizes,
            "encrypted": _needs_pass(p), "has_text": has_text,
            "metadata": md, "toc_count": len(toc),
            "size_bytes": _size(p)}


def _needs_pass(p):
    try:
        d = core.fitz.open(p)
        n = d.needs_pass
        d.close()
        return bool(n)
    except Exception:
        return False


def c_text(a):
    p = _in(a)
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    idx = _pages(a, doc.page_count)
    cap = int(getattr(a, "max_chars", 0) or 0)
    per, total = [], 0
    for i in idx:
        t = doc[i].get_text() or ""
        if cap and total + len(t) > cap:
            t = t[:max(0, cap - total)]
        per.append({"page": i + 1, "text": t})
        total += len(t)
        if cap and total >= cap:
            break
    doc.close()
    out = {"input": p, "pages_read": len(per), "chars": total,
           "pages": per}
    if getattr(a, "output", None):
        op = os.path.abspath(a.output)
        with open(op, "w", encoding="utf-8") as f:
            f.write("\n".join(x["text"] for x in per))
        out["output"] = op
    return out


def c_merge(a):
    files = _ins(a)
    out = _out(a, "合并")
    n = core.do_merge(files, out)
    return {"output": out, "inputs": len(files), "pages": n,
            "size_bytes": _size(out)}


def c_split(a):
    p = _in(a)
    d = _outdir(a, "拆分结果")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    total = doc.page_count
    doc.close()
    if getattr(a, "every", 0):
        n = max(1, int(a.every))
        ranges = [(s, min(s + n - 1, total - 1)) for s in range(0, total, n)]
    else:
        if not getattr(a, "ranges", None):
            raise ToolError("需要 --ranges（如 \"1-3,5-9\"）或 --every N")
        ranges = core.parse_ranges_expr(a.ranges, total)
    outs = core.do_split(p, ranges, d,
                         os.path.splitext(os.path.basename(p))[0],
                         getattr(a, "password", None))
    return {"output_dir": d, "outputs": outs, "parts": len(outs)}


def c_extract(a):
    p = _in(a)
    out = _out(a, "提取")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    idx = _pages(a, doc.page_count)
    doc.close()
    core.do_extract(p, idx, out, getattr(a, "password", None))
    return {"output": out, "pages": len(idx), "size_bytes": _size(out)}


def c_delete(a):
    p = _in(a)
    out = _out(a, "删页")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    idx = _pages(a, doc.page_count)
    if len(idx) >= doc.page_count:
        doc.close()
        raise ToolError("不能删除全部页面，至少要保留 1 页")
    doc.close()
    n = core.do_delete_pages(p, idx, out, getattr(a, "password", None))
    return {"output": out, "deleted": len(idx), "pages_left": n,
            "size_bytes": _size(out)}


def c_rotate(a):
    p = _in(a)
    out = _out(a, "旋转")
    ang = int(getattr(a, "angle", 90) or 90)
    if ang not in (90, 180, 270):
        raise ToolError("--angle 只能是 90 / 180 / 270")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    idx = _pages(a, doc.page_count)
    doc.close()
    core.do_rotate_pages(p, idx, ang, out, getattr(a, "password", None))
    return {"output": out, "pages": len(idx), "angle": ang,
            "size_bytes": _size(out)}


def c_reorder(a):
    p = _in(a)
    out = _out(a, "重排")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    total = doc.page_count
    doc.close()
    order = getattr(a, "order", None)
    if order:
        idx = core.parse_page_expr(order, total)
    else:
        idx = list(range(total))
    core.do_reorder(p, out, [("page", i) for i in idx], getattr(a, "password", None))
    return {"output": out, "pages": len(idx), "order": [i + 1 for i in idx],
            "size_bytes": _size(out)}


def c_pagenum(a):
    p = _in(a)
    out = _out(a, "加页码")
    core.do_add_page_numbers(
        p, out,
        getattr(a, "position", "底部居中"),
        int(getattr(a, "start", 1) or 1),
        getattr(a, "format", "第 X 页 / 共 N 页"),
        int(getattr(a, "fontsize", 10) or 10),
        getattr(a, "password", None))
    return {"output": out, "size_bytes": _size(out)}


def c_hf(a):
    p = _in(a)
    out = _out(a, "页眉页脚")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    cfg = {"header": getattr(a, "header", "") or "",
           "footer": getattr(a, "footer", "") or "",
           "size": int(getattr(a, "size", 10) or 10),
           "margin_mm": float(getattr(a, "margin_mm", 10) or 10),
           "align": getattr(a, "align", "居中"),
           "pages": pgs, "start": int(getattr(a, "start", 1) or 1)}
    if not cfg["header"] and not cfg["footer"]:
        raise ToolError("至少要给 --header 或 --footer 其中一个")
    core.do_header_footer(p, out, cfg, getattr(a, "password", None))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_wmtext(a):
    p = _in(a)
    out = _out(a, "水印")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    core.do_watermark_text(
        p, out, pgs, getattr(a, "text", "机密"),
        float(getattr(a, "fontsize", 48) or 48),
        _color(getattr(a, "color", "浅灰")),
        int(getattr(a, "angle", 45) or 45),
        float(getattr(a, "opacity", 0.2) or 0.2),
        getattr(a, "layout", "平铺整页"),
        getattr(a, "password", None))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_wmimage(a):
    p = _in(a)
    img = getattr(a, "image", None)
    if not img or not os.path.exists(img):
        raise ToolError("缺少或不存在的 --image 图片")
    out = _out(a, "图片水印")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    core.do_watermark_image(
        p, out, pgs, os.path.abspath(img),
        float(getattr(a, "scale", 30) or 30),
        float(getattr(a, "opacity", 0.2) or 0.2),
        getattr(a, "layout", "居中一个"),
        int(getattr(a, "angle", 0) or 0),
        getattr(a, "password", None))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_imgs2pdf(a):
    files = _ins(a, exts=(".png", ".jpg", ".jpeg", ".bmp", ".gif"))
    out = _out(a, "图片转PDF")
    rots = getattr(a, "rotate", None)
    auto_exif = not getattr(a, "no_exif", False)
    temps, final = [], []
    try:
        for i, f in enumerate(files):
            ang = 0
            if rots:
                parts = [x for x in str(rots).split(",") if x.strip()]
                if i < len(parts) and int(parts[i]) % 360:
                    ang = int(parts[i]) % 360
            elif auto_exif:
                ang = core.image_exif_rotation(f)
            if ang:
                t = core.rotate_image_file(f, ang)
                temps.append(t)
                final.append(t)
            else:
                final.append(f)
        core.do_images_to_pdf(final, out, bool(getattr(a, "a4", False)))
    finally:
        for t in temps:
            try:
                os.remove(t)
            except OSError:
                pass
    return {"output": out, "pages": len(files), "size_bytes": _size(out)}


def c_pdf2imgs(a):
    p = _in(a)
    d = _outdir(a, "导出图片")
    outs = core.do_pdf_to_images(p, d, getattr(a, "fmt", "png"),
                                 int(getattr(a, "dpi", 150) or 150),
                                 getattr(a, "password", None))
    return {"output_dir": d, "outputs": outs, "count": len(outs)}


def c_pdf2word(a):
    p = _in(a)
    out = os.path.abspath(getattr(a, "output", None)
                          or os.path.splitext(p)[0] + ".docx")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    core.do_pdf_to_docx(p, out, pgs, getattr(a, "password", None))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_pdf2excel(a):
    p = _in(a)
    out = os.path.abspath(getattr(a, "output", None)
                          or os.path.splitext(p)[0] + ".xlsx")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    npages, ntables = core.do_pdf_to_xlsx(p, out, pgs, getattr(a, "password", None))
    return {"output": out, "pages": npages, "tables": ntables,
            "size_bytes": _size(out)}


def c_extract_images(a):
    p = _in(a)
    d = _outdir(a, "文档内图片")
    outs = core.do_extract_images(p, d, getattr(a, "password", None),
                                  getattr(a, "fmt", "auto"),
                                  int(getattr(a, "min_px", 16) or 16))
    return {"output_dir": d, "outputs": outs, "count": len(outs)}


def c_compress(a):
    p = _in(a)
    out = _out(a, "压缩")
    old, new = core.do_compress(p, out,
                                int(getattr(a, "quality", 50) or 50),
                                int(getattr(a, "max_dpi", 0) or 0),
                                getattr(a, "password", None))
    return {"output": out, "size_before": old, "size_after": new,
            "reduced_pct": round((1 - new / old) * 100, 1) if old else 0,
            "size_bytes": new}


def c_compress_to(a):
    p = _in(a)
    out = _out(a, "压缩")
    mb = float(getattr(a, "target_mb", 0) or 0)
    if mb <= 0:
        raise ToolError("需要 --target-mb，如 5 表示压到 5MB 以内")
    old, new, ok = core.do_compress_to_target(
        p, out, int(mb * 1048576), getattr(a, "password", None))
    return {"output": out, "size_before": old, "size_after": new,
            "target_bytes": int(mb * 1048576), "target_reached": ok,
            "size_bytes": new}


def c_crop(a):
    p = _in(a)
    out = _out(a, "裁剪")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    margins = (0, 0, 0, 0)
    if getattr(a, "margins", None):
        parts = [x for x in str(a.margins).replace("，", ",").split(",") if x.strip()]
        if len(parts) != 4:
            raise ToolError('--margins 应为 "左,上,右,下"（毫米），如 "10,10,10,10"')
        margins = tuple(float(x) for x in parts)
    n, changed = core.do_crop_pages(
        p, out, pgs, getattr(a, "mode", "自动去黑边"),
        float(getattr(a, "keep_mm", 3) or 0), margins,
        getattr(a, "password", None))
    return {"output": out, "pages": n, "pages_cropped": changed,
            "size_bytes": _size(out)}


def c_office2pdf(a):
    files = _ins(a)
    d = _outdir(a, "Office转PDF结果")
    outs, errs = core.office_batch_to_pdf(files, d)
    res = {"output_dir": d, "outputs": outs, "converted": len(outs),
           "failed": len(errs)}
    if errs:
        res["errors"] = [{"file": n, "error": e} for n, e in errs]
    return res


def c_encrypt(a):
    p = _in(a)
    pw = getattr(a, "password", None)
    if not pw:
        raise ToolError("需要 --password 设置打开密码")
    out = _out(a, "加密")
    core.do_encrypt(p, out, pw, getattr(a, "owner_pw", None))
    return {"output": out, "encrypted": True, "size_bytes": _size(out)}


def c_decrypt(a):
    p = _in(a)
    out = _out(a, "无密码")
    core.do_decrypt(p, out, getattr(a, "password", None))
    return {"output": out, "encrypted": False, "size_bytes": _size(out)}


def c_redact(a):
    p = _in(a)
    out = _out(a, "涂黑")
    areas = _jsonarg(getattr(a, "areas", None), "--areas")
    n = core.do_redact(p, out, areas=areas,
                       pages=None, keyword=getattr(a, "keyword", None),
                       pw=getattr(a, "password", None))
    return {"output": out, "redacted": n, "size_bytes": _size(out)}


def c_annotate(a):
    p = _in(a)
    out = _out(a, "标注")
    n = core.do_annotate(p, out, getattr(a, "kind", "高亮"),
                         keyword=getattr(a, "keyword", None),
                         pw=getattr(a, "password", None))
    return {"output": out, "annotations": n, "size_bytes": _size(out)}


def c_insert_text(a):
    p = _in(a)
    out = _out(a, "加文字")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    xy = None
    if getattr(a, "xy", None):
        parts = [x for x in str(a.xy).replace("，", ",").split(",") if x.strip()]
        if len(parts) != 2:
            raise ToolError('--xy 应为 "x,y"（毫米，原点页面左上角）')
        xy = (float(parts[0]), float(parts[1]))
    core.do_insert_text(p, out, pgs, getattr(a, "text", ""),
                        float(getattr(a, "fontsize", 12) or 12),
                        _color(getattr(a, "color", "黑")),
                        getattr(a, "position", "页面居中"), xy,
                        int(getattr(a, "angle", 0) or 0),
                        getattr(a, "password", None),
                        bool(getattr(a, "vector", False)))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_insert_image(a):
    p = _in(a)
    img = getattr(a, "image", None)
    if not img or not os.path.exists(img):
        raise ToolError("缺少或不存在的 --image 图片")
    out = _out(a, "加图片")
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    xy = None
    if getattr(a, "xy", None):
        parts = [x for x in str(a.xy).replace("，", ",").split(",") if x.strip()]
        if len(parts) != 2:
            raise ToolError('--xy 应为 "x,y"（毫米，原点页面左上角）')
        xy = (float(parts[0]), float(parts[1]))
    core.do_insert_image(p, out, os.path.abspath(img), pgs,
                         getattr(a, "position", "页面居中"), xy,
                         float(getattr(a, "width_pct", 20) or 20),
                         int(getattr(a, "rotate", 0) or 0),
                         getattr(a, "password", None))
    return {"output": out, "pages": len(pgs), "size_bytes": _size(out)}


def c_toc(a):
    p = _in(a)
    toc = core.get_toc_list(p, getattr(a, "password", None))
    return {"input": p,
            "toc": [{"level": t[0], "title": t[1], "page": t[2]} for t in toc]}


def c_toc_set(a):
    p = _in(a)
    out = _out(a, "书签")
    data = _jsonarg(getattr(a, "toc", None), "--toc")
    if data is None:
        raise ToolError('需要 --toc，格式 \'[{"level":1,"title":"第一章","page":1}]\'')
    rows = [[int(x["level"]), str(x["title"]), int(x["page"])] for x in data]
    core.do_set_toc(p, out, rows, getattr(a, "password", None))
    return {"output": out, "entries": len(rows), "size_bytes": _size(out)}


def c_form(a):
    p = _in(a)
    out = _out(a, "已填表")
    vals = _jsonarg(getattr(a, "values", None), "--values")
    if not vals:
        raise ToolError('需要 --values，格式 \'{"字段名": "值"}\'')
    n = core.do_fill_form(p, out, vals, getattr(a, "password", None))
    return {"output": out, "filled": n, "size_bytes": _size(out)}


def c_print(a):
    p = _in(a)
    doc = core.open_pdf_doc(p, getattr(a, "password", None))
    pgs = _pages(a, doc.page_count)
    doc.close()
    n = core.do_print(p, getattr(a, "printer", ""), pgs,
                      int(getattr(a, "copies", 1) or 1),
                      getattr(a, "output", None),
                      None, getattr(a, "password", None))
    return {"printed_pages": n}


def c_batch(a):
    files = _ins(a)
    d = _outdir(a, "批量结果")
    params = _jsonarg(getattr(a, "params", None), "--params") or {}
    res = core.do_batch(getattr(a, "op", ""), files, d, params)
    rows = [{"input": r[0], "output": r[1], "error": r[2]} for r in res]
    return {"output_dir": d, "results": rows,
            "ok_count": sum(1 for r in rows if not r["error"]),
            "fail_count": sum(1 for r in rows if r["error"])}


# ---------- 命令表（capabilities 与 argparse 共用一份） ----------

SPECS = [
    ("capabilities", "列出全部命令与参数（Agent 第一步先调它）", [], c_capabilities),
    ("info", "查看 PDF 基本信息（页数/尺寸/是否加密/元数据/书签数）",
     [IN_PDF, PW], c_info),
    ("text", "抽取正文文本（Agent 读 PDF 内容的主力命令）",
     [IN_PDF, PW, PAGES,
      _o(["--max-chars"], "max_chars", "最多返回多少字符（0=不限）", 0, {"type": int}),
      _o(["--output"], "output", "同时把文本写成的 .txt 路径")], c_text),
    ("merge", "合并多个 PDF（按 --input 的给出顺序）", [MANY_PDF, OUT], c_merge),
    ("split", "拆分 PDF（按页码范围组或每 N 页一份）",
     [IN_PDF, PW, OUTDIR,
      _o(["--ranges"], "ranges", '页码范围组，如 "1-3,5-9,10-"'),
      _o(["--every"], "every", "每 N 页拆一份", 0, {"type": int})], c_split),
    ("extract", "按页码提取页面另存为新 PDF", [IN_PDF, PW, OUT, PAGES], c_extract),
    ("delete", "删除指定页", [IN_PDF, PW, OUT, PAGES], c_delete),
    ("rotate", "永久旋转页面（顺时针）",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--angle"], "angle", "旋转角度", 90, {"type": int}, enum=[90, 180, 270])],
     c_rotate),
    ("reorder", "按新顺序重排页面（--order 可重复引用同一页实现复制）",
     [IN_PDF, PW, OUT, _o(["--order"], "order", '新顺序，如 "3,1,2,2"')], c_reorder),
    ("pagenum", "添加页码",
     [IN_PDF, PW, OUT,
      _o(["--position"], "position", "位置", "底部居中",
         enum=["底部居中", "底部左侧", "底部右侧", "顶部左侧", "顶部右侧"]),
      _o(["--start"], "start", "起始页号", 1, {"type": int}),
      _o(["--format"], "format", "格式，X=当前页 N=总页数", "第 X 页 / 共 N 页"),
      _o(["--fontsize"], "fontsize", "字号", 10, {"type": int})], c_pagenum),
    ("header-footer", "添加页眉/页脚（{page} {total} 占位符）",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--header"], "header", "页眉文字"),
      _o(["--footer"], "footer", "页脚文字"),
      _o(["--size"], "size", "字号", 10, {"type": int}),
      _o(["--margin-mm"], "margin_mm", "边距（毫米）", 10, {"type": float}),
      _o(["--align"], "align", "对齐", "居中", enum=["左侧", "居中", "右侧"]),
      _o(["--start"], "start", "起始页号", 1, {"type": int})], c_hf),
    ("watermark-text", "添加文字水印",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--text"], "text", "水印文字", "机密"),
      _o(["--fontsize"], "fontsize", "字号", 48, {"type": float}),
      _o(["--color"], "color", "颜色：0.8,0.1,0.1 或 黑/红/蓝/绿/浅灰", "浅灰"),
      _o(["--angle"], "angle", "角度", 45, {"type": int}),
      _o(["--opacity"], "opacity", "不透明度 0~1", 0.2, {"type": float}),
      _o(["--layout"], "layout", "排布", "平铺整页", enum=["平铺整页", "居中一个"])],
     c_wmtext),
    ("watermark-image", "添加图片水印",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--image"], "image", "水印图片路径"),
      _o(["--scale"], "scale", "缩放百分比", 30, {"type": float}),
      _o(["--opacity"], "opacity", "不透明度 0~1", 0.2, {"type": float}),
      _o(["--layout"], "layout", "排布", "居中一个", enum=["平铺整页", "居中一个"]),
      _o(["--angle"], "angle", "角度", 0, {"type": int})], c_wmimage),
    ("images-to-pdf", "图片转 PDF（每张一页）",
     [_o(["--input", "-i"], "input", "图片，可多个", kw={"nargs": "+"},
         required=True), OUT,
      _o(["--a4"], "a4", "放入 A4 纸居中（适合打印）", False,
         {"action": "store_true"}),
      _o(["--rotate"], "rotate", "逐张顺时针角度，如 90,0,270"),
      _o(["--no-exif"], "no_exif", "不自动纠正手机照片方向", False,
         {"action": "store_true"})], c_imgs2pdf),
    ("pdf-to-images", "PDF 每页导出为图片",
     [IN_PDF, PW, OUTDIR,
      _o(["--dpi"], "dpi", "导出分辨率", 150, {"type": int}),
      _o(["--fmt"], "fmt", "格式", "png", enum=["png", "jpg"])], c_pdf2imgs),
    ("pdf-to-word", "PDF 转 Word（.docx）", [IN_PDF, PW, OUT, PAGES], c_pdf2word),
    ("pdf-to-excel", "PDF 转 Excel（.xlsx）", [IN_PDF, PW, OUT, PAGES], c_pdf2excel),
    ("extract-images", "导出 PDF 内嵌的原始图片",
     [IN_PDF, PW, OUTDIR,
      _o(["--fmt"], "fmt", "输出格式", "auto", enum=["auto", "png", "jpg"]),
      _o(["--min-px"], "min_px", "小于该像素的图片跳过", 16, {"type": int})],
     c_extract_images),
    ("compress", "压缩体积（按质量，可选降分辨率）",
     [IN_PDF, PW, OUT,
      _o(["--quality"], "quality", "JPEG 质量 10~90", 50, {"type": int}),
      _o(["--max-dpi"], "max_dpi", "分辨率上限 dpi，0=不动清晰度", 0, {"type": int})],
     c_compress),
    ("compress-to", "压缩到指定 MB 以内（自动二分试压）",
     [IN_PDF, PW, OUT,
      _o(["--target-mb"], "target_mb", "目标大小（MB）", 0, {"type": float})],
     c_compress_to),
    ("crop", "裁剪页面 / 自动去扫描黑边",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--mode"], "mode", "模式", "自动去黑边",
         enum=["自动去黑边", "四边留白", "恢复原始尺寸"]),
      _o(["--keep-mm"], "keep_mm", "自动裁切后保留边距（毫米）", 3, {"type": float}),
      _o(["--margins"], "margins", '四边留白模式的边距 "左,上,右,下"（毫米）')],
     c_crop),
    ("office-to-pdf", "Office 转 PDF（doc/docx/xls/xlsx/ppt/pptx/csv）",
     [_o(["--input", "-i"], "input", "Office 文件，可多个", kw={"nargs": "+"},
         required=True), OUTDIR], c_office2pdf),
    ("encrypt", "加打开密码（AES-256）",
     [IN_PDF, OUT, _o(["--password"], "password", "打开密码"),
      _o(["--owner-pw"], "owner_pw", "源 PDF 的密码（若已加密）")], c_encrypt),
    ("decrypt", "去除打开密码", [IN_PDF, PW, OUT], c_decrypt),
    ("redact", "涂黑遮盖（按关键词，同时抹除底层文字与文档残留）",
     [IN_PDF, PW, OUT,
      _o(["--keyword"], "keyword", "要遮盖的关键词"),
      _o(["--areas"], "areas", '手动框选 JSON，如 [[0,[100,100,200,130]]]')],
     c_redact),
    ("annotate", "添加标注/批注",
     [IN_PDF, PW, OUT,
      _o(["--kind"], "kind", "类型", "高亮",
         enum=["高亮", "下划线", "删除线", "波浪线", "便签", "文本框"]),
      _o(["--keyword"], "keyword", "按关键词定位")], c_annotate),
    ("ink", "画笔标记：自由手绘笔迹写进页面（内容级墨迹，任何阅读器可见）",
     [IN_PDF, PW, OUT,
      _o(["--strokes"], "strokes",
         'JSON 数组，如 [{"page":0,"color":"#E11D48","width":3,'
         '"shape":"free","points":[[72,100],[200,150]]}]；'
         'shape 可选 free/line/rect/ellipse（未旋转页面坐标，PDF 点）',
         required=True)], c_ink),
    ("save-as", "另存为新文件（原文件不动），可选顺带写画笔笔迹",
     [IN_PDF, PW, OUT,
      _o(["--strokes"], "strokes", "同 ink 命令的笔迹 JSON，可省略")],
     c_save_as),
    ("save", "覆盖保存：直接写回 --input 自身（可选顺带写画笔笔迹）",
     [IN_PDF, PW,
      _o(["--strokes"], "strokes", "同 ink 命令的笔迹 JSON，可省略")],
     c_save),
    ("insert-text", "插入文字",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--text"], "text", "要插入的文字"),
      _o(["--fontsize"], "fontsize", "字号", 12, {"type": float}),
      _o(["--position"], "position", "位置", "页面居中",
         enum=["页面居中", "页眉", "页脚", "四角", "自定义"]),
      _o(["--xy"], "xy", '自定义模式的坐标 "x,y"（毫米，原点左上）'),
      _o(["--angle"], "angle", "顺时针角度", 0, {"type": int}),
      _o(["--color"], "color", "颜色", "黑"),
      _o(["--vector"], "vector", "用可选文本（可复制）而非贴纸", False,
         {"action": "store_true"})], c_insert_text),
    ("insert-image", "插入图片",
     [IN_PDF, PW, OUT, PAGES,
      _o(["--image"], "image", "图片路径"),
      _o(["--width-pct"], "width_pct", "宽度百分比（100=图片原始物理尺寸）",
         20, {"type": float}),
      _o(["--position"], "position", "位置", "页面居中",
         enum=["页面居中", "页眉", "页脚", "四角", "自定义"]),
      _o(["--xy"], "xy", '自定义模式的坐标 "x,y"（毫米）'),
      _o(["--rotate"], "rotate", "旋转 90 的倍数", 0, {"type": int})],
     c_insert_image),
    ("toc", "读取书签/目录", [IN_PDF, PW], c_toc),
    ("toc-set", "写入书签/目录",
     [IN_PDF, PW, OUT,
      _o(["--toc"], "toc", 'JSON，如 [{"level":1,"title":"第一章","page":1}]')],
     c_toc_set),
    ("form-fill", "填写 PDF 表单字段", [IN_PDF, PW, OUT,
                                    _o(["--values"], "values", 'JSON 字段值')],
     c_form),
    ("print", "打印到物理打印机",
     [IN_PDF, PW, PAGES,
      _o(["--printer"], "printer", "打印机名（缺省用系统默认）"),
      _o(["--copies"], "copies", "份数", 1, {"type": int}),
      _o(["--output"], "output", "虚拟打印机输出到的文件")], c_print),
    ("batch", "对多个 PDF 执行同一操作",
     [MANY_PDF, OUTDIR,
      _o(["--op"], "op", "操作", None,
         enum=["添加文字水印", "添加页码", "压缩体积", "加密保护",
               "解密 / 去密码", "旋转页面"], required=True),
      _o(["--params"], "params", "该操作的参数 JSON")], c_batch),
]

RESULT = {name: fn for name, _h, _o_, fn in SPECS}
OPTS = {name: opts for name, _h, opts, _fn in SPECS}


def _namespace(name, params):
    ns = argparse.Namespace()
    for o in OPTS[name]:
        setattr(ns, o["dest"], o["default"])
    for k, v in (params or {}).items():
        setattr(ns, k, v)
    return ns


def execute(name, params):
    t0 = time.time()
    if name not in RESULT:
        return {"ok": False, "command": name, "error": "未知命令：%s" % name}
    try:
        extra = RESULT[name](_namespace(name, params)) or {}
        out = {"ok": True, "command": name}
        out.update(extra)
    except ToolError as e:
        return {"ok": False, "command": name, "error": str(e)}
    except Exception as e:
        return {"ok": False, "command": name,
                "error": "%s: %s" % (type(e).__name__, e)}
    out["elapsed_ms"] = int((time.time() - t0) * 1000)
    return out


def _emit(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False))
    sys.stdout.write("\n")
    sys.stdout.flush()


class _Parser(argparse.ArgumentParser):
    """把 argparse 的英文报错翻成中文并包进 JSON 信封。
    Agent 调错参数时不应该拿到一堆 usage 文本，而应该拿到一个能判断的 JSON。"""

    def error(self, message):
        msg = message
        if "invalid choice" in message:
            import re as _re
            m = _re.search(r"invalid choice: '([^']+)'", message)
            msg = "未知命令：%s" % (m.group(1) if m else "?")
        elif "the following arguments are required" in message:
            msg = "缺少必填参数：%s" % message.split(":")[-1].strip()
        elif "unrecognized arguments" in message:
            msg = "多余的参数：%s" % message.split(":")[-1].strip()
        elif "invalid int value" in message or "invalid float value" in message:
            msg = "参数类型不对：%s" % message.split(":")[-1].strip()
        _emit({"ok": False, "command": getattr(self, "_cmdname", None),
               "error": "%s（可用 capabilities 查看该命令的参数）" % msg})
        raise SystemExit(2)


def main(argv=None):
    ap = _Parser(
        prog="pdf_cli.py",
        description="3Q PDF开源免费工具 · Agent 入口（输出一行 JSON）。"
                    "先跑 `capabilities` 看能做什么。")
    ap._cmdname = None
    sub = ap.add_subparsers(dest="cmd", parser_class=_Parser)
    for name, help_, opts, _fn in SPECS:
        sp = sub.add_parser(name, help=help_)
        sp._cmdname = name
        for o in opts:
            sp.add_argument(*o["flags"], dest=o["dest"], default=o["default"],
                            help=o["help"] or "", **o["kw"])
    sp = sub.add_parser("run", help="从 JSON 文件/stdin 批量执行多个命令")
    sp._cmdname = "run"
    sp.add_argument("job", nargs="?",
                    help='作业文件；省略则从 stdin 读。'
                         '格式 {"jobs":[{"cmd":"merge","params":{...}}]}')

    try:
        a = ap.parse_args(argv)
    except SystemExit as e:
        return int(e.code or 0)
    if not a.cmd:
        _emit({"ok": False, "command": None,
               "error": "缺少命令。先跑 `capabilities` 看有哪些命令。"})
        return 2

    if a.cmd == "run":
        try:
            if a.job:
                with open(a.job, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                data = json.load(sys.stdin)
        except Exception as e:
            _emit({"ok": False, "command": "run", "error": "作业 JSON 读取失败：%s" % e})
            return 1
        jobs = data.get("jobs") or [data]
        results = [execute(j.get("cmd", ""), j.get("params", {})) for j in jobs]
        ok = all(r["ok"] for r in results)
        _emit({"ok": ok, "command": "run", "count": len(results),
               "results": results})
        return 0 if ok else 1

    res = execute(a.cmd, {k: v for k, v in vars(a).items() if k != "cmd"})
    _emit(res)
    return 0 if res["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
