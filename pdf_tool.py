# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (C) 2026 三扣科技 (3Q Tech)
# 本作品基于 PyMuPDF（Copyright (C) Artifex Software, Inc.，AGPL-3.0 或商业许可），
# 依 AGPL-3.0 第 5 条整体以 AGPL-3.0 授权。完整文本见同目录 LICENSE，第三方声明见 NOTICE。
"""
3Q PDF开源免费工具 V2.0（文档中心式 + 所见即所得编辑）
本文件只是**给人用的界面层**。真正的处理逻辑在同目录 pdf_core.py，
Agent 入口 pdf_cli.py 也 import 同一份 —— 改逻辑只改 core，人与 Agent 不会漂移到两份。
布局参照 Adobe Acrobat / 福昕 / WPS PDF：先打开 PDF，左侧缩略图 + 中央阅读区，
顶部工具栏所有操作均作用于当前文档。
V1.3 新增：
- 插入文字/图片改为可视化编辑：对象放置在页面上，可拖动位置、调整大小与角度，实时预览，确认后才另存新文件
- 图片转 PDF（多图合成）、PDF 转图片（导出 PNG/JPG）
- 删除页面、旋转页面、添加页码、压缩体积、加密保护
- 阅读区文字搜索（Ctrl+F）、旋转视图、打开加密 PDF、最近打开文件
V1.4 新增：
- 解密/去密码：去除打开密码与权限限制
- 添加水印：文字/图片水印，支持角度、透明度、平铺或居中、页范围
- 页面管理（重排）：上移/下移/复制/删除页面、插入空白页
- 插入文字可选「可选文本」：以矢量方式插入，文字可搜索/复制（默认为贴纸式，所见即所得）
V1.5 新增：
- PDF 转 Word（.docx）：抽取文字与图片，按页生成段落
- PDF 转 Excel（.xlsx）：抽取可识别表格，每张表一个工作表；无表格的页按文本行写入
- 书签/目录管理：查看、新增、编辑、删除、调整层级与目标页
- 表单填写：检测可填写字段并填入内容，另存为可正常显示填写结果的副本
- 批量处理：对多个 PDF 一次执行同一操作（水印/页码/压缩/加密/解密/旋转）
V1.6 新增：
- 标注/批注：高亮、下划线、删除线、波浪线，以及便签、文本框批注，可按关键词自动定位或手动框选
- 涂黑遮盖：框选或按关键词遮盖敏感内容，同时抹除下方文字（不可复制）
- 页眉页脚：自定义页眉/页脚文字，支持 {page}/{total} 页码占位符、位置与页范围
- 提取文档内图片：导出每页嵌入的原始图片到文件夹
V1.7 新增：
- 全新现代视觉：深色主视觉顶栏、品牌蓝主按钮、卡片式面板、扁平控件与悬停反馈、
  统一微软雅黑 UI 字体与应用图标；功能逻辑与坐标换算完全不变
V1.8 新增/修复：
- 修复：旋转页（/Rotate 90/180/270）上页码与页眉页脚错位、文字竖排的问题，
  改为在「显示坐标系」定位后反算回未旋转坐标插入
- 修复：图片转 PDF 与插入图片忽略图片自身 DPI，300dpi 扫描件被放大约 3 倍
- 修复：涂黑遮盖只抹页面文字，文档属性/书签/XMP 里仍残留原文（遮盖≠净化）
- 修复：PDF 转 Excel 把"12000"写成文本，无法求和
- 修复：合并时单个文件出错会漏关文件句柄
- 新增：阅读区文字选择与复制（Ctrl+C / 右键菜单），旋转页同样正确命中
- 新增：手写签名（鼠标/数位笔书写 → 透明 PNG → 页面内拖动缩放定位）
- 修复：全文搜索只往后扫不回绕，前进到末页就报「未找到」
- 修复：菜单按钮的下拉箭头在高 DPI 下被画成巨大的实心三角
V1.9 新增：
- 拖放：PDF / 图片直接拖进窗口 —— 单个直接打开、多个自动进入合并、
  纯图片进「图片转 PDF」、图文混合先把图片合成 PDF 再一起合并
- 左侧新增「功能」页签：按「把资料变成 PDF / 改页面 / 加内容 / 发出去 / 转换与导出」
  分组，不用再翻「更多工具」二级菜单；未打开文档时依赖文档的按钮自动置灰
- 拖放是可选依赖（tkinterdnd2），未安装时自动降级，不影响启动
- 修复：文字宽度口径错误 —— fitz.Font("china-ss").text_length 量得 70.47，
  实际落墨 130.0（低报约 1.8 倍），导致页码/页眉页脚「居中」右偏约 30pt、
  长文字不触发缩字号而溢出页边；改用 fitz.get_text_length 后逐点吻合
- 修复：拖放路径解析增加字面分词兜底，下发方漏转义时不再静默丢文件
V2.0 新增：
- 压缩到指定大小：填一个目标 MB，自动在「画质 → 分辨率」两档参数上二分试压，
  取仍能装进目标的最高画质；文字型 PDF 压不下去时如实告知，不假装成功
- 修复：压缩形同虚设 —— PyMuPDF 1.25.2 根本没有 Document.rewrite_images，
  老代码用 hasattr 挡着又把异常 pass 掉，实测 5.33MB 压完只瘦 344 字节。
  改为自建逐图重压（按显示尺寸算有效 dpi 降采样 + JPEG 重编码），同一样本压到 0.63MB；
  「压缩体积」新增分辨率上限，0 = 只重编码不动清晰度
- Office 转 PDF：doc/docx/xls/xlsx/ppt/pptx/csv 走本机 Office（COM）批量转换，
  排版与 Office 另存为 PDF 一致；缺 pywin32 或没装 Office 时给出明确处置建议
- 图片转 PDF 增强：缩略图预览、单张旋转（左/右转 90°、归正）、按住行拖动排序，
  并自动纠正手机照片的 EXIF 方向（MuPDF 贴图不应用 EXIF，竖拍照片会横躺）
- 页面裁剪 / 去黑边：按内容边界自动裁掉扫描黑边（可保留边距）、四边留白内缩、
  以及「恢复原始尺寸」撤销以前的裁剪
纯本地运行，不修改原文件，所有结果另存新文件。
"""
import io
import json
import math
import os
import queue
import re
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import fitz  # PyMuPDF

# ---- 核心逻辑在同目录 pdf_core.py（GUI 与 Agent CLI 共用一份，改动只改一处）----
from pdf_core import *
from pdf_core import (  # import * 不带上划线私有名，这里显式补齐
    _NUM_CELL_RE, _TEXT_CACHE, _all_pages,
    _apply_alpha, _cjk_font_bytes, _fields_from_doc,
    _fit_fontsize, _image_display_rects, _is_cjk,
    _pil_font, _recompress_images, _rot_map,
    _rot_unmap, _watermark_tile,
)

# 拖放支持为可选依赖：装了 tkinterdnd2 就启用拖入文件，没装也不影响启动
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:  # pragma: no cover - 环境相关
    TkinterDnD = None
    DND_FILES = None
    HAS_DND = False

# ---------------- 现代主题（统一视觉，只改观感不动逻辑） ----------------
class Theme:
    APP_BG = "#EDF1F8"      # 窗口底色（浅蓝灰）
    CARD = "#FFFFFF"        # 卡片 / 面板
    BORDER = "#D5DEEC"      # 边框
    TEXT = "#1B2537"        # 主文字
    MUTED = "#7E8AA0"       # 次要文字
    PRIMARY = "#2563EB"     # 主色（品牌蓝）
    PRIMARY_H = "#1D4ED8"   # 悬停
    PRIMARY_P = "#1E40AF"   # 按下
    PRIMARY_S = "#DBEAFE"   # 选中底色
    OK = "#0E9F6E"          # 成功 / 状态提示
    DANGER = "#DC2626"      # 高危操作警示（物理打印等不可撤销动作）
    DARK = "#0F1A2E"        # 顶栏 / 状态栏（更深，稳重）
    DARK2 = "#1B2C4C"       # 顶栏渐变浅端
    DARK_H = "#26344F"      # 深色按钮悬停
    DARK_P = "#1A2740"      # 深色按钮按下
    CANVAS_BG = "#2E3644"   # 阅读区底色（更深，白页更突出）
    EDITOR_BG = "#EAF1FD"   # 编辑栏底色
    GOLD = "#C9A25E"        # 高级金（赞助 / 点缀）
    GOLD_S = "#E7D2A6"      # 浅金
    GOLD_D = "#9C7B3E"      # 深金

# 工具面板分组强调色：六组轮流取色，整屏不再只有一种蓝
_GROUP_ACCENTS = ("#2563EB", "#0E9F6E", "#7C3AED",
                  "#D97706", "#E11D48", "#0891B2")

_UI_FONTS = {}

# 右侧工具面板中「不需要打开文档」也能用的工具（_sync_task_buttons 联动置灰用）
_PANEL_NO_DOC = frozenset(("合并…", "图片转 PDF…", "Office 转 PDF…", "批量处理…"))


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "#%02X%02X%02X" % tuple(
        max(0, min(255, int(round(v)))) for v in rgb)


def _ramp(colors, t):
    """多段色带：t∈[0,1] 处插值出的颜色（用于模拟渐变）。"""
    if t <= 0:
        return colors[0]
    if t >= 1:
        return colors[-1]
    seg = len(colors) - 1
    x = t * seg
    i = int(x)
    f = x - i
    a = _hex_to_rgb(colors[i])
    b = _hex_to_rgb(colors[i + 1])
    return _rgb_to_hex([a[k] + (b[k] - a[k]) * f for k in range(3)])


def _gradient_bar(parent, colors, height=3, steps=96):
    """纯 Tk 渐变细条：用 place 均分色块，自适应任意宽度。

    Tk 没有原生渐变，这里用若干细色块视觉合成，成本极低（每个只占
    1/96 宽度），用于顶栏棱彩线、金色分隔线等点缀。
    """
    bar = tk.Frame(parent, height=height, bd=0, highlightthickness=0,
                   bg=colors[0])
    bar.pack(fill="x")
    bar.pack_propagate(False)
    for i in range(steps):
        c = _ramp(colors, i / (steps - 1))
        tk.Frame(bar, bg=c, bd=0, highlightthickness=0).place(
            relx=i / steps, rely=0, relwidth=1.02 / steps, relheight=1.0)
    return bar


def ui_font_family(root):
    """优先使用现代无衬线中文字体；找不到回退默认。"""
    try:
        from tkinter import font as tkfont
        fams = set(tkfont.families(root))
    except Exception:
        return ""
    for f in ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC",
              "Noto Sans CJK SC", "Segoe UI"):
        if f in fams:
            return f
    return ""


def _entry_font_ok(style, root, f):
    """探测当前 Tk 是否支持用 style 配置 ttk 输入框字体。"""
    try:
        style.configure("TEntry", font=f)
        e = ttk.Entry(root)
        ok = f[0] in str(e.cget("font"))
        e.destroy()
        return ok
    except Exception:
        return False


def setup_style(root):
    """clam 基座 + 统一配色 / 字体 / 扁平控件（悬停反馈）。"""
    fam = ui_font_family(root)
    s = ttk.Style(root)
    try:
        s.theme_use("clam")
    except tk.TclError:
        pass
    # 输入框字体可样式化时正文 10 号，否则全局面 9 号保持一致
    size = 10 if (fam and _entry_font_ok(s, root, (fam, 10))) else 9
    base = (fam, size) if fam else None
    small = (fam, max(8, size - 1)) if fam else None
    bold = (fam, size, "bold") if fam else None
    title = (fam, size + 1, "bold") if fam else None
    _UI_FONTS.update(base=base, small=small, bold=bold, title=title)
    if fam:
        root.option_add("*Font", base)  # tk 控件：Listbox / Menu / 弹窗输入
        root.option_add("*TCombobox*Listbox.font", small)
        root.option_add("*TCombobox*Listbox.background", Theme.CARD)
        root.option_add("*TCombobox*Listbox.foreground", Theme.TEXT)
        root.option_add("*TCombobox*Listbox.selectBackground", Theme.PRIMARY_S)
        root.option_add("*TCombobox*Listbox.selectForeground", Theme.TEXT)
        s.configure(".", font=base)
    s.configure(".", background=Theme.CARD, foreground=Theme.TEXT,
                bordercolor=Theme.BORDER, focuscolor=Theme.PRIMARY)

    # ---- 帧与标签 ----
    s.configure("TFrame", background=Theme.CARD)
    s.configure("App.TFrame", background=Theme.APP_BG)
    s.configure("Card.TFrame", background=Theme.CARD)
    s.configure("Toolbar.TFrame", background=Theme.DARK)
    s.configure("Status.TFrame", background=Theme.DARK)
    s.configure("Editor.TFrame", background=Theme.EDITOR_BG)
    s.configure("TLabel", background=Theme.CARD, foreground=Theme.TEXT)
    s.configure("Muted.TLabel", background=Theme.CARD, foreground=Theme.MUTED)
    s.configure("App.TLabel", background=Theme.APP_BG, foreground=Theme.MUTED)
    s.configure("Title.TLabel", background=Theme.CARD, foreground=Theme.TEXT)
    s.configure("Toolbar.TLabel", background=Theme.DARK, foreground="#9AAECF")
    s.configure("ToolbarBadge.TLabel", background="#22314F", foreground="#C3D2EE",
                padding=(10, 3))
    s.configure("Status.TLabel", background=Theme.DARK, foreground="#9AAECF")
    s.configure("StatusBright.TLabel", background=Theme.DARK, foreground="#E9EEF9")
    s.configure("Nav.TLabel", background=Theme.CARD, foreground=Theme.PRIMARY)
    if bold:
        s.configure("Nav.TLabel", font=bold)
    s.configure("Editor.TLabel", background=Theme.EDITOR_BG, foreground=Theme.TEXT)
    if small:
        s.configure("Muted.TLabel", font=small)
    if title:
        s.configure("Title.TLabel", font=title)

    # ---- 按钮：扁平 + 悬停反馈 ----
    s.configure("TButton", background=Theme.CARD, foreground=Theme.TEXT,
                bordercolor=Theme.BORDER, borderwidth=1, relief="flat",
                focusthickness=0, padding=(10, 4))
    s.map("TButton",
          background=[("disabled", "#F0F3F9"), ("pressed", "#E4EAF5"),
                      ("active", "#F2F5FB")],
          foreground=[("disabled", "#A9B3C6")],
          bordercolor=[("pressed", Theme.PRIMARY), ("active", "#B9C7DE")])
    s.configure("Accent.TButton", background=Theme.PRIMARY, foreground="#FFFFFF",
                borderwidth=0, padding=(11, 4))
    if bold:
        s.configure("Accent.TButton", font=bold)
    s.map("Accent.TButton",
          background=[("disabled", "#A6C1F5"), ("pressed", Theme.PRIMARY_P),
                      ("active", Theme.PRIMARY_H)],
          foreground=[("disabled", "#ECF2FE")])
    s.configure("Ghost.TButton", background=Theme.DARK, foreground="#D8E0F0",
                borderwidth=0, focuscolor="#FFFFFF", padding=(9, 3))
    s.map("Ghost.TButton",
          background=[("disabled", Theme.DARK), ("pressed", Theme.DARK_P),
                      ("active", Theme.DARK_H)],
          foreground=[("disabled", "#5E6D8C")])
    # 赞助入口：香槟金，在深色工具栏上一眼能看到，又不至于刺眼
    s.configure("Donate.TButton", background=Theme.GOLD, foreground="#1F1606",
                borderwidth=0, focuscolor="#FFFFFF", padding=(11, 4))
    if bold:
        s.configure("Donate.TButton", font=bold)
    s.map("Donate.TButton",
          background=[("disabled", "#5E4A28"), ("pressed", Theme.GOLD_D),
                      ("active", Theme.GOLD_S)],
          foreground=[("disabled", "#C9B99B")])
    # 白色卡片上的小按钮（工具面板「折叠 / 收起」用它，别用 Ghost——那是深色底专用）
    s.configure("PanelGhost.TButton", background=Theme.CARD,
                foreground=Theme.MUTED, borderwidth=0, focuscolor=Theme.CARD,
                padding=(6, 3))
    s.map("PanelGhost.TButton",
          background=[("pressed", "#E4ECFA"), ("active", "#EFF4FC")],
          foreground=[("pressed", Theme.PRIMARY), ("active", Theme.PRIMARY)])
    # 工具面板专用滚动条：加宽、白槽、灰蓝滑块，白色卡片上看得出来
    s.configure("Panel.Vertical.TScrollbar", background="#C3CEE2",
                troughcolor=Theme.CARD, bordercolor=Theme.CARD,
                arrowcolor="#6B7891", arrowsize=12, gripcount=0,
                relief="flat", width=11)
    s.map("Panel.Vertical.TScrollbar",
          background=[("pressed", "#93A2BC"), ("active", "#AAB7CD")])

    # ---- 菜单按钮 ----
    s.configure("TMenubutton", background=Theme.CARD, foreground=Theme.TEXT,
                bordercolor=Theme.BORDER, borderwidth=1, relief="flat",
                focusthickness=0, arrowsize=12, padding=(12, 5))
    s.map("TMenubutton",
          background=[("disabled", "#F0F3F9"), ("pressed", "#E4EAF5"),
                      ("active", "#F2F5FB")],
          foreground=[("disabled", "#A9B3C6")],
          arrowcolor=[("disabled", "#A9B3C6")])
    s.configure("Ghost.TMenubutton", background=Theme.DARK, foreground="#D8E0F0",
                borderwidth=0, arrowsize=12, padding=(12, 5))
    s.map("Ghost.TMenubutton",
          background=[("pressed", Theme.DARK_P), ("active", Theme.DARK_H)],
          arrowcolor=[("active", "#FFFFFF")])
    # 去掉主题自带的下拉箭头元素：它在高 DPI 下会被画成巨大的实心三角，
    # 这里改用文字里的小箭头，视觉才干净。菜单弹出行为不受影响。
    try:
        s.layout("Ghost.TMenubutton", [
            ("Menubutton.padding", {"sticky": "nswe", "children": [
                ("Menubutton.label", {"sticky": "nswe"})]})])
    except Exception:
        pass

    # ---- 左侧页签（页面 / 功能） ----
    s.configure("Side.TNotebook", background=Theme.CARD, borderwidth=0,
                tabmargins=(4, 5, 4, 0))
    s.configure("Side.TNotebook.Tab", padding=(13, 4), background="#EDF2FA",
                foreground=Theme.MUTED, borderwidth=0)
    s.map("Side.TNotebook.Tab",
          background=[("selected", Theme.CARD)],
          foreground=[("selected", Theme.PRIMARY)])
    # 侧栏任务按钮：浅底 + 细边，才有「可以点」的观感
    s.configure("Task.TButton", background="#F5F8FD", foreground=Theme.TEXT,
                bordercolor="#DCE5F2", borderwidth=1, relief="flat",
                padding=(9, 5), anchor="w")
    s.map("Task.TButton",
          background=[("pressed", "#E3ECFA"), ("active", "#EBF2FC")],
          bordercolor=[("active", Theme.PRIMARY)],
          foreground=[("active", Theme.PRIMARY)])

    # ---- 输入类 ----
    s.configure("TEntry", fieldbackground=Theme.CARD, foreground=Theme.TEXT,
                insertcolor=Theme.TEXT, bordercolor=Theme.BORDER,
                lightcolor=Theme.BORDER, darkcolor=Theme.BORDER,
                selectbackground=Theme.PRIMARY_S, selectforeground=Theme.TEXT,
                padding=(6, 4))
    s.map("TEntry",
          bordercolor=[("focus", Theme.PRIMARY)],
          lightcolor=[("focus", Theme.PRIMARY)],
          darkcolor=[("focus", Theme.PRIMARY)],
          fieldbackground=[("disabled", "#F0F3F9")],
          foreground=[("disabled", "#A9B3C6")])
    s.configure("TCombobox", fieldbackground=Theme.CARD, background="#E6ECF6",
                foreground=Theme.TEXT, bordercolor=Theme.BORDER,
                lightcolor=Theme.BORDER, darkcolor=Theme.BORDER,
                arrowcolor=Theme.TEXT, arrowsize=13, padding=(6, 4))
    s.map("TCombobox",
          bordercolor=[("focus", Theme.PRIMARY)],
          lightcolor=[("focus", Theme.PRIMARY)],
          darkcolor=[("focus", Theme.PRIMARY)],
          fieldbackground=[("readonly", Theme.CARD), ("disabled", "#F0F3F9")],
          foreground=[("disabled", "#A9B3C6")],
          arrowcolor=[("disabled", "#A9B3C6")])
    s.configure("TSpinbox", fieldbackground=Theme.CARD, background="#E6ECF6",
                foreground=Theme.TEXT, bordercolor=Theme.BORDER,
                lightcolor=Theme.BORDER, darkcolor=Theme.BORDER,
                arrowcolor=Theme.TEXT, insertcolor=Theme.TEXT, padding=(6, 3))
    s.map("TSpinbox",
          bordercolor=[("focus", Theme.PRIMARY)],
          lightcolor=[("focus", Theme.PRIMARY)],
          darkcolor=[("focus", Theme.PRIMARY)],
          fieldbackground=[("disabled", "#F0F3F9")],
          foreground=[("disabled", "#A9B3C6")],
          arrowcolor=[("disabled", "#A9B3C6")])

    # ---- 单选 / 复选 ----
    for st_ in ("TRadiobutton", "TCheckbutton"):
        s.configure(st_, background=Theme.CARD, foreground=Theme.TEXT,
                    focuscolor=Theme.PRIMARY, indicatorbackground=Theme.CARD,
                    indicatorforeground=Theme.PRIMARY, padding=(4, 3))
        s.map(st_,
              background=[("active", "#F2F5FB")],
              indicatorbackground=[("pressed", Theme.PRIMARY),
                                   ("active", Theme.PRIMARY_S),
                                   ("disabled", "#F0F3F9")],
              foreground=[("disabled", "#A9B3C6")])
    s.configure("Editor.TCheckbutton", background=Theme.EDITOR_BG,
                foreground=Theme.TEXT, focuscolor=Theme.PRIMARY,
                indicatorbackground=Theme.EDITOR_BG,
                indicatorforeground=Theme.PRIMARY)
    s.map("Editor.TCheckbutton",
          background=[("active", "#D9E8FD")],
          indicatorbackground=[("pressed", Theme.PRIMARY),
                               ("active", Theme.PRIMARY_S)])

    # ---- 滚动条 / 分隔线 / 进度 / 分组框 ----
    s.configure("TScrollbar", background="#DCE4F0", troughcolor=Theme.APP_BG,
                bordercolor=Theme.APP_BG, arrowcolor="#8291AB", arrowsize=13,
                gripcount=0, relief="flat")
    s.map("TScrollbar",
          background=[("pressed", "#B7C4DA"), ("active", "#C6D1E4")],
          arrowcolor=[("active", Theme.TEXT)])
    s.configure("TSeparator", background="#E3E9F3")
    s.configure("Toolbar.TSeparator", background="#35436A")
    s.configure("TProgressbar", background=Theme.PRIMARY, troughcolor="#DFE6F2",
                borderwidth=0, lightcolor=Theme.PRIMARY, darkcolor=Theme.PRIMARY)
    s.configure("TLabelframe", background=Theme.CARD, bordercolor=Theme.BORDER,
                borderwidth=1, relief="flat")
    s.configure("TLabelframe.Label", background=Theme.CARD, foreground=Theme.TEXT)
    if bold:
        s.configure("TLabelframe.Label", font=bold)


def style_menu(m):
    """统一弹出菜单观感。"""
    kw = dict(background=Theme.CARD, foreground=Theme.TEXT,
              activebackground=Theme.PRIMARY_S, activeforeground=Theme.PRIMARY,
              bd=0, relief="flat", activeborderwidth=0)
    if _UI_FONTS.get("small"):
        kw["font"] = _UI_FONTS["small"]
    try:
        m.configure(**kw)
    except tk.TclError:
        pass


def themed_listbox(master, **kw):
    """统一列表框观感。"""
    return tk.Listbox(master, bg=Theme.CARD, fg=Theme.TEXT,
                      selectbackground=Theme.PRIMARY_S,
                      selectforeground="#1E3A8A", activestyle="none",
                      relief="flat", bd=0, highlightthickness=1,
                      highlightbackground=Theme.BORDER,
                      highlightcolor=Theme.PRIMARY, **kw)


def make_icon_png():
    """PIL 现场绘制应用图标（蓝色圆角 + PDF 字样），返回 PNG 字节，无新依赖。"""
    try:
        from PIL import Image, ImageDraw
        s_ = 8
        img = Image.new("RGBA", (32 * s_, 32 * s_), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([2 * s_, 0, 30 * s_, 32 * s_], radius=6 * s_,
                            fill=(37, 99, 235, 255))
        f = _pil_font(12 * s_)
        l, t, r, b = f.getbbox("PDF")
        d.text(((32 * s_ - (r - l)) / 2 - l, (32 * s_ - (b - t)) / 2 - t),
               "PDF", font=f, fill=(255, 255, 255, 255))
        img = img.resize((32, 32), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, "PNG")
        return buf.getvalue()
    except Exception:
        return None


def load_recent():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("recent", [])
    except Exception:
        return []


def add_recent(path):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        items = [p for p in load_recent() if p != path]
        items.insert(0, path)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"recent": items[:10]}, f, ensure_ascii=False)
    except Exception:
        pass


# ============================================================
# 界面层（文档中心式）
# ============================================================

class Runner:
    """子线程执行 + 状态提示 + 结果弹窗（SR-04）。
    线程安全：tk 控件/变量只在主线程访问；子线程结果经队列由主线程轮询消费。"""
    app = None
    run_btn = None

    def set_status(self, text):
        if self.status_var is not None:
            self.status_var.set(text)

    def on_progress(self, done, total):
        """子任务进度回调（主线程执行）。需要进度显示的宿主可覆盖。"""

    def run_in_thread(self, work, on_done):
        """work() 在子线程执行；on_done(result) 返回 (标题, 消息, 文件路径, 文件夹)。
        work 内可用 self.report_progress(d, t) 上报进度。"""
        self._q = queue.Queue()
        if self.run_btn is not None:
            self.run_btn.config(state="disabled", text="处理中…")
        self.set_status("正在处理，请稍候…")

        def worker():
            try:
                self._q.put(("ok", work()))
            except Exception as e:
                self._q.put(("err", str(e)))

        threading.Thread(target=worker, daemon=True).start()
        self._poll(on_done)

    def report_progress(self, done, total):
        """供 work() 在子线程中调用，转发进度到主线程。"""
        self._q.put(("prog", (done, total)))

    def _poll(self, on_done):
        if not self.app.root.winfo_exists():
            return
        try:
            kind, payload = self._q.get_nowait()
        except queue.Empty:
            self.app.root.after(80, lambda: self._poll(on_done))
            return
        if kind == "prog":
            self.on_progress(*payload)
            self._poll(on_done)
            return
        if kind == "err":
            self._on_error(payload)
        else:
            self._on_done(payload, on_done)

    def _restore_btn(self):
        """恢复按钮状态；宿主窗口已关闭时静默跳过。"""
        if self.run_btn is None:
            return
        try:
            self.run_btn.config(state="normal", text=self._run_text)
        except tk.TclError:
            pass

    def _on_error(self, msg):
        self._restore_btn()
        self.set_status("")
        if self.app.root.winfo_exists():
            messagebox.showerror("处理失败", msg, parent=self.app.root)

    def _on_done(self, result, on_done):
        self._restore_btn()
        self.set_status("处理完成")
        title, message, file_path, folder = on_done(result)
        if self.app.root.winfo_exists():
            ResultDialog(self.app.root, title, message, file_path, folder)


class ResultDialog(tk.Toplevel):
    """SR-04 完成提示：显示输出位置 + 打开文件 / 打开所在文件夹。"""

    def __init__(self, master, title, message, file_path=None, folder=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        self.configure(bg=Theme.APP_BG)
        strip = tk.Frame(self, bg=Theme.OK, height=3, bd=0, highlightthickness=0)
        strip.pack(fill="x")
        strip.pack_propagate(False)
        card = ttk.Frame(self, style="Card.TFrame", padding=(20, 16))
        card.pack(fill="both", expand=True)
        ttk.Label(card, text=message, style="Title.TLabel", wraplength=420,
                  justify="left").pack(anchor="w")
        show_path = file_path or folder or ""
        if show_path:
            ttk.Label(card, text=show_path, style="Muted.TLabel", wraplength=420,
                      justify="left").pack(anchor="w", pady=(8, 0))
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(pady=(16, 2))
        if file_path:
            ttk.Button(btns, text="打开文件", style="Accent.TButton",
                       command=lambda: os.startfile(file_path)).pack(
                side="left", padx=(0, 8))
        if folder:
            ttk.Button(btns, text="打开所在文件夹",
                       command=lambda: os.startfile(folder)).pack(side="left", padx=8)
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="left", padx=8)
        self.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


class HighRiskDialog(tk.Toplevel):
    """高危操作确认框（MCP 智能体触发，人工裁决）。

    出现场景：仅本机档的服务收到 print 之类会真的出纸的调用。
    在 Tk 主线程创建并模态运行；结果写在 self.result。
    """

    def __init__(self, master, name, args, reason):
        super().__init__(master)
        self.title("高危操作确认")
        self.resizable(False, False)
        self.transient(master)
        self.configure(bg=Theme.APP_BG)
        self.result = False
        strip = tk.Frame(self, bg=Theme.DANGER, height=3, bd=0,
                         highlightthickness=0)
        strip.pack(fill="x")
        strip.pack_propagate(False)
        card = ttk.Frame(self, style="Card.TFrame", padding=(20, 16))
        card.pack(fill="both", expand=True)
        ttk.Label(card, text="智能体请求执行高危操作",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(card, text=reason, style="Muted.TLabel", wraplength=420,
                  justify="left").pack(anchor="w", pady=(4, 8))
        detail = "命令：%s\n参数：%s" % (
            name, json.dumps(args, ensure_ascii=False, default=str)
            if args else "（无）")
        ttk.Label(card, text=detail, style="Muted.TLabel", wraplength=420,
                  justify="left").pack(anchor="w")
        ttk.Label(card, text="请在确认你了解后果后再放行。",
                  style="Muted.TLabel", wraplength=420,
                  justify="left").pack(anchor="w", pady=(8, 0))
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(pady=(16, 2))
        ttk.Button(btns, text="取消（推荐）", style="Accent.TButton",
                   command=self._deny).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="仍然执行", command=self._allow).pack(
            side="left", padx=8)
        self.protocol("WM_DELETE_WINDOW", self._deny)
        self.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width()
                                    - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height()
                                    - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _allow(self):
        self.result = True
        self.destroy()

    def _deny(self):
        self.result = False
        self.destroy()


# ---------------- 智能体接入（MCP）面板 ----------------
try:
    import pdf_mcp
    import pdf_cli
    _MCP_ERR = ""
except Exception as _e:      # MCP 不可用不能把整个界面拖垮
    pdf_mcp = None
    pdf_cli = None
    _MCP_ERR = "%s: %s" % (type(_e).__name__, _e)


class AboutDialog(tk.Toplevel):
    """「建议与赞助」窗口：品牌、开发者声明、联系方式、意见反馈与打赏入口。

    打赏完全自愿，不是使用本软件的前提（AGPL 不允许把赞助做成门槛）。
    若程序目录 / _MEIPASS 放了「打赏码.png」，直接在窗口里展示；
    没有则显示文字指引，不报错。
    """

    BRAND = "三扣科技"
    MEDIA = "3Q科技"
    QQ = "397218026"

    def __init__(self, app, auto_close=0):
        """auto_close>0：N 秒后自动关闭（欢迎弹窗用）；0 = 常驻（手动打开）。"""
        super().__init__(app.root)
        self.title("建议与赞助 · 3Q PDF开源免费工具")
        self.resizable(False, False)
        self.transient(app.root)
        self.configure(bg=Theme.APP_BG)

        fam = ui_font_family(self) or "Microsoft YaHei UI"

        # 先把打赏码读进来（找不到就退回文字指引，不报错）
        self._qr = None
        qr_path = self._find_reward_qr()
        if qr_path:
            try:
                self._qr = tk.PhotoImage(file=qr_path)
            except Exception:
                self._qr = None

        # ---- 顶部主视觉：渐变横幅（深蓝→靛→紫，底部压一道金线）----
        hero = tk.Canvas(self, height=92, bd=0, highlightthickness=0,
                         bg=Theme.DARK)
        hero.pack(fill="x")

        def _paint_hero(_e=None):
            hero.delete("all")
            w = max(hero.winfo_width(), 440)
            h = 92
            cols = (Theme.DARK, "#1E3A8A", "#4F46E5", "#7C3AED")
            for i in range(120):
                hero.create_rectangle(
                    i * w / 120, 0, (i + 1) * w / 120 + 1, h,
                    outline="", fill=_ramp(cols, i / 119))
            hero.create_rectangle(0, h - 3, w, h, outline="", fill=Theme.GOLD)
            hero.create_text(26, 33, anchor="w", fill="#FFFFFF",
                             text="3Q PDF开源免费工具",
                             font=(fam, 17, "bold"))
            hero.create_text(26, 61, anchor="w", fill="#DCE6FF",
                             text="三扣科技 出品 · 本地运行，文件不上传",
                             font=(fam, 10))

        hero.bind("<Configure>", _paint_hero)
        _paint_hero()

        body = ttk.Frame(self, style="Card.TFrame", padding=(24, 18, 24, 14))
        body.pack(fill="both", expand=True)

        cols = ttk.Frame(body, style="Card.TFrame")
        cols.pack(fill="both", expand=True)

        # ---- 左栏：关于 / 意见与建议 / 联系方式 ----
        left = ttk.Frame(cols, style="Card.TFrame")
        left.pack(side="left", fill="both", expand=True, padx=(0, 16))

        ttk.Label(left, text="关于与建议", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="这是一款本地处理的 PDF 开源免费工具：文件只在你电脑上转换，"
                 "不会上传到任何服务器。基于 PyMuPDF（AGPL-3.0），"
                 "整体以 AGPL-3.0 开源，源码随安装包附带。",
            style="Muted.TLabel", wraplength=300, justify="left",
        ).pack(anchor="w", pady=(6, 12))

        ttk.Label(left, text="有建议、遇到问题，或想要新功能？",
                  style="Muted.TLabel").pack(anchor="w")
        ttk.Label(left, text="欢迎加 QQ，或关注公众号「3Q科技」留言。",
                  style="Muted.TLabel", wraplength=300,
                  justify="left").pack(anchor="w", pady=(0, 10))

        box = ttk.Frame(left, style="Card.TFrame")
        box.pack(anchor="w")
        ttk.Label(box, text="开发者：", style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", pady=2)
        ttk.Label(box, text=self.BRAND, style="Card.TLabel").grid(
            row=0, column=1, sticky="w")
        ttk.Label(box, text="联系 QQ：", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=2)
        ttk.Label(box, text=self.QQ, style="Card.TLabel").grid(
            row=1, column=1, sticky="w")
        ttk.Button(box, text="复制", width=5, style="Ghost.TButton",
                   command=self._copy_qq).grid(row=1, column=2, padx=(8, 0))
        ttk.Label(box, text="公众号：", style="Muted.TLabel").grid(
            row=2, column=0, sticky="w", pady=2)
        ttk.Label(box, text=self.MEDIA, style="Card.TLabel").grid(
            row=2, column=1, sticky="w")

        # ---- 右栏：赞助卡片（暖色，突出但不强推）----
        warm_bg = "#FFF8EC"
        right = tk.Frame(cols, bg=warm_bg, bd=0,
                         highlightthickness=1, highlightbackground=Theme.GOLD)
        right.pack(side="right", fill="y")
        inner = tk.Frame(right, bg=warm_bg)
        inner.pack(padx=18, pady=16)

        tk.Label(inner, text="请作者喝杯咖啡", bg=warm_bg, fg="#B45309",
                 font=(fam, 13, "bold")).pack(anchor="center")
        tk.Label(inner, text="如果它替你\n省下了几分钟", bg=warm_bg,
                 fg="#8A5A16", font=(fam, 9), justify="center").pack(
                     anchor="center", pady=(4, 10))

        if self._qr is not None:
            qwrap = tk.Frame(inner, bg="#FFFFFF", bd=0,
                             highlightthickness=1,
                             highlightbackground="#E7C99B")
            qwrap.pack(anchor="center")
            tk.Label(qwrap, image=self._qr, bg="#FFFFFF", bd=0).pack(
                padx=7, pady=7)
        else:
            tk.Label(inner, text="（打赏码见发布页或「%s」主页）" % self.MEDIA,
                     bg=warm_bg, fg="#A97B3C", font=(fam, 9),
                     wraplength=210, justify="center").pack(anchor="center")

        tk.Label(inner, text="¥5.20 · ¥8.88 · ¥20 · 任意金额",
                 bg=warm_bg, fg="#8A5A16", font=(fam, 9)).pack(
                     anchor="center", pady=(10, 0))
        tk.Label(inner, text="打赏完全自愿 · 不影响任何功能",
                 bg=warm_bg, fg="#A97B3C", font=(fam, 8)).pack(
                     anchor="center", pady=(3, 0))
        tk.Label(inner, text="谢谢你的支持", bg=warm_bg, fg="#B45309",
                 font=(fam, 9, "bold")).pack(anchor="center", pady=(6, 0))

        # 倒计时（欢迎弹窗用）
        self._auto = int(auto_close)
        if self._auto > 0:
            self._cd_lbl = ttk.Label(body, text="", style="Muted.TLabel")
            self._cd_lbl.pack(anchor="w", pady=(10, 0))
            self._tick()

        self.update_idletasks()
        x = app.root.winfo_rootx() + (app.root.winfo_width()
                                      - self.winfo_width()) // 2
        y = app.root.winfo_rooty() + (app.root.winfo_height()
                                      - self.winfo_height()) // 2
        self.geometry("+%d+%d" % (max(0, x), max(0, y)))

    def _tick(self):
        """倒计时到 0 自动关闭；窗口已销毁/应用已退出则静默退出。"""
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        if self._auto <= 0:
            self.destroy()
            return
        self._cd_lbl.config(text="（%d 秒后自动关闭…）" % self._auto)
        self._auto -= 1
        self.after(1000, self._tick)

    @staticmethod
    def _find_reward_qr():
        for base in (getattr(sys, "_MEIPASS", ""),
                     os.path.dirname(os.path.abspath(__file__))):
            if not base:
                continue
            p = os.path.join(base, "打赏码.png")
            if os.path.isfile(p):
                return p
        return ""

    def _copy_qq(self):
        self.clipboard_clear()
        self.clipboard_append(self.QQ)


class McpDialog(tk.Toplevel):
    """「智能体接入」：关闭 / 仅本机 / 远程智能体接入 / 手动导出。

    四档为什么这么定：
    - 仅本机：服务只听 127.0.0.1，全量工具。智能体和 PDF 都在你机器上，
      文件一步不动。
    - 远程智能体接入：开一个**只出代码**的 HTTP MCP 服务（code_only，
      仅 pdf-code / pdf_source 两个工具，不读写本机任何文件），绑定到
      局域网地址。远程智能体连上来自己 tools/list 发现工具，需要处理
      PDF 时调 pdf-code 拿到可执行脚本，保存到**它自己**机器上执行——
      文件不经过你的电脑，也不需要你手动发代码。因为不碰文件，监听
      非回环地址是安全的。
    - 手动导出（备用）：把生成的代码或整套分发包存成文件，自己发给
      对方。这是「分发」，AGPL 第 13 条不触发。适用于对方连不进你
      网络的场景。
    """

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("智能体接入（MCP）")
        self.resizable(False, False)
        self.transient(app.root)
        self.configure(bg=Theme.APP_BG)

        strip = tk.Frame(self, bg=Theme.PRIMARY, height=3, bd=0,
                         highlightthickness=0)
        strip.pack(fill="x")
        strip.pack_propagate(False)
        card = ttk.Frame(self, style="Card.TFrame", padding=(20, 16))
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="智能体接入", style="Title.TLabel").pack(anchor="w")
        ttk.Label(card, text="开启后 AI 智能体能调用本工具处理 PDF。"
                             "默认关闭，随时可以关掉。",
                  style="Muted.TLabel", wraplength=460,
                  justify="left").pack(anchor="w", pady=(4, 10))

        self.mode = tk.StringVar(value="off")
        for val, text in (
                ("off", "关闭（默认）"),
                ("local", "仅本机智能体可调用（全量工具，127.0.0.1）"),
                ("remote", "远程智能体接入（开启 HTTP 服务，按需下发代码）"),
                ("export", "手动导出代码 / 分发包（备用）")):
            ttk.Radiobutton(card, text=text, value=val, variable=self.mode,
                            command=self._on_mode).pack(anchor="w", pady=2)

        boxes = ttk.Frame(card, style="Card.TFrame")
        boxes.pack(fill="x", pady=(10, 0))
        self.local_box = ttk.Frame(boxes, style="Card.TFrame")
        self.local_box.grid(row=0, column=0, sticky="ew")
        self.remote_box = ttk.Frame(boxes, style="Card.TFrame")
        self.remote_box.grid(row=1, column=0, sticky="ew")
        self.export_box = ttk.Frame(boxes, style="Card.TFrame")
        self.export_box.grid(row=2, column=0, sticky="ew")

        self.status = ttk.Label(self.local_box, text="未启动",
                                style="Muted.TLabel", wraplength=460,
                                justify="left")
        self.status.pack(anchor="w")
        row = ttk.Frame(self.local_box, style="Card.TFrame")
        row.pack(anchor="w", pady=(8, 0))
        self.toggle_btn = ttk.Button(row, text="启动", style="Accent.TButton",
                                     command=self._toggle)
        self.toggle_btn.pack(side="left")
        ttk.Button(row, text="复制配置",
                   command=self._copy_config).pack(side="left", padx=8)
        self.cfg = tk.Text(self.local_box, height=8, width=64, wrap="word",
                           relief="solid", bd=1)
        self.cfg.pack(fill="x", pady=(8, 0))
        self.confirm_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.local_box,
                        text="高危操作（物理打印）执行前弹窗确认",
                        variable=self.confirm_var).pack(anchor="w",
                                                        pady=(8, 0))

        # ---- 远程智能体接入：只出代码的 HTTP MCP 服务（code_only）----
        ttk.Label(self.remote_box,
                  text="开启一个只提供代码的 HTTP MCP 服务：远程智能体连上来"
                       "自己发现 pdf-code 工具，需要处理 PDF 时按需取得可准确"
                       "执行的代码，保存到它自己的机器上运行——文件全程不出"
                       "它的机器，不需要你手动发代码。本模式不开放任何读写"
                       "你这台电脑文件的工具，因此可以放心监听局域网地址。",
                  style="Muted.TLabel", wraplength=460,
                  justify="left").pack(anchor="w")
        self.rstatus = ttk.Label(self.remote_box, text="未启动",
                                 style="Muted.TLabel", wraplength=460,
                                 justify="left")
        self.rstatus.pack(anchor="w", pady=(6, 0))
        rrow = ttk.Frame(self.remote_box, style="Card.TFrame")
        rrow.pack(anchor="w", pady=(8, 0))
        self.rtoggle_btn = ttk.Button(rrow, text="启动",
                                      style="Accent.TButton",
                                      command=self._toggle)
        self.rtoggle_btn.pack(side="left")
        ttk.Button(rrow, text="复制接入配置",
                   command=self._copy_remote_config).pack(side="left", padx=8)
        self.rcfg = tk.Text(self.remote_box, height=8, width=64, wrap="word",
                            relief="solid", bd=1)
        self.rcfg.pack(fill="x", pady=(8, 0))

        ttk.Label(self.export_box,
                  text="不把软件装包发给对方，而是生成一段自包含的 Python 脚本"
                       "通过 MCP 交给远程智能体：它保存到自己机器上执行，"
                       "不需要 Windows、不用装本程序，文件全程不出它的机器。",
                  style="Muted.TLabel", wraplength=460,
                  justify="left").pack(anchor="w")
        r1 = ttk.Frame(self.export_box, style="Card.TFrame")
        r1.pack(fill="x", pady=(8, 0))
        ttk.Label(r1, text="操作", style="Muted.TLabel").pack(side="left")
        self.gen_cmd = ttk.Combobox(
            r1, width=13, state="readonly",
            values=([n for n, _h, _o, _f in pdf_cli.SPECS] if pdf_cli else []))
        self.gen_cmd.set("compress")
        self.gen_cmd.pack(side="left", padx=4)
        ttk.Label(r1, text="参数", style="Muted.TLabel").pack(side="left")
        self.gen_params = ttk.Entry(r1, width=30)
        self.gen_params.insert(0, '{"input": "a.pdf", "target_mb": 5}')
        self.gen_params.pack(side="left", padx=4)
        ttk.Button(r1, text="生成代码", style="Accent.TButton",
                   command=self._gen_code).pack(side="left", padx=4)
        self.gen_out = tk.Text(self.export_box, height=6, width=64, wrap="word",
                               relief="solid", bd=1)
        self.gen_out.pack(fill="x", pady=(8, 0))
        r2 = ttk.Frame(self.export_box, style="Card.TFrame")
        r2.pack(fill="x", pady=(6, 0))
        ttk.Button(r2, text="保存为 .py",
                   command=self._save_code).pack(side="left")
        ttk.Button(r2, text="导出整套分发包（.zip）",
                   command=self._export).pack(side="left", padx=8)
        self.gen_status = ttk.Label(self.export_box, text="",
                                    style="Muted.TLabel", wraplength=460,
                                    justify="left")
        self.gen_status.pack(anchor="w", pady=(6, 0))
        self._last_code = ""

        ttk.Label(card, text="许可：AGPL-3.0（基于 PyMuPDF）。"
                            "处理产出的 PDF 不受该许可约束。",
                  style="Muted.TLabel", wraplength=460,
                  justify="left").pack(anchor="w", pady=(14, 0))

        self._on_mode()
        self.update_idletasks()
        x = app.root.winfo_rootx() + (app.root.winfo_width()
                                      - self.winfo_width()) // 2
        y = app.root.winfo_rooty() + (app.root.winfo_height()
                                      - self.winfo_height()) // 2
        self.geometry("+%d+%d" % (x, y))

    def _on_mode(self):
        m = self.mode.get()
        self.local_box.grid() if m == "local" else self.local_box.grid_remove()
        self.remote_box.grid() if m == "remote" else self.remote_box.grid_remove()
        self.export_box.grid() if m == "export" else self.export_box.grid_remove()
        # 正在跑的服务档位与所选档位不一致 → 先停掉，避免两套服务并存
        if getattr(self.app, "mcp_kind", "") not in ("", m):
            self._stop()
        if m == "off":
            self._stop()
        self._refresh()

    def _toggle(self):
        if self.app.mcp_srv is not None:
            self._stop()
        else:
            self._start(self.mode.get())

    def _start(self, kind):
        if pdf_mcp is None:
            msg = "MCP 模块不可用：" + _MCP_ERR
            (self.rstatus if kind == "remote" else self.status).config(text=msg)
            return
        try:
            if kind == "remote":
                # 只出代码的服务：绑 0.0.0.0 让局域网/可达网络都能连，
                # code_only 模式下没有文件工具，不需要路径沙箱
                srv, port = pdf_mcp.start_http_background("0.0.0.0", 0,
                                                          code_only=True)
            else:
                srv, port = pdf_mcp.start_http_background("127.0.0.1", 0)
        except Exception as e:
            msg = "启动失败：%s: %s" % (type(e).__name__, e)
            (self.rstatus if kind == "remote" else self.status).config(text=msg)
            return
        self.app.mcp_srv, self.app.mcp_port, self.app.mcp_kind = srv, port, kind
        if kind == "local":
            # 登记高危确认钩子：智能体调 print 时弹框问人，不默默出纸
            # （远程档没有文件工具，不涉及打印，无需钩子）
            pdf_mcp.CONFIRM_HOOK = self._confirm_hook
        self._refresh()

    def _stop(self):
        if self.app.mcp_srv is not None:
            try:
                pdf_mcp.stop_http(self.app.mcp_srv)
            except Exception:
                pass
            self.app.mcp_srv, self.app.mcp_port = None, 0
        if getattr(self.app, "mcp_kind", ""):
            self.app.mcp_kind = ""
        # 摘钩子：服务都停了，别再让本面板拦别的调用方
        if pdf_mcp is not None:
            try:
                hook = pdf_mcp.CONFIRM_HOOK
                if hook is not None and getattr(
                        hook, "__self__", None) is self:
                    pdf_mcp.CONFIRM_HOOK = None
            except Exception:
                pass
        self._refresh()

    def _confirm_hook(self, name, args, reason):
        """pdf_mcp 在 MCP 服务器线程里调它。

        弹模态框必须回 Tk 主线程，所以用 after(0, ...) 派发，
        再用 Event 阻塞服务器线程等用户答复。
        确认闸被用户关掉 → 直接放行；对话框已销毁 / 派发失败 /
        60 秒没答复 → 一律拒绝（安全侧倒）。
        """
        if not self.confirm_var.get():
            return True
        try:
            if not self.winfo_exists():
                return False        # 面板已销毁：没人可问，拒绝
        except tk.TclError:
            return False
        box = {}
        ev = threading.Event()

        def _ask():
            try:
                box["ok"] = HighRiskDialog(self, name, args, reason).result
            except Exception:
                box["ok"] = False
            finally:
                ev.set()

        if threading.current_thread() is threading.main_thread():
            _ask()                      # 本就在主线程，直接问
            return bool(box.get("ok"))
        try:
            self.after(0, _ask)
        except tk.TclError:
            return False
        if not ev.wait(60):
            return False
        return bool(box.get("ok"))

    def _refresh(self):
        import json as _json
        kind = getattr(self.app, "mcp_kind", "")
        running_local = self.app.mcp_srv is not None and kind == "local"
        running_remote = self.app.mcp_srv is not None and kind == "remote"
        if running_local:
            self.status.config(text="已启动 · 仅本机 127.0.0.1:%d，"
                                    "网络上的其他机器连不进来"
                                % self.app.mcp_port)
            self.toggle_btn.config(text="停止")
        else:
            self.status.config(text="未启动")
            self.toggle_btn.config(text="启动")
        if running_remote:
            ip = pdf_mcp.lan_ip() if pdf_mcp else "127.0.0.1"
            self.rstatus.config(
                text="已启动 · 本机 http://127.0.0.1:%d/mcp · "
                     "给远程智能体用 http://%s:%d/mcp"
                     "（同一服务，两个地址任选可达的）"
                     % (self.app.mcp_port, ip, self.app.mcp_port))
            self.rtoggle_btn.config(text="停止")
        else:
            self.rstatus.config(text="未启动")
            self.rtoggle_btn.config(text="启动")
        snippets = (pdf_mcp.client_config_snippets(self.app.mcp_port or 8765)
                    if pdf_mcp else {})
        self.cfg.delete("1.0", "end")
        self.cfg.insert("1.0", _json.dumps(snippets, ensure_ascii=False,
                                           indent=2))
        rcfg = (pdf_mcp.remote_config_snippet(self.app.mcp_port or 8765)
                if pdf_mcp else {})
        self.rcfg.delete("1.0", "end")
        self.rcfg.insert("1.0", _json.dumps(rcfg, ensure_ascii=False,
                                            indent=2))

    def _copy_remote_config(self):
        txt = self.rcfg.get("1.0", "end").strip()
        if not txt:
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.rstatus.config(text="接入配置已复制到剪贴板")

    def _copy_config(self):
        txt = self.cfg.get("1.0", "end").strip()
        if not txt:
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.status.config(text="配置已复制到剪贴板")

    def _export(self):
        from tkinter import filedialog, messagebox
        if pdf_mcp is None:
            messagebox.showerror("无法导出", "MCP 模块不可用：" + _MCP_ERR,
                                 parent=self)
            return
        p = filedialog.asksaveasfilename(
            parent=self, defaultextension=".zip",
            filetypes=[("ZIP 压缩包", "*.zip")],
            initialfile="3QPDF-MCP-分发包.zip")
        if not p:
            return
        try:
            z = self._export_to(p)
        except Exception as e:
            messagebox.showerror("导出失败", "%s: %s" % (type(e).__name__, e),
                                 parent=self)
            return
        ResultDialog(self, "已导出",
                     "分发包已生成。把它发给对方，他解压后按"
                     "「给远程方看这里.txt」运行即可——他的文件不会离开他的机器。",
                     file_path=z, folder=os.path.dirname(z))

    def _export_to(self, path):
        """真正的打包动作。拆出来是为了让回归测试直接调（不必弹对话框）。"""
        return pdf_mcp.export_distribution(path)

    def _gen_code(self):
        import json as _json
        if pdf_mcp is None:
            self.gen_status.config(text="MCP 模块不可用：" + _MCP_ERR)
            return
        cmd = self.gen_cmd.get().strip()
        raw = self.gen_params.get().strip() or "{}"
        try:
            params = _json.loads(raw)
        except Exception as e:
            self.gen_status.config(text="参数不是合法 JSON：%s" % e)
            return
        r = pdf_mcp.gen_pdf_code(cmd, params)
        if not r.get("ok"):
            self.gen_status.config(text=r.get("error", "生成失败"))
            return
        self._last_code = r["code"]
        self.gen_out.delete("1.0", "end")
        self.gen_out.insert("1.0", r["code"])
        self.gen_status.config(
            text="已生成 %d 字节 · sha256 %s… · 依赖：%s"
                 % (r["code_bytes"], r["code_sha256"][:12], r["pip"]))

    def _save_code(self):
        from tkinter import filedialog
        code = getattr(self, "_last_code", "")
        if not code:
            self.gen_status.config(text="先生成代码")
            return
        p = filedialog.asksaveasfilename(
            parent=self, defaultextension=".py",
            filetypes=[("Python 脚本", "*.py")],
            initialfile="pdf_%s.py" % self.gen_cmd.get().replace("-", "_"))
        if not p:
            return
        with open(p, "w", encoding="utf-8") as f:
            f.write(code)
        self.gen_status.config(text="已保存：%s" % p)


# ---------------- 页码校验 mixin（SR-02 红字实时校验） ----------------
class ExprValidator:
    def bind_expr(self, var, get_total, err_label, on_valid=None):
        def check(*_):
            expr = var.get()
            total = get_total()
            if not expr.strip():
                err_label.config(text="")
                if on_valid:
                    on_valid(True)
                return
            if total <= 0:
                err_label.config(text="请先打开文档", foreground="#c0392b")
                if on_valid:
                    on_valid(False)
                return
            try:
                self.parse_expr(expr, total)
                err_label.config(text="", foreground="#c0392b")
                if on_valid:
                    on_valid(True)
            except ValueError as e:
                err_label.config(text=str(e), foreground="#c0392b")
                if on_valid:
                    on_valid(False)
        var.trace_add("write", check)
        return check


# ---------------- 工具对话框基类 ----------------
class ToolDialog(tk.Toplevel, ExprValidator, Runner):
    """作用于当前文档的工具对话框公共基类。"""

    def __init__(self, app, title, run_text):
        super().__init__(app.root)
        self.app = app
        self._run_text = run_text
        self.title(title)
        self.resizable(False, False)
        self.transient(app.root)
        self.configure(bg=Theme.APP_BG)
        strip = tk.Frame(self, bg=Theme.PRIMARY, height=3, bd=0,
                         highlightthickness=0)
        strip.pack(fill="x")
        strip.pack_propagate(False)
        self.status_var = tk.StringVar(value="")
        self.body = ttk.Frame(self, style="Card.TFrame", padding=(18, 14, 18, 12))
        self.body.pack(fill="both", expand=True)
        self.build(self.body)
        bottom = ttk.Frame(self, style="App.TFrame", padding=(18, 10, 18, 16))
        bottom.pack(fill="x")
        self.run_btn = ttk.Button(bottom, text=self._run_text, style="Accent.TButton",
                                  command=self.run)
        self.run_btn.pack(side="left")
        ttk.Button(bottom, text="关闭", command=self.destroy).pack(side="left", padx=8)
        ttk.Label(bottom, textvariable=self.status_var, style="App.TLabel",
                  foreground=Theme.OK).pack(side="left", padx=10)

    def build(self, body):
        raise NotImplementedError

    def run(self):
        raise NotImplementedError

    # ---- 通用小部件 ----
    def doc_head(self, body, row, total_pages):
        ttk.Label(body, text=f"当前文档：{os.path.basename(self.app.src_path)}"
                             f"（共 {total_pages} 页）",
                  style="Muted.TLabel").grid(row=row, column=0, columnspan=4,
                                             sticky="w", pady=(0, 10))

    def out_file_row(self, body, row, var, default_name, label="输出文件",
                     ext=".pdf", filetypes=None):
        ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=var, width=48).grid(
            row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
        ttk.Button(body, text="浏览…",
                   command=lambda: self.pick_save_file(var, default_name, ext,
                                                       filetypes)).grid(
            row=row, column=3, pady=4)

    def out_dir_row(self, body, row, var, label="输出位置"):
        ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=var, width=48).grid(
            row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
        ttk.Button(body, text="浏览…", command=lambda: self.pick_dir(var)).grid(
            row=row, column=3, pady=4)

    def scope_widgets(self, body, row, include_expr=True):
        """「全部页面 / 仅当前页 / 指定页码」选择行。返回 (scope_var, expr_var, err_label)。"""
        scope = tk.StringVar(value="全部页面")
        expr = tk.StringVar()
        ttk.Label(body, text="目标页").grid(row=row, column=0, sticky="nw", pady=4)
        fr = ttk.Frame(body)
        fr.grid(row=row, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(fr, text=f"全部页面（共 {self.app.total} 页）", variable=scope,
                        value="全部页面").pack(anchor="w")
        ttk.Radiobutton(fr, text=f"仅当前页（第 {self.app.cur + 1} 页）", variable=scope,
                        value="当前页").pack(anchor="w")
        err = ttk.Label(fr, text="", foreground="#c0392b")
        if include_expr:
            row2 = ttk.Frame(fr)
            row2.pack(anchor="w")
            ttk.Radiobutton(row2, text="指定页码：", variable=scope,
                            value="指定页码").pack(side="left")
            ttk.Entry(row2, textvariable=expr, width=18).pack(side="left", padx=4)
            err.pack(anchor="w")
        return scope, expr, err

    def resolve_scope(self, scope, expr, err):
        if scope.get() == "全部页面":
            return list(range(self.app.total))
        if scope.get() == "当前页":
            return [self.app.cur]
        return parse_page_expr(expr.get(), self.app.total)

    def pick_dir(self, var):
        d = filedialog.askdirectory(title="选择输出位置", parent=self)
        if d:
            var.set(d)

    def pick_save_file(self, var, default_name, ext=".pdf", filetypes=None):
        path = filedialog.asksaveasfilename(
            title="选择输出文件", defaultextension=ext, parent=self,
            initialfile=default_name,
            filetypes=filetypes or [("PDF 文件", "*.pdf")])
        if path:
            var.set(path)


# ---------------- 拆分 ----------------
class SplitDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.mode = tk.StringVar(value="按页码范围")
        self.ranges_expr = tk.StringVar()
        self.n_per = tk.IntVar(value=10)
        base = os.path.splitext(os.path.basename(app.src_path))[0]
        self.out_dir = tk.StringVar(value=os.path.join(
            os.path.dirname(app.src_path), f"{base}_拆分结果"))
        super().__init__(app, "PDF 拆分", "开始拆分")
        self.parse_expr = parse_ranges_expr
        self.bind_expr(self.ranges_expr, lambda: self.total_pages, self.range_err)
        self.refresh_mode()

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="拆分方式").grid(row=1, column=0, sticky="w", pady=4)
        mode_cb = ttk.Combobox(body, textvariable=self.mode, state="readonly", width=28,
                               values=["按页码范围", "每页一个文件", "按固定页数分段"])
        mode_cb.grid(row=1, column=1, sticky="w", padx=6)
        mode_cb.bind("<<ComboboxSelected>>", lambda *_: self.refresh_mode())
        self.range_row = ttk.Frame(body)
        ttk.Label(self.range_row, text="页码范围").pack(side="left")
        ttk.Entry(self.range_row, textvariable=self.ranges_expr, width=36).pack(side="left", padx=6)
        ttk.Label(self.range_row, text="如 1-10，11-30，31-60", foreground="#999").pack(side="left")
        self.range_row.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        self.range_err = ttk.Label(self.range_row, text="", foreground="#c0392b")
        self.range_err.pack(side="left", padx=6)
        self.n_row = ttk.Frame(body)
        ttk.Label(self.n_row, text="每").pack(side="left")
        ttk.Spinbox(self.n_row, from_=1, to=9999, textvariable=self.n_per,
                    width=6).pack(side="left", padx=4)
        ttk.Label(self.n_row, text="页拆分为一个文件").pack(side="left")
        self.out_dir_row(body, 3, self.out_dir)
        body.columnconfigure(1, weight=1)

    def refresh_mode(self):
        self.range_row.grid_remove()
        self.n_row.grid_remove()
        if self.mode.get() == "按页码范围":
            self.range_row.grid()
        elif self.mode.get() == "按固定页数分段":
            self.n_row.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)

    def run(self):
        base = os.path.splitext(os.path.basename(self.app.src_path))[0]
        try:
            if self.mode.get() == "按页码范围":
                ranges = parse_ranges_expr(self.ranges_expr.get(), self.total_pages)
            elif self.mode.get() == "每页一个文件":
                ranges = [(i, i) for i in range(self.total_pages)]
            else:
                n = max(1, self.n_per.get())
                ranges = [(s, min(s + n, self.total_pages) - 1)
                          for s in range(0, self.total_pages, n)]
        except ValueError as e:
            messagebox.showerror("页码范围有误", str(e), parent=self)
            return
        out_dir = self.out_dir.get().strip()
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: do_split(src, ranges, out_dir, base, pw)
        self.run_in_thread(
            work,
            lambda outs: ("拆分完成", f"已生成 {len(outs)} 个文件。",
                          None, os.path.dirname(outs[0])))


# ---------------- 页面提取 ----------------
class ExtractDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.expr = tk.StringVar(value=str(app.cur + 1))
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_提取页.pdf")
        super().__init__(app, "页面提取", "开始提取")
        self.parse_expr = parse_page_expr
        self.bind_expr(self.expr, lambda: self.total_pages, self.expr_err)

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="提取页码").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.expr, width=36).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="如 3，5，12-18，40（按书写顺序合成新文件，已预填当前页）",
                  foreground="#999").grid(row=1, column=2, sticky="w")
        self.expr_err = ttk.Label(body, text="", foreground="#c0392b")
        self.expr_err.grid(row=2, column=1, sticky="w", padx=6)
        self.out_file_row(body, 3, self.out_file, "提取页.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        try:
            pages = parse_page_expr(self.expr.get(), self.total_pages)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_extract(src, pages, out_path, pw), len(pages))
        self.run_in_thread(
            work,
            lambda res: ("提取完成", f"已提取 {res[2]} 页，生成 {os.path.basename(res[0])}",
                         res[0], os.path.dirname(res[0])))


# ---------------- 合并 ----------------
class MergeDialog(ToolDialog):
    def __init__(self, app, files=None):
        self.files = []
        self._initial_files = list(files) if files else None
        self.out_dir = tk.StringVar()
        self.out_name = tk.StringVar(value="合并结果.pdf")
        super().__init__(app, "PDF 合并", "开始合并")
        if self._initial_files:
            self.add_paths(self._initial_files)
        elif app.doc:
            self._add_one(app.src_path)

    def build(self, body):
        ttk.Label(body, text="待合并文件（按列表顺序拼接）").grid(
            row=0, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=1, column=0, columnspan=3, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=9, width=58)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        btns = ttk.Frame(body)
        btns.grid(row=1, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="添加文件…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="移除选中", command=self.remove_sel).pack(fill="x", pady=2)
        ttk.Button(btns, text="上移", command=lambda: self.move_sel(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="下移", command=lambda: self.move_sel(1)).pack(fill="x", pady=2)
        self.out_dir_row(body, 2, self.out_dir)
        ttk.Label(body, text="输出文件名").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.out_name, width=30).grid(
            row=3, column=1, sticky="w", padx=6)
        body.columnconfigure(1, weight=1)

    def _add_one(self, path):
        ok, msg, _ = check_pdf_readable(path)
        if ok:
            self.files.append(path)
            self.listbox.insert("end", os.path.basename(path))
            if not self.out_dir.get():
                self.out_dir.set(os.path.dirname(path))
        else:
            self.listbox.insert("end", f"[无法读取] {os.path.basename(path)}（{msg}）")

    def add_paths(self, paths):
        """供拖放使用：一次加入多个文件，输出目录跟随第一个。"""
        for p in paths:
            self._add_one(p)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件（可多选）", parent=self,
            filetypes=[("PDF 文件", "*.pdf")])
        for p in paths:
            self._add_one(p)

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            if i < len(self.files):
                self.files.pop(i)
            self.listbox.delete(i)

    def move_sel(self, delta):
        sel = self.listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= self.listbox.size() or i >= len(self.files) or j >= len(self.files):
            return
        self.files[i], self.files[j] = self.files[j], self.files[i]
        t = self.listbox.get(i)
        self.listbox.delete(i)
        self.listbox.insert(j, t)
        self.listbox.selection_set(j)

    def run(self):
        if len(self.files) < 2:
            messagebox.showwarning("无法合并", "请至少添加 2 个可读取的 PDF 文件。", parent=self)
            return
        out_dir = self.out_dir.get().strip() or os.path.dirname(self.files[0])
        name = self.out_name.get().strip() or "合并结果.pdf"
        if not name.lower().endswith(".pdf"):
            name += ".pdf"
        out_path = unique_out_path(os.path.join(out_dir, name))
        files = list(self.files)
        work = lambda: (out_path, do_merge(files, out_path))
        self.run_in_thread(
            work,
            lambda res: ("合并完成", f"已生成 {os.path.basename(res[0])}（共 {res[1]} 页）",
                         res[0], os.path.dirname(res[0])))


# ---------------- 图片转 PDF ----------------
class ImagesToPdfDialog(ToolDialog):
    """图片转 PDF：缩略图预览 + 单张旋转 + 拖动排序 + 手机照片方向自动纠正。"""

    def __init__(self, app):
        self.items = []          # [[路径, 顺时针角度], ...]
        self._thumb = None       # 保留引用，防止 PhotoImage 被 GC
        self._drag_from = None
        self.page_mode = tk.StringVar(value="跟随图片尺寸")
        self.auto_exif = tk.BooleanVar(value=True)
        base_dir = os.path.dirname(app.src_path) if app.doc else ""
        self.out_file = tk.StringVar(value=os.path.join(base_dir, "图片转PDF.pdf")
                                     if base_dir else "图片转PDF.pdf")
        super().__init__(app, "图片转 PDF", "开始转换")

    @property
    def files(self):
        return [p for p, _ in self.items]

    # ---- 列表渲染 ----
    @staticmethod
    def _row_text(p, ang):
        return os.path.basename(p) + ("  ↻%d°" % ang if ang else "")

    def _resync(self, keep_sel=None):
        self.listbox.delete(0, "end")
        for p, a in self.items:
            self.listbox.insert("end", self._row_text(p, a))
        if keep_sel is not None and 0 <= keep_sel < len(self.items):
            self.listbox.selection_set(keep_sel)
        self.show_preview()

    def build(self, body):
        ttk.Label(body, text="图片列表（每张一页，按住行可直接拖动改顺序）").grid(
            row=0, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=1, column=0, columnspan=2, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=9, width=40)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for ev, fn in (("<<ListboxSelect>>", lambda *_: self.show_preview()),
                       ("<ButtonPress-1>", self._drag_start),
                       ("<B1-Motion>", self._drag_motion),
                       ("<ButtonRelease-1>", self._drag_end)):
            self.listbox.bind(ev, fn)

        pv = ttk.Frame(body)
        pv.grid(row=1, column=2, sticky="nswe", padx=(8, 4))
        self.preview = tk.Canvas(pv, width=170, height=170, bg="#EEF2F8",
                                 highlightthickness=1, highlightbackground="#DCE5F2")
        self.preview.pack()
        self.pv_label = ttk.Label(pv, text="未选择", style="Muted.TLabel",
                                  wraplength=170, justify="center")
        self.pv_label.pack(pady=(4, 0))

        btns = ttk.Frame(body)
        btns.grid(row=1, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="添加图片…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="移除选中", command=self.remove_sel).pack(fill="x", pady=2)
        ttk.Button(btns, text="上移", command=lambda: self.move_sel(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="下移", command=lambda: self.move_sel(1)).pack(fill="x", pady=2)
        ttk.Separator(btns, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(btns, text="左转 90°", command=lambda: self.rotate_sel(-90)).pack(fill="x", pady=2)
        ttk.Button(btns, text="右转 90°", command=lambda: self.rotate_sel(90)).pack(fill="x", pady=2)
        ttk.Button(btns, text="归正方向", command=lambda: self.rotate_sel(0)).pack(fill="x", pady=2)

        ttk.Checkbutton(body, text="添加时自动纠正手机照片方向（EXIF）",
                        variable=self.auto_exif).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(body, text="页面方式").grid(row=3, column=0, sticky="w", pady=4)
        fr = ttk.Frame(body)
        fr.grid(row=3, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(fr, text="跟随图片尺寸（适合屏幕查看）", variable=self.page_mode,
                        value="跟随图片尺寸").pack(anchor="w")
        ttk.Radiobutton(fr, text="A4 纸适配居中（适合打印）", variable=self.page_mode,
                        value="A4").pack(anchor="w")
        self.out_file_row(body, 4, self.out_file, "图片转PDF.pdf")
        body.columnconfigure(1, weight=1)

    # ---- 增删改序 ----
    def add_paths(self, paths):
        for p in paths:
            ang = image_exif_rotation(p) if self.auto_exif.get() else 0
            self.items.append([p, ang])
            self.listbox.insert("end", self._row_text(p, ang))
            if len(self.items) == 1 and not os.path.dirname(self.out_file.get()):
                self.out_file.set(os.path.join(
                    os.path.dirname(p), "图片转PDF.pdf"))
        if self.items:
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(len(self.items) - 1)
        self.show_preview()

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择图片（可多选）", parent=self,
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if paths:
            self.add_paths(list(paths))

    def remove_sel(self):
        sel = list(self.listbox.curselection())
        if not sel:
            return
        for i in reversed(sel):
            if i < len(self.items):
                self.items.pop(i)
        self._resync(keep_sel=min(sel[0], len(self.items) - 1))

    def move_sel(self, delta):
        sel = self.listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if j < 0 or j >= len(self.items):
            return
        self.items[i], self.items[j] = self.items[j], self.items[i]
        self._resync(keep_sel=j)

    def rotate_sel(self, delta):
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.items):
            messagebox.showinfo("提示", "请先在左侧选中一张图片。", parent=self)
            return
        i = sel[0]
        if delta == 0:
            self.items[i][1] = 0
        else:
            self.items[i][1] = (self.items[i][1] + delta) % 360
        self._resync(keep_sel=i)

    # ---- 拖动排序 ----
    def _drag_start(self, event):
        i = self.listbox.nearest(event.y)
        if 0 <= i < len(self.items):
            self._drag_from = i
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(i)
            self.listbox.activate(i)
            self.show_preview()

    def _drag_motion(self, event):
        if self._drag_from is None:
            return
        to = self.listbox.nearest(event.y)
        if to == self._drag_from or not (0 <= to < len(self.items)):
            return
        item = self.items.pop(self._drag_from)
        self.items.insert(to, item)
        self._drag_from = to
        self._resync(keep_sel=to)

    def _drag_end(self, event):
        self._drag_from = None

    # ---- 预览 ----
    def show_preview(self):
        self.preview.delete("all")
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.items):
            self.pv_label.config(text="未选择")
            return
        p, ang = self.items[sel[0]]
        try:
            from PIL import Image as PILImage, ImageTk
            with PILImage.open(p) as src:
                im = src.copy()
                w, h = im.size
            if ang == 90:
                im = im.transpose(PILImage.ROTATE_270)
            elif ang == 180:
                im = im.transpose(PILImage.ROTATE_180)
            elif ang == 270:
                im = im.transpose(PILImage.ROTATE_90)
            im = im.convert("RGB")
            im.thumbnail((158, 158))
            photo = ImageTk.PhotoImage(im)
            self._thumb = photo
            self.preview.create_image(85, 85, image=photo, anchor="center")
            self.pv_label.config(text="%d×%d px" % (w, h)
                                 + (" · 顺时针 %d°" % ang if ang else ""))
        except Exception as e:
            self.pv_label.config(text="预览失败：" + str(e)[:60])

    def run(self):
        if not self.items:
            messagebox.showwarning("缺少图片", "请先添加至少 1 张图片。", parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path if out_path.lower().endswith(".pdf")
                                   else out_path + ".pdf")
        items = list(self.items)
        a4 = self.page_mode.get() == "A4"

        def work():
            files, temps = [], []
            for p, ang in items:
                if ang:
                    t = rotate_image_file(p, ang)
                    temps.append(t)
                    files.append(t)
                else:
                    files.append(p)
            try:
                do_images_to_pdf(files, out_path, a4)
            finally:
                for t in temps:
                    try:
                        os.remove(t)
                    except OSError:
                        pass
            return out_path, len(files)

        self.run_in_thread(
            work,
            lambda res: ("转换完成", f"已生成 {os.path.basename(res[0])}"
                                    f"（共 {res[1]} 页）",
                         res[0], os.path.dirname(res[0])))


# ---------------- PDF 转图片 ----------------
class PdfToImagesDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.fmt = tk.StringVar(value="png")
        self.dpi = tk.IntVar(value=150)
        base = os.path.splitext(app.src_path)[0]
        self.out_dir = tk.StringVar(value=f"{base}_图片")
        super().__init__(app, "PDF 转图片", "开始导出")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="图片格式").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.fmt, state="readonly", width=10,
                     values=["png", "jpg"]).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="分辨率").grid(row=2, column=0, sticky="w", pady=4)
        row = ttk.Frame(body)
        row.grid(row=2, column=1, sticky="w", padx=6)
        ttk.Spinbox(row, from_=72, to=300, textvariable=self.dpi, width=6).pack(side="left")
        ttk.Label(row, text="DPI（150 清晰够用，300 打印级）",
                  foreground="#999").pack(side="left", padx=6)
        self.out_dir_row(body, 3, self.out_dir, label="输出文件夹")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_dir = self.out_dir.get().strip()
        src, pw = self.app.src_path, self.app.open_pw
        fmt, dpi = self.fmt.get(), max(72, min(300, self.dpi.get()))
        work = lambda: do_pdf_to_images(src, out_dir, fmt, dpi, pw)
        self.run_in_thread(
            work,
            lambda outs: ("导出完成", f"已导出 {len(outs)} 张图片。",
                          None, os.path.dirname(outs[0])))


# ---------------- 删除页面 ----------------
class DeletePagesDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.expr = tk.StringVar(value=str(app.cur + 1))
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_删页后.pdf")
        super().__init__(app, "删除页面", "开始删除")
        self.parse_expr = parse_page_expr
        self.bind_expr(self.expr, lambda: self.total_pages, self.expr_err)

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="删除页码").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.expr, width=36).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="如 2，5，10-12（已预填当前页）",
                  foreground="#999").grid(row=1, column=2, sticky="w")
        self.expr_err = ttk.Label(body, text="", foreground="#c0392b")
        self.expr_err.grid(row=2, column=1, sticky="w", padx=6)
        self.out_file_row(body, 3, self.out_file, "删页后.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        try:
            pages = parse_page_expr(self.expr.get(), self.total_pages)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        if len(set(pages)) >= self.total_pages:
            messagebox.showerror("无法删除", "不能删除全部页面，至少保留 1 页。", parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_delete_pages(src, pages, out_path, pw))
        self.run_in_thread(
            work,
            lambda res: ("删除完成", f"已生成 {os.path.basename(res[0])}（剩 {self.total_pages - len(set(pages))} 页）",
                         res[0], os.path.dirname(res[0])))


# ---------------- 旋转页面 ----------------
class RotatePagesDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.angle = tk.StringVar(value="90")
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_已旋转.pdf")
        super().__init__(app, "旋转页面（保存到新文件）", "开始旋转")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        self.scope, self.expr, self.err = self.scope_widgets(body, 1)
        ttk.Label(body, text="旋转角度").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.angle, state="readonly", width=10,
                     values=["90", "180", "270"]).grid(row=2, column=1, sticky="w", padx=6)
        ttk.Label(body, text="度（顺时针，永久保存）", foreground="#999").grid(
            row=2, column=2, sticky="w")
        self.out_file_row(body, 3, self.out_file, "已旋转.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        ang = int(self.angle.get())  # 主线程先取值，子线程只用普通 Python 值
        work = lambda: (out_path, do_rotate_pages(src, pages, ang, out_path, pw))
        self.run_in_thread(
            work,
            lambda res: ("旋转完成", f"已生成 {os.path.basename(res[0])}",
                         res[0], os.path.dirname(res[0])))


# ---------------- 页面管理（重排 / 复制 / 删除 / 插入空白页） ----------------
class ReorderDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        r0 = app.doc[app.cur].rect
        self.blank_w, self.blank_h = r0.width, r0.height  # 新空白页沿用当前页尺寸
        self.plan = [("page", i) for i in range(app.total)]
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_重排.pdf")
        super().__init__(app, "页面管理（重排）", "生成新文件")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="页面顺序（可上移/下移/复制/删除/插入空白页）").grid(
            row=1, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=2, column=0, columnspan=3, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=10, width=44)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        btns = ttk.Frame(body)
        btns.grid(row=2, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="上移", command=lambda: self._move(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="下移", command=lambda: self._move(1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="复制选中", command=self._dup).pack(fill="x", pady=2)
        ttk.Button(btns, text="插入空白页", command=self._blank).pack(fill="x", pady=2)
        ttk.Button(btns, text="删除选中", command=self._del).pack(fill="x", pady=2)
        ttk.Button(btns, text="恢复原始顺序", command=self._reset).pack(fill="x", pady=(12, 2))
        self.out_file_row(body, 3, self.out_file, "重排.pdf")
        body.columnconfigure(1, weight=1)
        self._reload()

    def _reload(self):
        self.listbox.delete(0, "end")
        for item in self.plan:
            self.listbox.insert(
                "end", f"第 {item[1] + 1} 页" if item[0] == "page" else "（空白页）")

    def _sel(self):
        s = self.listbox.curselection()
        return s[0] if s else -1

    def _move(self, delta):
        i = self._sel()
        if i < 0:
            return
        j = i + delta
        if j < 0 or j >= len(self.plan):
            return
        self.plan[i], self.plan[j] = self.plan[j], self.plan[i]
        self._reload()
        self.listbox.selection_set(j)

    def _dup(self):
        i = self._sel()
        if i < 0:
            return
        self.plan.insert(i + 1, self.plan[i])
        self._reload()
        self.listbox.selection_set(i + 1)

    def _blank(self):
        i = self._sel()
        at = i + 1 if i >= 0 else len(self.plan)
        self.plan.insert(at, ("blank", self.blank_w, self.blank_h))
        self._reload()
        self.listbox.selection_set(at)

    def _del(self):
        i = self._sel()
        if i < 0:
            return
        if len(self.plan) <= 1:
            messagebox.showwarning("无法删除", "至少保留 1 页。", parent=self)
            return
        self.plan.pop(i)
        self._reload()

    def _reset(self):
        self.plan = [("page", i) for i in range(self.total_pages)]
        self._reload()

    def run(self):
        if not self.plan:
            messagebox.showwarning("没有页面", "页面列表为空。", parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        plan = list(self.plan)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_reorder(src, out_path, plan, pw))
        self.run_in_thread(
            work,
            lambda res: ("重排完成",
                         f"已生成 {os.path.basename(res[0])}（共 {res[1]} 页）",
                         res[0], os.path.dirname(res[0])))


# ---------------- 添加页码 ----------------
class PageNumbersDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.position = tk.StringVar(value="底部居中")
        self.start = tk.IntVar(value=1)
        self.fmt = tk.StringVar(value="第 X 页 / 共 N 页")
        self.fontsize = tk.IntVar(value=10)
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_加页码.pdf")
        super().__init__(app, "添加页码", "开始添加")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="位置").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.position, state="readonly", width=12,
                     values=["底部居中", "底部左侧", "底部右侧", "顶部左侧", "顶部右侧"]).grid(
            row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="起始页号").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Spinbox(body, from_=1, to=9999, textvariable=self.start, width=6).grid(
            row=2, column=1, sticky="w", padx=6)
        ttk.Label(body, text="格式").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.fmt, state="readonly", width=20,
                     values=["第 X 页 / 共 N 页", "第 X 页", "X / N", "X"]).grid(
            row=3, column=1, sticky="w", padx=6)
        ttk.Label(body, text="字号").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Spinbox(body, from_=6, to=36, textvariable=self.fontsize, width=6).grid(
            row=4, column=1, sticky="w", padx=6)
        self.out_file_row(body, 5, self.out_file, "加页码.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        # 主线程先取值，子线程只用普通 Python 值（tk 变量不能跨线程读取）
        pos = self.position.get()
        start = max(1, self.start.get())
        fmt = self.fmt.get()
        fs = self.fontsize.get()
        work = lambda: (out_path, do_add_page_numbers(src, out_path, pos, start, fmt, fs, pw))
        self.run_in_thread(
            work,
            lambda res: ("页码添加完成", f"已生成 {os.path.basename(res[0])}",
                         res[0], os.path.dirname(res[0])))


# ---------------- 压缩 ----------------
class CompressDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.quality = tk.IntVar(value=50)
        self.max_dpi = tk.IntVar(value=0)
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_压缩.pdf")
        super().__init__(app, "压缩体积", "开始压缩")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="压缩质量").grid(row=1, column=0, sticky="w", pady=4)
        row = ttk.Frame(body)
        row.grid(row=1, column=1, sticky="w", padx=6)
        ttk.Spinbox(row, from_=10, to=90, textvariable=self.quality, width=6).pack(side="left")
        ttk.Label(row, text="数字越小体积越小、越模糊（50 兼顾清晰度）",
                  foreground="#999").pack(side="left", padx=6)
        ttk.Label(body, text="分辨率上限").grid(row=2, column=0, sticky="w", pady=4)
        row2 = ttk.Frame(body)
        row2.grid(row=2, column=1, sticky="w", padx=6)
        ttk.Spinbox(row2, from_=0, to=600, increment=24, textvariable=self.max_dpi,
                    width=6).pack(side="left")
        ttk.Label(row2, text="dpi（0 = 保持原分辨率，只重编码；150 适合屏幕、96 适合邮件）",
                  foreground="#999").pack(side="left", padx=6)
        ttk.Label(body, text="说明").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Label(body, text="主要压缩文档内图片；文字部分不受影响。原文件保持不变。",
                  foreground="#999").grid(row=3, column=1, columnspan=2, sticky="w", padx=6)
        self.out_file_row(body, 4, self.out_file, "压缩.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        q = max(10, min(90, self.quality.get()))
        dpi = max(0, int(self.max_dpi.get()))
        work = lambda: (out_path, do_compress(src, out_path, q, dpi, pw))
        self.run_in_thread(
            work,
            lambda res: self._msg(res))

    def _msg(self, res):
        old, new = res[1], res[2]
        pct = (1 - new / old) * 100 if old else 0
        return ("压缩完成",
                f"{old / 1048576:.2f} MB → {new / 1048576:.2f} MB"
                f"（缩小 {pct:.0f}%）" if pct > 0.5 else
                f"已完成优化，{old / 1048576:.2f} MB → {new / 1048576:.2f} MB"
                f"（该文档已较紧凑）",
                res[0], os.path.dirname(res[0]))


class CompressToTargetDialog(ToolDialog):
    """压缩到指定大小：在「画质 → 分辨率」两档参数上二分试压。"""

    def __init__(self, app):
        self.total_pages = app.total
        self.target_mb = tk.DoubleVar(value=5.0)
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=base + "_压缩.pdf")
        super().__init__(app, "压缩到指定大小", "开始压缩")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        try:
            cur = os.path.getsize(self.app.src_path) / 1048576.0
        except OSError:
            cur = 0.0
        ttk.Label(body, text="当前大小").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Label(body, text="%.2f MB" % cur, style="Muted.TLabel").grid(
            row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="目标大小").grid(row=2, column=0, sticky="w", pady=4)
        row = ttk.Frame(body)
        row.grid(row=2, column=1, sticky="w", padx=6)
        ttk.Spinbox(row, from_=0.1, to=500, increment=0.5,
                    textvariable=self.target_mb, width=7).pack(side="left")
        ttk.Label(row, text="MB（以内）", foreground="#999").pack(side="left", padx=4)
        ttk.Label(body, text="说明").grid(row=3, column=0, sticky="nw", pady=4)
        ttk.Label(body, text="会依次在「画质」「分辨率」两档参数上二分试压，取仍能装进目标的最高画质。\n"
                            "文字型 PDF 没有图片可压，达不到目标时会保留最小的一份并如实说明。",
                  foreground="#999", justify="left").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=6)
        self.out_file_row(body, 4, self.out_file, "压缩.pdf")
        body.columnconfigure(1, weight=1)

    def on_progress(self, done, total):
        self.set_status("正在试压…（第 %d 次）" % done)

    def run(self):
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        target = int(max(0.05, self.target_mb.get()) * 1048576)

        def work():
            old, new, ok = do_compress_to_target(src, out_path, target, pw,
                                                 self.report_progress)
            return out_path, old, new, ok

        self.run_in_thread(work, self._msg)

    def _msg(self, res):
        out, old, new, ok = res
        pct = (1 - new / old) * 100 if old else 0
        if ok:
            msg = ("%.2f MB → %.2f MB（缩小 %.0f%%，已达成目标）"
                   % (old / 1048576, new / 1048576, pct))
        else:
            msg = ("%.2f MB → %.2f MB（缩小 %.0f%%）\n仍大于目标：该文档可压的部分已压到最小，"
                   "可调大目标或先拆分文档。" % (old / 1048576, new / 1048576, pct))
        return ("压缩完成", msg, out, os.path.dirname(out))


class OfficeToPdfDialog(ToolDialog):
    """Word / Excel / PowerPoint → PDF（Windows + 本机 Office，批量）。"""

    def __init__(self, app):
        self.items = []
        base_dir = (os.path.dirname(app.src_path) if app.doc
                    else os.path.expanduser("~"))
        self.out_dir = tk.StringVar(value=os.path.join(base_dir, "Office转PDF结果"))
        super().__init__(app, "Office 转 PDF", "开始转换")

    @property
    def files(self):
        return list(self.items)

    def build(self, body):
        ttk.Label(body, text="待转换文件（doc/docx/xls/xlsx/ppt/pptx/csv）").grid(
            row=0, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=1, column=0, columnspan=3, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=8, width=54)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        btns = ttk.Frame(body)
        btns.grid(row=1, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="添加文件…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="移除选中", command=self.remove_sel).pack(fill="x", pady=2)
        ttk.Button(btns, text="清空", command=self.clear_all).pack(fill="x", pady=2)
        ttk.Label(body, text="说明").grid(row=2, column=0, sticky="nw", pady=4)
        ttk.Label(body, text="走本机 Office（COM）转换，排版与 Office「另存为 PDF」一致；\n"
                            "需要 Windows 且已装 Word / Excel / PowerPoint（WPS 需开启 COM 兼容）。\n"
                            "每个文件单独转换，某个失败不影响其余。",
                  foreground="#999", justify="left").grid(
            row=2, column=1, columnspan=2, sticky="w", padx=6)
        self.out_dir_row(body, 3, self.out_dir, "输出文件夹")
        body.columnconfigure(1, weight=1)

    def add_paths(self, paths):
        for p in paths:
            if office_kind(p) is None:
                continue
            if p not in self.items:
                self.items.append(p)
                self.listbox.insert("end", os.path.basename(p))

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Office 文件（可多选）", parent=self,
            filetypes=[("Office 文件",
                        "*.doc *.docx *.xls *.xlsx *.ppt *.pptx *.csv")])
        if paths:
            self.add_paths(list(paths))

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            if i < len(self.items):
                self.items.pop(i)
            self.listbox.delete(i)

    def clear_all(self):
        self.items = []
        self.listbox.delete(0, "end")

    def on_progress(self, done, total):
        self.set_status("正在转换…（%d/%d）" % (done, total))

    def run(self):
        if not self.items:
            messagebox.showwarning("缺少文件", "请先添加至少一个 Office 文件。", parent=self)
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("缺少输出", "请选择输出文件夹。", parent=self)
            return
        files = list(self.items)

        def work():
            outs, errs = office_batch_to_pdf(files, out_dir, self.report_progress)
            return outs, errs, out_dir

        self.run_in_thread(work, self._msg)

    def _msg(self, res):
        outs, errs, folder = res
        if not outs:
            detail = "\n".join("%s：%s" % (n, e) for n, e in errs[:6])
            return ("转换失败", detail or "没有文件被转换。", None, folder)
        msg = "成功 %d 个" % len(outs)
        if errs:
            msg += "，失败 %d 个：\n" % len(errs)
            msg += "\n".join("%s：%s" % (n, e) for n, e in errs[:6])
        return ("转换完成", msg, outs[0] if len(outs) == 1 else None, folder)


class CropPagesDialog(ToolDialog):
    """裁剪页面：自动去扫描黑边 / 四边留白 / 恢复原始尺寸。"""

    def __init__(self, app):
        self.total_pages = app.total
        self.mode = tk.StringVar(value="自动去黑边")
        self.keep_mm = tk.DoubleVar(value=3.0)
        self.m_l = tk.DoubleVar(value=10.0)
        self.m_t = tk.DoubleVar(value=10.0)
        self.m_r = tk.DoubleVar(value=10.0)
        self.m_b = tk.DoubleVar(value=10.0)
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=base + "_裁剪.pdf")
        super().__init__(app, "裁剪页面 / 去黑边", "开始裁剪")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        self.scope, self.expr, self.err = self.scope_widgets(body, 1)
        ttk.Label(body, text="裁剪方式").grid(row=2, column=0, sticky="nw", pady=4)
        fr = ttk.Frame(body)
        fr.grid(row=2, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Radiobutton(fr, text="自动去黑边（按内容边界裁，扫描件首选）",
                        variable=self.mode, value="自动去黑边").pack(anchor="w")
        rr = ttk.Frame(fr)
        rr.pack(anchor="w")
        ttk.Radiobutton(rr, text="自动裁切后保留边距", variable=self.mode,
                        value="自动去黑边").pack(side="left")
        ttk.Spinbox(rr, from_=0, to=30, increment=1, textvariable=self.keep_mm,
                    width=6).pack(side="left", padx=4)
        ttk.Label(rr, text="mm", foreground="#999").pack(side="left")
        ttk.Radiobutton(fr, text="四边留白（按给定边距内缩）", variable=self.mode,
                        value="四边留白").pack(anchor="w")
        mr = ttk.Frame(fr)
        mr.pack(anchor="w", pady=(2, 0))
        for lab, var in (("左", self.m_l), ("上", self.m_t),
                         ("右", self.m_r), ("下", self.m_b)):
            ttk.Label(mr, text=lab).pack(side="left")
            ttk.Spinbox(mr, from_=0, to=200, increment=1, textvariable=var,
                        width=5).pack(side="left", padx=(1, 6))
        ttk.Label(mr, text="mm", foreground="#999").pack(side="left")
        ttk.Radiobutton(fr, text="恢复原始尺寸（撤销以前的裁剪）", variable=self.mode,
                        value="恢复原始尺寸").pack(anchor="w", pady=(4, 0))
        ttk.Label(body, text="说明", foreground="#999").grid(
            row=3, column=0, sticky="nw", pady=4)
        ttk.Label(body, text="裁剪只改显示/打印的可视范围，页面内容仍在（可再「恢复原始尺寸」）。\n"
                            "内容已铺满整页时不会裁，避免误裁。",
                  foreground="#999", justify="left").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=6)
        self.out_file_row(body, 4, self.out_file, "裁剪.pdf")
        body.columnconfigure(1, weight=1)

    def on_progress(self, done, total):
        self.set_status("正在分析/裁剪…（%d/%d）" % (done, total))

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except Exception as e:
            messagebox.showwarning("页码有误", str(e), parent=self)
            return
        if not pages:
            messagebox.showwarning("页码有误", "没有选中任何页面。", parent=self)
            return
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        mode = self.mode.get()
        keep = self.keep_mm.get()
        margins = (self.m_l.get(), self.m_t.get(), self.m_r.get(), self.m_b.get())

        def work():
            n, changed = do_crop_pages(src, out_path, pages, mode, keep, margins,
                                       pw, self.report_progress)
            return out_path, n, changed

        self.run_in_thread(work, self._msg)

    def _msg(self, res):
        out, n, changed = res
        return ("裁剪完成",
                "已处理 %d 页，其中 %d 页被裁剪。" % (n, changed)
                + ("" if changed else "\n（内容已铺满整页，无需裁剪）"),
                out, os.path.dirname(out))


# ---------------- 加密 ----------------
class EncryptDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.pw1 = tk.StringVar()
        self.pw2 = tk.StringVar()
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_加密.pdf")
        super().__init__(app, "加密保护", "开始加密")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="打开密码").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.pw1, show="*", width=24).grid(
            row=1, column=1, sticky="w", padx=6)
        ttk.Label(body, text="确认密码").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.pw2, show="*", width=24).grid(
            row=2, column=1, sticky="w", padx=6)
        ttk.Label(body, text="下次打开该副本需输入密码；保留打印与复制权限。",
                  foreground="#999").grid(row=3, column=0, columnspan=3, sticky="w", pady=4)
        self.out_file_row(body, 4, self.out_file, "加密.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        p1, p2 = self.pw1.get(), self.pw2.get()
        if not p1:
            messagebox.showwarning("缺少密码", "请输入打开密码。", parent=self)
            return
        if p1 != p2:
            messagebox.showwarning("密码不一致", "两次输入的密码不一致。", parent=self)
            return
        out_path = unique_out_path(self.out_file.get().strip())
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_encrypt(src, out_path, p1, pw))
        self.run_in_thread(
            work,
            lambda res: ("加密完成", f"已生成 {os.path.basename(res[0])}，打开需密码。",
                         res[0], os.path.dirname(res[0])))


# ---------------- 解密 / 去除限制 ----------------
class DecryptDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_无密码.pdf")
        super().__init__(app, "解密 / 去除限制", "开始解密")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="说明").grid(row=1, column=0, sticky="nw", pady=4)
        ttk.Label(body, text="去除当前文档的打开密码与权限限制（如禁止复制/打印），\n"
                             "生成无需密码即可打开的副本。原文件保持不变。",
                  justify="left", foreground="#999").grid(
            row=1, column=1, columnspan=2, sticky="w", padx=6, pady=4)
        self.out_file_row(body, 2, self.out_file, "无密码.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_decrypt(src, out_path, pw))
        self.run_in_thread(
            work,
            lambda res: ("解密完成",
                         f"已生成 {os.path.basename(res[0])}，打开无需密码。",
                         res[0], os.path.dirname(res[0])))


# ---------------- 水印 ----------------
class WatermarkDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.kind = tk.StringVar(value="文字水印")
        self.text = tk.StringVar(value="机密")
        self.fontsize = tk.IntVar(value=48)
        self.color = tk.StringVar(value="浅灰")
        self.angle = tk.IntVar(value=45)
        self.img_path = ""
        self.scale = tk.IntVar(value=50)
        self.opacity = tk.IntVar(value=20)
        self.layout = tk.StringVar(value="平铺整页")
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_水印.pdf")
        super().__init__(app, "添加水印", "开始添加")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="水印类型").grid(row=1, column=0, sticky="w", pady=4)
        cb = ttk.Combobox(body, textvariable=self.kind, state="readonly", width=12,
                          values=["文字水印", "图片水印"])
        cb.grid(row=1, column=1, sticky="w", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda *_: self._refresh())

        self.text_frame = ttk.Frame(body)
        self.text_frame.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(self.text_frame, text="水印文字").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.text_frame, textvariable=self.text, width=16).grid(
            row=0, column=1, padx=6)
        ttk.Label(self.text_frame, text="字号").grid(row=0, column=2, padx=(10, 0))
        ttk.Spinbox(self.text_frame, from_=8, to=200, textvariable=self.fontsize,
                    width=5).grid(row=0, column=3, padx=4)
        ttk.Label(self.text_frame, text="颜色").grid(row=0, column=4, padx=(10, 0))
        ttk.Combobox(self.text_frame, textvariable=self.color, state="readonly", width=5,
                     values=list(COLOR_PRESETS.keys())).grid(row=0, column=5, padx=4)
        ttk.Label(self.text_frame, text="角度°").grid(row=0, column=6, padx=(10, 0))
        ttk.Spinbox(self.text_frame, from_=0, to=359, textvariable=self.angle,
                    width=5).grid(row=0, column=7, padx=4)

        self.img_frame = ttk.Frame(body)
        ttk.Button(self.img_frame, text="选择水印图片…", command=self._pick_img).grid(
            row=0, column=0)
        self.img_lbl = ttk.Label(self.img_frame, text="未选择", foreground="#c0392b",
                                 width=22)
        self.img_lbl.grid(row=0, column=1, padx=6, sticky="w")
        ttk.Label(self.img_frame, text="缩放%").grid(row=0, column=2, padx=(10, 0))
        ttk.Spinbox(self.img_frame, from_=5, to=200, textvariable=self.scale,
                    width=5).grid(row=0, column=3, padx=4)

        ttk.Label(body, text="透明度%").grid(row=4, column=0, sticky="w", pady=4)
        row = ttk.Frame(body)
        row.grid(row=4, column=1, columnspan=2, sticky="w", padx=6)
        ttk.Spinbox(row, from_=5, to=100, textvariable=self.opacity, width=5).pack(side="left")
        ttk.Label(row, text="数字越小越淡（文字水印建议 10~30）",
                  foreground="#999").pack(side="left", padx=6)

        ttk.Label(body, text="排版").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.layout, state="readonly", width=12,
                     values=["平铺整页", "居中一个"]).grid(row=5, column=1, sticky="w", padx=6)

        self.scope, self.expr, self.err = self.scope_widgets(body, 6)
        self.out_file_row(body, 7, self.out_file, "水印.pdf")
        body.columnconfigure(1, weight=1)
        self._refresh()

    def _refresh(self):
        if self.kind.get() == "文字水印":
            self.img_frame.grid_remove()
            self.text_frame.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        else:
            self.text_frame.grid_remove()
            self.img_frame.grid(row=3, column=0, columnspan=4, sticky="w", pady=4)

    def _pick_img(self):
        p = filedialog.askopenfilename(
            title="选择水印图片", parent=self,
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp")])
        if p:
            self.img_path = p
            self.img_lbl.config(text=os.path.basename(p), foreground="#333")

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        # 主线程先取值，子线程只用普通 Python 值
        src, pw = self.app.src_path, self.app.open_pw
        op = max(5, min(100, self.opacity.get())) / 100.0
        layout = self.layout.get()
        if self.kind.get() == "文字水印":
            text = self.text.get().strip()
            if not text:
                messagebox.showwarning("缺少文字", "请输入水印文字。", parent=self)
                return
            fs = max(8, min(200, self.fontsize.get()))
            color = COLOR_PRESETS.get(self.color.get(), (0.75, 0.75, 0.75))
            ang = float(self.angle.get()) % 360
            work = lambda: (out_path, do_watermark_text(src, out_path, pages, text,
                                                        fs, color, ang, op, layout, pw))
        else:
            if not self.img_path:
                messagebox.showwarning("缺少图片", "请先选择水印图片。", parent=self)
                return
            img = self.img_path
            sc = max(5, min(200, self.scale.get()))
            work = lambda: (out_path, do_watermark_image(src, out_path, pages, img,
                                                         sc, op, layout, 0, pw))
        self.run_in_thread(
            work,
            lambda res: ("水印完成",
                         f"已生成 {os.path.basename(res[0])}（{len(pages)} 页）",
                         res[0], os.path.dirname(res[0])))


# ---------------- PDF 转 Word ----------------
class PdfToWordDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}.docx")
        super().__init__(app, "PDF 转 Word", "开始转换")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        self.scope, self.expr, self.err = self.scope_widgets(body, 1)
        ttk.Label(body, text="说明：适用于文字型 / 图文型 PDF，图片会一并抽出嵌入；\n"
                             "复杂版式（多栏、表格、公式）保真度有限，转换后建议人工校对。\n"
                             "原文件保持不变。",
                  justify="left", foreground="#999").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=6, pady=4)
        self.out_file_row(body, 4, self.out_file, "转换结果.docx",
                          ext=".docx", filetypes=[("Word 文档", "*.docx")])
        body.columnconfigure(1, weight=1)

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_pdf_to_docx(src, out_path, pages, pw))
        self.run_in_thread(
            work,
            lambda res: ("转换完成",
                         f"已生成 {os.path.basename(res[0])}（{res[1]} 页）。建议打开校对排版。",
                         res[0], os.path.dirname(res[0])))


# ---------------- PDF 转 Excel ----------------
class PdfToExcelDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}.xlsx")
        super().__init__(app, "PDF 转 Excel", "开始转换")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        self.scope, self.expr, self.err = self.scope_widgets(body, 1)
        ttk.Label(body, text="说明：自动识别文档中的表格，每张表生成一个工作表；\n"
                             "无表格的页面按文本行写入一列。仅对可识别表格的 PDF 有效，\n"
                             "扫描件需先做 OCR。原文件保持不变。",
                  justify="left", foreground="#999").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=6, pady=4)
        self.out_file_row(body, 4, self.out_file, "转换结果.xlsx",
                          ext=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        body.columnconfigure(1, weight=1)

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path,) + do_pdf_to_xlsx(src, out_path, pages, pw)
        self.run_in_thread(
            work,
            lambda res: ("转换完成",
                         f"已转换 {res[1]} 页，识别到 {res[2]} 个表格，"
                         f"生成 {os.path.basename(res[0])}。",
                         res[0], os.path.dirname(res[0])))


# ---------------- 书签 / 目录管理 ----------------
class TocDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        try:
            self.toc = [list(x) for x in get_toc_list(app.src_path, app.open_pw)]
        except Exception:
            self.toc = []
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_书签.pdf")
        super().__init__(app, "书签 / 目录管理", "生成新文件")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="书签条目（缩进表示层级，目标页为 1 基页码；双击可编辑）").grid(
            row=1, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=2, column=0, columnspan=3, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=10, width=46)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.listbox.bind("<Double-Button-1>", lambda *_: self._edit())
        btns = ttk.Frame(body)
        btns.grid(row=2, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="新增", command=self._add).pack(fill="x", pady=2)
        ttk.Button(btns, text="编辑选中", command=self._edit).pack(fill="x", pady=2)
        ttk.Button(btns, text="删除选中", command=self._del).pack(fill="x", pady=2)
        ttk.Button(btns, text="上移", command=lambda: self._move(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="下移", command=lambda: self._move(1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="降级", command=lambda: self._level(1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="升级", command=lambda: self._level(-1)).pack(fill="x", pady=2)
        ttk.Button(btns, text="读取原始书签", command=self._reload).pack(fill="x", pady=(12, 2))
        ttk.Button(btns, text="清空", command=self._clear).pack(fill="x", pady=2)
        self.out_file_row(body, 3, self.out_file, "书签.pdf")
        body.columnconfigure(1, weight=1)
        self._refresh()

    def _refresh(self):
        self.listbox.delete(0, "end")
        for lvl, title, page in self.toc:
            self.listbox.insert(
                "end", f"{'    ' * (max(1, int(lvl)) - 1)}{title}  ·  第 {int(page)} 页")

    def _sel(self):
        s = self.listbox.curselection()
        return s[0] if s else -1

    def _add(self):
        item = self._ask_item(1, "", 1)
        if item:
            i = self._sel()
            at = i + 1 if i >= 0 else len(self.toc)
            self.toc.insert(at, list(item))
            self._refresh()
            self.listbox.selection_set(at)

    def _edit(self):
        i = self._sel()
        if i < 0:
            return
        item = self._ask_item(self.toc[i][0], self.toc[i][1], self.toc[i][2])
        if item:
            self.toc[i] = list(item)
            self._refresh()
            self.listbox.selection_set(i)

    def _del(self):
        i = self._sel()
        if i < 0:
            return
        self.toc.pop(i)
        self._refresh()

    def _move(self, delta):
        i = self._sel()
        if i < 0:
            return
        j = i + delta
        if j < 0 or j >= len(self.toc):
            return
        self.toc[i], self.toc[j] = self.toc[j], self.toc[i]
        self._refresh()
        self.listbox.selection_set(j)

    def _level(self, delta):
        i = self._sel()
        if i < 0:
            return
        self.toc[i][0] = max(1, min(6, int(self.toc[i][0]) + delta))
        self._refresh()
        self.listbox.selection_set(i)

    def _reload(self):
        try:
            self.toc = [list(x) for x in get_toc_list(self.app.src_path, self.app.open_pw)]
        except Exception as e:
            messagebox.showerror("读取失败", str(e), parent=self)
            return
        self._refresh()

    def _clear(self):
        if self.toc and not messagebox.askyesno("清空书签", "确定清空所有书签条目吗？",
                                                parent=self):
            return
        self.toc = []
        self._refresh()

    def _ask_item(self, level, title, page):
        dlg = tk.Toplevel(self)
        dlg.title("书签条目")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.configure(bg=Theme.CARD)
        ttk.Label(dlg, text="标题").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        t_var = tk.StringVar(value=str(title))
        ttk.Entry(dlg, textvariable=t_var, width=34).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(dlg, text="目标页").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        p_var = tk.IntVar(value=max(1, min(self.total_pages, int(page or 1))))
        ttk.Spinbox(dlg, from_=1, to=max(1, self.total_pages), textvariable=p_var,
                    width=8).grid(row=1, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(dlg, text="层级").grid(row=2, column=0, sticky="w", padx=10, pady=6)
        l_var = tk.IntVar(value=max(1, min(6, int(level))))
        ttk.Spinbox(dlg, from_=1, to=6, textvariable=l_var, width=8).grid(
            row=2, column=1, sticky="w", padx=6, pady=6)
        result = {"v": None}

        def ok():
            t = t_var.get().strip()
            if not t:
                messagebox.showwarning("缺少标题", "请输入书签标题。", parent=dlg)
                return
            result["v"] = (l_var.get(), t, p_var.get())
            dlg.destroy()

        bar = ttk.Frame(dlg)
        bar.grid(row=3, column=0, columnspan=2, pady=(4, 12))
        ttk.Button(bar, text="确定", command=ok).pack(side="left", padx=8)
        ttk.Button(bar, text="取消", command=dlg.destroy).pack(side="left", padx=8)
        dlg.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dlg.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        dlg.grab_set()
        self.wait_window(dlg)
        return result["v"]

    def run(self):
        if not self.toc:
            messagebox.showwarning("没有书签", "请至少新增 1 条书签。", parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        toc = [list(x) for x in self.toc]
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_set_toc(src, out_path, toc, pw))
        self.run_in_thread(
            work,
            lambda res: ("书签已写入",
                         f"已为 {os.path.basename(res[0])} 写入 {res[1]} 条书签。",
                         res[0], os.path.dirname(res[0])))


# ---------------- 表单填写 ----------------
class FormFillDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        try:
            self.fields = _fields_from_doc(app.doc)
        except Exception:
            self.fields = []
        self.vars = []
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_已填写.pdf")
        super().__init__(app, "表单填写", "保存填写结果")
        if not self.fields:
            self.run_btn.config(state="disabled")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        if not self.fields:
            ttk.Label(body, text="未检测到可填写的表单字段。\n"
                                 "该文档可能不是交互式表单（AcroForm），或字段已被扁平化。\n"
                                 "可改用「插入文字」直接在页面上填写。",
                      justify="left", foreground="#c0392b").grid(
                row=1, column=0, columnspan=4, sticky="w", pady=6)
            body.columnconfigure(1, weight=1)
            return
        ttk.Label(body, text=f"共检测到 {len(self.fields)} 个可填写字段（复选框以勾选框显示，"
                             f"下拉项以列表选择）").grid(
            row=1, column=0, columnspan=4, sticky="w", pady=(0, 4))
        wrap = ttk.Frame(body)
        wrap.grid(row=2, column=0, columnspan=4, sticky="nwe", pady=4)
        canvas = tk.Canvas(wrap, height=300, highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda *_: canvas.configure(scrollregion=canvas.bbox("all")))
        for r, f in enumerate(self.fields):
            label = f"{f['name']}（{f['type']} · 第 {f['page']} 页）"
            ttk.Label(inner, text=label).grid(row=r, column=0, sticky="w",
                                              padx=(0, 8), pady=3)
            if "CheckBox" in f["type"]:
                v = tk.BooleanVar(value=str(f["value"]).strip().lower()
                                  in ("true", "1", "on", "yes"))
                ttk.Checkbutton(inner, variable=v).grid(row=r, column=1,
                                                        sticky="w", pady=3)
            elif f["choices"]:
                v = tk.StringVar(value=f["value"])
                ttk.Combobox(inner, textvariable=v, values=f["choices"],
                             width=30).grid(row=r, column=1, sticky="w", pady=3)
            else:
                v = tk.StringVar(value=f["value"])
                ttk.Entry(inner, textvariable=v, width=34).grid(
                    row=r, column=1, sticky="w", pady=3)
            self.vars.append((f["name"], v))
        ttk.Label(body, text="填写结果另存为新文件，原文件保持不变。",
                  foreground="#999").grid(row=3, column=0, columnspan=4, sticky="w")
        self.out_file_row(body, 4, self.out_file, "已填写.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        values = {}
        for name, v in self.vars:
            try:
                values[name] = v.get()
            except tk.TclError:
                values[name] = ""
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_fill_form(src, out_path, values, pw))
        self.run_in_thread(
            work,
            lambda res: ("填写完成",
                         f"已填写 {res[1]} 个字段，生成 {os.path.basename(res[0])}。",
                         res[0], os.path.dirname(res[0])))


# ---------------- 批量处理 ----------------
class BatchDialog(ToolDialog):
    def __init__(self, app):
        self.files = []
        self._batch_out = ""
        self.op = tk.StringVar(value="添加文字水印")
        # 水印参数
        self.wm_text = tk.StringVar(value="机密")
        self.wm_fontsize = tk.IntVar(value=48)
        self.wm_angle = tk.IntVar(value=45)
        self.wm_opacity = tk.IntVar(value=20)
        self.wm_color = tk.StringVar(value="浅灰")
        self.wm_layout = tk.StringVar(value="平铺整页")
        # 页码参数
        self.pn_position = tk.StringVar(value="底部居中")
        self.pn_fmt = tk.StringVar(value="第 X 页 / 共 N 页")
        self.pn_start = tk.IntVar(value=1)
        # 压缩 / 加密 / 旋转
        self.compress_q = tk.IntVar(value=50)
        self.enc_pw = tk.StringVar()
        self.dec_pw = tk.StringVar()
        self.rot_angle = tk.StringVar(value="90")
        self.out_dir = tk.StringVar(value="")
        super().__init__(app, "批量处理", "开始批量处理")

    def build(self, body):
        ttk.Label(body, text="待处理的 PDF 文件（可一次选择多个）").grid(
            row=0, column=0, columnspan=4, sticky="w")
        frame = ttk.Frame(body)
        frame.grid(row=1, column=0, columnspan=3, sticky="nwe", pady=4)
        self.listbox = themed_listbox(frame, height=8, width=46)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        btns = ttk.Frame(body)
        btns.grid(row=1, column=3, sticky="n", padx=(6, 0))
        ttk.Button(btns, text="添加文件…", command=self._add).pack(fill="x", pady=2)
        ttk.Button(btns, text="移除选中", command=self._del).pack(fill="x", pady=2)
        ttk.Button(btns, text="清空", command=self._clear).pack(fill="x", pady=2)

        ttk.Label(body, text="处理操作").grid(row=2, column=0, sticky="w", pady=4)
        cb = ttk.Combobox(body, textvariable=self.op, state="readonly", width=16,
                          values=BATCH_OPS)
        cb.grid(row=2, column=1, sticky="w", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda *_: self._refresh_op())

        self.param_holder = ttk.Frame(body)
        self.param_holder.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)
        self._build_param_frames()
        self.out_dir_row(body, 4, self.out_dir, label="输出文件夹")
        ttk.Label(body, text="所有文件只读处理，结果按原文件名另存到输出文件夹，重名自动加序号。",
                  foreground="#999").grid(row=5, column=0, columnspan=4, sticky="w")
        body.columnconfigure(1, weight=1)
        self._refresh_op()

    def _build_param_frames(self):
        h = self.param_holder
        self.f_wm = ttk.Frame(h)
        ttk.Label(self.f_wm, text="水印文字").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.f_wm, textvariable=self.wm_text, width=16).grid(
            row=0, column=1, padx=6)
        ttk.Label(self.f_wm, text="字号").grid(row=0, column=2, padx=(10, 0))
        ttk.Spinbox(self.f_wm, from_=8, to=200, textvariable=self.wm_fontsize,
                    width=5).grid(row=0, column=3, padx=4)
        ttk.Label(self.f_wm, text="角度°").grid(row=0, column=4, padx=(10, 0))
        ttk.Spinbox(self.f_wm, from_=0, to=359, textvariable=self.wm_angle,
                    width=5).grid(row=0, column=5, padx=4)
        ttk.Label(self.f_wm, text="颜色").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(self.f_wm, textvariable=self.wm_color, state="readonly", width=6,
                     values=list(COLOR_PRESETS.keys())).grid(
            row=1, column=1, padx=6, pady=(6, 0), sticky="w")
        ttk.Label(self.f_wm, text="透明度%").grid(row=1, column=2, padx=(10, 0),
                                                pady=(6, 0))
        ttk.Spinbox(self.f_wm, from_=5, to=100, textvariable=self.wm_opacity,
                    width=5).grid(row=1, column=3, padx=4, pady=(6, 0))
        ttk.Label(self.f_wm, text="排版").grid(row=1, column=4, padx=(10, 0), pady=(6, 0))
        ttk.Combobox(self.f_wm, textvariable=self.wm_layout, state="readonly", width=9,
                     values=["平铺整页", "居中一个"]).grid(
            row=1, column=5, padx=4, pady=(6, 0), sticky="w")

        self.f_pn = ttk.Frame(h)
        ttk.Label(self.f_pn, text="位置").grid(row=0, column=0, sticky="w")
        ttk.Combobox(self.f_pn, textvariable=self.pn_position, state="readonly",
                     width=10, values=["底部居中", "底部左侧", "底部右侧",
                                       "顶部左侧", "顶部右侧"]).grid(
            row=0, column=1, padx=6)
        ttk.Label(self.f_pn, text="格式").grid(row=0, column=2, padx=(10, 0))
        ttk.Combobox(self.f_pn, textvariable=self.pn_fmt, state="readonly", width=16,
                     values=["第 X 页 / 共 N 页", "第 X 页", "X / N", "X"]).grid(
            row=0, column=3, padx=4)
        ttk.Label(self.f_pn, text="起始页号").grid(row=0, column=4, padx=(10, 0))
        ttk.Spinbox(self.f_pn, from_=1, to=9999, textvariable=self.pn_start,
                    width=6).grid(row=0, column=5, padx=4)

        self.f_cp = ttk.Frame(h)
        ttk.Label(self.f_cp, text="压缩质量").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(self.f_cp, from_=10, to=90, textvariable=self.compress_q,
                    width=6).grid(row=0, column=1, padx=6)
        ttk.Label(self.f_cp, text="数字越小体积越小、越模糊（50 兼顾清晰度）",
                  foreground="#999").grid(row=0, column=2, padx=6)

        self.f_enc = ttk.Frame(h)
        ttk.Label(self.f_enc, text="打开密码").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.f_enc, textvariable=self.enc_pw, show="*", width=20).grid(
            row=0, column=1, padx=6)
        ttk.Label(self.f_enc, text="所有文件使用同一密码", foreground="#999").grid(
            row=0, column=2, padx=6)

        self.f_rot = ttk.Frame(h)
        ttk.Label(self.f_rot, text="旋转角度").grid(row=0, column=0, sticky="w")
        ttk.Combobox(self.f_rot, textvariable=self.rot_angle, state="readonly", width=8,
                     values=["90", "180", "270"]).grid(row=0, column=1, padx=6)
        ttk.Label(self.f_rot, text="对每个文件的所有页顺时针旋转", foreground="#999").grid(
            row=0, column=2, padx=6)

        self.f_none = ttk.Frame(h)
        ttk.Label(self.f_none, text="打开密码（文件未加密时留空）").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.f_none, textvariable=self.dec_pw, show="*", width=20).grid(
            row=0, column=1, padx=6)
        ttk.Label(self.f_none, text="已加密的文件需填对密码才能解密", foreground="#999").grid(
            row=0, column=2, padx=6)

        self._op_frames = {
            "添加文字水印": self.f_wm, "添加页码": self.f_pn, "压缩体积": self.f_cp,
            "加密保护": self.f_enc, "解密 / 去密码": self.f_none, "旋转页面": self.f_rot,
        }

    def _refresh_op(self):
        for fr in self._op_frames.values():
            fr.pack_forget()
        self._op_frames[self.op.get()].pack(anchor="w", fill="x")

    def _refresh(self):
        self.listbox.delete(0, "end")
        for p in self.files:
            self.listbox.insert("end", os.path.basename(p))

    def _add(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件", parent=self, filetypes=[("PDF 文件", "*.pdf")])
        if not paths:
            return
        for p in paths:
            if p.lower().endswith(".pdf") and p not in self.files:
                self.files.append(p)
        self._refresh()
        if not self.out_dir.get() and self.files:
            self.out_dir.set(os.path.join(os.path.dirname(self.files[0]), "批量结果"))

    def _del(self):
        s = self.listbox.curselection()
        if not s:
            return
        self.files.pop(s[0])
        self._refresh()

    def _clear(self):
        self.files = []
        self._refresh()

    def on_progress(self, done, total):
        self.set_status(f"处理中… {done}/{total}")

    def run(self):
        if not self.files:
            messagebox.showwarning("没有文件", "请先添加待处理的 PDF 文件。", parent=self)
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("缺少输出", "请选择输出文件夹。", parent=self)
            return
        op = self.op.get()
        params = {}
        if op == "添加文字水印":
            t = self.wm_text.get().strip()
            if not t:
                messagebox.showwarning("缺少文字", "请输入水印文字。", parent=self)
                return
            params = {"text": t,
                      "fontsize": max(8, min(200, self.wm_fontsize.get())),
                      "color": COLOR_PRESETS.get(self.wm_color.get(), (0.75, 0.75, 0.75)),
                      "angle": float(self.wm_angle.get()) % 360,
                      "opacity": max(5, min(100, self.wm_opacity.get())) / 100.0,
                      "layout": self.wm_layout.get()}
        elif op == "添加页码":
            params = {"position": self.pn_position.get(), "fmt": self.pn_fmt.get(),
                      "start": max(1, self.pn_start.get()), "fontsize": 10}
        elif op == "压缩体积":
            params = {"quality": max(10, min(90, self.compress_q.get()))}
        elif op == "加密保护":
            p = self.enc_pw.get()
            if not p:
                messagebox.showwarning("缺少密码", "请输入打开密码。", parent=self)
                return
            params = {"password": p}
        elif op == "旋转页面":
            params = {"rotate": int(self.rot_angle.get())}
        elif op == "解密 / 去密码":
            params = {"pw": self.dec_pw.get().strip()}
        files = list(self.files)
        self._batch_out = out_dir
        work = lambda: do_batch(op, files, out_dir, params, self.report_progress)
        self.run_in_thread(work, self._summary)

    def _summary(self, results):
        ok = [r for r in results if not r[2]]
        fail = [r for r in results if r[2]]
        msg = f"成功 {len(ok)} 个，失败 {len(fail)} 个。"
        if fail:
            msg += f"\n首个失败：{os.path.basename(fail[0][0])} —— {fail[0][2]}"
        return ("批量处理完成", msg, None, self._batch_out)


# ---------------- 页内取区控件（标注/遮盖共用） ----------------
class RegionPicker(ttk.Frame):
    """页内取区控件：显示页面位图，鼠标拖拽框选矩形，坐标换算为 PDF 点。
    regions 形如 [(页号0基, fitz.Rect), ...]，坐标为未旋转页面坐标。"""

    MAX_W = 560
    MAX_H = 400

    def __init__(self, master, app, on_change=None):
        super().__init__(master)
        self.app = app
        self.on_change = on_change
        self.regions = []
        self.pi = app.cur if 0 <= app.cur < app.total else 0
        self.zoom = 1.0
        self._drag0 = None
        self._photo = None

        nav = ttk.Frame(self)
        nav.pack(fill="x")
        ttk.Button(nav, text="上一页", width=7,
                   command=lambda: self.step(-1)).pack(side="left")
        self.lbl = ttk.Label(nav, text="", width=14, anchor="center")
        self.lbl.pack(side="left", padx=4)
        ttk.Button(nav, text="下一页", width=7,
                   command=lambda: self.step(1)).pack(side="left")
        ttk.Button(nav, text="清除本页选区", command=self.clear_page).pack(
            side="left", padx=(10, 0))
        ttk.Button(nav, text="清除全部", command=self.clear_all).pack(side="left", padx=4)
        self.cnt = ttk.Label(nav, text="已选 0 处", foreground=Theme.OK)
        self.cnt.pack(side="right")

        self.canvas = tk.Canvas(self, bg=Theme.CANVAS_BG, highlightthickness=0, bd=0)
        self.canvas.pack(pady=4)
        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._motion)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        ttk.Label(self, text="在页面上按住鼠标拖拽即可框选区域；可翻页在多页分别框选，"
                             "所选区域会以橙色虚线示意。", style="Muted.TLabel",
                  wraplength=self.MAX_W).pack(anchor="w")
        self.load_page()

    # ---- 坐标换算 ----
    def _page(self):
        return self.app.doc[self.pi]

    def _dims(self):
        page = self._page()
        return page.mediabox.width, page.mediabox.height, (page.rotation or 0) % 360

    def _region_canvas_rect(self, rect):
        w, h, rot = self._dims()
        pts = [_rot_map(rot, w, h, rect.x0, rect.y0),
               _rot_map(rot, w, h, rect.x1, rect.y0),
               _rot_map(rot, w, h, rect.x0, rect.y1),
               _rot_map(rot, w, h, rect.x1, rect.y1)]
        xs = [p[0] * self.zoom for p in pts]
        ys = [p[1] * self.zoom for p in pts]
        return min(xs), min(ys), max(xs), max(ys)

    def load_page(self):
        if not self.app.doc or self.pi >= self.app.total:
            return
        page = self._page()
        r = page.rect
        self.zoom = min(self.MAX_W / r.width, self.MAX_H / r.height)
        pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
        self._photo = tk.PhotoImage(data=pix.tobytes("png"))
        self.canvas.delete("all")
        self.canvas.config(width=pix.width, height=pix.height)
        self.canvas.create_image(0, 0, anchor="nw", image=self._photo)
        self.lbl.config(text=f"第 {self.pi + 1} / {self.app.total} 页")
        for pno, rect in self.regions:
            if pno == self.pi:
                x0, y0, x1, y1 = self._region_canvas_rect(rect)
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="#ff9800",
                                             width=2, dash=(4, 3))
        self._update_count()

    def step(self, d):
        if not self.app.doc:
            return
        self.pi = max(0, min(self.app.total - 1, self.pi + d))
        self.load_page()

    def clear_page(self):
        self.regions = [(p, r) for (p, r) in self.regions if p != self.pi]
        self.load_page()

    def clear_all(self):
        self.regions = []
        self.load_page()

    def _update_count(self):
        self.cnt.config(text=f"已选 {len(self.regions)} 处")

    def _press(self, e):
        self._drag0 = (e.x, e.y)

    def _motion(self, e):
        if self._drag0 is None:
            return
        self.canvas.delete("rubber")
        self.canvas.create_rectangle(self._drag0[0], self._drag0[1], e.x, e.y,
                                     outline="#ff5722", width=2, tags="rubber")

    def _release(self, e):
        if self._drag0 is None:
            return
        x0, y0 = self._drag0
        self._drag0 = None
        self.canvas.delete("rubber")
        if abs(e.x - x0) < 4 or abs(e.y - y0) < 4:
            return
        w, h, rot = self._dims()
        u0, v0 = _rot_unmap(rot, w, h, min(x0, e.x) / self.zoom, min(y0, e.y) / self.zoom)
        u1, v1 = _rot_unmap(rot, w, h, max(x0, e.x) / self.zoom, max(y0, e.y) / self.zoom)
        rect = fitz.Rect(min(u0, u1), min(v0, v1), max(u0, u1), max(v0, v1))
        self.regions.append((self.pi, rect))
        self.load_page()
        if self.on_change:
            self.on_change()


# ---------------- 标注 / 批注 ----------------
class AnnotateDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.kind = tk.StringVar(value="高亮")
        self.color = tk.StringVar(value="黄")
        self.text = tk.StringVar(value="")
        self.fontsize = tk.IntVar(value=12)
        self.mode = tk.StringVar(value="按关键词定位")
        self.keyword = tk.StringVar()
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_标注.pdf")
        super().__init__(app, "标注 / 批注", "开始标注")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        row = ttk.Frame(body)
        row.grid(row=1, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(row, text="标注类型").pack(side="left")
        cb = ttk.Combobox(row, textvariable=self.kind, state="readonly", width=7,
                          values=ANNOT_TYPES)
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda *_: self._refresh())
        ttk.Label(row, text="颜色").pack(side="left")
        ttk.Combobox(row, textvariable=self.color, state="readonly", width=5,
                     values=list(ANNOT_COLORS.keys())).pack(side="left", padx=6)
        self.txt_lbl = ttk.Label(row, text="批注文字")
        self.txt_e = ttk.Entry(row, textvariable=self.text, width=20)
        self.txt_lbl.pack(side="left")
        self.txt_e.pack(side="left", padx=6)
        self.fs_lbl = ttk.Label(row, text="字号")
        self.fs_sp = ttk.Spinbox(row, from_=8, to=48, textvariable=self.fontsize, width=4)
        self.fs_lbl.pack(side="left")
        self.fs_sp.pack(side="left", padx=4)

        mrow = ttk.Frame(body)
        mrow.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(mrow, text="定位方式").pack(side="left")
        ttk.Radiobutton(mrow, text="按关键词定位", variable=self.mode,
                        value="按关键词定位", command=self._refresh).pack(side="left", padx=4)
        ttk.Radiobutton(mrow, text="手动框选区域", variable=self.mode,
                        value="手动框选", command=self._refresh).pack(side="left", padx=4)

        self.kw_frame = ttk.Frame(body)
        self.kw_frame.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)
        ttk.Label(self.kw_frame, text="关键词").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.kw_frame, textvariable=self.keyword, width=28).grid(
            row=0, column=1, padx=6, sticky="w")
        ttk.Label(self.kw_frame, text="查找该文字并自动标注（扫描件无效）", foreground="#999").grid(
            row=0, column=2, sticky="w")
        self.scope, self.expr, self.err = self.scope_widgets(self.kw_frame, 1)

        self.pick_frame = ttk.Frame(body)
        ttk.Label(self.pick_frame, text="在页面上框选需要标注的区域：", foreground="#666").pack(
            anchor="w")
        self.picker = RegionPicker(self.pick_frame, self.app)

        self.out_file_row(body, 4, self.out_file, "标注.pdf")
        body.columnconfigure(1, weight=1)
        self._refresh()

    def _refresh(self):
        is_text = self.kind.get() in ("便签", "文本框")
        if is_text:
            self.txt_lbl.pack(side="left")
            self.txt_e.pack(side="left", padx=6)
        else:
            self.txt_lbl.pack_forget()
            self.txt_e.pack_forget()
        if self.kind.get() == "文本框":
            self.fs_lbl.pack(side="left")
            self.fs_sp.pack(side="left", padx=4)
        else:
            self.fs_lbl.pack_forget()
            self.fs_sp.pack_forget()
        self.kw_frame.grid_remove()
        self.pick_frame.grid_remove()
        if self.mode.get() == "按关键词定位":
            self.kw_frame.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)
        else:
            self.pick_frame.grid(row=3, column=0, columnspan=4, sticky="w", pady=4)

    def run(self):
        kind = self.kind.get()
        color = ANNOT_COLORS.get(self.color.get(), (1.0, 0.90, 0.0))
        text = self.text.get().strip()
        fs = max(8, min(48, self.fontsize.get()))
        if kind in ("便签", "文本框") and not text:
            messagebox.showwarning("缺少内容", "请输入批注文字内容。", parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        if self.mode.get() == "按关键词定位":
            kw = self.keyword.get().strip()
            if not kw:
                messagebox.showwarning("缺少关键词", "请输入要查找的关键词。", parent=self)
                return
            try:
                pages = self.resolve_scope(self.scope, self.expr, self.err)
            except ValueError as e:
                messagebox.showerror("页码有误", str(e), parent=self)
                return
            work = lambda: (out_path, do_annotate(src, out_path, kind, pages=pages,
                                                  keyword=kw, text=text, color=color,
                                                  fontsize=fs, pw=pw))
        else:
            areas = [(p, (r.x0, r.y0, r.x1, r.y1)) for p, r in self.picker.regions]
            if not areas:
                messagebox.showwarning("没有选区", "请先在页面上框选至少一个区域。",
                                       parent=self)
                return
            work = lambda: (out_path, do_annotate(src, out_path, kind, areas=areas,
                                                  text=text, color=color,
                                                  fontsize=fs, pw=pw))
        self.run_in_thread(
            work,
            lambda res: ("标注完成",
                         f"已添加 {res[1]} 处标注，生成 {os.path.basename(res[0])}",
                         res[0], os.path.dirname(res[0])))


# ---------------- 涂黑遮盖 ----------------
class RedactDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.fill = tk.StringVar(value="黑色")
        self.mode = tk.StringVar(value="手动框选")
        self.keyword = tk.StringVar()
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_遮盖.pdf")
        super().__init__(app, "涂黑遮盖", "开始遮盖")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        row = ttk.Frame(body)
        row.grid(row=1, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(row, text="遮盖颜色").pack(side="left")
        ttk.Combobox(row, textvariable=self.fill, state="readonly", width=6,
                     values=["黑色", "白色"]).pack(side="left", padx=6)
        ttk.Label(row, text="遮盖后下方文字会被抹除，无法复制", foreground="#999").pack(
            side="left", padx=6)
        ttk.Label(row, text="并清除文档属性/书签/XMP（防原文残留被反查）",
                  foreground="#c0392b").pack(side="left", padx=6)

        mrow = ttk.Frame(body)
        mrow.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(mrow, text="定位方式").pack(side="left")
        ttk.Radiobutton(mrow, text="手动框选区域", variable=self.mode,
                        value="手动框选", command=self._refresh).pack(side="left", padx=4)
        ttk.Radiobutton(mrow, text="按关键词遮盖", variable=self.mode,
                        value="按关键词遮盖", command=self._refresh).pack(side="left", padx=4)

        self.kw_frame = ttk.Frame(body)
        self.kw_frame.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)
        ttk.Label(self.kw_frame, text="关键词").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.kw_frame, textvariable=self.keyword, width=28).grid(
            row=0, column=1, padx=6, sticky="w")
        ttk.Label(self.kw_frame, text="查找该文字并遮盖（扫描件无效）", foreground="#999").grid(
            row=0, column=2, sticky="w")
        self.scope, self.expr, self.err = self.scope_widgets(self.kw_frame, 1)

        self.pick_frame = ttk.Frame(body)
        ttk.Label(self.pick_frame, text="在页面上框选需要遮盖的区域：", foreground="#666").pack(
            anchor="w")
        self.picker = RegionPicker(self.pick_frame, self.app)

        self.out_file_row(body, 4, self.out_file, "遮盖.pdf")
        body.columnconfigure(1, weight=1)
        self._refresh()

    def _refresh(self):
        self.kw_frame.grid_remove()
        self.pick_frame.grid_remove()
        if self.mode.get() == "手动框选":
            self.pick_frame.grid(row=3, column=0, columnspan=4, sticky="w", pady=4)
        else:
            self.kw_frame.grid(row=3, column=0, columnspan=4, sticky="we", pady=4)

    def run(self):
        fill = (0, 0, 0) if self.fill.get() == "黑色" else (1, 1, 1)
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        src, pw = self.app.src_path, self.app.open_pw
        if self.mode.get() == "手动框选":
            areas = [(p, (r.x0, r.y0, r.x1, r.y1)) for p, r in self.picker.regions]
            if not areas:
                messagebox.showwarning("没有选区", "请先在页面上框选至少一个区域。",
                                       parent=self)
                return
            work = lambda: (out_path, do_redact(src, out_path, areas=areas, fill=fill, pw=pw))
        else:
            kw = self.keyword.get().strip()
            if not kw:
                messagebox.showwarning("缺少关键词", "请输入要遮盖的关键词。", parent=self)
                return
            try:
                pages = self.resolve_scope(self.scope, self.expr, self.err)
            except ValueError as e:
                messagebox.showerror("页码有误", str(e), parent=self)
                return
            work = lambda: (out_path, do_redact(src, out_path, pages=pages, keyword=kw,
                                                fill=fill, pw=pw))
        self.run_in_thread(
            work,
            lambda res: ("遮盖完成",
                         f"已遮盖 {res[1]} 处，生成 {os.path.basename(res[0])}"
                         "（遮盖处文字已抹除）",
                         res[0], os.path.dirname(res[0])))


# ---------------- 页眉页脚 ----------------
class HeaderFooterDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.header = tk.StringVar(value="")
        self.footer = tk.StringVar(value="第 {page} 页 / 共 {total} 页")
        self.size = tk.IntVar(value=10)
        self.margin = tk.IntVar(value=10)
        self.align = tk.StringVar(value="居中")
        self.start = tk.IntVar(value=1)
        base = os.path.splitext(app.src_path)[0]
        self.out_file = tk.StringVar(value=f"{base}_页眉页脚.pdf")
        super().__init__(app, "页眉页脚", "开始添加")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="页眉文字").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.header, width=44).grid(
            row=1, column=1, columnspan=2, sticky="we", padx=6, pady=4)
        ttk.Label(body, text="页脚文字").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(body, textvariable=self.footer, width=44).grid(
            row=2, column=1, columnspan=2, sticky="we", padx=6, pady=4)
        ttk.Label(body, text="占位符：{page}=当前页号，{total}=总页数；留空则不添加该项",
                  foreground="#999").grid(row=3, column=0, columnspan=4, sticky="w")

        row = ttk.Frame(body)
        row.grid(row=4, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(row, text="字号").pack(side="left")
        ttk.Spinbox(row, from_=6, to=48, textvariable=self.size, width=5).pack(
            side="left", padx=4)
        ttk.Label(row, text="页边距(mm)").pack(side="left", padx=(12, 0))
        ttk.Spinbox(row, from_=3, to=40, textvariable=self.margin, width=5).pack(
            side="left", padx=4)
        ttk.Label(row, text="对齐").pack(side="left", padx=(12, 0))
        ttk.Combobox(row, textvariable=self.align, state="readonly", width=6,
                     values=["居中", "左侧", "右侧"]).pack(side="left", padx=4)
        ttk.Label(row, text="起始页号").pack(side="left", padx=(12, 0))
        ttk.Spinbox(row, from_=1, to=9999, textvariable=self.start, width=6).pack(
            side="left", padx=4)

        self.scope, self.expr, self.err = self.scope_widgets(body, 5)
        self.out_file_row(body, 6, self.out_file, "页眉页脚.pdf")
        body.columnconfigure(1, weight=1)

    def run(self):
        if not self.header.get().strip() and not self.footer.get().strip():
            messagebox.showwarning("缺少内容", "请至少填写页眉或页脚文字。", parent=self)
            return
        try:
            pages = self.resolve_scope(self.scope, self.expr, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        out_path = self.out_file.get().strip()
        if not out_path:
            messagebox.showwarning("缺少输出", "请选择输出文件位置。", parent=self)
            return
        out_path = unique_out_path(out_path)
        cfg = {"header": self.header.get(), "footer": self.footer.get(),
               "size": max(6, min(48, self.size.get())),
               "margin_mm": max(3, min(40, self.margin.get())),
               "align": self.align.get(), "pages": pages,
               "start": max(1, self.start.get()), "color": (0, 0, 0)}
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: (out_path, do_header_footer(src, out_path, cfg, pw))
        self.run_in_thread(
            work,
            lambda res: ("添加完成",
                         f"已为 {res[1]} 页添加页眉/页脚，生成 {os.path.basename(res[0])}",
                         res[0], os.path.dirname(res[0])))


# ---------------- 提取文档内图片 ----------------
class ExtractImagesDialog(ToolDialog):
    def __init__(self, app):
        self.total_pages = app.total
        self.fmt = tk.StringVar(value="原格式")
        base = os.path.splitext(os.path.basename(app.src_path))[0]
        self.out_dir = tk.StringVar(value=os.path.join(
            os.path.dirname(app.src_path), f"{base}_图片"))
        super().__init__(app, "提取文档内图片", "开始提取")

    def build(self, body):
        self.doc_head(body, 0, self.total_pages)
        ttk.Label(body, text="图片格式").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(body, textvariable=self.fmt, state="readonly", width=10,
                     values=["原格式", "PNG", "JPG"]).grid(
            row=1, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(body, text="「原格式」保持 PDF 内图片的原始编码，不做转换",
                  foreground="#999").grid(row=1, column=2, sticky="w")
        self.out_dir_row(body, 2, self.out_dir, label="输出文件夹")
        ttk.Label(body, text="导出每一页内嵌的图片（同一张图片多次出现只导出一份）。",
                  foreground="#999").grid(row=3, column=0, columnspan=4, sticky="w")
        body.columnconfigure(1, weight=1)

    def run(self):
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("缺少输出", "请选择输出文件夹。", parent=self)
            return
        want = {"原格式": "auto", "PNG": "png", "JPG": "jpg"}[self.fmt.get()]
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: do_extract_images(src, out_dir, pw=pw, want_fmt=want,
                                         progress=self.report_progress)
        self.run_in_thread(
            work,
            lambda outs: ("提取完成",
                          f"共导出 {len(outs)} 张图片到文件夹。",
                          None, out_dir))

    def on_progress(self, done, total):
        self.set_status(f"正在提取… {done}/{total} 页")


# ---------------- 打印 ----------------
class PrintDialog(ToolDialog):
    def __init__(self, app, names, default):
        self.printer_var = tk.StringVar(value=default if default in names else names[0])
        self.names = names
        self.copies_var = tk.IntVar(value=1)
        super().__init__(app, "打印", "开始打印")

    def build(self, body):
        self.scope, self.expr_var, self.err = self.scope_widgets(body, 1)
        ttk.Label(body, text="打印机").grid(row=0, column=0, sticky="w", padx=4, pady=5)
        ttk.Combobox(body, textvariable=self.printer_var, values=self.names,
                     state="readonly", width=34).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Label(body, text="份数").grid(row=2, column=0, sticky="w", padx=4, pady=5)
        ttk.Spinbox(body, from_=1, to=99, textvariable=self.copies_var, width=6).grid(
            row=2, column=1, sticky="w", pady=5)
        ttk.Label(body, text="页面将按纸张可打印区域等比缩放并居中",
                  foreground="#999").grid(row=3, column=0, columnspan=2,
                                          sticky="w", padx=4, pady=5)

    def run(self):
        try:
            pages = self.resolve_scope(self.scope, self.expr_var, self.err)
        except ValueError as e:
            messagebox.showerror("页码有误", str(e), parent=self)
            return
        printer = self.printer_var.get()
        copies = max(1, self.copies_var.get())
        src, pw = self.app.src_path, self.app.open_pw
        work = lambda: do_print(src, printer, pages, copies,
                                progress=self.report_progress, pw=pw)
        self.run_in_thread(
            work,
            lambda res: ("打印完成", f"已发送 {res} 页到打印机「{printer}」。", None, None))

    def on_progress(self, done, total):
        self.set_status(f"正在打印 {done}/{total} …")


# ---------------- 手写签名 ----------------
class SignDialog(ToolDialog):
    """手写签名：鼠标/数位笔书写 → 导出透明 PNG → 交给可视化编辑器拖动定位后另存。"""

    W, H = 560, 200
    SCALE = 3          # 导出超采样倍数（保证放大不糊）
    TARGET_MM = 60     # 100% 时签名宽度（毫米）

    def __init__(self, app):
        self.pen = tk.IntVar(value=4)
        self.strokes = []      # [(points, width), ...]
        self._cur = None
        super().__init__(app, "手写签名", "使用这个签名")

    def build(self, body):
        ttk.Label(body, text="在下面方框里手写签名（支持鼠标 / 数位笔 / 触摸屏）").grid(
            row=0, column=0, columnspan=4, sticky="w")
        wrap = tk.Frame(body, bg="#FFFFFF", highlightthickness=1,
                        highlightbackground=Theme.BORDER)
        wrap.grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 8))
        self.canvas = tk.Canvas(wrap, width=self.W, height=self.H, bg="#FFFFFF",
                                highlightthickness=0, cursor="pencil")
        self.canvas.pack(padx=1, pady=1)
        self.canvas.bind("<ButtonPress-1>", self._down)
        self.canvas.bind("<B1-Motion>", self._move)
        self.canvas.bind("<ButtonRelease-1>", self._up)

        row = ttk.Frame(body)
        row.grid(row=2, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(row, text="笔画粗细").pack(side="left")
        ttk.Spinbox(row, from_=1, to=12, textvariable=self.pen, width=4).pack(
            side="left", padx=6)
        ttk.Button(row, text="撤销一笔", command=self.undo).pack(side="left", padx=(14, 4))
        ttk.Button(row, text="清空", command=self.clear).pack(side="left", padx=4)

        ttk.Label(body, text="签好后点「使用这个签名」，再在页面上拖动 / 缩放定位，"
                             "确认后另存为新文件（原文件不变）。",
                  style="Muted.TLabel", wraplength=520, justify="left").grid(
            row=3, column=0, columnspan=4, sticky="w")

    # ---- 书写 ----
    def _down(self, e):
        self._cur = [(e.x, e.y)]
        self.strokes.append((self._cur, max(1, self.pen.get())))

    def _move(self, e):
        if not self._cur:
            return
        x0, y0 = self._cur[-1]
        self._cur.append((e.x, e.y))
        self.canvas.create_line(x0, y0, e.x, e.y, fill="#111111",
                                width=max(1, self.pen.get()), capstyle="round",
                                smooth=True, tags="ink")

    def _up(self, _e):
        if self._cur and len(self._cur) == 1:      # 单击 = 点状笔画
            x, y = self._cur[0]
            r = max(1, self.pen.get()) / 2
            self.canvas.create_oval(x - r, y - r, x + r, y + r,
                                    fill="#111111", outline="", tags="ink")
        self._cur = None

    def _repaint(self):
        self.canvas.delete("ink")
        for pts, wd in self.strokes:
            if len(pts) == 1:
                x, y = pts[0]
                r = wd / 2
                self.canvas.create_oval(x - r, y - r, x + r, y + r,
                                        fill="#111111", outline="", tags="ink")
            for i in range(1, len(pts)):
                self.canvas.create_line(*pts[i - 1], *pts[i], fill="#111111",
                                        width=wd, capstyle="round", smooth=True,
                                        tags="ink")

    def undo(self):
        if self.strokes:
            self.strokes.pop()
            self._repaint()

    def clear(self):
        self.strokes = []
        self._repaint()

    # ---- 导出 + 进入可视化定位 ----
    def run(self):
        if not self.strokes:
            messagebox.showwarning("还没有签名", "请先在方框里写一个签名。", parent=self)
            return
        path = self._export()
        if not path:
            return
        self.destroy()
        self.app.start_editor("image", preset_image=path)

    def _export(self):
        from PIL import Image as PILImage, ImageDraw
        s = self.SCALE
        im = PILImage.new("RGBA", (self.W * s, self.H * s), (0, 0, 0, 0))
        dr = ImageDraw.Draw(im)
        for pts, wd in self.strokes:
            xy = [(x * s, y * s) for x, y in pts]
            wpx = max(1, int(round(wd * s)))
            if len(xy) == 1:
                x, y = xy[0]
                r = wpx / 2
                dr.ellipse([x - r, y - r, x + r, y + r], fill=(17, 17, 17, 255))
            else:
                dr.line(xy, fill=(17, 17, 17, 255), width=wpx, joint="curve")
        # 写入 DPI，使 100% 缩放大致等于 TARGET_MM 宽（之后还能在编辑器里缩放）
        dpi = max(72, int(round(im.width * 72.0 / (self.TARGET_MM * MM))))
        path = os.path.join(tempfile.gettempdir(), f"pdf_sign_{os.getpid()}.png")
        try:
            im.save(path, dpi=(dpi, dpi))
        except Exception as e:
            messagebox.showerror("导出失败", f"无法生成签名图片：{e}", parent=self)
            return None
        return path


# ============================================================
# 可视化编辑器（插入文字 / 插入图片）
# ============================================================

class Editor(Runner):
    """页面上直接放置文字/图片对象：拖动移动，控件实时调整大小与角度，确认后另存新文件。
    坐标状态一律保存在 PDF 点（页面左上角为原点），渲染时乘 zoom 转画布坐标。"""

    def __init__(self, app, kind, preset_image=None):
        self.app = app
        self.kind = kind  # "text" / "image"
        self._run_text = "保存为新文件"
        self.range_var = tk.StringVar(value="当前页")
        self.status_var = tk.StringVar(value="在页面上拖动可移动位置")
        r = app.doc[app.cur].rect
        self.fs = 24
        self.angle = 0.0        # 顺时针度数（与输出一致）
        self.pct = 100 if preset_image else 20
        self.rotate = 0         # 图片 90 度倍数
        self.img_path = ""
        self._img_pil = None
        self._img_photo = None
        self._img_dpi = (96.0, 96.0)
        self._drag0 = None
        self._resizing = False
        self.text_var = tk.StringVar(value="批注内容")
        self.color = tk.StringVar(value="红")
        self.fs_var = tk.IntVar(value=24)
        self.angle_var = tk.DoubleVar(value=0.0)
        self.pct_var = tk.IntVar(value=self.pct)
        self.vector_var = tk.BooleanVar(value=False)  # 可选文本（矢量插入，可搜索/复制）
        # 初始放置：文字水平居中、垂直中部；图片居中
        self.x_pt = r.width / 2
        self.y_pt = r.height / 2
        if kind == "text":
            col = COLOR_PRESETS.get(self.color.get(), (0.75, 0.75, 0.75))
            w = text_image(self.text_var.get(), self.fs, col).width / TEXT_SS
            self.x_pt = max(6.0, (r.width - w) / 2)
        else:
            self.x_pt = r.width * 0.4
        self._build_bar()
        self.run_btn = self.save_btn
        self.redraw()
        if kind == "image":
            if preset_image:
                self.root.after(0, lambda: self._load_image(preset_image))
            else:
                self.root.after(60, self.choose_image)

    @property
    def root(self):
        return self.app.root

    # ---------------- 编辑栏 ----------------
    def _build_bar(self):
        s = self.app.editor_slot
        for w in s.winfo_children():
            w.destroy()
        kind_lbl = "插入文字" if self.kind == "text" else "插入图片"
        ttk.Label(s, text=f"{kind_lbl}（编辑中）：", style="Editor.TLabel",
                  font=("", 10, "bold")).pack(side="left", padx=(8, 2))

        if self.kind == "text":
            ttk.Label(s, text="内容", style="Editor.TLabel").pack(side="left")
            e = ttk.Entry(s, textvariable=self.text_var, width=22)
            e.pack(side="left", padx=3)
            ttk.Label(s, text="字号", style="Editor.TLabel").pack(side="left")
            ttk.Spinbox(s, from_=8, to=200, textvariable=self.fs_var, width=5).pack(
                side="left", padx=3)
            ttk.Label(s, text="颜色", style="Editor.TLabel").pack(side="left")
            ttk.Combobox(s, textvariable=self.color, state="readonly", width=5,
                         values=list(COLOR_PRESETS.keys())).pack(side="left", padx=3)
            ttk.Label(s, text="角度°", style="Editor.TLabel").pack(side="left")
            ttk.Spinbox(s, from_=0, to=359, textvariable=self.angle_var, width=5).pack(
                side="left", padx=3)
            ttk.Checkbutton(s, text="可选文本", style="Editor.TCheckbutton",
                            variable=self.vector_var).pack(side="left", padx=(6, 0))
        else:
            ttk.Button(s, text="选择图片…", command=self.choose_image).pack(
                side="left", padx=2)
            self.img_lbl = ttk.Label(s, text="未选择", style="Editor.TLabel",
                                     foreground="#c0392b", width=16)
            self.img_lbl.pack(side="left")
            ttk.Label(s, text="缩放%", style="Editor.TLabel").pack(side="left")
            ttk.Spinbox(s, from_=5, to=300, textvariable=self.pct_var, width=5).pack(
                side="left", padx=3)
            self.rot_btn = ttk.Button(s, text="旋转90°(0)", command=self._rot90)
            self.rot_btn.pack(side="left", padx=3)
            ttk.Label(s, text="100=图片原始大小，拖右下角蓝块也可缩放",
                      style="Editor.TLabel", foreground="#5B7BB8").pack(
                side="left", padx=4)

        ttk.Label(s, text="应用范围", style="Editor.TLabel").pack(side="left")
        ttk.Combobox(s, textvariable=self.range_var, state="readonly", width=7,
                     values=["当前页", "所有页"]).pack(side="left", padx=3)
        self.save_btn = ttk.Button(s, text=self._run_text, style="Accent.TButton",
                                   command=self.save)
        self.save_btn.pack(side="left", padx=(10, 2))
        ttk.Button(s, text="取消", command=lambda: self.cancel()).pack(
            side="left", padx=2)
        ttk.Label(s, textvariable=self.status_var, style="Editor.TLabel",
                  foreground=Theme.OK).pack(side="left", padx=8)

        self.text_var.trace_add("write", lambda *_: self.redraw())
        self.fs_var.trace_add("write", lambda *_: self._on_fs())
        self.angle_var.trace_add("write", lambda *_: self._on_angle())
        self.pct_var.trace_add("write", lambda *_: self._on_pct())
        self.color.trace_add("write", lambda *_: self.redraw())

    def _on_fs(self):
        try:
            self.fs = max(8, min(200, int(self.fs_var.get())))
        except Exception:
            return
        self.redraw()

    def _on_angle(self):
        try:
            self.angle = float(self.angle_var.get()) % 360
        except Exception:
            return
        self.redraw()

    def _on_pct(self):
        try:
            self.pct = max(5, min(300, int(self.pct_var.get())))
        except Exception:
            return
        self.redraw()

    def _rot90(self):
        self.rotate = (self.rotate + 90) % 360
        self.rot_btn.config(text=f"旋转90°({self.rotate})")
        self.redraw()

    def choose_image(self):
        path = filedialog.askopenfilename(
            title="选择图片", parent=self.app.root,
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp")])
        if not path:
            if not self.img_path:
                self.set_status("已取消选择，可继续选图或点「取消」退出编辑")
            return
        self._load_image(path)

    def _load_image(self, path):
        from PIL import Image as PILImage
        try:
            self._img_pil = PILImage.open(path)
            self._img_pil.load()
        except Exception:
            messagebox.showwarning("无法读取", "图片格式不支持或文件已损坏。",
                                   parent=self.app.root)
            return
        self.img_path = path
        # 记录图片 DPI：预览与输出共用，100% = 图片真实物理尺寸
        self._img_dpi = (96.0, 96.0)
        try:
            dpi = self._img_pil.info.get("dpi")
            if dpi:
                dx, dy = float(dpi[0]), float(dpi[1])
                if 10 <= dx <= 2400:
                    self._img_dpi = (dx, self._img_dpi[1])
                if 10 <= dy <= 2400:
                    self._img_dpi = (self._img_dpi[0], dy)
        except Exception:
            pass
        try:
            self.img_lbl.config(text=os.path.basename(path), foreground="#333")
        except Exception:
            pass
        self.redraw()

    # ---------------- 预览渲染 ----------------
    def _rotated_pil(self):
        # PDF rotate=90 为顺时针；PIL rotate 正值为逆时针，故取负
        return self._img_pil.rotate(-self.rotate, expand=True) if self.rotate else self._img_pil

    @staticmethod
    def _map_pt(page_rot, pw, ph, u, v):
        """未旋转页面坐标 (u,v) → 旋转后（显示）坐标（旋转为顺时针）。"""
        if page_rot == 90:
            return (ph - v, u)
        if page_rot == 180:
            return (pw - u, ph - v)
        if page_rot == 270:
            return (v, pw - u)
        return (u, v)

    def _display_topleft(self, page_rot, pw, ph, x, y, w, h):
        """未旋转矩形 (x,y,w,h) 旋转后，显示包围盒的左上角坐标。
        必须映射四角再取最小，否则 90/270 会整体偏移一个对象尺寸。"""
        pts = (self._map_pt(page_rot, pw, ph, x, y),
               self._map_pt(page_rot, pw, ph, x + w, y),
               self._map_pt(page_rot, pw, ph, x, y + h),
               self._map_pt(page_rot, pw, ph, x + w, y + h))
        return min(p[0] for p in pts), min(p[1] for p in pts)

    def redraw(self):
        cv = self.app.canvas
        if not cv.winfo_exists():
            return
        cv.delete("ovl")
        # 用页面位图的真实显示比例（整数像素对齐后的比例），而非名义 zoom，
        # 保证叠加层与页面位图逐像素对齐，消除微抖。
        zx = self.app._scale_x or 1.0
        zy = self.app._scale_y or 1.0
        # 页面自身 rotation 会被渲染时应用；预览坐标需转到旋转后坐标系，
        # 否则预览方向/位置与保存后阅读器显示不一致（"跳角度"）。
        page = self.app.doc[self.app.cur]
        page_rot = page.rotation or 0
        # 注意：page.rect 对 rotation=90/270 的页面返回已旋转的宽高，
        # 必须用 mediabox（未旋转）的宽高做坐标变换，否则位置会偏移。
        pw, ph = page.mediabox.width, page.mediabox.height
        x, y = self.x_pt, self.y_pt
        if self.kind == "text":
            t = self.text_var.get() or "预览文字"
            col = COLOR_PRESETS.get(self.color.get(), (0.75, 0.75, 0.75))
            # 页面空间尺寸（仅含文字自身角度），用于旋转后包围盒定位
            base = text_image(t, self.fs, col, self.angle)
            w, h = base.width / TEXT_SS, base.height / TEXT_SS
            # 显示用图：在页面空间基础上叠加页面 rotation
            img = base.rotate(-page_rot, expand=True) if page_rot else base
            tx, ty = self._display_topleft(page_rot, pw, ph, x, y, w, h)
            cx = 12 + tx * zx
            cy = 12 + ty * zy
            disp = img.resize((max(1, round(img.width * zx / TEXT_SS)),
                               max(1, round(img.height * zy / TEXT_SS))))
            buf = io.BytesIO()
            disp.save(buf, "PNG")
            self._img_photo = tk.PhotoImage(data=buf.getvalue())
            self._item = cv.create_image(cx, cy, anchor="nw",
                                         image=self._img_photo, tags="ovl")
        else:
            if self._img_pil is None:
                return
            # 图片在页面空间的占位尺寸（点）：按图片自身 DPI 还原物理尺寸，
            # 与 do_insert_image 完全同一套换算，保证预览=输出。90/270 旋转时宽高互换。
            iw, ih = self._img_pil.size
            kx = (72.0 / self._img_dpi[0]) * self.pct / 100.0
            ky = (72.0 / self._img_dpi[1]) * self.pct / 100.0
            if self.rotate in (90, 270):
                kx, ky = ky, kx
            w, h = iw * kx, ih * ky
            if self.rotate in (90, 270):
                w, h = h, w
            # 页面 rotation 需叠加到图片预览旋转上
            disp = self._rotated_pil()
            if page_rot:
                disp = disp.rotate(-page_rot, expand=True)
            kx, ky = kx * zx, ky * zy
            disp = disp.resize((max(1, round(disp.width * kx)),
                                max(1, round(disp.height * ky))))
            buf = io.BytesIO()
            disp.save(buf, "PNG")
            self._img_photo = tk.PhotoImage(data=buf.getvalue())
            tx, ty = self._display_topleft(page_rot, pw, ph, x, y, w, h)
            cx = 12 + tx * zx
            cy = 12 + ty * zy
            self._item = cv.create_image(cx, cy, anchor="nw",
                                         image=self._img_photo, tags="ovl")
            bb = cv.bbox(self._item)
            if bb:
                cv.create_rectangle(bb[2] - 9, bb[3] - 9, bb[2], bb[3],
                                    fill="#4a90d9", outline="white",
                                    tags=("ovl", "hnd"))
        cv.tag_bind("ovl", "<ButtonPress-1>", self._press)
        cv.tag_bind("ovl", "<B1-Motion>", self._motion)
        cv.tag_bind("ovl", "<ButtonRelease-1>", self._release)
        cv.tag_bind("hnd", "<ButtonPress-1>", self._hnd_press)
        cv.tag_bind("hnd", "<B1-Motion>", self._hnd_motion)

    # ---------------- 交互 ----------------
    def _press(self, e):
        if self._resizing:
            return
        self._drag0 = (e.x, e.y)
        self.set_status("拖动中…松开完成")

    def _motion(self, e):
        if self._drag0 is None or self._resizing:
            return
        dx, dy = e.x - self._drag0[0], e.y - self._drag0[1]
        self._drag0 = (e.x, e.y)
        z = self.app.zoom or 1.0
        # 画布位移转换为页面坐标系位移（考虑页面 rotation 与真实显示比例）
        page_rot = (self.app.doc[self.app.cur].rotation or 0) % 360
        dxp = dx / (self.app._scale_x or z)
        dyp = dy / (self.app._scale_y or z)
        if page_rot == 90:
            self.x_pt += dyp
            self.y_pt -= dxp
        elif page_rot == 180:
            self.x_pt -= dxp
            self.y_pt -= dyp
        elif page_rot == 270:
            self.x_pt -= dyp
            self.y_pt += dxp
        else:
            self.x_pt += dxp
            self.y_pt += dyp
        self.app.canvas.move("ovl", dx, dy)

    def _release(self, _e):
        self._drag0 = None
        self._resizing = False

    def _hnd_press(self, e):
        self._resizing = True
        bb = self.app.canvas.bbox(self._item)
        self._hnd0 = (e.x, e.y, self.pct, bb)

    def _hnd_motion(self, e):
        bb = self._hnd0[3]
        x0, y0 = bb[0], bb[1]
        d0 = max(1.0, math.hypot(self._hnd0[0] - x0, self._hnd0[1] - y0))
        d1 = math.hypot(e.x - x0, e.y - y0)
        self.pct = int(max(5, min(300, self._hnd0[2] * d1 / d0)))
        self.pct_var.set(self.pct)
        self.redraw()

    # ---------------- 保存 / 退出 ----------------
    def save(self):
        app = self.app
        # 在主线程先捕获全部 tk 变量值，子线程只用普通 Python 值（线程安全）
        pages = (list(range(app.total)) if self.range_var.get() == "所有页"
                 else [app.cur])
        base = os.path.splitext(app.src_path)[0]
        if self.kind == "text":
            text = self.text_var.get().strip()
            if not text:
                messagebox.showwarning("缺少文字", "请输入要插入的文字内容。",
                                       parent=app.root)
                return
            out = unique_out_path(f"{base}_已加文字.pdf")
            # 直接从 tk 变量读取最新值，避免 trace 异步更新导致取值滞后
            ang = float(self.angle_var.get()) % 360
            fs = max(8, min(200, int(self.fs_var.get())))
            color = COLOR_PRESETS.get(self.color.get(), (0.75, 0.75, 0.75))
            xy = (self.x_pt / MM, self.y_pt / MM)
            src, pw = app.src_path, app.open_pw
            vec = bool(self.vector_var.get())
            work = lambda: do_insert_text(src, out, pages, text, fs, color,
                                          "自定义坐标", xy, ang, pw=pw, vector=vec)
        else:
            if not self.img_path:
                messagebox.showwarning("缺少图片", "请先选择要插入的图片。",
                                       parent=app.root)
                return
            out = unique_out_path(f"{base}_已加图片.pdf")
            xy = (self.x_pt / MM, self.y_pt / MM)
            src, pw = app.src_path, app.open_pw
            img = self.img_path
            pct = max(5, min(300, int(self.pct_var.get())))
            rot = self.rotate
            work = lambda: do_insert_image(src, out, img, pages, "自定义坐标",
                                           xy, pct, rot, pw=pw)
        self.run_in_thread(
            work,
            lambda _res: ("插入完成",
                          f"已生成 {os.path.basename(out)}"
                          f"（{len(pages)} 页）", out, os.path.dirname(out)))

    def _on_done(self, result, on_done):
        """保存成功：退出编辑态并自动切换打开新文件，可继续在新文件上编辑。"""
        out = result
        self.cancel()
        if self.app.root.winfo_exists():
            self.app.open_file(out)

    def cancel(self):
        self.app.editor = None
        self.app.editor_slot.pack_forget()
        self.app.canvas.delete("ovl")


# ============================================================
# 主窗口（文档中心）
# ============================================================

class App:
    def __init__(self, show_about=True):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(f"3Q PDF开源免费工具 {VERSION}")
        self.root.geometry("1320x820")
        self.root.minsize(1040, 680)

        # 当前文档状态
        self.doc = None
        self.src_path = ""
        self.open_pw = ""
        self.total = 0
        self.cur = 0
        self.zoom = 1.0
        self._scale_x = 1.0
        self._scale_y = 1.0
        self.fit = True
        self.editor = None
        self._photo = None
        self._render_job = None
        self._thumb_gen = 0
        self._thumb_queue = None
        self._thumb_widgets = []
        # 搜索状态
        self._hit_page = -1
        self._hits = []
        self._hit_idx = -1
        # 文字选择状态
        self._words = []        # [(fitz.Rect 显示坐标, 词), ...]
        self._sel0 = None       # 拖拽起点（画布像素）
        self._sel_text = ""
        # 画笔（自由手绘标记）：笔迹先留在内存 + 画布覆盖层，
        # Ctrl+S 覆盖保存时才写进文档（写进去就撤不回，所以保存前可撤销/清空）
        self.ink_mode = False
        self.ink_color = INK_PRESETS["红"]
        self.ink_width = 3
        self.ink_shape = "free"  # free / line / rect / ellipse
        self.ink_strokes = []   # [{"page","color","width","shape","points":[..]}]
        self._ink_drawing = False
        self._ink_item = None
        self._ink_pts = []
        self._close_after_save = False
        self._saving = False    # 存盘中：render 让路，避免和子线程抢同一个 doc
        # MCP 智能体接入（默认关闭，必须用户主动开）
        self.mcp_srv = None
        self.mcp_port = 0
        self.mcp_kind = ""      # 正在跑的 MCP 服务档位："local" / "remote"

        # ---- 全局视觉 ----
        setup_style(self.root)
        self.root.configure(bg=Theme.APP_BG)
        self._icon = None
        icon_png = make_icon_png()
        if icon_png:
            try:
                self._icon = tk.PhotoImage(data=icon_png)
                self.root.iconphoto(True, self._icon)
            except tk.TclError:
                self._icon = None

        # ---- 顶部棱彩线 + 工具栏（深色主视觉）----
        _gradient_bar(self.root, (Theme.PRIMARY, "#7C3AED", Theme.GOLD,
                                  "#0E9F6E"), height=3, steps=110)
        tb = ttk.Frame(self.root, style="Toolbar.TFrame", padding=(10, 4))
        tb.pack(fill="x")
        # 右侧固定簇先 pack：窗口再窄也不会被左侧按钮挤掉（此前「关于」默认被挤没）
        tb_right = ttk.Frame(tb, style="Toolbar.TFrame")
        tb_right.pack(side="right")
        ttk.Button(tb_right, text="建议与赞助", style="Donate.TButton",
                   command=self.open_about).pack(side="right", padx=(8, 0))
        ttk.Button(tb, text="打开 PDF…", style="Accent.TButton",
                   command=self.open_file).pack(side="left")
        self.recent_btn = ttk.Menubutton(tb, text="最近打开", style="Ghost.TMenubutton")
        self.recent_btn.pack(side="left", padx=(10, 0))
        self._rebuild_recent_menu()
        self.print_btn = ttk.Button(tb, text="打印…", style="Ghost.TButton",
                                    command=self.open_print, state="disabled")
        self.print_btn.pack(side="left")
        ttk.Separator(tb, orient="vertical", style="Toolbar.TSeparator").pack(
            side="left", fill="y", padx=10, pady=2)
        self.tool_btns = {}
        for label, cmd in [("插入文字", lambda: self.start_editor("text")),
                           ("插入图片", lambda: self.start_editor("image")),
                           ("签名", self.open_sign),
                           ("画笔", self.toggle_ink),
                           ("保存", self.save_inplace)]:
            b = ttk.Button(tb, text=label, style="Ghost.TButton", command=cmd,
                           state="disabled")
            b.pack(side="left", padx=2)
            self.tool_btns[label] = b
        ttk.Separator(tb, orient="vertical", style="Toolbar.TSeparator").pack(
            side="left", fill="y", padx=10, pady=2)
        ttk.Button(tb, text="智能体接入…", style="Ghost.TButton",
                   command=self.open_mcp).pack(side="left")
        self.panel_btn = ttk.Button(tb, text="收起工具面板", style="Ghost.TButton",
                                    command=self.toggle_tool_panel)
        self.panel_btn.pack(side="left", padx=(10, 0))
        tools = tk.Menu(self.root, tearoff=0)
        style_menu(tools)

        def sub(label, items):
            sm = tk.Menu(tools, tearoff=0)
            style_menu(sm)
            for text, cmd in items:
                sm.add_command(label=text, command=cmd)
            tools.add_cascade(label=label, menu=sm)

        sub("页面整理", [("页面提取…", self.open_extract),
                        ("拆分…", self.open_split),
                        ("页面管理（重排）…", self.open_reorder),
                        ("删除页面…", self.open_delete),
                        ("旋转页面…", self.open_rotate),
                        ("裁剪页面 / 去黑边…", self.open_crop)])
        sub("编辑与标注", [("添加页码…", self.open_pagenum),
                          ("添加水印…", self.open_watermark),
                          ("页眉页脚…", self.open_hf),
                          ("标注/批注…", self.open_annotate),
                          ("画笔标记（手绘/直线/矩形/椭圆）", self.toggle_ink),
                          ("保存到原文件（Ctrl+S）", self.save_inplace),
                          ("另存为新文件…", self.save_as),
                          ("涂黑遮盖…", self.open_redact)])
        sub("转换与导出", [("PDF 转 Word…", self.open_pdf2word),
                          ("PDF 转 Excel…", self.open_pdf2excel),
                          ("PDF 转图片…", self.open_pdf2img),
                          ("提取文档内图片…", self.open_extract_img)])
        sub("安全与优化", [("压缩体积…", self.open_compress),
                          ("压缩到指定大小…", self.open_compress_target),
                          ("加密保护…", self.open_encrypt),
                          ("解密/去密码…", self.open_decrypt)])
        sub("文档增强", [("书签/目录管理…", self.open_toc),
                        ("表单填写…", self.open_form)])
        sub("文件工具", [("合并…", self.open_merge),
                        ("图片转 PDF…", self.open_img2pdf),
                        ("Office 转 PDF…", self.open_office2pdf),
                        ("批量处理…", self.open_batch)])
        ttk.Menubutton(tb, text="更多工具", menu=tools,
                       style="Ghost.TMenubutton").pack(side="left", padx=(10, 0))
        self._tools_menu = tools
        # 顶栏与内容之间的金色细线（渐变，比纯色更有质感）
        _gradient_bar(self.root, (Theme.DARK, Theme.GOLD_D, Theme.GOLD,
                                  Theme.GOLD_D, Theme.DARK),
                      height=2, steps=140)

        # ---- 翻页 / 缩放栏（白色工具条，页码用品牌蓝强调）----
        nav = ttk.Frame(self.root, style="Card.TFrame", padding=(10, 6))
        nav.pack(fill="x")
        ttk.Button(nav, text="|<", width=3, command=lambda: self.goto(0)).pack(side="left")
        ttk.Button(nav, text="<", width=3,
                   command=lambda: self.goto(self.cur - 1)).pack(side="left", padx=(2, 6))
        self.page_var = tk.StringVar(value="-")
        pe = ttk.Entry(nav, textvariable=self.page_var, width=6, justify="center")
        pe.pack(side="left")
        pe.bind("<Return>", lambda *_: self.jump_from_entry())
        self.total_lbl = ttk.Label(nav, text="/ 0", style="Nav.TLabel")
        self.total_lbl.pack(side="left", padx=(2, 6))
        ttk.Button(nav, text=">", width=3,
                   command=lambda: self.goto(self.cur + 1)).pack(side="left")
        ttk.Button(nav, text=">|", width=3,
                   command=lambda: self.goto(self.total - 1)).pack(side="left", padx=(2, 12))
        ttk.Button(nav, text="－", width=3,
                   command=lambda: self.set_zoom(self.zoom / 1.25)).pack(side="left")
        self.zoom_lbl = ttk.Label(nav, text="适应窗口", style="Muted.TLabel",
                                  width=8, anchor="center")
        self.zoom_lbl.pack(side="left", padx=2)
        ttk.Button(nav, text="＋", width=3,
                   command=lambda: self.set_zoom(self.zoom * 1.25)).pack(side="left")
        ttk.Button(nav, text="适应窗口", command=self.set_fit).pack(side="left", padx=(6, 0))
        ttk.Button(nav, text="旋转视图", command=self.rotate_view).pack(side="left", padx=(6, 0))
        ttk.Button(nav, text="搜索", command=lambda: self.toggle_search(True)).pack(
            side="left", padx=(6, 0))
        ttk.Label(nav, text="Ctrl+滚轮缩放 · ←→翻页 · Ctrl+F 搜索",
                  style="Muted.TLabel").pack(side="right")

        # ---- 编辑栏插槽（编辑器激活时显示） ----
        self.editor_slot = ttk.Frame(self.root, style="Editor.TFrame", padding=(10, 6))

        # ---- 底部状态栏（深色） ----
        sb = ttk.Frame(self.root, style="Status.TFrame", padding=(10, 5))
        sb.pack(fill="x", side="bottom")
        self.st_path = ttk.Label(sb, text="未打开文件", style="Status.TLabel", anchor="w")
        self.st_path.pack(side="left", fill="x", expand=True)
        self.st_page = ttk.Label(sb, text="", style="StatusBright.TLabel",
                                 width=18, anchor="e")
        self.st_page.pack(side="right")
        self.st_zoom = ttk.Label(sb, text="", style="StatusBright.TLabel",
                                 width=10, anchor="e")
        self.st_zoom.pack(side="right")

        # ---- 中部：左缩略图 + 右阅读区 ----
        mid = ttk.Frame(self.root, style="App.TFrame")
        mid.pack(fill="both", expand=True)
        self.mid = mid

        side = ttk.Frame(mid, style="Card.TFrame", width=208)
        side.pack(side="left", fill="y", padx=(10, 0), pady=(0, 8))
        side.pack_propagate(False)
        self.side_nb = ttk.Notebook(side, style="Side.TNotebook")
        self.side_nb.pack(fill="both", expand=True, padx=2, pady=2)
        page_tab = ttk.Frame(self.side_nb, style="Card.TFrame")
        self.side_nb.add(page_tab, text="页面")
        twrap = ttk.Frame(page_tab, style="Card.TFrame")
        twrap.pack(fill="both", expand=True)
        self.thumb_canvas = tk.Canvas(twrap, bg="#EDF1F8", highlightthickness=0, bd=0)
        tsb = ttk.Scrollbar(twrap, orient="vertical", command=self.thumb_canvas.yview)
        self.thumb_canvas.configure(yscrollcommand=tsb.set)
        self.thumb_canvas.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        self.thumb_inner = ttk.Frame(self.thumb_canvas)
        self._thumb_win = self.thumb_canvas.create_window(
            (0, 0), window=self.thumb_inner, anchor="nw")
        self.thumb_inner.bind("<Configure>", lambda *_: self.thumb_canvas.configure(
            scrollregion=self.thumb_canvas.bbox("all")))
        self.thumb_canvas.bind("<Configure>", lambda e: self.thumb_canvas.itemconfig(
            self._thumb_win, width=e.width))
        self.thumb_canvas.bind("<MouseWheel>", lambda e: self.thumb_canvas.yview_scroll(
            int(-e.delta / 120), "units"))
        self.thumb_inner.bind("<MouseWheel>", lambda e: self.thumb_canvas.yview_scroll(
            int(-e.delta / 120), "units"))
        self._build_task_tab()

        # ---- 右侧工具面板：全部工具平铺点选，免翻「更多工具」二级菜单 ----
        # 必须先于阅读区 pack（side="right"），否则阅读区 expand 会把它挤没。
        self._build_tool_panel(mid)

        right = ttk.Frame(mid, style="App.TFrame")
        right.pack(side="left", fill="both", expand=True, padx=(10, 10), pady=(0, 8))
        self.viewer = right
        self.canvas = tk.Canvas(right, bg=Theme.CANVAS_BG, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(right, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        self.canvas.pack(fill="both", expand=True)

        # ---- 搜索栏（浮动在阅读区右上角，白色卡片） ----
        self.search_bar = tk.Frame(right, bg=Theme.CARD, bd=0, highlightthickness=1,
                                   highlightbackground="#B9C6DB")
        self.search_var = tk.StringVar()
        ttk.Label(self.search_bar, text="查找：").pack(side="left", padx=(8, 0), pady=6)
        se = ttk.Entry(self.search_bar, textvariable=self.search_var, width=22)
        se.pack(side="left", padx=6, pady=5)
        se.bind("<Return>", lambda *_: self.search_next())
        se.bind("<Shift-Return>", lambda *_: self.search_next(True))
        ttk.Button(self.search_bar, text="上一个", width=6,
                   command=lambda: self.search_next(True)).pack(side="left", padx=2)
        ttk.Button(self.search_bar, text="下一个", width=6,
                   command=lambda: self.search_next()).pack(side="left", padx=2)
        ttk.Button(self.search_bar, text="×", width=3,
                   command=lambda: self.toggle_search(False)).pack(
            side="left", padx=(2, 8))

        # ---- 画笔工具条（浮动在阅读区左上角，进入画笔模式才显示）----
        self.ink_bar = tk.Frame(right, bg=Theme.CARD, bd=0, highlightthickness=1,
                                highlightbackground="#B9C6DB")
        ttk.Label(self.ink_bar, text="画笔", style="Title.TLabel").pack(
            side="left", padx=(8, 4), pady=6)
        self.ink_swatches = {}
        for _nm, _hx in INK_PRESETS.items():
            sw = tk.Label(self.ink_bar, text="", bg=_hx, width=2, height=1,
                          cursor="hand2", bd=1, relief="flat")
            sw.pack(side="left", padx=2)
            sw.bind("<Button-1>",
                    lambda _e, n=_nm, h=_hx: self._pick_ink_color(n, h))
            self.ink_swatches[_nm] = sw
        ttk.Separator(self.ink_bar, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=4)
        self.ink_shape_btns = {}
        for _sh, _lb in (("free", "手绘"), ("line", "直线"),
                         ("rect", "矩形"), ("ellipse", "椭圆")):
            sb = ttk.Button(self.ink_bar, text=_lb, width=4,
                            style="PanelGhost.TButton",
                            command=lambda s=_sh: self._pick_ink_shape(s))
            sb.pack(side="left", padx=1)
            self.ink_shape_btns[_sh] = sb
        ttk.Separator(self.ink_bar, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=4)
        self.ink_w_btns = {}
        for _pt, _lb in ((2, "细"), (4, "中"), (8, "粗")):
            wb = ttk.Button(self.ink_bar, text=_lb, width=3,
                            style="PanelGhost.TButton",
                            command=lambda p=_pt: self._pick_ink_width(p))
            wb.pack(side="left", padx=1)
            self.ink_w_btns[_pt] = wb
        ttk.Separator(self.ink_bar, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=4)
        ttk.Button(self.ink_bar, text="撤销", width=4, style="PanelGhost.TButton",
                   command=self.ink_undo).pack(side="left", padx=1)
        ttk.Button(self.ink_bar, text="清空本页", width=6,
                   style="PanelGhost.TButton",
                   command=self.ink_clear_page).pack(side="left", padx=1)
        ttk.Button(self.ink_bar, text="退出", width=4, style="PanelGhost.TButton",
                   command=lambda: self.toggle_ink(False)).pack(side="left", padx=1)
        ttk.Label(self.ink_bar, text="Ctrl+S 保存",
                  style="Muted.TLabel").pack(side="left", padx=(6, 8))
        self._sync_ink_bar()

        # ---- 事件 ----
        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.bind("<Enter>", lambda *_: self.canvas.focus_set())
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.canvas.bind("<Control-MouseWheel>", self._ctrl_wheel)
        # 文字选择：拖拽框选，Ctrl+C 复制，右键菜单
        self.canvas.bind("<ButtonPress-1>", self._sel_press)
        self.canvas.bind("<B1-Motion>", self._sel_motion)
        self.canvas.bind("<ButtonRelease-1>", self._sel_release)
        self.canvas.bind("<Control-c>", lambda *_: self._copy_selection())
        self.canvas.bind("<Button-3>", self._sel_menu)
        for key in ("<Left>", "<Prior>"):
            self.canvas.bind(key, lambda *_: self.goto(self.cur - 1))
        for key in ("<Right>", "<Next>"):
            self.canvas.bind(key, lambda *_: self.goto(self.cur + 1))
        self.canvas.bind("<Home>", lambda *_: self.goto(0))
        self.canvas.bind("<End>", lambda *_: self.goto(self.total - 1))
        self.root.bind("<Control-o>", lambda *_: self.open_file())
        self.root.bind("<Control-p>", lambda *_: self.open_print())
        self.root.bind("<Control-f>", lambda *_: self.toggle_search(True))
        # 画笔模式下 Esc 先退出画笔，否则才关搜索栏
        self.root.bind("<Escape>", self._on_escape)
        self.root.bind("<Control-s>", lambda *_: self.save_inplace())
        self.root.bind("<Control-z>", lambda *_: self.ink_undo())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self._init_dnd()
        self.root.after(80, self._draw_hint)
        # 欢迎弹窗：每次启动都展示「关于」（声明 + 联系方式 + 打赏码），
        # 5 秒后自动关闭；延迟 500ms 等主窗口完成布局，居中计算才准。
        if show_about:
            self.root.after(500, lambda: self.open_about(auto_close=5))

    # ================= 最近文件 =================
    def _rebuild_recent_menu(self):
        m = tk.Menu(self.recent_btn, tearoff=0)
        style_menu(m)
        recent = load_recent()
        if recent:
            for p in recent:
                m.add_command(label=p, command=lambda pp=p: self.open_file(pp))
            m.add_separator()
            m.add_command(label="清除记录", command=self._clear_recent)
        else:
            m.add_command(label="（暂无记录）", state="disabled")
        self.recent_btn.config(menu=m)

    def _clear_recent(self):
        try:
            if os.path.exists(CONFIG_PATH):
                os.remove(CONFIG_PATH)
        except Exception:
            pass
        self._rebuild_recent_menu()

    # ================= 打开与关闭 =================
    def open_file(self, path=None, page=None):
        if self.editor and not self._confirm_leave():
            return
        was_empty = self.doc is None
        if not path:
            path = filedialog.askopenfilename(
                title="打开 PDF 或图片", filetypes=[
                    ("PDF 与图片", "*.pdf *.png *.jpg *.jpeg"),
                    ("PDF 文件", "*.pdf"), ("图片文件", "*.png *.jpg *.jpeg")])
            if not path:
                return
        # 图片直接进入「图片转 PDF」
        if path.lower().endswith((".png", ".jpg", ".jpeg")):
            dlg = ImagesToPdfDialog(self)
            dlg.add_paths([path])
            return
        try:
            doc = fitz.open(path)
        except Exception:
            messagebox.showwarning("无法读取",
                                   f"{os.path.basename(path)}：文件损坏或不是有效的 PDF",
                                   parent=self.root)
            return
        if doc.needs_pass:
            for _ in range(3):
                pw = simpledialog.askstring(
                    "加密文件", "该 PDF 已加密，请输入打开密码：", show="*",
                    parent=self.root)
                if pw is None:
                    doc.close()
                    return
                if doc.authenticate(pw):
                    self.open_pw = pw
                    break
            else:
                messagebox.showerror("密码错误", "连续 3 次密码错误，已取消打开。",
                                     parent=self.root)
                doc.close()
                return
        else:
            self.open_pw = ""
        # 覆盖保存会把 self.doc 关掉再写回，这里必须能安全接住「已关闭的 doc」
        _old = self.doc
        self.doc = None
        if _old is not None:
            try:
                _old.close()
            except Exception:
                pass
        self.doc = doc
        self.src_path = path
        self.total = doc.page_count
        self.cur = 0 if page is None else max(0, min(int(page), self.total - 1))
        self.fit = True
        self._hit_page, self._hits, self._hit_idx = -1, [], -1
        add_recent(path)
        self._rebuild_recent_menu()
        self.st_path.config(text=path)
        self._enable_tools(True)
        self._start_thumbs()
        if was_empty:
            try:
                self.side_nb.select(0)   # 首次打开回到「页面」页签
            except Exception:
                pass
        self.render()

    # ================= 拖放打开 / 拖入合并 =================
    IMG_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".gif")

    def _init_dnd(self):
        if not HAS_DND:
            return
        for w in (self.root, self.canvas, self.thumb_canvas, self.side_nb):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _on_drop(self, event):
        paths = self._parse_drop_paths(event.data)
        if paths:
            self.dispatch_paths(paths)
        return "break"

    def _parse_drop_paths(self, data):
        """解析拖放数据为存在的文件路径列表（保序、去重）。
        tkdnd 下发的是 Tcl 列表，Windows 路径里的反斜杠是转义符（\\\\1 会被读成八进制），
        所以：① 先用 Tcl splitlist（正常路径）；② 再用字面分词兜底，把反斜杠、
        花括号、引号都当普通字符，专治下发方漏转义导致的路径残缺。
        实测（2026-10-17）：只走 splitlist 时，漏转义的 `C:\\a\\b c.pdf C:\\a\\d.pdf`
        第二个元素被吃成 `C:ad.pdf` 而静默丢失；加字面兜底后 4 种写法全通。"""
        raw = []
        try:
            raw = [str(x) for x in self.root.tk.splitlist(data)]
        except Exception:
            raw = [str(data)]

        text = str(data)
        # 字面分词：{...} / "..." / 非空白连续串；不做任何转义还原
        literal = []
        for m in re.finditer(r"\{([^}]*)\}|\"([^\"]*)\"|(\S+)", text):
            tok = m.group(1) or m.group(2) or m.group(3) or ""
            tok = tok.strip()
            if tok:
                literal.append(tok)

        def variants(s):
            s = str(s).strip().strip("{}\"").strip()
            if not s:
                return []
            cands = [s]
            if len(s) > 1:
                cands.append(s.replace("\\\\", "\\"))      # 双重反斜杠还原
            if os.name == "nt":
                cands.append(s.replace("/", "\\"))
                cands.append(s.replace("\\\\", "\\").replace("/", "\\"))
            else:
                cands.append(s.replace("\\", "/"))
            return cands

        out, seen = [], set()

        def add(p):
            p = os.path.normpath(p)
            k = os.path.normcase(p)
            if k not in seen:
                seen.add(k)
                out.append(p)

        for s in raw:
            for c in variants(s):
                if os.path.exists(c):
                    add(c)
                    break
        # 字面兜底：贪心最长匹配（3/2/1 个词拼一段），这样未加花括号的
        # "C:\a\b c.pdf" 也能拼回来，而不是被拆成两个不存在的碎片。
        i, n = 0, len(literal)
        while i < n:
            hit = None
            for span in (3, 2, 1):
                if i + span > n:
                    continue
                for c in variants(" ".join(literal[i:i + span])):
                    if os.path.exists(c):
                        hit = (c, span)
                        break
                if hit:
                    break
            if hit:
                add(hit[0])
                i += hit[1]
            else:
                i += 1
        return out

    def dispatch_paths(self, paths):
        """拖入分流：1 个 PDF=打开；多个 PDF=合并；图片=转 PDF；图文混合=先转图片再合并。"""
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        imgs = [p for p in paths if p.lower().endswith(self.IMG_EXTS)]
        if not pdfs and not imgs:
            messagebox.showwarning("无法处理",
                                   "只支持拖入 PDF 或图片（PNG / JPG / BMP / GIF）。",
                                   parent=self.root)
            return
        if self.editor and not self._confirm_leave():
            return
        if pdfs and not imgs:
            if len(pdfs) == 1:
                self.open_file(pdfs[0])
            else:
                MergeDialog(self, files=pdfs)
            return
        if imgs and not pdfs:
            dlg = ImagesToPdfDialog(self)
            dlg.add_paths(imgs)
            return
        # 图文混合：先把图片合成一份临时 PDF，再连同 PDF 一起进入合并
        tmp = os.path.join(tempfile.gettempdir(), f"拖入图片_{os.getpid()}.pdf")
        try:
            do_images_to_pdf(imgs, tmp)
        except Exception as e:
            messagebox.showwarning("图片转换失败", f"{e}\n\n可先单独拖入图片处理。",
                                   parent=self.root)
            return
        MergeDialog(self, files=pdfs + [tmp])

    def _enable_tools(self, on):
        state = "normal" if on else "disabled"
        self.print_btn.config(state=state)
        for b in self.tool_btns.values():
            b.config(state=state)
        self._sync_task_buttons(on)

    def _sync_task_buttons(self, on):
        """没有打开文档时，把依赖文档的任务按钮置灰，避免点了只弹警告。"""
        for b, no_doc in getattr(self, "task_btns", []):
            usable = on or no_doc
            try:
                b.config(state="normal" if usable else "disabled",
                         fg=Theme.TEXT if usable else "#A9B3C6",
                         bg="#F5F8FD" if usable else "#F2F4F8",
                         activebackground="#EBF2FC" if usable else "#F2F4F8",
                         cursor="hand2" if usable else "arrow")
            except Exception:
                pass

    def open_mcp(self):
        McpDialog(self)

    def open_about(self, auto_close=0):
        AboutDialog(self, auto_close=auto_close)

    def on_close(self):
        if self.ink_mode and self.ink_strokes:
            r = messagebox.askyesnocancel(
                "未保存的笔迹",
                f"还有 {len(self.ink_strokes)} 处画笔标记没有保存。\n\n"
                "是 = 覆盖保存到原文件后退出；否 = 直接丢弃退出；取消 = 不退出。",
                parent=self.root)
            if r is None:
                return
            if r:
                self._close_after_save = True
                self.save_inplace()
                return          # 保存成功后由 poll 里再次关闭
        self._thumb_gen += 1
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass
        # MCP 服务跑在后台线程，退出前显式停掉，别把端口留在那儿
        if self.mcp_srv is not None:
            try:
                import pdf_mcp as _m
                _m.CONFIRM_HOOK = None
                _m.CODE_ONLY = False
                _m.stop_http(self.mcp_srv)
                self.mcp_srv, self.mcp_port, self.mcp_kind = None, 0, ""
            except Exception:
                pass
        self.root.destroy()

    # ================= 翻页与缩放 =================
    def jump_from_entry(self):
        s = self.page_var.get().strip()
        if s.isdigit():
            self.goto(int(s) - 1)
        else:
            self.page_var.set(str(self.cur + 1))

    def goto(self, i):
        if not self.doc:
            return
        if self.editor and not self._confirm_leave():
            return
        i = max(0, min(i, self.total - 1))
        if i != self.cur:
            self.cur = i
            self.render()

    def set_zoom(self, z):
        if not self.doc:
            return
        self.zoom = max(0.2, min(5.0, z))
        self.fit = False
        self.render()

    def set_fit(self):
        if not self.doc:
            return
        self.fit = True
        self.render()

    def rotate_view(self):
        if not self._need_doc():
            return
        for i in range(self.total):
            self.doc[i].set_rotation((self.doc[i].rotation + 90) % 360)
        self.render()

    def _ctrl_wheel(self, e):
        self.set_zoom(self.zoom * (1.25 if e.delta > 0 else 1 / 1.25))

    def on_resize(self, *_):
        if not self.doc:
            self._draw_hint()
            return
        if self.fit:
            if self._render_job:
                self.root.after_cancel(self._render_job)
            self._render_job = self.root.after(120, self.render)

    def _draw_hint(self):
        if self.doc:
            return
        self.canvas.delete("hint")
        w = max(self.canvas.winfo_width(), 200)
        h = max(self.canvas.winfo_height(), 200)
        self.canvas.create_text(
            w / 2, h / 2, tags="hint", fill="#98A3B6", font=("", 12),
            justify="center",
            text=("把 PDF 或图片直接拖进窗口（多个文件会自动进入合并）\n"
                  if HAS_DND else "") +
                 "或点击左上角「打开 PDF…」选择文件\n"
                 "左侧「功能」页签可按任务找工具")

    def render(self):
        # 存盘子线程正在操作同一个 doc，MuPDF 不是线程安全的，这里让路
        if not self.doc or getattr(self, "_saving", False):
            return
        page = self.doc[self.cur]
        r = page.rect
        if self.fit:
            cw = max(self.canvas.winfo_width() - 24, 100)
            ch = max(self.canvas.winfo_height() - 24, 100)
            self.zoom = min(cw / r.width, ch / r.height)
            ztxt = "适应窗口"
        else:
            ztxt = f"{int(round(self.zoom * 100))}%"
        self.zoom_lbl.config(text=ztxt)
        self.st_zoom.config(text=ztxt)
        png, w, h = page_pixmap_png(self.doc, self.cur, self.zoom)
        # 渲染位图按整数像素对齐，实际显示比例与名义 zoom 有微差；
        # 记录真实比例，供叠加层预览用，保证与页面位图逐像素对齐。
        self._scale_x = w / r.width if r.width else self.zoom
        self._scale_y = h / r.height if r.height else self.zoom
        self._photo = tk.PhotoImage(data=png)
        self.canvas.delete("all")
        self.canvas.create_image(12, 12, anchor="nw", image=self._photo)
        self.canvas.config(scrollregion=(0, 0, w + 24, h + 24))
        self.page_var.set(str(self.cur + 1))
        self.total_lbl.config(text=f"/ {self.total}")
        self.st_page.config(text=f"第 {self.cur + 1} / 共 {self.total} 页")
        self._sel_text = ""
        self._cache_words(page)
        self._draw_search_hits()
        self._redraw_ink()
        if self.editor:
            self.editor.redraw()
        else:
            self._highlight_thumb(self.cur)

    # ================= 页面缩略图 =================
    def _start_thumbs(self):
        self._thumb_gen += 1
        gen = self._thumb_gen
        self._thumb_queue = queue.Queue()
        for f, _nl in self._thumb_widgets:
            f.destroy()
        self._thumb_widgets = []
        self.thumb_canvas.yview_moveto(0)
        pw = self.open_pw or None
        threading.Thread(target=self._thumb_worker,
                         args=(self.src_path, gen, pw), daemon=True).start()
        self.root.after(60, lambda: self._poll_thumbs(gen))

    def _thumb_worker(self, path, gen, pw):
        try:
            doc = fitz.open(path)
            if doc.needs_pass:
                doc.authenticate(pw or "")
        except Exception:
            self._thumb_queue.put((gen, None, None))
            return
        for i in range(doc.page_count):
            if gen != self._thumb_gen:
                break
            z = 120.0 / doc[i].rect.width
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            self._thumb_queue.put((gen, i, pix.tobytes("png")))
        doc.close()
        self._thumb_queue.put((gen, None, None))

    def _poll_thumbs(self, gen):
        done = False
        try:
            while True:
                g, i, png = self._thumb_queue.get_nowait()
                if g != gen:
                    continue
                if i is None:
                    done = True
                else:
                    self._add_thumb(i, png)
        except queue.Empty:
            pass
        if not done and gen == self._thumb_gen:
            self.root.after(80, lambda: self._poll_thumbs(gen))

    def _add_thumb(self, i, png):
        if not self.root.winfo_exists():
            return
        photo = tk.PhotoImage(data=png)
        frame = tk.Frame(self.thumb_inner, bg=Theme.CARD, cursor="hand2",
                         highlightthickness=1, highlightbackground="#DFE6F1")
        il = tk.Label(frame, image=photo, bg="white", bd=0, cursor="hand2")
        il.image = photo
        il.pack(padx=3, pady=(3, 0))
        nl = tk.Label(frame, text=str(i + 1), bg=Theme.CARD, fg=Theme.MUTED,
                      font=("", 9), cursor="hand2")
        nl.pack(pady=2)
        frame.pack(pady=3, padx=7, fill="x")
        wheel = lambda e: self.thumb_canvas.yview_scroll(int(-e.delta / 120), "units")
        for wgt in (frame, il, nl):
            wgt.bind("<Button-1>", lambda _e, i=i: self.goto(i))
            wgt.bind("<MouseWheel>", wheel)
        self._thumb_widgets.append((frame, nl))
        if i == self.cur and not self.editor:
            self._highlight_thumb(i)

    def _highlight_thumb(self, i):
        for j, (f, nl) in enumerate(self._thumb_widgets):
            cur = (j == i)
            f.configure(bg="#EFF5FF" if cur else Theme.CARD,
                        highlightthickness=2 if cur else 1,
                        highlightbackground=Theme.PRIMARY if cur else "#DFE6F1")
            nl.configure(bg="#EFF5FF" if cur else Theme.CARD,
                         fg=Theme.PRIMARY if cur else Theme.MUTED)
        self._thumb_scroll_to(i)

    def _thumb_scroll_to(self, i):
        if i >= len(self._thumb_widgets):
            return
        self.root.update_idletasks()
        f = self._thumb_widgets[i][0]
        h = self.thumb_inner.winfo_height()
        if h > 1:
            self.thumb_canvas.yview_moveto(max(0.0, (f.winfo_y() - 30) / h))

    # ================= 任务式功能导航（左侧「功能」页签） =================
    def _build_task_tab(self):
        """按「用户想干什么」分组，而不是按工具分类。轻度用户不必再翻二级菜单。"""
        self.task_btns = []          # [(按钮, 无需文档也能用), ...]
        tab = ttk.Frame(self.side_nb, style="Card.TFrame")
        self.side_nb.add(tab, text="功能")
        cv = tk.Canvas(tab, bg=Theme.CARD, highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(tab, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        inner = ttk.Frame(cv, style="Card.TFrame")
        win = cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda *_: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        wheel = lambda e: cv.yview_scroll(int(-e.delta / 120), "units")
        cv.bind("<MouseWheel>", wheel)
        inner.bind("<MouseWheel>", wheel)
        try:
            btn_font = ttk.Style().lookup("TButton", "font") or ""
        except Exception:
            btn_font = ""

        no_doc = ("打开 PDF…", "图片转 PDF…", "合并多个 PDF…", "批量处理…",
                  "Office 转 PDF…")

        def group(title, items):
            _gi = getattr(group, "_i", 0)
            group._i = _gi + 1
            _accent = _GROUP_ACCENTS[_gi % len(_GROUP_ACCENTS)]
            hrow = tk.Frame(inner, bg=Theme.CARD)
            hrow.pack(fill="x", padx=6, pady=(10, 3))
            tk.Frame(hrow, bg=_accent, width=3, height=15).pack(
                side="left", padx=(3, 7))
            _hf = _UI_FONTS.get("bold")
            _hl = tk.Label(hrow, text=title, bg=Theme.CARD, fg=_accent,
                           anchor="w")
            if _hf:
                _hl.configure(font=_hf)
            _hl.pack(side="left")
            for label, cmd in items:
                # 用 tk.Button 才能左对齐（ttk 按钮不支持 -anchor），做成「列表项」观感
                b = tk.Button(inner, text=label, command=cmd, anchor="w",
                              font=btn_font, bg="#F5F8FD", fg=Theme.TEXT,
                              activebackground="#EBF2FC",
                              activeforeground=_accent,
                              relief="flat", bd=0, highlightthickness=1,
                              highlightbackground="#DCE5F2",
                              highlightcolor=_accent,
                              padx=9, pady=4, cursor="hand2")
                b.pack(fill="x", padx=5, pady=1)
                self.task_btns.append((b, label in no_doc))

        group("把资料变成 PDF", [
            ("打开 PDF…", self.open_file),
            ("图片转 PDF…", self.open_img2pdf),
            ("Office 转 PDF…", self.open_office2pdf),
            ("合并多个 PDF…", self.open_merge),
            ("拆分…", self.open_split),
            ("提取页面…", self.open_extract),
        ])
        group("改页面", [
            ("页面管理（重排）…", self.open_reorder),
            ("删除页面…", self.open_delete),
            ("旋转页面…", self.open_rotate),
            ("裁剪页面 / 去黑边…", self.open_crop),
            ("添加页码…", self.open_pagenum),
            ("页眉页脚…", self.open_hf),
        ])
        group("加内容", [
            ("插入文字", lambda: self.start_editor("text")),
            ("插入图片", lambda: self.start_editor("image")),
            ("手写签名", self.open_sign),
            ("水印…", self.open_watermark),
            ("标注 / 批注…", self.open_annotate),
            ("涂黑遮盖…", self.open_redact),
        ])
        group("发出去", [
            ("压缩体积…", self.open_compress),
            ("压缩到指定大小…", self.open_compress_target),
            ("加密保护…", self.open_encrypt),
            ("解密 / 去密码…", self.open_decrypt),
            ("打印…", self.open_print),
        ])
        group("转换与导出", [
            ("PDF 转 Word…", self.open_pdf2word),
            ("PDF 转 Excel…", self.open_pdf2excel),
            ("PDF 转图片…", self.open_pdf2img),
            ("提取文档内图片…", self.open_extract_img),
            ("表单填写…", self.open_form),
            ("书签 / 目录…", self.open_toc),
            ("批量处理…", self.open_batch),
        ])
        self._sync_task_buttons(False)

    def _build_tool_panel(self, mid):
        """右侧工具面板：分组可折叠 + 常驻滚动条 + 收起后留竖向导轨。

        四条都对应实际使用反馈，别改回去：
        1. 六组 26 个按钮，1320×820 一屏放不下 → 每组标题可点折叠，
           顶部还有「折叠/展开」一键切全部。
        2. 滚动条必须**先 pack**：canvas 带 expand=True，先 pack 会吃光宽度，
           后 pack 的滚动条只剩 0px —— 之前「明明放不下却没有滚动条」就是这原因。
        3. 收起后右侧留一条 24px 竖向导轨（工/具/面/板 竖排），点一下回来，
           不必去顶栏找按钮。
        4. 按钮注册进 task_btns，随文档打开状态自动置灰（_sync_task_buttons）。
        """
        self.tool_panel_on = True
        self._panel_groups = []

        panel = ttk.Frame(mid, style="Card.TFrame", width=196)
        panel.pack(side="right", fill="y", padx=(0, 10), pady=(0, 8))
        panel.pack_propagate(False)
        self.tool_panel = panel

        head = ttk.Frame(panel, style="Card.TFrame")
        head.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(head, text="全部工具", style="Title.TLabel").pack(side="left")
        ttk.Button(head, text="收起 ›", width=6, style="PanelGhost.TButton",
                   command=self.toggle_tool_panel).pack(side="right")
        self.panel_fold_btn = ttk.Button(
            head, text="折叠", width=4, style="PanelGhost.TButton",
            command=self.toggle_all_groups)
        self.panel_fold_btn.pack(side="right", padx=(0, 2))

        ttk.Separator(panel, orient="horizontal").pack(fill="x", padx=8)

        # 顺序铁律：滚动条先 pack，canvas 后 pack 才能 share 到宽度
        sb = ttk.Scrollbar(panel, orient="vertical",
                           style="Panel.Vertical.TScrollbar")
        sb.pack(side="right", fill="y", pady=(4, 6), padx=(0, 2))
        self.panel_sb = sb
        cv = tk.Canvas(panel, bg=Theme.CARD, highlightthickness=0, bd=0,
                       yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True, pady=(4, 6))
        sb.configure(command=cv.yview)
        self._panel_cv = cv
        inner = ttk.Frame(cv, style="Card.TFrame")
        win = cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda *_: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        wheel = lambda e: cv.yview_scroll(int(-e.delta / 120), "units")
        cv.bind("<MouseWheel>", wheel)
        inner.bind("<MouseWheel>", wheel)
        try:
            btn_font = ttk.Style().lookup("TButton", "font") or ""
        except Exception:
            btn_font = ""
        _hf = _UI_FONTS.get("bold")

        def group(title, items):
            _gi = getattr(group, "_i", 0)
            group._i = _gi + 1
            _accent = _GROUP_ACCENTS[_gi % len(_GROUP_ACCENTS)]
            g = tk.Frame(inner, bg=Theme.CARD)
            g.pack(fill="x")
            hrow = tk.Frame(g, bg=Theme.CARD, cursor="hand2")
            hrow.pack(fill="x", padx=6, pady=(9, 2))
            bar = tk.Frame(hrow, bg=_accent, width=3, height=15)
            bar.pack(side="left", padx=(3, 6))
            chev = tk.Label(hrow, text="▾", bg=Theme.CARD, fg=_accent,
                            cursor="hand2", width=2)
            if _hf:
                chev.configure(font=_hf)
            chev.pack(side="left")
            tl = tk.Label(hrow, text=title, bg=Theme.CARD, fg=_accent,
                          anchor="w", cursor="hand2")
            if _hf:
                tl.configure(font=_hf)
            tl.pack(side="left")
            cnt = tk.Label(hrow, text=str(len(items)), bg=Theme.CARD,
                           fg=Theme.MUTED, cursor="hand2")
            cnt.pack(side="right", padx=(0, 5))

            body = tk.Frame(g, bg=Theme.CARD)
            body.pack(fill="x")
            for label, cmd in items:
                b = tk.Button(body, text=label, command=cmd, anchor="w",
                              font=btn_font, bg="#F5F8FD", fg=Theme.TEXT,
                              activebackground="#EBF2FC",
                              activeforeground=_accent,
                              relief="flat", bd=0, highlightthickness=1,
                              highlightbackground="#DCE5F2",
                              highlightcolor=_accent,
                              padx=9, pady=4, cursor="hand2")
                b.pack(fill="x", padx=5, pady=1)
                self.task_btns.append((b, label in _PANEL_NO_DOC))
            rec = {"title": title, "body": body, "chev": chev, "folded": False}
            self._panel_groups.append(rec)
            for w in (hrow, bar, chev, tl, cnt):
                w.bind("<Button-1>", lambda _e, r=rec: self._toggle_group(r))
            hrow.bind("<Enter>", lambda _e, c=chev: c.configure(fg=_accent))
            return rec

        group("页面整理", [
            ("页面提取…", self.open_extract),
            ("拆分…", self.open_split),
            ("页面管理（重排）…", self.open_reorder),
            ("删除页面…", self.open_delete),
            ("旋转页面…", self.open_rotate),
            ("裁剪页面 / 去黑边…", self.open_crop),
        ])
        group("编辑与标注", [
            ("添加页码…", self.open_pagenum),
            ("添加水印…", self.open_watermark),
            ("页眉页脚…", self.open_hf),
            ("标注/批注…", self.open_annotate),
            ("画笔标记…", self.toggle_ink),
            ("涂黑遮盖…", self.open_redact),
        ])
        group("转换与导出", [
            ("PDF 转 Word…", self.open_pdf2word),
            ("PDF 转 Excel…", self.open_pdf2excel),
            ("PDF 转图片…", self.open_pdf2img),
            ("提取文档内图片…", self.open_extract_img),
        ])
        group("安全与优化", [
            ("压缩体积…", self.open_compress),
            ("压缩到指定大小…", self.open_compress_target),
            ("加密保护…", self.open_encrypt),
            ("解密/去密码…", self.open_decrypt),
        ])
        group("文档增强", [
            ("书签/目录管理…", self.open_toc),
            ("表单填写…", self.open_form),
        ])
        group("文件工具", [
            ("合并…", self.open_merge),
            ("图片转 PDF…", self.open_img2pdf),
            ("Office 转 PDF…", self.open_office2pdf),
            ("保存（覆盖原文件）…", self.save_inplace),
            ("另存为新文件…", self.save_as),
            ("批量处理…", self.open_batch),
        ])
        # Tk 的 MouseWheel 只沿 bindtags（widget→class→toplevel→all）传递，
        # 不会往上冒泡到父 Frame；面板里几乎全是 tk.Button，所以必须逐个
        # 子控件绑定，否则鼠标停在按钮上滚轮完全没反应（看起来像面板死了）。
        def _bind_wheel_deep(widget):
            try:
                widget.bind("<MouseWheel>", wheel)
            except Exception:
                pass
            for _ch in widget.winfo_children():
                _bind_wheel_deep(_ch)

        _bind_wheel_deep(inner)
        self._panel_wheel_bound = True

        # 收起后右侧留一条竖向导轨，点一下展开（不用去顶栏找按钮）
        rail = tk.Frame(mid, bg=Theme.PRIMARY, width=24, cursor="hand2")
        rail.pack_propagate(False)
        rl = tk.Label(rail, text="工\n具\n面\n板", bg=Theme.PRIMARY,
                      fg="#FFFFFF", justify="center", cursor="hand2", pady=8)
        if _hf:
            rl.configure(font=_hf)
        rl.pack(expand=True, fill="both")

        def _rail_hover(on):
            bg = Theme.PRIMARY_H if on else Theme.PRIMARY
            rail.configure(bg=bg)
            rl.configure(bg=bg)

        for _w in (rail, rl):
            _w.bind("<Button-1>", lambda *_: self.toggle_tool_panel())
            _w.bind("<Enter>", lambda *_: _rail_hover(True))
            _w.bind("<Leave>", lambda *_: _rail_hover(False))
        self.panel_rail = rail

        self._sync_fold_btn()
        self._sync_task_buttons(self.doc is not None)

    # ---------- 分组折叠 ----------
    def _toggle_group(self, rec, fold=None):
        """折叠/展开一个分组。fold=None 表示取反。"""
        rec["folded"] = (not rec["folded"]) if fold is None else bool(fold)
        if rec["folded"]:
            rec["body"].pack_forget()
            rec["chev"].config(text="▸")
        else:
            rec["body"].pack(fill="x")
            rec["chev"].config(text="▾")
        try:
            self._panel_cv.update_idletasks()
            self._panel_cv.configure(scrollregion=self._panel_cv.bbox("all"))
        except Exception:
            pass
        self._sync_fold_btn()

    def toggle_all_groups(self):
        """一键折叠 / 展开全部分组。"""
        any_open = any(not r["folded"] for r in self._panel_groups)
        for r in self._panel_groups:
            self._toggle_group(r, fold=any_open)
        self._sync_fold_btn()

    def _sync_fold_btn(self):
        any_open = any(not r["folded"] for r in self._panel_groups)
        try:
            self.panel_fold_btn.config(text="折叠" if any_open else "展开")
        except Exception:
            pass

    def toggle_tool_panel(self):
        self.tool_panel_on = not self.tool_panel_on
        if self.tool_panel_on:
            try:
                self.panel_rail.pack_forget()
            except Exception:
                pass
            # before=阅读区：保证重新显示时仍在阅读区右侧占位，不被 expand 挤掉
            self.tool_panel.pack(side="right", fill="y", padx=(0, 10),
                                 pady=(0, 8), before=self.viewer)
            self.panel_btn.config(text="收起工具面板")
        else:
            self.tool_panel.pack_forget()
            # 收起后立刻留一条竖向导轨，视觉上「还能点回来」
            try:
                self.panel_rail.pack(side="right", fill="y", padx=(0, 6),
                                     pady=(0, 8), before=self.viewer)
            except Exception:
                pass
            self.panel_btn.config(text="展开工具面板")

    # ================= 搜索 =================
    def toggle_search(self, show):
        if show:
            self.search_bar.place(relx=1.0, y=6, anchor="ne", x=-12)
            for w in self.search_bar.winfo_children():
                if isinstance(w, ttk.Entry):
                    w.focus_set()
                    break
        else:
            self.search_bar.place_forget()
            self._hit_page, self._hits, self._hit_idx = -1, [], -1
            self.canvas.delete("srch")

    def search_next(self, back=False):
        text = self.search_var.get().strip()
        if not text or not self.doc:
            return
        # 先尝试同页内移动
        if self._hit_page == self.cur and self._hits:
            self._hit_idx = (self._hit_idx + (-1 if back else 1)) % len(self._hits)
            self._draw_search_hits()
            self._scroll_to_hit()
            return
        # 换页搜索：从当前页的下一/上一页起，扫到文件头尾后回绕到另一端
        if back:
            order = (list(range(self.cur - 1, -1, -1))
                     + list(range(self.total - 1, self.cur, -1)))
        else:
            order = (list(range(self.cur + 1, self.total))
                     + list(range(0, self.cur)))
        for p in order:
            hits = self.doc[p].search_for(text)
            if hits:
                self._hit_page, self._hits = p, hits
                self._hit_idx = len(hits) - 1 if back else 0
                self.goto(p)
                if self.cur == p:  # 编辑器拦截等导致未跳页时手动渲染
                    self._draw_search_hits()
                    self._scroll_to_hit()
                return
        self.set_status_bar(f"全文未找到「{text}」")

    def _draw_search_hits(self):
        self.canvas.delete("srch")
        if self._hit_page != self.cur or not self._hits:
            return
        page = self.doc[self.cur]
        zx = self._scale_x or self.zoom
        zy = self._scale_y or self.zoom
        for r in self._hits:
            d = rect_display(page, r)   # 旋转页必须映射到显示坐标系
            self.canvas.create_rectangle(
                12 + d.x0 * zx, 12 + d.y0 * zy, 12 + d.x1 * zx, 12 + d.y1 * zy,
                outline="#ff9800", width=2, tags="srch")
        if 0 <= self._hit_idx < len(self._hits):
            d = rect_display(page, self._hits[self._hit_idx])
            self.canvas.create_rectangle(
                12 + d.x0 * zx, 12 + d.y0 * zy, 12 + d.x1 * zx, 12 + d.y1 * zy,
                outline="#f57c00", width=3, tags="srch")
        if self._hits:
            self.set_status_bar(
                f"第 {self.cur + 1} 页找到 {len(self._hits)} 处"
                f"（当前第 {self._hit_idx + 1} 处）")

    def _scroll_to_hit(self):
        if not (0 <= self._hit_idx < len(self._hits)):
            return
        d = rect_display(self.doc[self.cur], self._hits[self._hit_idx])
        zy = self._scale_y or self.zoom
        total_h = self.canvas.bbox("all")
        if total_h:
            self.canvas.yview_moveto(max(0.0, (12 + d.y0 * zy - 60) / (total_h[3] - total_h[1])))

    # ================= 文字选择与复制 =================
    def _cache_words(self, page):
        """缓存本页所有词及其「显示坐标系」矩形，供框选命中。"""
        _dw, _dh, rot, mbw, mbh = page_geometry(page)
        out = []
        try:
            raw = page.get_text("words")
        except Exception:
            raw = []
        for wd in raw:
            x0, y0, x1, y1 = wd[:4]
            if rot:
                pts = (_rot_map(rot, mbw, mbh, x0, y0), _rot_map(rot, mbw, mbh, x1, y0),
                       _rot_map(rot, mbw, mbh, x0, y1), _rot_map(rot, mbw, mbh, x1, y1))
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                r = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
            else:
                r = fitz.Rect(x0, y0, x1, y1)
            out.append((r, wd[4]))
        self._words = out

    def _canvas_to_page(self, x, y):
        zx = self._scale_x or self.zoom or 1.0
        zy = self._scale_y or self.zoom or 1.0
        return (x - 12) / zx, (y - 12) / zy

    def _sel_press(self, e):
        if self.ink_mode:
            self._ink_press(e)
            return
        if self.editor or not self.doc:
            return
        self._sel0 = (e.x, e.y)
        self.canvas.delete("selrubber")

    def _sel_motion(self, e):
        if self.ink_mode:
            self._ink_motion(e)
            return
        if self._sel0 is None or self.editor:
            return
        self.canvas.delete("selrubber")
        self.canvas.create_rectangle(self._sel0[0], self._sel0[1], e.x, e.y,
                                     outline=Theme.PRIMARY, width=1, dash=(3, 2),
                                     tags="selrubber")

    def _sel_release(self, e):
        if self.ink_mode:
            self._ink_release(e)
            return
        if self._sel0 is None or self.editor:
            return
        x0, y0 = self._sel0
        self._sel0 = None
        self.canvas.delete("selrubber")
        if abs(e.x - x0) < 3 and abs(e.y - y0) < 3:
            self.clear_selection()
            return
        px0, py0 = self._canvas_to_page(min(x0, e.x), min(y0, e.y))
        px1, py1 = self._canvas_to_page(max(x0, e.x), max(y0, e.y))
        self._apply_selection(fitz.Rect(px0, py0, px1, py1))

    # ================= 画笔（自由手绘标记） =================
    def toggle_ink(self, on=None):
        """进入/退出画笔模式。笔迹在内存里，Ctrl+S 覆盖保存时才写进文档。"""
        if not self._need_doc():
            return
        if on is None:
            on = not self.ink_mode
        if on and self.editor and not self._confirm_leave():
            return
        self.ink_mode = bool(on)
        if self.ink_mode:
            self.clear_selection()
            self.toggle_search(False)
            self.ink_bar.lift()
            self.ink_bar.place(relx=0.0, x=14, y=8, anchor="nw")
            self.canvas.config(cursor="crosshair")
            self.set_status_bar("画笔模式：在页面上按住拖动即可做记号；"
                                "Ctrl+Z 撤销上一笔，Ctrl+S 覆盖保存，Esc 退出")
        else:
            self._ink_drawing = False
            self.canvas.delete("inkcur")
            self.ink_bar.place_forget()
            self.canvas.config(cursor="")
            if self._ink_dropped():
                self.set_status_bar("已退出画笔（笔迹未保存，Ctrl+S 可写回原文件）")
        self._sync_ink_bar()

    def _ink_dropped(self):
        return bool(self.ink_strokes)

    def _pick_ink_color(self, name, hexv):
        self.ink_color = hexv
        self._sync_ink_bar()
        self.set_status_bar(f"画笔颜色：{name}")

    def _pick_ink_width(self, pt):
        self.ink_width = pt
        self._sync_ink_bar()
        self.set_status_bar(f"画笔粗细：{ {'2':'细','4':'中','8':'粗'}.get(str(pt), pt) }")

    def _pick_ink_shape(self, shape):
        self.ink_shape = shape
        self._sync_ink_bar()
        self.set_status_bar("画笔形状：%s" % INK_SHAPE_LABELS.get(shape, shape))

    def _sync_ink_bar(self):
        """同步工具条选中态：色块加粗边框、形状/粗细按钮换 Accent 样式。"""
        for name, sw in getattr(self, "ink_swatches", {}).items():
            try:
                cur = INK_PRESETS.get(name) == self.ink_color
                sw.configure(relief="solid" if cur else "flat",
                             bd=2 if cur else 1)
            except Exception:
                pass
        for pt, b in getattr(self, "ink_w_btns", {}).items():
            try:
                b.configure(style="Accent.TButton" if pt == self.ink_width
                            else "PanelGhost.TButton")
            except Exception:
                pass
        for sh, b in getattr(self, "ink_shape_btns", {}).items():
            try:
                b.configure(style="Accent.TButton" if sh == self.ink_shape
                            else "PanelGhost.TButton")
            except Exception:
                pass

    # ---- 坐标换算：画布 → 页面（显示）→ 未旋转页面 ----
    def _ink_canvas_to_pdf(self, cx, cy):
        dx, dy = self._canvas_to_page(self.canvas.canvasx(cx),
                                      self.canvas.canvasy(cy))
        page = self.doc[self.cur]
        _dw, _dh, rot, mbw, mbh = page_geometry(page)
        if rot:
            return _rot_unmap(rot, mbw, mbh, dx, dy)
        return (dx, dy)

    def _ink_pdf_to_canvas(self, u, v):
        page = self.doc[self.cur]
        _dw, _dh, rot, mbw, mbh = page_geometry(page)
        if rot:
            dx, dy = _rot_map(rot, mbw, mbh, u, v)
        else:
            dx, dy = u, v
        zx = self._scale_x or self.zoom or 1.0
        zy = self._scale_y or self.zoom or 1.0
        return 12 + dx * zx, 12 + dy * zy

    def _ink_new_item(self, x, y):
        """按当前形状造一个预览图元（拖拽时实时更新 coords）。"""
        col, wd = self.ink_color, self.ink_width
        if self.ink_shape == "rect":
            return self.canvas.create_rectangle(x, y, x, y, outline=col,
                                                width=wd, tags="inkcur")
        if self.ink_shape == "ellipse":
            return self.canvas.create_oval(x, y, x, y, outline=col,
                                           width=wd, tags="inkcur")
        if self.ink_shape == "line":
            # create_line 只认 fill，传 outline 会 TclError
            return self.canvas.create_line(x, y, x, y, fill=col, width=wd,
                                           capstyle=tk.ROUND, tags="inkcur")
        return self.canvas.create_line(x, y, x, y, fill=col, width=wd,
                                       capstyle=tk.ROUND, joinstyle=tk.ROUND,
                                       smooth=True, splinesteps=12,
                                       tags="inkcur")

    def _ink_press(self, e):
        if not self.doc:
            return
        self._ink_drawing = True
        self.canvas.delete("inkcur")
        self._ink_pts = [(e.x, e.y)]
        self._ink_item = self._ink_new_item(e.x, e.y)

    def _ink_motion(self, e):
        if not self._ink_drawing or self._ink_item is None:
            return
        x0, y0 = self._ink_pts[0]
        if self.ink_shape == "free":
            self._ink_pts.append((e.x, e.y))
            flat = []
            for x, y in self._ink_pts:
                flat.extend((x, y))
            self.canvas.coords(self._ink_item, *flat)
        else:
            # 直线/矩形/椭圆只认起点与当前点，拖动全程一个图元
            self.canvas.coords(self._ink_item, x0, y0, e.x, e.y)

    def _ink_release(self, e):
        if not self._ink_drawing:
            return
        self._ink_drawing = False
        if self.ink_shape == "free":
            if self._ink_pts:
                self._ink_pts.append((e.x, e.y))
            pts = [self._ink_canvas_to_pdf(x, y) for x, y in self._ink_pts]
        else:
            x0, y0 = self._ink_pts[0]
            if abs(e.x - x0) < 3 and abs(e.y - y0) < 3:
                # 形状类拖得太小视为误点，不落笔
                self.canvas.delete("inkcur")
                self._ink_item = None
                return
            pts = [self._ink_canvas_to_pdf(x0, y0),
                   self._ink_canvas_to_pdf(e.x, e.y)]
        # 单点也保留（落笔点），apply_ink 会画成小圆点
        self.ink_strokes.append({"page": self.cur, "color": self.ink_color,
                                 "width": self.ink_width,
                                 "shape": self.ink_shape, "points": pts})
        if self._ink_item is not None:
            self.canvas.itemconfig(self._ink_item, tags="ink")
            self._ink_item = None
        self.set_status_bar(
            f"已画 {len(self.ink_strokes)} 笔（"
            f"{INK_SHAPE_LABELS.get(self.ink_shape, self.ink_shape)}"
            f" · Ctrl+Z 撤销 · Ctrl+S 保存）")

    def _redraw_ink(self):
        """按当前页/缩放重画内存笔迹（翻页、缩放后覆盖层要跟着变）。"""
        self.canvas.delete("ink")
        for s in self.ink_strokes:
            if s["page"] != self.cur:
                continue
            kind = s.get("shape", "free")
            flat = []
            for u, v in s["points"]:
                x, y = self._ink_pdf_to_canvas(u, v)
                flat.extend((x, y))
            col = s["color"]
            wd = s["width"]
            if kind == "rect" and len(flat) == 4:
                self.canvas.create_rectangle(*flat, outline=col, width=wd,
                                             tags="ink")
                continue
            if kind == "ellipse" and len(flat) == 4:
                self.canvas.create_oval(*flat, outline=col, width=wd,
                                        tags="ink")
                continue
            if kind == "line" and len(flat) == 4:
                self.canvas.create_line(*flat, fill=col, width=wd, tags="ink")
                continue
            if len(flat) < 4:
                x, y = flat[0], flat[1]
                r = max(2.0, wd * 0.6)
                self.canvas.create_oval(x - r, y - r, x + r, y + r,
                                        fill=col, outline=col, tags="ink")
                continue
            self.canvas.create_line(
                *flat, fill=col, width=wd,
                capstyle=tk.ROUND, joinstyle=tk.ROUND, smooth=True,
                splinesteps=12, tags="ink")

    def ink_undo(self):
        """撤销最后一笔（仅当前页）。"""
        if not self.ink_mode:
            return
        for i in range(len(self.ink_strokes) - 1, -1, -1):
            if self.ink_strokes[i]["page"] == self.cur:
                self.ink_strokes.pop(i)
                break
        self._redraw_ink()
        self.set_status_bar(f"已撤销，剩 {len(self.ink_strokes)} 笔")

    def ink_clear_page(self):
        if not self.ink_mode:
            return
        self.ink_strokes = [s for s in self.ink_strokes
                            if s["page"] != self.cur]
        self._redraw_ink()
        self.set_status_bar("已清空本页笔迹")

    def _ink_core_strokes(self):
        """GUI 内部笔迹 → pdf_core.apply_ink 的入参（0基页号 + 未旋转坐标）。"""
        return [{"page": int(s["page"]),
                 "color": s["color"],
                 "width": float(s["width"]),
                 "shape": s.get("shape", "free"),
                 "points": [[float(x), float(y)] for x, y in s["points"]]}
                for s in self.ink_strokes]

    def save_as(self, path=None):
        """另存为新文件：原文件不动，当前内容（含画笔笔迹）写到新路径。"""
        if not self._need_doc():
            return
        if self.editor and not self._confirm_leave():
            return
        src = self.src_path
        if not path:
            base = os.path.splitext(os.path.basename(src))[0] if src else "文档"
            path = filedialog.asksaveasfilename(
                title="另存为新文件", parent=self.root,
                defaultextension=".pdf",
                initialdir=os.path.dirname(src) if src else "",
                initialfile="%s_已标记.pdf" % base,
                filetypes=[("PDF 文件", "*.pdf")])
        if not path:
            return
        if src and os.path.abspath(path) == os.path.abspath(src):
            messagebox.showwarning(
                "另存不能覆盖原文件",
                "所选路径就是当前文档。要用当前内容覆盖它，请按 Ctrl+S 选「是」。",
                parent=self.root)
            return
        gui_strokes = list(self.ink_strokes)
        core_strokes = self._ink_core_strokes() if gui_strokes else None
        if core_strokes:
            try:
                apply_ink(self.doc, core_strokes)
            except Exception as e:
                messagebox.showerror("写入失败",
                                     f"画笔标记写入文档时出错：{e}",
                                     parent=self.root)
                return
            self.ink_strokes = []
            self.canvas.delete("ink")
        out = path
        self._saving = True
        self.set_status_bar("正在另存…")
        q = queue.Queue()

        def work():
            try:
                self.doc.save(out, garbage=3, deflate=True)
                q.put(("ok", out))
            except Exception as e:  # noqa: BLE001
                q.put(("err", str(e)))

        threading.Thread(target=work, daemon=True).start()
        n = len(gui_strokes)

        def poll():
            if not self.root.winfo_exists():
                return
            try:
                kind, payload = q.get_nowait()
            except queue.Empty:
                self.root.after(80, poll)
                return
            self._saving = False
            if kind == "err":
                self.ink_strokes = gui_strokes
                self._redraw_ink()
                messagebox.showerror("另存失败", payload, parent=self.root)
                self.set_status_bar("另存失败")
                return
            self.render()      # 笔迹已进文档内容，重画一次即可见
            self.set_status_bar(f"已另存 {os.path.basename(out)}"
                                f"（{n} 处笔迹已写入）")
            messagebox.showinfo("已另存", f"已保存到：\n{out}", parent=self.root)

        self.root.after(80, poll)

    def save_inplace(self):
        """覆盖保存：把内存笔迹写进当前文档并写回原路径。

        这是「直接保存」：不再逼用户另存新文件。写回前一定问一次，
        因为原文件会被替换（PDF 改动不可撤销）。
        注意走的是「当前 self.doc」而不是重新开文件：视图旋转等内存态
        改动也要一起写回，否则会出现「旋转了却存不住」的假保存。
        """
        if not self._need_doc():
            return
        if self.editor and not self._confirm_leave():
            return
        if not self.src_path:
            messagebox.showwarning("无法保存", "当前文档没有对应磁盘文件。",
                                   parent=self.root)
            return
        n = len(self.ink_strokes)
        tip = f"共 {n} 处画笔标记" if n else "当前无未保存的画笔标记"
        ans = messagebox.askyesnocancel(
            "保存文档",
            f"要怎么保存？\n{self.src_path}\n\n{tip}\n\n"
            "是 = 覆盖原文件（原文件会被替换，建议先备份）\n"
            "否 = 另存为新文件（原文件不动）",
            parent=self.root)
        if ans is None:      # 取消：什么都不做，笔迹留在内存里
            return
        if ans is False:     # 否：直接进另存为
            self.save_as()
            return
        gui_strokes = list(self.ink_strokes)
        core_strokes = self._ink_core_strokes() if gui_strokes else None
        if core_strokes:
            try:
                # 先写进内存文档：保存后界面内容即文件内容，不会「存了却看不见」
                apply_ink(self.doc, core_strokes)
            except Exception as e:
                messagebox.showerror("写入失败",
                                     f"画笔标记写入文档时出错：{e}",
                                     parent=self.root)
                return
            self.ink_strokes = []
            self.canvas.delete("ink")
        src = self.src_path
        self._thumb_gen += 1      # 让缩略图线程撒手，别占着旧文件句柄
        self._saving = True
        self.set_status_bar("正在保存…")
        _quit_after = getattr(self, "_close_after_save", False)
        self._close_after_save = False
        _keep_page = self.cur
        q = queue.Queue()

        def work():
            try:
                q.put(("ok", save_doc_overwrite(self.doc, src)))
            except Exception as e:  # noqa: BLE001 —— 原样带回主线程提示
                q.put(("err", str(e)))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            if not self.root.winfo_exists():
                return
            try:
                kind, payload = q.get_nowait()
            except queue.Empty:
                self.root.after(80, poll)
                return
            self._saving = False
            if kind == "err":
                # 存盘失败：重新打开原文件，并把笔迹还原成内存覆盖层
                self.open_file(src, page=_keep_page)
                self.ink_strokes = gui_strokes
                self._redraw_ink()
                messagebox.showerror("保存失败", payload, parent=self.root)
                self.set_status_bar("保存失败")
                return
            self.open_file(src, page=_keep_page)
            self.set_status_bar(f"已保存 {os.path.basename(src)}"
                                f"（{n} 处笔迹已写入）")
            if _quit_after and self.root.winfo_exists():
                self.on_close()

        self.root.after(80, poll)

    def _on_escape(self, *_):
        if self.ink_mode:
            self.toggle_ink(False)
            return "break"
        self.toggle_search(False)

    def _apply_selection(self, box):
        """把显示坐标系里的框选矩形落到具体词上：画高亮 + 拼出纯文本。"""
        self.canvas.delete("sel")
        picked = []
        for r, t in self._words:
            inter = fitz.Rect(r) & box
            if inter.is_empty:
                continue
            # 词面积过半被框住才算选中，避免擦边误选
            if inter.get_area() >= r.get_area() * 0.45:
                picked.append((r, t))
        if not picked:
            self._sel_text = ""
            return
        zx = self._scale_x or self.zoom
        zy = self._scale_y or self.zoom
        # 同一行合并成一个框，避免逐词小方块把版面切成碎片
        lines = []
        for r, _t in picked:
            if lines and abs(r.y0 - lines[-1][0]) < max(2.0, (r.y1 - r.y0) * 0.6):
                b = lines[-1]
                b[1], b[2], b[3] = min(b[1], r.x0), max(b[2], r.x1), max(b[3], r.y1)
            else:
                lines.append([r.y0, r.x0, r.x1, r.y1])
        for y0, x0, x1, y1 in lines:
            self.canvas.create_rectangle(12 + x0 * zx, 12 + y0 * zy,
                                         12 + x1 * zx, 12 + y1 * zy,
                                         fill="", outline="#2563EB",
                                         width=1, tags="sel")
        self._sel_text = self._join_words(picked)
        self.set_status_bar(f"已选中 {len(self._sel_text)} 字（Ctrl+C 复制）")

    @staticmethod
    def _join_words(picked):
        if not picked:
            return ""
        out = [picked[0][1]]
        prev = picked[0][0]
        for r, t in picked[1:]:
            h = max(1.0, prev.y1 - prev.y0)
            if abs(r.y0 - prev.y0) < h * 0.6:      # 同一行
                cjk = _is_cjk(out[-1][-1:]) or _is_cjk(t[:1])
                gap = r.x0 - prev.x1
                out.append(t if (cjk and gap < h * 1.5) else " " + t)
            else:
                out.append("\n" + t)
            prev = r
        return "".join(out)

    def _copy_selection(self):
        if not self._sel_text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self._sel_text)
        self.set_status_bar(f"已复制 {len(self._sel_text)} 字到剪贴板")

    def clear_selection(self):
        self.canvas.delete("sel")
        self._sel_text = ""

    def select_all_page(self):
        if not self.doc or not self._words:
            return
        self._apply_selection(fitz.Rect(-1e6, -1e6, 1e6, 1e6))

    def _sel_menu(self, e):
        if self.editor or not self.doc:
            return
        m = tk.Menu(self.root, tearoff=0)
        style_menu(m)
        m.add_command(label="复制选中文本", command=self._copy_selection,
                      state="normal" if self._sel_text else "disabled")
        m.add_command(label="全选本页文本", command=self.select_all_page)
        m.add_separator()
        m.add_command(label="清除选择", command=self.clear_selection)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def set_status_bar(self, text):
        self.st_path.config(text=text)

    # ================= 编辑器 =================
    def _need_doc(self):
        if self.doc:
            return True
        messagebox.showwarning("未打开文件", "请先打开 PDF 文件，再使用该功能。",
                               parent=self.root)
        return False

    def _confirm_leave(self):
        if messagebox.askyesno("放弃编辑？", "当前有未保存的编辑内容，确定放弃吗？",
                               parent=self.root):
            self.editor.cancel()
            return True
        return False

    def start_editor(self, kind, preset_image=None):
        if not self._need_doc():
            return
        if self.editor and not self._confirm_leave():
            return
        if self.editor:
            self.editor.cancel()
        self.editor_slot.pack(fill="x", before=self.mid)
        self.editor = Editor(self, kind, preset_image)

    def open_sign(self):
        if self._need_doc():
            SignDialog(self)

    # ================= 工具入口 =================
    def open_text(self):
        self.start_editor("text")

    def open_image(self):
        self.start_editor("image")

    def open_extract(self):
        if self._need_doc():
            ExtractDialog(self)

    def open_split(self):
        if self._need_doc():
            SplitDialog(self)

    def open_delete(self):
        if self._need_doc():
            DeletePagesDialog(self)

    def open_rotate(self):
        if self._need_doc():
            RotatePagesDialog(self)

    def open_pagenum(self):
        if self._need_doc():
            PageNumbersDialog(self)

    def open_compress(self):
        if self._need_doc():
            CompressDialog(self)

    def open_compress_target(self):
        if self._need_doc():
            CompressToTargetDialog(self)

    def open_office2pdf(self):
        OfficeToPdfDialog(self)

    def open_crop(self):
        if self._need_doc():
            CropPagesDialog(self)

    def open_encrypt(self):
        if self._need_doc():
            EncryptDialog(self)

    def open_decrypt(self):
        if self._need_doc():
            DecryptDialog(self)

    def open_watermark(self):
        if self._need_doc():
            WatermarkDialog(self)

    def open_annotate(self):
        if self._need_doc():
            AnnotateDialog(self)

    def open_redact(self):
        if self._need_doc():
            RedactDialog(self)

    def open_hf(self):
        if self._need_doc():
            HeaderFooterDialog(self)

    def open_extract_img(self):
        if self._need_doc():
            ExtractImagesDialog(self)

    def open_reorder(self):
        if self._need_doc():
            ReorderDialog(self)

    def open_pdf2word(self):
        if self._need_doc():
            PdfToWordDialog(self)

    def open_pdf2excel(self):
        if self._need_doc():
            PdfToExcelDialog(self)

    def open_toc(self):
        if self._need_doc():
            TocDialog(self)

    def open_form(self):
        if self._need_doc():
            FormFillDialog(self)

    def open_batch(self):
        BatchDialog(self)

    def open_pdf2img(self):
        if self._need_doc():
            PdfToImagesDialog(self)

    def open_merge(self):
        MergeDialog(self)

    def open_img2pdf(self):
        ImagesToPdfDialog(self)

    def open_print(self):
        if not self._need_doc():
            return
        names, default = list_printers()
        if not names:
            messagebox.showwarning("没有打印机", "未检测到可用打印机。", parent=self.root)
            return
        PrintDialog(self, names, default)

    def run(self):
        self.root.mainloop()


def _mcp_cli(argv):
    """隐藏入口：让打包后的 exe 也能被 AI 客户端直接拉起当 MCP 服务器。

    --mcp-stdio        以 MCP stdio 服务器身份运行（客户端拉起子进程）
    --mcp-http H:P     以 MCP HTTP 服务器身份运行
    --mcp-self-test    自检后退出（打印 JSON）
    返回退出码；参数不匹配时返回 None（调用方走正常 GUI）。
    """
    import argparse
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--mcp-stdio", action="store_true")
    ap.add_argument("--mcp-http", default="")
    ap.add_argument("--mcp-self-test", action="store_true")
    ap.add_argument("--mcp-port", type=int, default=8765)
    ap.add_argument("--mcp-allow-dir", action="append", default=[])
    ap.add_argument("--mcp-source-url", default="")
    ap.add_argument("--mcp-allow-print", action="store_true")
    try:
        a, _rest = ap.parse_known_args(argv)
    except SystemExit:
        return None
    if not (a.mcp_stdio or a.mcp_http or a.mcp_self_test):
        return None
    if pdf_mcp is None:
        sys.stderr.write("MCP 模块不可用：%s\n" % _MCP_ERR)
        return 3
    if a.mcp_allow_print:
        pdf_mcp.ALLOW_HIGH_RISK = True
    pdf_mcp.SOURCE_URL = (a.mcp_source_url or "").rstrip("/")
    pdf_mcp.ALLOW_DIRS = [os.path.abspath(d) for d in a.mcp_allow_dir]
    if a.mcp_self_test:
        r = pdf_mcp._call_tool(1, {"name": pdf_mcp.SOURCE_TOOL_NAME,
                                   "arguments": {}})
        info = r["result"]["structuredContent"]
        sys.stdout.write(json.dumps({
            "tools": len(pdf_mcp.TOOLS),
            "protocol": pdf_mcp.PROTOCOL_VERSION,
            "license": info["license"],
            "frozen": bool(getattr(sys, "frozen", False)),
            "source_files": [f["file"] for f in info["source_files"]],
            "components": info["components"],
        }, ensure_ascii=False, indent=2) + "\n")
        return 0
    if a.mcp_stdio:
        pdf_mcp.serve_stdio()
        return 0
    host, _, port_s = a.mcp_http.partition(":")
    try:
        port = int(port_s or a.mcp_port)
    except ValueError:
        sys.stderr.write("--mcp-http 格式应为 HOST:PORT\n")
        return 2
    if not pdf_mcp._is_loopback(host):
        if not a.mcp_allow_dir:
            sys.stderr.write(
                "拒绝启动：%s 不是本机回环且未设 --mcp-allow-dir，\n"
                "这等于把本机文件系统对网络开放。\n" % host)
            return 2
        sys.stderr.write("[mcp] 注意：监听 %s，路径限制在 %s\n"
                         % (host, "、".join(a.mcp_allow_dir)))
    pdf_mcp.serve_http(host, port)
    return 0


if __name__ == "__main__":
    _rc = _mcp_cli(sys.argv[1:])
    if _rc is None:
        App().run()
    else:
        sys.exit(_rc)
