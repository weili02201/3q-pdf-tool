# -*- coding: utf-8 -*-
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (C) 2026 三扣科技 (3Q Tech)
"""3Q PDF开源免费工具 · exe 重打包（PyInstaller onefile，仅 Windows）。

用法：
    D:\\ProgramData\\Anaconda3\\envs\\wlpy310\\python.exe build_exe.py

为什么这么打包（每条都对应一个已发生的故障）：
1. pdf_core / pdf_cli / pdf_mcp 显式 hidden-import：
   pdf_tool 在 try 里 import 它们，漏打的话成品里「智能体接入」
   按钮只会显示「MCP 模块不可用」，而且不报错、很难发现。
2. 4 个 .py 源码 + LICENSE + NOTICE 必须当数据文件打进去：
   - gen_pdf_code 要读 pdf_core.py / pdf_cli.py 的原始源码文本；
   - HTTP 的 /source 端点（AGPL 第 13 条义务）要能真的供出文件。
   PyInstaller 默认只装编译后的 .pyc，不装 .py 原文。
3. --collect-all pymupdf / tkinterdnd2：
   fitz 是转发到 pymupdf 的垫片包，MuPDF 二进制 _mupdf.pyd 在
   pymupdf 里；tkdnd 的 Windows 二进制在 tkinterdnd2/tkdnd/win-x64/，
   漏了它 GUI 启动即崩（TkinterDnD.Tk() 找不到 tkdnd 包）。
4. --windowed：GUI 不要控制台。stdio 入口靠 pdf_mcp._stdio_streams
   的 fd 0/1 兜底，被客户端用管道拉起时协议帧不丢。

环境注记（两条都踩过）：
1. wlpy310 是数据科学全家桶环境（torch/transformers/cv2/sqlalchemy/...）。
   pymupdf/__init__.py 的 `from .table import find_tables` 会把 pandas
   拖进依赖图，pandas 的 hook 再 collect_submodules 连 pandas.tests 都不
   放过，于是一路雪崩到 torch/sentry_sdk；而本机 sentry_sdk 装了却是
   坏的（httpcore 链断），它的 PyInstaller hook 一加载就 JSONDecodeError，
   整个构建崩掉。所以在项目内建一个干净构建 venv（.venv-build），
   只装运行时真正需要的 7 个包，依赖图立刻收敛。
2. wlpy310 里还装着 pathlib 过时垫底包，PyInstaller 启动时的
   check_requirements() 会据此拒绝运行（只查 dist-info 元数据）。
   干净 venv 没这个问题；这里的子进程补丁是双保险，对构建结果零影响。

构建 venv 的重建方法（依赖版本与 wlpy310 实测一致，勿随意升）：
    D:\\ProgramData\\Anaconda3\\envs\\wlpy310\\python.exe -m venv .venv-build
    .venv-build\\Scripts\\python.exe -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple ^
        pymupdf==1.25.2 pillow==10.4.0 python-docx==1.1.2 openpyxl==3.1.5 ^
        pywin32==312 tkinterdnd2==0.6.4 pyinstaller==6.3.0
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))

# 干净构建 venv 优先；不存在则退回当前解释器（wlpy310 需靠下面的补丁）
VENV_PY = os.path.join(ROOT, ".venv-build", "Scripts", "python.exe")
BUILD_PY = VENV_PY if os.path.exists(VENV_PY) else sys.executable

# (源文件, 包内目标目录)。目标 "." = 打平到 _MEIPASS 根，
# 与 pdf_mcp.BASE_DIR 的假设一致。
DATA_FILES = [
    "pdf_core.py",
    "pdf_cli.py",
    "pdf_mcp.py",
    "pdf_tool.py",
    "LICENSE",
    "NOTICE",
]

# 打赏码（可选）：AboutDialog 会在 _MEIPASS 里找「打赏码.png」，有就嵌入展示
if os.path.exists(os.path.join(ROOT, "打赏码.png")):
    DATA_FILES.append("打赏码.png")

pyi_args = [
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", "3Q PDF开源免费工具",
    "--distpath", os.path.join(ROOT, "dist"),
    "--workpath", os.path.join(ROOT, "build"),
    "--specpath", os.path.join(ROOT, "build"),
    "--hidden-import", "pdf_core",
    "--hidden-import", "pdf_cli",
    "--hidden-import", "pdf_mcp",
    "--collect-all", "pymupdf",
    "--collect-all", "tkinterdnd2",
    # dist-info 也要打进包：pdf_source 工具的「组件许可清单」是运行时扫
    # dist-info METADATA 现读的，缺了它成品 exe 里这一栏会是空的。
    "--copy-metadata", "pymupdf",
]
for name in DATA_FILES:
    src = os.path.join(ROOT, name)
    if not os.path.exists(src):
        sys.stderr.write("缺少数据文件：%s\n" % src)
        sys.exit(1)
    pyi_args += ["--add-data", src + os.pathsep + "."]
pyi_args.append(os.path.join(ROOT, "pdf_tool.py"))

# 子进程内做两件事，再走 PyInstaller 正规 CLI 路径：
# 1) 把 base 环境的 Library\\bin 与 DLLs 加进 PATH——venv 子进程默认不带
#    conda 激活时的 PATH，PyInstaller 因此解析不到 ffi.dll / libcrypto-3-x64 /
#    libssl-3-x64 / liblzma / LIBBZ2 / tcl86t / tk86t 这些依赖 DLL，
#    而 hashlib（pdf_mcp 的 sha256）与 tkinter（GUI）都指着它们活；
#    base 环境路径取 sys.base_prefix，venv 由哪个解释器建的就指向哪。
# 2) 跳过 pathlib 垫底包检查（干净 venv 里本就是空操作，双保险）。
_WRAP = (
    "import os, sys;"
    "os.environ['PATH'] = os.pathsep.join("
    "    [p for p in (os.path.join(sys.base_prefix, 'Library', 'bin'),"
    "                 sys.base_prefix + os.sep + 'DLLs')"
    "     if os.path.isdir(p)] + [os.environ.get('PATH', '')]);"
    "from PyInstaller import __main__ as _pm;"
    "_pm.compat.check_requirements = lambda: None;"
    "_pm.run(sys.argv[1:])"
)
cmd = [BUILD_PY, "-c", _WRAP] + pyi_args

print("构建解释器：%s" % BUILD_PY)
print(" ".join(cmd))
rc = subprocess.call(cmd, cwd=ROOT)
if rc == 0:
    exe = os.path.join(ROOT, "dist", "3Q PDF开源免费工具.exe")
    print("\n打包完成：%s（%d 字节）" % (exe, os.path.getsize(exe)))
sys.exit(rc)
