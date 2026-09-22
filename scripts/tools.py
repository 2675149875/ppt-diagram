# -*- coding: utf-8 -*-
"""外部工具发现:LibreOffice (soffice) 与 poppler (pdftocairo)。

导出链需要这两个外部程序,但它们的安装位置因机器而异,所以不写死路径。
查找顺序:

    1. 环境变量   PPT_DIAGRAM_SOFFICE / PPT_DIAGRAM_PDFTOCAIRO
    2. PATH       (shutil.which)
    3. 候选目录   (见下方 *_CANDIDATES)

第 1 条是给你本机那种非标准安装位置用的 —— 装在哪儿就把环境变量指过去,
不用改代码。第 3 条只是尽量猜得准一点,猜不到就让 require_* 报出装法。

用法:
    from tools import require_soffice, require_pdftocairo
    soffice = require_soffice()          # 找不到会抛出带安装指引的 RuntimeError

注意 export_figure.ps1 里有一份**独立**的等价逻辑。那里刻意不 import 本模块 ——
导出链不该为了找路径反过来依赖 Python。
"""
import os
import shutil
import sys

ENV_SOFFICE = "PPT_DIAGRAM_SOFFICE"
ENV_PDFTOCAIRO = "PPT_DIAGRAM_PDFTOCAIRO"

_IS_WIN = sys.platform.startswith("win")

# ---- 候选目录 ----------------------------------------------------------------
# 顺序即优先级。Windows 在前(本 skill 的导出链主要在 Windows 上验证过),
# POSIX 路径是尽力而为,未实测。
_SOFFICE_WIN = [
    r"%ProgramFiles%\LibreOffice\program\soffice.com",
    r"%ProgramFiles(x86)%\LibreOffice\program\soffice.com",
    r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.com",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
]
_SOFFICE_POSIX = [
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

_PDFTOCAIRO_WIN = [
    r"%USERPROFILE%\.claude\skills\shared\tools\poppler\pdftocairo.exe",
    r"%LOCALAPPDATA%\poppler\pdftocairo.exe",
    r"C:\Program Files\poppler\bin\pdftocairo.exe",
    r"C:\ProgramData\chocolatey\bin\pdftocairo.exe",
]
_PDFTOCAIRO_POSIX = [
    "/usr/bin/pdftocairo",
    "/usr/local/bin/pdftocairo",
    "/opt/homebrew/bin/pdftocairo",
]


def _expand(path):
    return os.path.expandvars(os.path.expanduser(path))


def _find(env_var, names, candidates):
    """按 环境变量 → PATH → 候选目录 的顺序找一个可执行文件。"""
    override = os.environ.get(env_var)
    if override:
        path = _expand(override)
        # 环境变量是显式意图:指错了要报出来,不能静默回退到别处
        if os.path.isfile(path):
            return path
        raise RuntimeError(
            f"环境变量 {env_var} 指向的文件不存在:{path}\n"
            f"请改成实际路径,或清空该变量改用自动查找。")

    for name in names:
        found = shutil.which(name)
        if found:
            return found

    for cand in candidates:
        path = _expand(cand)
        if os.path.isfile(path):
            return path
    return None


def find_soffice():
    """返回 soffice 的完整路径,找不到返回 None。"""
    names = ["soffice.com", "soffice"] if _IS_WIN else ["soffice"]
    cands = _SOFFICE_WIN + _SOFFICE_POSIX if _IS_WIN else _SOFFICE_POSIX
    return _find(ENV_SOFFICE, names, cands)


def find_pdftocairo():
    """返回 pdftocairo 的完整路径,找不到返回 None。"""
    names = ["pdftocairo.exe", "pdftocairo"] if _IS_WIN else ["pdftocairo"]
    cands = _PDFTOCAIRO_WIN + _PDFTOCAIRO_POSIX if _IS_WIN else _PDFTOCAIRO_POSIX
    return _find(ENV_PDFTOCAIRO, names, cands)


_HOW_TO_INSTALL = {
    "soffice": (
        "LibreOffice —— 用来把 PPTX 转成矢量 PDF。\n"
        "  Windows : winget install TheDocumentFoundation.LibreOffice\n"
        "            或到 https://www.libreoffice.org/download/ 下载\n"
        "  macOS   : brew install --cask libreoffice\n"
        "  Linux   : sudo apt install libreoffice\n"
        "  已装在非标准位置时,设环境变量 {}\n"
    ).format(ENV_SOFFICE),
    "pdftocairo": (
        "pdftocairo —— poppler 的一部分,用来把 PDF 转成高 DPI PNG / 矢量 SVG。\n"
        "  Windows : 下载 poppler for Windows,把 bin/ 加入 PATH\n"
        "  macOS   : brew install poppler\n"
        "  Linux   : sudo apt install poppler-utils\n"
        "  已装在非标准位置时,设环境变量 {}\n"
    ).format(ENV_PDFTOCAIRO),
}


def require_soffice():
    path = find_soffice()
    if not path:
        raise RuntimeError("找不到 LibreOffice (soffice)。\n" + _HOW_TO_INSTALL["soffice"])
    return path


def require_pdftocairo():
    path = find_pdftocairo()
    if not path:
        raise RuntimeError("找不到 pdftocairo。\n" + _HOW_TO_INSTALL["pdftocairo"])
    return path


if __name__ == "__main__":
    # 诊断入口:python tools.py —— 直接看这两个工具被解析到了哪里
    for label, finder in (("soffice", find_soffice), ("pdftocairo", find_pdftocairo)):
        try:
            print(f"{label:12} {finder() or '** 未找到 **'}")
        except RuntimeError as exc:
            print(f"{label:12} ** 配置错误 **\n{exc}")
