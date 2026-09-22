# -*- coding: utf-8 -*-
"""
风格对比表:同一版式的图,用 4 种 theme 各渲染一遍,拼成一张对比图。

为什么要固定几何:同一张图只换 theme、画布与字号不变,比的才是**风格**;
如果连尺寸一起变,比的是尺寸,选不出想要的东西。

用法:
    python style_gallery.py [输出目录]
    # 默认输出到 %TEMP%\\diagram_style_gallery\\

产出:
    style_<theme>.pptx / .png   —— 单张
    风格对比.png                —— 2×2 拼版,拿去挑

选好之后在正式脚本里写 theme="tinted" 即可(theme 默认就是 tinted)。

导出这一步不走 PowerShell(避免执行策略问题),直接调 soffice + pdftocairo,
与 export_figure.ps1 是同一套参数。这两个程序的路径由 tools.py 自动发现,
装在非标准位置就设环境变量 PPT_DIAGRAM_SOFFICE / PPT_DIAGRAM_PDFTOCAIRO。
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from diagram_kit import Diagram, DARK          # noqa: E402
from tools import require_soffice, require_pdftocairo  # noqa: E402

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    tempfile.gettempdir(), "diagram_style_gallery")
os.makedirs(OUT, exist_ok=True)

DPI = 200

THEMES = [
    ("tinted", "① 浅色底（默认）",
     "原色向白混 82% 的浅底 + 同色系深字。层级靠色块区分，比实心柔和"),
    ("academic", "② 白底细框",
     "白底 + 细彩框 + 近黑常规字。最素净，最像期刊插图"),
    ("mono", "③ 纯黑白",
     "统一黑框 + 浅灰底区分次级框。不依赖颜色，黑白印刷不会丢信息"),
    ("presentation", "④ 汇报风",
     "实心饱和色块 + 白色粗体字。视觉冲击强，适合答辩 PPT"),
]

# ---- 固定几何:5.77 英寸版心、8pt、0.04 内边距,四个主题完全一致 ----
W, H = 5.77, 2.60
SIZE, PAD = 8, 0.04
Y, BH, GAPX, LBLW = 0.45, 0.55, 0.45, 0.45


def build(theme):
    d = Diagram(width_in=W, height_in=H, theme=theme)
    bx = lambda *a, **k: d.box(*a, size=SIZE, pad=PAD, **k)   # noqa: E731

    cells = d.grid(3, start=0.15, end=W - 0.15, gap=GAPX)
    for i, (txt, st) in enumerate([("数据输入\n与预处理", "blue"),
                                   ("模型建立\n与求解", "green"),
                                   ("结果分析\n与验证", "blue")]):
        bx(cells[i][0], Y, cells[i][1], BH, txt, style=st)

    for i in range(2):
        x1 = cells[i][0] + cells[i][1]
        x2 = cells[i + 1][0]
        d.arrow(x1, Y + BH / 2, x2, Y + BH / 2)
        d.label((x1 + x2) / 2 - LBLW / 2, Y + BH / 2 + 0.05, LBLW,
                "特征提取" if i == 0 else "误差评估", size=7)

    bx_ = cells[2][0] + cells[2][1] - 0.2
    d.arrow(bx_, Y + BH, bx_, Y + BH + 0.45, dashed=True, head=False)
    d.arrow(bx_, Y + BH + 0.45, cells[1][0] + cells[1][1] / 2, Y + BH + 0.45, dashed=True)
    d.label(cells[1][0], Y + BH + 0.52, cells[1][1], "参数迭代优化",
            size=7, color=DARK, italic=False)

    bx(0.15, Y + BH + 0.95, W - 0.30, 0.45,
       "约束条件：计算资源有限、数据含噪声、需在 72 小时内完成",
       style="gray", filled=False)
    return d


def export_png(pptx, outdir):
    """PPTX -> PDF(LibreOffice) -> PNG(pdftocairo),与 export_figure.ps1 同参数。"""
    soffice = require_soffice()
    pdftocairo = require_pdftocairo()
    profile = "file:///" + os.path.join(tempfile.gettempdir(), "lo_export_profile").replace("\\", "/")
    subprocess.run([soffice, "--headless", "--norestore",
                    f"-env:UserInstallation={profile}",
                    "--convert-to", "pdf", "--outdir", outdir, pptx],
                   capture_output=True, text=True)

    base = os.path.splitext(os.path.basename(pptx))[0]
    pdf = os.path.join(outdir, base + ".pdf")
    if not os.path.exists(pdf):
        sys.exit(f"LibreOffice 未能生成 PDF: {pdf}")

    # pdftocairo 对非 ASCII 路径支持不佳,统一在 ASCII 临时目录里转换再搬回
    with tempfile.TemporaryDirectory(prefix="pdftocairo_") as tmp:
        tmp_pdf = os.path.join(tmp, "in.pdf")
        shutil.copy2(pdf, tmp_pdf)
        subprocess.run([pdftocairo, "-png", "-r", str(DPI), "-singlefile",
                        tmp_pdf, os.path.join(tmp, "out")], capture_output=True)
        made = os.path.join(tmp, "out.png")
        if not os.path.exists(made):
            sys.exit(f"pdftocairo 未能生成 PNG: {base}")
        shutil.copy2(made, os.path.join(outdir, base + ".png"))
    return os.path.join(outdir, base + ".png")


pptxes = []
for key, _, _ in THEMES:
    p = os.path.join(OUT, f"style_{key}.pptx")
    build(key).save(p)
    pptxes.append(p)
    print("SAVED", p)

for p in pptxes:
    print("  ->", export_png(p, OUT))

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print("\n没装 Pillow,跳过拼版。单张图在:", OUT)
    sys.exit(0)

FONT = r"C:\Windows\Fonts\msyh.ttc"
FONT_B = r"C:\Windows\Fonts\msyhbd.ttc"
f_title = ImageFont.truetype(FONT_B, 26)
f_desc = ImageFont.truetype(FONT, 19)

PAD, CGX, CGY, BAR_H, DESC_H = 26, 26, 30, 44, 44
panels = [(Image.open(os.path.join(OUT, f"style_{k}.png")).convert("RGB"), t, s)
          for k, t, s in THEMES]

pw = max(p[0].width for p in panels)
ph = max(p[0].height for p in panels)
cw, ch = pw + 2 * PAD, BAR_H + DESC_H + ph + 2 * PAD
sheet = Image.new("RGB", (2 * cw + 3 * CGX, 2 * ch + 3 * CGY), (250, 250, 250))
dr = ImageDraw.Draw(sheet)

for idx, (im, title, desc) in enumerate(panels):
    cx = CGX + (idx % 2) * (cw + CGX)
    cy = CGY + (idx // 2) * (ch + CGY)
    dr.rectangle([cx, cy, cx + cw, cy + ch], fill="white", outline=(214, 214, 214))
    dr.rectangle([cx, cy, cx + cw, cy + BAR_H], fill=(38, 38, 38))
    dr.text((cx + PAD, cy + 8), title, font=f_title, fill="white")
    dr.text((cx + PAD, cy + BAR_H + 10), desc, font=f_desc, fill=(70, 70, 70))
    ix, iy = cx + (cw - im.width) // 2, cy + BAR_H + DESC_H + PAD
    sheet.paste(im, (ix, iy))
    dr.rectangle([ix, iy, ix + im.width, iy + im.height], outline=(220, 220, 220))

path = os.path.join(OUT, "风格对比.png")
sheet.save(path)
print("\n对比表:", path)
