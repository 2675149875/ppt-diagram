# -*- coding: utf-8 -*-
"""
示例:三阶段方法框架图 —— 演示 diagram_kit 的典型用法(期刊插图风)。

要点:
  · 画布宽度 = 论文版心(A4 + 3.17cm 边距 = 5.77 英寸),字号即最终印刷字号
  · 不传 theme 时是 tinted(浅色底 + 同色系深字);换 academic / mono / presentation 见 SKILL.md
  · 直角走线用多段 arrow() 拼,不用 elbow()

运行:python example_pipeline.py [输出路径]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from diagram_kit import Diagram, DARK  # noqa: E402

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.environ.get("TEMP", "."), "ppt_diagram_example.pptx")

d = Diagram(width_in=5.77, height_in=2.60)          # theme 默认 tinted

# ---- 三个主阶段,网格自动等距排列 ----
# gap 要留够:箭头标注放在间隙里,间隙太窄标签会压到两侧盒子上
Y, H, GAPX, LBLW = 0.45, 0.55, 0.45, 0.45
d.rbox(0, Y, H, "数据输入\n与预处理", 3, style="blue")
d.rbox(1, Y, H, "模型建立\n与求解",   3, style="green")
d.rbox(2, Y, H, "结果分析\n与验证",   3, style="blue")

# ---- 阶段之间的箭头 + 间隙里的标注 ----
cells = d.grid(3, start=0.15, end=5.62, gap=GAPX)
for i in range(2):
    x1 = cells[i][0] + cells[i][1]
    x2 = cells[i + 1][0]
    d.arrow(x1, Y + H / 2, x2, Y + H / 2)
    d.label((x1 + x2) / 2 - LBLW / 2, Y + H / 2 + 0.05, LBLW,
            "特征提取" if i == 0 else "误差评估", size=7)

# ---- 反馈回路:虚线,多段 arrow 拼直角(elbow 走线由渲染器决定,不可控) ----
bx = cells[2][0] + cells[2][1] - 0.2
d.arrow(bx, Y + H, bx, Y + H + 0.45, dashed=True, head=False)
d.arrow(bx, Y + H + 0.45, cells[1][0] + cells[1][1] / 2, Y + H + 0.45, dashed=True)
d.label(cells[1][0], Y + H + 0.52, cells[1][1], "参数迭代优化",
        size=7, color=DARK, italic=False)

# ---- 下方补充说明框:非主流程用灰框 + 浅色填充 ----
d.box(0.15, Y + H + 0.95, 5.47, 0.45,
      "约束条件：计算资源有限、数据含噪声、需在 72 小时内完成",
      style="gray", filled=False, size=7)

# 注意:不要加 d.caption("图 1  方法总体框架")。
# 图注会被烤进导出的 PNG/SVG,插进论文后跟 Word 题注重复。详见 SKILL.md 铁律。

d.save(out)
print("SAVED", out)
print("提示:用 PowerPoint 打开微调后,运行 export_figure.ps1 导出成图")
