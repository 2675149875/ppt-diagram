# -*- coding: utf-8 -*-
"""示例:生成一张完整的学术示意图(迭代式状态估计方法总体框架)。

图和 `style_gallery.py` 共用同一个构造函数 `example_figure.build()`,
所以这里看到的就是四风格对比里那张图 —— 不会出现"示例和对比图不是同一张"。

想学版式怎么搭,直接读 `example_figure.py`,那里逐条对着 SKILL.md 的
「图形质量要求」写了注释。这个文件只负责出 pptx。

运行:python example_pipeline.py [输出路径]
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from example_figure import build  # noqa: E402

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.environ.get("TEMP", "."), "ppt_diagram_example.pptx")

d = build()

# 自查:audit() 查盒子重叠、箭头杆过短、斜箭头、内容出画布。
# 这几样都是"缩略图上看不出来、印出来才发现"的瑕疵,值得每次画完跑一遍。
issues = d.audit()
if issues:
    print("⚠️  版面自查发现问题:", file=sys.stderr)
    for i in issues:
        print("   - " + i, file=sys.stderr)
else:
    print("版面自查:无问题")

# 注意:不要加 d.caption("图 1  xxx")。
# 图注一旦画进画布就会被烤进导出的 PNG/SVG,插进论文后跟 Word 题注重复。
# 详见 SKILL.md 铁律 —— 编号和标题交给 Word 的"引用 → 插入题注"。

d.save(out)
print("SAVED", out)
