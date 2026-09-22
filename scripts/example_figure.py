# -*- coding: utf-8 -*-
"""示例图:迭代式状态估计方法的总体框架。

这是本 skill 的**参考版式** —— 演示一张"内容充实、箭头贴边框、层次清楚"的
学术示意图该怎么搭。`example_pipeline.py`(出 pptx)和 `style_gallery.py`
(四风格对比)都调这里的 build(),保证两者永远是同一张图。

版式要点(逐条对照 SKILL.md「图形质量要求」):
  · 纵向主链:数据 → 模型 → 求解 → 判据 → 输出 → 结论
  · 求解层用 bus() **树形分叉**成三个并行步骤,再用 fan_in() **汇聚**到判据
  · 判据用**菱形**(标准里的"判断框"),两个出口各带 是/否 标注
  · "否"出口用**虚线回路绕左侧**返回第一步 —— 绕在框外走,不穿任何框
  · 三个求解步骤套一个**虚线分组容器**
  · **所有箭头都由 Node 的边框坐标算出**,不存在手算的半宽
  · 框内写具体公式/算子,不写"数据输入与预处理"这类泛称
  · 同层级元素**等宽等距**
  · **不画图注**(SKILL.md 铁律:图注归 Word 题注管)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from diagram_kit import Diagram, DARK  # noqa: E402
from pptx.enum.text import PP_ALIGN     # noqa: E402

# 画布 = 论文版心宽(A4 + 3.17cm 边距)。这样 8pt 就是最终印刷字号,所见即所得。
W, H = 5.77, 4.26

MX = 0.16           # 右边距
MX_LEFT = 0.55      # 左边距 —— 留宽一点,给"否"的虚线回路让出一条通道
CW = W - MX_LEFT - MX

# 各层的 y / 高度。集中放这里,调整版式时一眼能看全,不用在几十个数字里找。
L1_Y, L1_H = 0.12, 0.38     # 数据与先验
L2_Y, L2_H = 0.66, 0.42     # 模型
BUS_Y = 1.22                # 模型 → 求解 的分叉总线。显式指定是因为下面要放
                            # 容器标签,默认的中点位置会和标签撞上
L3_Y, L3_H = 1.62, 0.46     # 三个并行求解步骤
GRP_PAD = 0.10              # 分组容器的四周留白
GATE_Y, GATE_W, GATE_H = 2.50, 1.00, 0.46   # 收敛判据(菱形)
L4_Y, L4_H = 3.16, 0.38     # 结果输出
L5_Y, L5_H = 3.74, 0.38     # 结论与推广
LOOP_X = 0.28               # "否"回路走左侧这条竖线的 x


def build(theme="tinted", width_in=W, height_in=H):
    """画好并返回 Diagram。theme 见 SKILL.md「风格预设」。"""
    d = Diagram(width_in=width_in, height_in=height_in, theme=theme)

    # ---------------- 第一层:数据与先验(两个平行输入) ----------------
    gap = 0.28
    bw = (CW - gap) / 2
    obs = d.box(MX_LEFT, L1_Y, bw, L1_H, "y_{1:T}\n观测序列", style="sky")
    pri = d.box(MX_LEFT + bw + gap, L1_Y, bw, L1_H,
                "x_0 ~ p(x_0)\n初始状态先验", style="sky")

    # ---------------- 第二层:模型 ----------------
    model = d.box(MX_LEFT, L2_Y, CW, L2_H,
                  "状态空间模型    x_{t+1} = f(x_t, u_t) + w_t,    "
                  "y_t = h(x_t) + v_t",
                  style="blue", size=7.5)

    # ---------------- 第三层:三个并行求解步骤(等宽等距) ----------------
    g = 0.26
    sw = (CW - 2 * g) / 3
    steps = [
        d.box(MX_LEFT + 0 * (sw + g), L3_Y, sw, L3_H,
              "E 步\n求后验 p(x_{1:T}|y,θ)", style="orange", size=7.5),
        d.box(MX_LEFT + 1 * (sw + g), L3_Y, sw, L3_H,
              "M 步\nθ ← argmax Q(θ)", style="orange", size=7.5),
        d.box(MX_LEFT + 2 * (sw + g), L3_Y, sw, L3_H,
              "似然评估\nℓ(θ) = log p(y|θ)", style="orange", size=7.5),
    ]
    # 分组容器。标签放容器左上角的外侧、并且做得窄 —— 放正上方或放宽了,
    # 都会被 bus 的支线从中间穿过(支线落在各步骤上边框的中点)。
    grp = d.group(steps, pad=GRP_PAD, style="gray")
    d.label(grp.x, L2_Y + L2_H + 0.24, 0.72, "E–M 交替迭代",
            size=7.5, color=DARK, align=PP_ALIGN.LEFT)

    # ---------------- 第四层:收敛判据(菱形判断框) ----------------
    gate = d.box(MX_LEFT + (CW - GATE_W) / 2, GATE_Y, GATE_W, GATE_H,
                 "收敛?", style="purple", shape="diamond", size=7.5)

    # ---------------- 第五 / 六层:输出与结论 ----------------
    out = d.box(MX_LEFT, L4_Y, CW, L4_H,
                "状态估计 x̂_{1:T}      误差评估 RMSE = √(Σ‖x̂_t − x_t‖²/T)",
                style="green", size=7.5)
    concl = d.box(MX_LEFT, L5_Y, CW, L5_H,
                  "灵敏度分析:σ_w 扰动      模型推广至非线性观测 h(·)",
                  style="green", size=7.5)

    # ================= 连线:端点全部来自 Node 的边框坐标 =================

    # 数据 + 先验 → 模型。这两个输入并排,各引一条箭头落到模型上边框的不同位置
    # 就够了 —— 这儿不用 fan_in,是因为层间距只有 0.18",摆一条总线会让带箭头的
    # 那段短到 0.1" 以下,退化成孤立小三角。
    # port() 的落点特意取在**正对源框中线**的位置,这样箭头是完全竖直的;
    # 随便取个落点就会变成斜箭头,而斜箭头正是"看起来乱"的主要来源。
    for src in (obs, pri):
        d.arrow(*src.bottom, *model.port("top", (src.cx - model.x) / model.w))

    # 模型 → 三个步骤:树形分叉
    d.bus(model, steps, bus_at=BUS_Y)

    # 三个步骤 → 判据:汇聚
    d.fan_in(gate, steps)

    # 判据"是" → 输出
    d.connect(gate, out)
    d.label(gate.cx + 0.09, gate.y + gate.h + 0.02, 0.4, "是",
            size=7.5, color=DARK, align=PP_ALIGN.LEFT)

    # 输出 → 结论
    d.connect(out, concl)

    # 判据"否" → 第一步:虚线,绕左侧通道上行
    d.arrow(*gate.left, LOOP_X, gate.cy, dashed=True, head=False)
    d.arrow(LOOP_X, gate.cy, LOOP_X, steps[0].cy, dashed=True, head=False)
    d.arrow(LOOP_X, steps[0].cy, *steps[0].left, dashed=True)
    d.label(LOOP_X + 0.08, gate.cy + 0.04, 1.2, "否,继续迭代",
            size=7.5, color=DARK, align=PP_ALIGN.LEFT)

    return d
