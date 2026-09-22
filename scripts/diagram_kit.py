# -*- coding: utf-8 -*-
"""
PPT 示意图脚手架 —— 用简单 API 生成可编辑的学术示意图 (PPTX)。

配色与字体遵循 scientific-visualization 模块的出版级规范:
  · 配色:Okabe-Ito 色盲友好色板 (Okabe & Ito 2008)
  · 字体:拉丁 Arial / 中文 微软雅黑

用法:
    from diagram_kit import Diagram
    d = Diagram()
    d.box(0.5, 1.5, 2.0, 0.8, "数据输入", style="blue")
    d.arrow(2.5, 1.9, 3.0, 1.9)
    d.label(2.5, 1.4, 0.6, "预处理")
    d.save(r"C:\\path\\out.pptx")

生成后用 PowerPoint 打开手动微调,再用 scripts/export_figure.ps1 导出成图。
"""
import os
import re
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ============================================================================
# 配色:Okabe-Ito 色盲友好色板
# 来源:Okabe, M. & Ito, K. (2008). Color Universal Design (CUD):
#       How to Make Figures and Presentations That Are Friendly to
#       Colorblind People.  https://jfly.uni-koeln.de/color/
# 说明:科学出版最广泛推荐的色盲友好配色,约 8% 男性有色觉障碍。
#       换色请从 OKABE_ITO 里取,不要随意引入新色。
# ============================================================================
OKABE_ITO = {
    "orange":     "E69F00",
    "sky":        "56B4E9",
    "green":      "009E73",
    "yellow":     "F0E442",
    "blue":       "0072B2",
    "vermillion": "D55E00",
    "purple":     "CC79A7",
    "black":      "000000",
}


def _strip_style(shp):
    """删掉 <p:style>。

    否则 a:effectRef 会引用主题里的投影效果。PowerPoint 认空 <a:effectLst/>
    覆盖,LibreOffice 不认 —— 导出成 PNG/SVG 时框上照样挂着投影,一眼像 PPT
    截图而不是期刊插图。填充/描边我们都已显式写在 spPr 里,不依赖这个 style。
    """
    sp = shp._element
    style_el = sp.find(qn("p:style"))
    if style_el is not None:
        sp.remove(style_el)
    return shp


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _tint(hexstr, ratio=0.82):
    """把颜色向白色混合,生成浅色填充版。ratio = 混入的白色比例。"""
    r, g, b = int(hexstr[0:2], 16), int(hexstr[2:4], 16), int(hexstr[4:6], 16)
    r = int(r + (255 - r) * ratio)
    g = int(g + (255 - g) * ratio)
    b = int(b + (255 - b) * ratio)
    return RGBColor(r, g, b)


# style 名 → Okabe-Ito 原色作描边/实心;soft 为自动派生的浅色填充
STYLES = {
    "blue":       {"fill": _rgb(OKABE_ITO["blue"]),       "soft": _tint(OKABE_ITO["blue"])},
    "orange":     {"fill": _rgb(OKABE_ITO["orange"]),     "soft": _tint(OKABE_ITO["orange"])},
    "green":      {"fill": _rgb(OKABE_ITO["green"]),      "soft": _tint(OKABE_ITO["green"])},
    "sky":        {"fill": _rgb(OKABE_ITO["sky"]),        "soft": _tint(OKABE_ITO["sky"])},
    "vermillion": {"fill": _rgb(OKABE_ITO["vermillion"]), "soft": _tint(OKABE_ITO["vermillion"])},
    "purple":     {"fill": _rgb(OKABE_ITO["purple"]),     "soft": _tint(OKABE_ITO["purple"])},
    # 黄色本身很亮,描边用原色、填充调得更淡
    "yellow":     {"fill": _rgb(OKABE_ITO["yellow"]),     "soft": _tint(OKABE_ITO["yellow"], 0.90)},
    # 中性灰(Okabe-Ito 无灰阶,用于补充说明类的非主流程框)
    "gray":       {"fill": RGBColor(0x66, 0x66, 0x66),    "soft": RGBColor(0xE8, 0xE8, 0xE8)},
}

# ============================================================================
# 风格预设(theme)
#
# 配色永远是 Okabe-Ito,theme 只决定「怎么用色」:
#   fill  : solid=实心原色 / soft=原色向白混 82% 的浅色 / white=白底
#   text  : white=白字 / dark=近黑字 / accented=用该框自己的原色作字色
#   mono  : True 时忽略 style,一律黑框黑字(黑白印刷 / 不依赖颜色)
#   size / pad / line_w / arrow_w / label_size : 各项默认值,调用时可逐个覆盖
#
# size 是**最终印刷字号**,所以画布宽度要取论文版心,见 SKILL.md「画布尺寸」。
# ============================================================================
THEMES = {
    # 浅色底 + 同色系深字(默认):论文主图常用,层级靠色块区分但比实心柔和
    "tinted": dict(fill="soft", text="accented", bold=False,
                   line_w=1.25, size=8, pad=0.04, arrow_w=1.0,
                   label_size=7.5, italic=False, mono=False),
    # 答辩 / 汇报 PPT:实心色块 + 白粗体。画布通常 10 英寸、字号 13
    "presentation": dict(fill="solid", text="white", bold=True,
                         line_w=1.5, size=13, pad=0.10, arrow_w=1.75,
                         label_size=10.5, italic=True, mono=False),
    # 白底 + 细彩框 + 近黑常规字:最素净的期刊插图风
    "academic": dict(fill="white", text="dark", bold=False,
                     line_w=1.0, size=8, pad=0.04, arrow_w=1.0,
                     label_size=7.5, italic=False, mono=False),
    # 纯黑白:所有框统一黑描边,次级框用浅灰底区分。印刷不依赖颜色时用
    "mono": dict(fill="white", text="dark", bold=False,
                 line_w=1.0, size=8, pad=0.04, arrow_w=1.0,
                 label_size=7.5, italic=False, mono=True),
}

DARK = _rgb(OKABE_ITO["black"])       # 正文/箭头/描边
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x55, 0x55, 0x55)    # 次级文字(箭头标注、说明)

# 字体:拉丁 Arial,中文 微软雅黑(与 publication.mplstyle 的 sans-serif 一致)
FONT_LATIN = "Arial"
FONT_EA = "微软雅黑"

# 带箭头的线段最小可见杆长(英寸)。低于这个值,箭头会退化成"孤立的小三角" ——
# 只看得见箭尖看不见杆,读者无法判断它从哪儿指过来。见 SKILL.md「图形质量要求」。
MIN_SHAFT_IN = 0.14

# 总线定位用的两个偏好值。总线两侧未必都带箭头 —— bus() 里带箭头的是支线,
# fan_in() 里带箭头的是干线。默认位置要**保证带箭头那一段够长**,另一段随它短。
_BUS_NEAR = 0.12                  # 不带箭头那一段的理想长度
_BUS_HEAD = MIN_SHAFT_IN + 0.02   # 带箭头那一段的硬下限


def _bus_y(y_src, y_dst):
    """总线该放在哪个 y。

    默认取中点(两侧长度均等,最好看)。但 bus() 的支线和 fan_in() 的干线都是
    **靠目标那一侧**带箭头,所以一旦中点到目标的距离短于最小杆长,就得把总线
    往源那头挪 —— 宁可让不带箭头的一段短,也不能让箭头退化成孤立小三角。
    """
    return max(y_src + 0.02, min((y_src + y_dst) / 2, y_dst - _BUS_HEAD))

_OPPOSITE = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}


class Node:
    """box() / label() / group() 的返回值:记住自己画在哪儿的句柄,单位英寸。

    存在的唯一理由:**让箭头精确落在边框上**。

    手写 `d.arrow(x + 0.60, ...)` 时,0.60 是盒子半宽 —— 一旦改了盒子宽度,
    箭头就和边框脱开、或者插进框里半截。这是示意图里最常见的低级瑕疵,而且在
    缩略图上看不出来。改用 Node 取点,结构上就不可能算错:

        a = d.box(0.5, 0.5, 1.6, 0.5, "输入")
        b = d.box(3.0, 0.5, 1.6, 0.5, "输出")
        d.arrow(*a.right, *b.left)      # a 的右边框中点 → b 的左边框中点

    更省事的是 d.connect(a, b) —— 连边都不用选。
    """

    __slots__ = ("x", "y", "w", "h", "shape", "is_container")

    def __init__(self, x, y, w, h, shape=None, is_container=False):
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.shape = shape          # 底层 pptx 形状,需要精细操作时用
        self.is_container = is_container   # group() 画的容器框,audit 不查它的重叠

    # ---- 中心 ----
    @property
    def cx(self):
        return self.x + self.w / 2

    @property
    def cy(self):
        return self.y + self.h / 2

    # ---- 四条边的中点,返回 (x, y) 元组,可直接用 * 展开给 arrow() ----
    @property
    def top(self):
        return (self.cx, self.y)

    @property
    def bottom(self):
        return (self.cx, self.y + self.h)

    @property
    def left(self):
        return (self.x, self.cy)

    @property
    def right(self):
        return (self.x + self.w, self.cy)

    def port(self, side, t=0.5):
        """边上任意一点。

        side: 'top' / 'bottom' / 'left' / 'right'
        t   : 沿这条边的比例。 top/bottom 按 x 方向 0→1(0 是左端);
              left/right 按 y 方向 0→1(0 是上端)。
        """
        if side == "top":
            return (self.x + self.w * t, self.y)
        if side == "bottom":
            return (self.x + self.w * t, self.y + self.h)
        if side == "left":
            return (self.x, self.y + self.h * t)
        if side == "right":
            return (self.x + self.w, self.y + self.h * t)
        raise ValueError(f"side 只能是 top/bottom/left/right,收到 {side!r}")

    def __repr__(self):
        return (f"Node(x={self.x:.3f}, y={self.y:.3f}, "
                f"w={self.w:.3f}, h={self.h:.3f})")


class Diagram:
    """一张 16:9 画布,坐标单位统一用英寸,原点在左上角。

    theme: THEMES 里的键 —— tinted(默认) / academic / mono / presentation。
    配色永远是 Okabe-Ito,theme 只决定怎么用色(见 THEMES 注释)。

    ⚠️ width_in / height_in 的默认值只是给"随手试一下"用的,**正式画图一律
    显式传 width_in=** —— 画布宽度要等于这张图在论文里的最终印刷宽度(版心宽),
    这样字号才所见即所得。10 英寸画布插进版心 5.77 英寸的 Word 会被缩到
    0.577 倍,图里 8pt 落到纸面只剩 4.6pt,低于所有期刊下限。
    量版心和常用取值见 SKILL.md「画布尺寸」。
    """

    def __init__(self, width_in=10.0, height_in=4.2,
                 font_latin=FONT_LATIN, font_ea=FONT_EA, theme="tinted"):
        self.prs = Presentation()
        self.prs.slide_width = Inches(width_in)
        self.prs.slide_height = Inches(height_in)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])  # 空白版式
        self.w = width_in
        self.h = height_in
        self.font_latin = font_latin
        self.font_ea = font_ea
        self.theme = theme if theme in THEMES else "tinted"
        self.th = THEMES[self.theme]
        self._nodes = []            # 所有 box()/group() 产生的 Node,供 audit() 自查

    # ---------- 基本元素 ----------

    def box(self, x, y, w, h, text, style="blue", shape="rounded",
            size=None, bold=None, filled=True, align=PP_ALIGN.CENTER,
            line_w=None, bg=None, text_color=None, pad=None):
        """加一个带文字的盒子。style 取 STYLES 的键;filled=False 用浅色填充。

        返回 **Node 句柄**(不是底层形状)—— 拿它的 .top/.bottom/.left/.right
        去连箭头,别手算半宽。需要底层 pptx 形状时用 node.shape。

        size / bold / line_w / bg / text_color / pad 不传时按 theme 取默认。

        pad 是左右内边距。pptx 形状默认 0.1in,**小尺寸图里这一项很吃宽度** ——
        按版心画图时务必用默认的 0.04(tinted/academic/mono 已是),否则文字会被挤到换行。
        """
        th = self.th
        st = STYLES.get(style, STYLES["blue"])
        mono = th["mono"]
        edge = DARK if mono else st["fill"]          # 描边色
        soft = RGBColor(0xE8, 0xE8, 0xE8) if mono else st["soft"]
        size = th["size"] if size is None else size
        bold = th["bold"] if bold is None else bold
        line_w = th["line_w"] if line_w is None else line_w
        pad = th["pad"] if pad is None else pad
        if bg is None:
            bg = soft if not filled else {
                "solid": st["fill"], "soft": st["soft"], "white": WHITE}[th["fill"]]
        if text_color is None:
            if mono:
                text_color = DARK
            elif not filled:
                # 次级框(filled=False)是浅色底,字必须压深,不能用 theme 的文字色 ——
                # 否则 presentation 的白字会落在浅灰底上,直接看不见
                text_color = st["fill"] if th["fill"] in ("solid", "soft") else DARK
            else:
                text_color = {"white": WHITE, "dark": DARK,
                              "accented": st["fill"]}[th["text"]]
        kind = {"rounded": MSO_SHAPE.ROUNDED_RECTANGLE,
                "rect": MSO_SHAPE.RECTANGLE,
                "ellipse": MSO_SHAPE.OVAL,
                "diamond": MSO_SHAPE.DIAMOND,
                "parallelogram": MSO_SHAPE.PARALLELOGRAM,
                "cylinder": MSO_SHAPE.CAN}.get(shape, MSO_SHAPE.ROUNDED_RECTANGLE)
        shp = self.slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = bg
        shp.line.color.rgb = edge
        shp.line.width = Pt(line_w)
        shp.shadow.inherit = False
        _strip_style(shp)
        tf = shp.text_frame
        tf.margin_left = tf.margin_right = Inches(pad)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        self._set_text(shp, text, size, bold, text_color, align)
        node = Node(x, y, w, h, shp)
        self._nodes.append(node)
        return node

    def arrow(self, x1, y1, x2, y2, style="dark", width=None, dashed=False, head=True):
        """直线连接符,可带三角箭头。dashed=True 为虚线(常用于反馈/可选路径)。"""
        width = self.th["arrow_w"] if width is None else width
        conn = self.slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        color = DARK if (style == "dark" or self.th["mono"]) \
            else STYLES.get(style, STYLES["gray"])["fill"]
        conn.line.color.rgb = color
        conn.line.width = Pt(width)
        ln = conn.line._get_or_add_ln()
        if dashed:
            dash = ln.makeelement(qn("a:prstDash"), {"val": "dash"})
            ln.append(dash)
        if head:
            ln.append(ln.makeelement(qn("a:tailEnd"),
                                     {"type": "triangle", "w": "med", "len": "med"}))
        _strip_style(conn)
        return conn

    def elbow(self, x1, y1, x2, y2, style="dark", width=None, dashed=False):
        """肘形连接符(直角折线),适合绕行的关系线。

        ⚠️ 走线由 PowerPoint / LibreOffice 各自的算法决定,两边可能不一致。
        要精确控制直角走线,改用多段 arrow() 拼接。
        """
        width = self.th["arrow_w"] if width is None else width
        conn = self.slide.shapes.add_connector(
            MSO_CONNECTOR.ELBOW, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        color = DARK if (style == "dark" or self.th["mono"]) \
            else STYLES.get(style, STYLES["gray"])["fill"]
        conn.line.color.rgb = color
        conn.line.width = Pt(width)
        ln = conn.line._get_or_add_ln()
        if dashed:
            ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
        _strip_style(conn)
        return conn

    # ---------- 连接(箭头精确贴边框) ----------

    @staticmethod
    def _pick_side(a, b):
        """按两个盒子中心的相对位置选边:水平距离大就走左右,否则走上下。"""
        dx = b.cx - a.cx
        dy = b.cy - a.cy
        if abs(dx) >= abs(dy):
            return "right" if dx >= 0 else "left"
        return "bottom" if dy >= 0 else "top"

    def connect(self, a, b, side_a=None, side_b=None, style="dark",
                width=None, dashed=False, head=True):
        """从 a 的**边框**连到 b 的**边框** —— 两端都由 Node 几何算出。

        这是"箭头连边框"的根治手段:端点不可能是手算的,所以既不会脱框、
        也不会插进框里。自动选边规则见 _pick_side();要强制走某条边就传
        side_a / side_b('top'/'bottom'/'left'/'right')。

        返回箭头形状。注意返回的不是 Node —— 箭头本身不做为连线端点使用。
        """
        if side_a is None:
            side_a = self._pick_side(a, b)
        if side_b is None:
            side_b = _OPPOSITE[side_a]
        x1, y1 = a.port(side_a)
        x2, y2 = b.port(side_b)
        return self.arrow(x1, y1, x2, y2, style=style, width=width,
                          dashed=dashed, head=head)

    def bus(self, source, targets, axis="v", bus_at=None, style="dark",
            width=None, dashed=False):
        """树形分叉:源框 → 干线 → 一条总线 → 每个目标一条支线贴上边框。

        复刻参考图里那种"一个算子分发给多个并行步骤"的画法 —— 比各自画直线
        干净得多,也符合"正交走线、避免斜箭头"的要求。

        axis='v' : 源在上、目标在下,总线水平(最常用)
        axis='h' : 源在左、目标在右,总线竖直
        bus_at   : 总线所在坐标(英寸)。不传就取源边框与目标边框的中点 ——
                   中点能保证干线长度和支线长度大致相当,不会出现"孤立小三角"。

        返回本次画出的所有箭头(干线 + 总线 + 各支线),便于需要时再调整。
        """
        if not targets:
            return []
        made = []
        kw = dict(style=style, width=width, dashed=dashed)

        if axis == "v":
            y_src = source.bottom[1]
            y_dst = min(t.top[1] for t in targets)
            if bus_at is None:
                bus_at = _bus_y(y_src, y_dst)
            # 干线:源框下边框 → 总线(无箭头,它只是导线)
            made.append(self.arrow(source.bottom[0], y_src,
                                   source.bottom[0], bus_at, head=False, **kw))
            # 总线:横跨所有支线落点。必须把干线的接入点也算进去,
            # 否则源框不在目标跨度内时总线会短一截,干线悬空。
            xs = [t.cx for t in targets] + [source.bottom[0]]
            made.append(self.arrow(min(xs), bus_at, max(xs), bus_at,
                                   head=False, **kw))
            # 支线:总线 → 各目标的上边框中点
            for t in targets:
                made.append(self.arrow(t.cx, bus_at, *t.top, **kw))
        else:
            x_src = source.right[0]
            x_dst = min(t.left[0] for t in targets)
            if bus_at is None:
                bus_at = (x_src + x_dst) / 2
            made.append(self.arrow(x_src, source.right[1],
                                   bus_at, source.right[1], head=False, **kw))
            ys = [t.cy for t in targets] + [source.right[1]]
            made.append(self.arrow(bus_at, min(ys), bus_at, max(ys),
                                   head=False, **kw))
            for t in targets:
                made.append(self.arrow(bus_at, t.cy, *t.left, **kw))
        return made

    def fan_in(self, dest, sources, axis="v", bus_at=None, style="dark",
               width=None, dashed=False):
        """`bus()` 的镜像:多个来源汇入同一个目标。

        各来源引出一条支线 → 汇到一条总线 → 一条干线带箭头进入目标边框。
        和 bus() 配合,就是「分叉 → 并行处理 → 汇聚」三板斧。

        axis / bus_at 含义同 bus()。返回本次画出的所有箭头。
        """
        if not sources:
            return []
        made = []
        kw = dict(style=style, width=width, dashed=dashed)

        if axis == "v":
            y_src = max(s.bottom[1] for s in sources)
            y_dst = dest.top[1]
            if bus_at is None:
                bus_at = _bus_y(y_src, y_dst)
            xs = [s.cx for s in sources] + [dest.top[0]]
            for s in sources:
                made.append(self.arrow(s.cx, s.bottom[1], s.cx, bus_at,
                                       head=False, **kw))
            made.append(self.arrow(min(xs), bus_at, max(xs), bus_at,
                                   head=False, **kw))
            made.append(self.arrow(dest.top[0], bus_at, *dest.top, **kw))
        else:
            x_src = max(s.right[0] for s in sources)
            x_dst = dest.left[0]
            if bus_at is None:
                bus_at = (x_src + x_dst) / 2
            ys = [s.cy for s in sources] + [dest.left[1]]
            for s in sources:
                made.append(self.arrow(s.right[0], s.cy, bus_at, s.cy,
                                       head=False, **kw))
            made.append(self.arrow(bus_at, min(ys), bus_at, max(ys),
                                   head=False, **kw))
            made.append(self.arrow(bus_at, dest.left[1], *dest.left, **kw))
        return made

    def label(self, x, y, w, text, size=None, color=None, italic=None,
              align=PP_ALIGN.CENTER, bold=False):
        """纯文字标签(无边框),常放在箭头上标注意义。

        注意:标签会按 w 换行(_set_text 里统一设了 word_wrap),w 要留够。
        """
        size = self.th["label_size"] if size is None else size
        italic = self.th["italic"] if italic is None else italic
        h = 0.32
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        self._set_text(tb, text, size, bold, color or MUTED, align, italic=italic, frame=tf)
        return Node(x, y, w, h, tb)

    def caption(self, text, size=12, y=None):
        """图注,默认贴底部居中。"""
        y = self.h - 0.5 if y is None else y
        return self.label(0, y, self.w, text, size=size, color=DARK, italic=False)

    def title(self, text, size=15, y=0.22):
        """图内标题(多数期刊要求放在 caption 而非图内,按需取用)。"""
        return self.label(0, y, self.w, text, size=size, color=DARK, italic=False, bold=True)

    # ---------- 布局辅助 ----------

    def grid(self, n, start=0.5, end=None, gap=0.35):
        """把可用宽度均分成 n 列,返回每列的 (x, w)。用于等距排列盒子。"""
        end = (self.w - 0.5) if end is None else end
        total = end - start
        w = (total - gap * (n - 1)) / n
        return [(start + i * (w + gap), w) for i in range(n)]

    def rbox(self, col, y, h, text, total_cols, **kw):
        """按网格放盒子:col 从 0 开始。省去手算坐标。"""
        cells = self.grid(total_cols)
        x, w = cells[col]
        return self.box(x, y, w, h, text, **kw)

    def group(self, nodes, label=None, pad=0.14, label_size=None,
              style="gray"):
        """给一组盒子套一个虚线容器并加标签。

        用来表达"这几个东西属于同一类"(参考图里的「基学习器集合」「预测输出」)。
        容器**无填充**,只画虚线描边,并且会移到最底层 —— 否则会盖住里面的盒子。

        pad 是容器相对最外层盒子的四周留白。标签画在容器上方。
        """
        if not nodes:
            return None
        x0 = min(n.x for n in nodes) - pad
        y0 = min(n.y for n in nodes) - pad
        x1 = max(n.x + n.w for n in nodes) + pad
        y1 = max(n.y + n.h for n in nodes) + pad
        w, h = x1 - x0, y1 - y0

        st = STYLES.get(style, STYLES["gray"])
        edge = DARK if self.th["mono"] else st["fill"]
        shp = self.slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x0), Inches(y0),
            Inches(w), Inches(h))
        shp.fill.background()                    # 无填充,只留描边
        shp.line.color.rgb = edge
        shp.line.width = Pt(self.th["line_w"])
        shp.shadow.inherit = False
        _strip_style(shp)
        ln = shp.line._get_or_add_ln()
        # 用**点线**而不是虚线:虚线已经用来表示"反馈/可选路径"这类流程语义了。
        # 容器的边框是分组语义,两者样式必须能一眼分开,否则读者会把容器边
        # 当成一条流程线。
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "sysDot"}))
        self._send_to_back(shp)

        if label:
            self.label(x0, y0 - 0.26, w, label,
                       size=label_size or self.th["label_size"],
                       color=DARK, italic=False)
        node = Node(x0, y0, w, h, shp, is_container=True)
        self._nodes.append(node)
        return node

    # ---------- 自查 ----------

    def audit(self, min_shaft=MIN_SHAFT_IN):
        """检查版面里的常见低级瑕疵,返回问题描述列表(空列表 = 通过)。

        查三件事:
          1. **盒子重叠** —— 两个 box() 的矩形相交
          2. **箭头杆太短** —— 带箭头的连线短于 min_shaft,会退化成孤立的小三角
          3. **盒子出画布** —— 内容跑到画布外,导出时被裁掉

        audit() 看不出"箭头有没有连到边框" —— 用 connect()/bus()/port() 画的话,
        那是结构上就成立的,不需要检查;手算坐标画的话它也查不出来,只能靠肉眼。
        所以真正该做的是:**别手算坐标**。
        """
        issues = []
        # 容器框天生会"包住"里面的盒子,那不是缺陷,是它的职责 —— 排除掉
        boxes = [n for n in self._nodes
                 if n.w > 0 and n.h > 0 and not n.is_container]

        for i, a in enumerate(boxes):
            for b in boxes[i + 1:]:
                ox = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
                oy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
                if ox > 0.01 and oy > 0.01:
                    issues.append(f"盒子重叠 {ox:.3f}\" x {oy:.3f}\":"
                                  f"{a!r} 与 {b!r}")

        tol = 0.005
        for shp in self.slide.shapes:
            # 只看连接符(arrow/bus 画的线);文本框和自动图形没有 begin_x
            for attr in ("begin_x", "begin_y", "end_x", "end_y"):
                if not hasattr(shp, attr):
                    break
            else:
                w = abs(shp.end_x.inches - shp.begin_x.inches)
                h = abs(shp.end_y.inches - shp.begin_y.inches)
                # 只有带箭头的线段才受"最小杆长"约束 —— 干线和总线是没有箭头的
                # 导线,再短也无所谓,否则会误报
                ln = shp.line._get_or_add_ln()
                if ln.find(qn("a:tailEnd")) is not None and max(w, h) < min_shaft:
                    issues.append(
                        f"箭头太短 ({w:.3f}\", {h:.3f}\") —— 低于最小杆长 "
                        f"{min_shaft}\",箭头会退化成孤立的小三角,读者看不出它从哪来")
                if w > tol and h > tol:
                    issues.append(
                        f"斜箭头 ({w:.3f}\", {h:.3f}\") —— 走线应统一到水平或垂直,"
                        f"斜率不一致会让整张图显得乱")

        for n in boxes:
            if n.x < -0.001 or n.y < -0.001 \
                    or n.x + n.w > self.w + 0.001 or n.y + n.h > self.h + 0.001:
                issues.append(f"盒子出画布: {n!r}(画布 {self.w}\" x {self.h}\")")
        return issues

    def _send_to_back(self, shp):
        """把形状移到最底层(spTree 的前两个子元素是分组属性,内容从索引 2 开始)。"""
        spTree = self.slide.shapes._spTree
        el = shp._element
        spTree.remove(el)
        spTree.insert(2, el)

    # ---------- 内部 ----------

    def _add_run(self, p, text, size, bold, color, italic, script=None):
        """往段落里加一个 run。script 为 'sub' / 'sup' 时设为真下标 / 上标。"""
        if not text:
            return
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        rPr = r._r.get_or_add_rPr()
        # 真·上下标:OOXML 的 baseline 属性(千分之一百分比)。
        # 比用 Unicode 上下标字符可靠 —— 那种字符集不全,且缺字时变方框。
        if script == "sub":
            rPr.set("baseline", "-25000")
        elif script == "sup":
            rPr.set("baseline", "30000")
        # 拉丁字体走 python-pptx 的 API:它会把 <a:latin> 插到 schema 规定的位置。
        # 手动 append 会在 rPr 里留下**两个** a:latin,且次序不对 —— Word 能容错,
        # 严格校验器不能。
        r.font.name = self.font_latin
        # 中文字符必须显式指定 east-asian 字体,否则部分环境回落成宋体。
        # CT_TextCharacterProperties 的元素次序是 ... latin, ea, cs ...,
        # 所以 <a:ea> 要紧跟 <a:latin> 之后插,不能直接 append 到末尾。
        latin = rPr.find(qn("a:latin"))
        ea = rPr.makeelement(qn("a:ea"), {"typeface": self.font_ea})
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)

    def _add_rich(self, p, line, size, bold, color, italic):
        """解析行内的上下标写法,生成带真实上下标的 run 序列。

        上下标:  X_t   X_{t-1}   a^2   a^{n+1}
        转义:    \\_  \\^  —— 输出字面的下划线/脱字符,不做上下标

        注意 `_` 和 `^` 是**默认生效**的,所以文本里想写字面下划线(如
        `model_v2`)必须写成 `model\\_v2`,否则 `_v` 会被当成下标。
        只认 `\\_` 和 `\\^` 两个转义;反斜杠后跟别的字符原样保留。
        """
        pattern = re.compile(r"\\([_^])|([_^])(?:\{([^}]*)\}|(.))")
        pos = 0
        for m in pattern.finditer(line):
            if m.start() > pos:
                self._add_run(p, line[pos:m.start()], size, bold, color, italic)
            if m.group(1) is not None:
                # 转义分支:字面字符,不带 baseline
                self._add_run(p, m.group(1), size, bold, color, italic)
            else:
                script = "sub" if m.group(2) == "_" else "sup"
                content = m.group(3) if m.group(3) is not None else m.group(4)
                self._add_run(p, content, size, bold, color, italic, script)
            pos = m.end()
        if pos < len(line):
            self._add_run(p, line[pos:], size, bold, color, italic)

    def _set_text(self, shape, text, size, bold, color, align, italic=False, frame=None):
        tf = frame if frame is not None else shape.text_frame
        tf.word_wrap = True
        if frame is None:
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        # 多行文本必须逐行建段落。若把 "\n" 塞进单个 run,python-pptx 会生成垂直
        # 制表符,导致各行对齐不一致(第一行左对齐、后续行居中)。
        for i, line in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            self._add_rich(p, line, size, bold, color, italic)

    # ---------- 输出 ----------

    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.prs.save(path)
        return path
