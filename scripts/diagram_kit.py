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

    # ---------- 基本元素 ----------

    def box(self, x, y, w, h, text, style="blue", shape="rounded",
            size=None, bold=None, filled=True, align=PP_ALIGN.CENTER,
            line_w=None, bg=None, text_color=None, pad=None):
        """加一个带文字的盒子。style 取 STYLES 的键;filled=False 用浅色填充。

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
        return shp

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

    def label(self, x, y, w, text, size=None, color=None, italic=None,
              align=PP_ALIGN.CENTER, bold=False):
        """纯文字标签(无边框),常放在箭头上标注意义。

        注意:标签会按 w 换行(_set_text 里统一设了 word_wrap),w 要留够。
        """
        size = self.th["label_size"] if size is None else size
        italic = self.th["italic"] if italic is None else italic
        tb = self.slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.32))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        self._set_text(tb, text, size, bold, color or MUTED, align, italic=italic, frame=tf)
        return tb

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
