---
name: ppt-diagram
description: 用代码生成**可编辑的学术示意图**(方法框架图、流程管线图、模块关系图、对比示意图),然后在 PowerPoint 里手动微调,最后导出成论文可用的高分辨率 PNG 或矢量 SVG。当用户要画论文配图、方法框架图、技术路线图、流程图、系统架构图、关系示意图,或说"帮我画个图""做个示意图""这张图要放进论文""导出成矢量图"时使用。触发词:画示意图、框架图、流程图、技术路线、论文配图、插图、矢量图、导出图片、PPT 画图。
---

# PPT 示意图工作流

**核心思路**:用代码生成骨架(精确对齐)→ 你在 PowerPoint 里拖拽微调 → 程序导出成图。

为什么用 PPT 而不是直接画?因为你**能改**。代码生成的位置不一定完全合你意,PPT 让你直接用鼠标调,调完再导。

---

## 三步工作流

```
① 我生成 PPTX  ──→  ② 你在 PowerPoint 里微调  ──→  ③ 我导出成图
   (代码，精确对齐)      (拖拽、改字、换色)          (600 DPI PNG / 矢量 SVG)
```

第 ③ 步你不能自己跑的话,把调好的 pptx 路径告诉我,我来导。

---

## ⚠️ 铁律:不要生成图注

**默认不调用 `caption()`,也不用 `title()`,画布上不出现「图 5-2  元学习器架构」这类编号+标题。**

原因:图注只要写进画布,导出时就会被一起烤进 PNG/SVG。插进论文后跟 Word 自己的题注段落**重复**,还得手动裁掉或重导 —— 白干一轮。

- 图的**编号和标题归 Word 题注管**(引用 → 插入题注),不归本 skill 管
- 需要说明这张图画的是什么,写在**给用户的回复正文**里,别写进画布
- `caption()` / `title()` 方法保留不删,但只在**用户明确要求「图注画进图里」**时才用
- 想在画布底部留白、让用户在 PPT 里自己加,是可以的 —— 留白 ≠ 写文字

---


## 画布尺寸:按最终印刷尺寸画

**画布宽度必须等于这张图在论文里的最终宽度(版心宽),不能拍脑袋用 10 英寸。**

理由:10 英寸的画布插进 Word 会被缩到 0.577 倍,图里的 11pt 落到纸面只剩 **6.3pt**,
低于所有期刊的下限(7-9pt,最小 6pt)。**画布 = 版心,字号才所见即所得。**

量论文版心(需要 `python-docx`):

```python
import docx
s = docx.Document(r"论文.docx").sections[0]
tw = s.page_width - s.left_margin - s.right_margin
print(tw / 914400, "英寸")      # A4 + 3.17cm 边距 → 5.77
```

| 纸张 | 左右页边距 | 版心宽 |
|---|---|---|
| A4 | 3.17cm(西北工大模板) | **5.77 英寸** |
| A4 | 2.5cm | 6.30 英寸 |
| Letter | 1 英寸 | 6.50 英寸 |

用户没给论文时,默认按 **5.77** 画,并在回复里说明这个假设。
画完必须自查:导出 PNG 的像素宽 ÷ DPI 应等于画布宽度。

---

## 依赖与安装

| 组件 | 用途 | 怎么装 |
|---|---|---|
| **Python 3.9+** + `python-pptx` | 生成 PPTX | `pip install python-pptx`(见 `requirements.txt`) |
| **LibreOffice** | PPTX → 矢量 PDF | Windows `winget install TheDocumentFoundation.LibreOffice` · macOS `brew install --cask libreoffice` · Linux `sudo apt install libreoffice` |
| **poppler / pdftocairo** | PDF → 高 DPI PNG / 矢量 SVG | Windows 下载 poppler 并把 `bin/` 加进 PATH · macOS `brew install poppler` · Linux `sudo apt install poppler-utils` |
| `Pillow`(可选) | 只有 `style_gallery.py` 拼对比图时用 | `pip install Pillow` |

本 skill 自己的文件:

| 文件 | 作用 |
|---|---|
| `scripts/diagram_kit.py` | 脚手架 —— 画图 API |
| `scripts/tools.py` | 找 soffice / pdftocairo 在哪(不用手改路径) |
| `scripts/export_figure.ps1` | 导出成图(Windows PowerShell) |
| `scripts/example_pipeline.py` | 完整可运行示例 |
| `scripts/style_gallery.py` | 一条命令出 4 种风格的拼版对比 |
| `export.cmd` | 拖拽导出入口(Windows) |

### 工具路径不用手改

`soffice` 和 `pdftocairo` 的安装位置因机器而异,所以**代码里不写死路径**。
查找顺序:

1. **环境变量** —— `PPT_DIAGRAM_SOFFICE` / `PPT_DIAGRAM_PDFTOCAIRO`
2. **PATH**
3. 几个常见的安装目录
4. **Windows 注册表** —— 仅用于找 LibreOffice 的安装路径

装在非标准位置就设环境变量指过去,别改代码:

```powershell
$env:PPT_DIAGRAM_SOFFICE = "D:\PortableApps\LibreOffice\program\soffice.com"
```

自检当前解析到了哪里:

```powershell
python "$HOME\.claude\skills\ppt-diagram\scripts\tools.py"
```

> ⚠️ **不要用 PowerPoint COM** 导出。`New-Object -ComObject PowerPoint.Application`
> 在受限/沙箱环境下常报 `80080005 CO_E_SERVER_EXEC_FAILURE`,而且会意外启动
> PowerPoint 并弹加载项错误框。**导出统一走 LibreOffice 无头模式** ——
> `export_figure.ps1` 已经这么做了。

> 📌 **导出链目前只在 Windows 上实测过。** `diagram_kit` 和 `tools.py` 是纯 Python,
> 跨平台没问题;`export_figure.ps1` / `export.cmd` 是 Windows 专用。macOS / Linux
> 用户可以直接用 `style_gallery.py`(它内部调 soffice + pdftocairo),或照它的
> `export_png()` 自己拼一条导出命令。

---

## 快速上手

**先读 `scripts/example_figure.py`** —— 那才是完整参考版式:五层主链 + 树形分叉 +
汇聚判据 + 虚线回路 + 分组容器,注释逐条对着「图形质量要求」。

下面是最小骨架,只演示怎么把箭头**连到边框上**:

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/ppt-diagram/scripts"))
from diagram_kit import Diagram

# 画布宽度取论文版心 —— 这样字号就是最终印刷字号,见「画布尺寸」
# 高度按内容算,别硬套:下面内容到 2.70 英寸,所以取 2.90
d = Diagram(width_in=5.77, height_in=2.90)          # theme 默认 tinted

# box() 返回 Node 句柄,后面靠它取边框坐标
a = d.box(0.16, 0.30, 2.60, 0.60, "y_{1:T}\n观测序列", style="sky")
b = d.box(3.01, 0.30, 2.60, 0.60, "x_0 ~ p(x_0)\n初始状态先验", style="sky")

# 两个输入汇入一个模型:各引一条竖直箭头落到上边框的不同位置。
# port() 的落点对准源框中线,箭头就完全竖直 —— 别随手取落点,会变成斜箭头
m = d.box(0.16, 1.30, 5.45, 0.55, "状态空间模型  x_{t+1} = f(x_t) + w_t",
          style="blue", size=7.5)
for src in (a, b):
    d.arrow(*src.bottom, *m.port("top", (src.cx - m.x) / m.w))

# 一个源分发给多个并行步骤:用 bus(),别各画一条线
steps = [d.box(0.16 + i * 1.86, 2.20, 1.60, 0.50,
               f"步骤 {i+1}\n具体算子", style="orange", size=7.5)
         for i in range(3)]
d.bus(m, steps)

# 担心箭头太短 / 盒子重叠 / 斜箭头?跑自查
for issue in d.audit():
    print("⚠️", issue)

# 不要加 d.caption(...) —— 见上面「铁律」
d.save("fig.pptx")
```

> 检查箭头有没有贴框,不用靠肉眼:端点坐标 == 对应 Node 的 `.top/.bottom/...` 即可。
> 用 `connect()` / `bus()` / `port()` 画的话,这是结构上成立的。

跑完把 pptx 交给用户微调。完整示例:`python scripts/example_pipeline.py`。

---

## 导出

```powershell
# 默认导出 600 DPI PNG + 矢量 SVG
& "$HOME\.claude\skills\ppt-diagram\scripts\export_figure.ps1" -Pptx ".\fig1.pptx"

# 只要高 DPI PNG,300 DPI
.\export_figure.ps1 -Pptx .\fig1.pptx -Formats png -Dpi 300

# 全都要
.\export_figure.ps1 -Pptx .\fig1.pptx -Formats png,svg,pdf -Dpi 600 -OutDir .\out

# 只导第 2 页(默认 -Page 0 = 全部页)
.\export_figure.ps1 -Pptx .\fig1.pptx -Page 2
```

### 拖拽导出(不用敲命令)

**把任何 `.pptx` 拖到 `export.cmd` 的图标上松手**,就会在 pptx 同目录自动生成图片。

- 多页 PPT → 每页一张(`fig-1.png`、`fig-2.png`…)
- 单页 PPT → 不带编号(`fig.png`)
- 默认 600 DPI PNG + 矢量 SVG
- 可以一次拖多个文件

**这个脚本跟 PPT 内容完全无关** —— 它只吃文件路径,任何 pptx 都能转(不止本 skill 生成的)。

> **维护注意:`.cmd` 必须保持纯 ASCII + CRLF 换行。** 批处理是逐字节解析的,
> 换成非 ASCII 就得保证编码和代码页一致(GBK 存、`chcp 936` 读),而任何人用
> 默认存 UTF-8 的编辑器打开再保存,就会让 `if "%~1"=="" (` 这类多行块被拆散、
> 脚本直接崩。纯 ASCII 让这一整类问题消失。`.gitattributes` 里已声明
> `*.cmd text eol=crlf`,保证任何客户端检出都是 CRLF。
>
> 另外 `export.cmd` 里 `chcp` / `powershell` 都写成绝对路径 —— 有些机器
> (比如装了精简 PATH 的)**PATH 里没有 `C:\Windows\System32`**,按名字调会找不到。

**实测产出**(10 英寸宽画布 @ 600 DPI):`6000 x 2520 px`。

### 导出格式怎么选

| 格式 | 矢量? | 能插 Word? | 何时用 |
|---|---|---|---|
| **SVG** | ✅ | ✅ Word 2016+ | **首选**。无限放大不糊,还能右键"转换为形状"编辑 |
| **PNG** | ❌ | ✅ 任何版本 | 期刊要求位图时用;600 DPI 足够印刷 |
| PDF | ✅ | ❌ | 作为中间产物,Ctrl+S 拿走投期刊 |
| EMF | ✅ | ✅ 原生 | 需要 Inkscape 之类的工具才能转,本 skill 不提供;SVG 已能满足需求 |

---

## diagram_kit API

坐标单位统一**英寸**,原点左上角。

| 方法 | 说明 |
|---|---|
| `Diagram(width_in, height_in, theme, font)` | 建画布。**宽度取论文版心**。代码默认 10 英寸只是为方便随手试用,**正式画图必须显式传 `width_in=`**(见「画布尺寸」)。theme 见「风格预设」:presentation / tinted / academic / mono |
| `box(x, y, w, h, text, style, shape, size, bold, filled, align, line_w, bg, text_color, pad)` | 带文字的盒子。后六个不传就按 theme 取默认。**返回 Node 句柄** |
| `arrow(x1, y1, x2, y2, style, width, dashed, head)` | 画一条给定坐标的直线箭头。**能用 Node 取点就别手算** |
| `connect(a, b, side_a, side_b, ...)` | **从 a 的边框连到 b 的边框**,自动选边。箭头贴框就靠它 |
| `bus(src, targets, axis, bus_at, ...)` | 树形分叉:干线 + 总线 + 逐条支线贴上各目标边框 |
| `fan_in(dest, sources, axis, bus_at, ...)` | `bus()` 的镜像:多来源汇聚进一个目标 |
| `group(nodes, label, pad, style)` | 给一组盒子套**点线**容器(分组语义),可加标签 |
| `audit()` | 版面自查:盒子重叠、箭头杆过短、斜箭头、内容出画布。**画完跑一遍** |
| `elbow(x1, y1, x2, y2, ...)` | 肘形折线。⚠️ 走线由渲染器决定,两边可能不一致 —— 优先用多段 `arrow()` 拼 |
| `label(x, y, w, text, size, color, italic, align, bold)` | 纯文字标签。返回 Node |
| `caption(text, size, y)` | 图注(默认贴底居中)。**默认不用**,见上面铁律 |
| `title(text, size, y)` | 图内标题。**默认不用**,见上面铁律 |
| `grid(n, start, end, gap)` | 把宽度均分 n 列,返回 `[(x, w), ...]` |
| `rbox(col, y, h, text, total_cols, **kw)` | 按网格放盒子(省去手算坐标) |
| `save(path)` | 存文件 |

### Node:别手算坐标

`box()` / `label()` / `group()` 返回的都是 **Node 句柄**,记住自己画在哪儿(英寸):

| 属性 | 含义 |
|---|---|
| `node.x` `.y` `.w` `.h` | 原始几何 |
| `node.cx` `.cy` | 中心 |
| `node.top` `.bottom` `.left` `.right` | **四条边的中点**,返回 `(x, y)` |
| `node.port(side, t=0.5)` | 边上任意一点,`t` 是沿边比例 |
| `node.shape` | 底层 pptx 形状(要精细操作时用) |

```python
obs = d.box(0.55, 0.12, 2.39, 0.38, "y_{1:T}\n观测序列")
model = d.box(0.55, 0.66, 5.06, 0.42, "状态空间模型")
d.arrow(*obs.bottom, *model.port("top", 0.24))   # 落在模型上边框的 24% 处

b1 = d.box(...); b2 = d.box(...)
d.connect(b1, b2)                                 # 连边都不用选
```

**为什么要这样**:手写 `d.arrow(x + 0.60, ...)` 时,`0.60` 是盒子半宽 ——
改动盒子尺寸后箭头就和边框脱开,或者插进框里半截。这种瑕疵在缩略图上看不出来,
印出来才发现。用 Node 取点,结构上就不可能算错。

> `port()` 的落点若取在**正对源框中线**的位置,箭头就是完全竖直的;
> 随便取个落点会变成斜箭头,而斜箭头是"看起来乱"的主要来源。

调用 `connect()` / `bus()` / `fan_in()` 时,自动选边和总线位置都有默认规则;
不满意可以传 `side_a` / `side_b` / `bus_at` 覆盖。

**`style`** 取值:`blue` / `orange` / `green` / `sky` / `vermillion` / `purple` / `yellow` / `gray`
**`shape`** 取值:`rounded`(默认) / `rect` / `ellipse` / `diamond` / `parallelogram` / `cylinder`
**`filled=False`** 用浅色填充(适合放约束条件、补充说明这类非主流程的框)

---

## 公式与上下标

文本里直接写 LaTeX 风格的上下标,`diagram_kit` 会渲染成**真·上下标** —— 用的是 OOXML 的 `baseline` 属性,**不是 Unicode 上下标字符**(那种字符集不全,缺字会变方框,且没法在 PowerPoint 里编辑)。

| 写法 | 效果 |
|---|---|
| `C_t` | C 带下标 t |
| `H_{t-1}` | H 带下标 t-1 |
| `a^2` | a 带上标 2 |
| `x^{n+1}` | x 带上标 n+1 |
| `model\_v2` | 字面下划线(转义)—— 渲染成 `model_v2`,不是下标 |

```python
d.box(0.6, 3.35, 5.1, 0.7, "C_t  =  f_t ⊙ C_{t-1}  +  i_t ⊙ g_t", style="green")
```

**注意**:
- ⚠️ **`_` 和 `^` 是默认生效的**,不需要任何标记。所以文本里想写字面下划线
  (变量名 `model_v2`、文件名 `train_loop.py`)必须**转义成 `\_`**,
  否则 `_v` 会被当成下标。脱字符同理用 `\^`。
  只认这两个转义,反斜杠后跟别的字符原样保留(路径 `C:\Users` 不受影响)
- 多字符下标**必须加花括号**:`H_{t-1}` ✓,写成 `H_t-1` 只会把 `t` 变下标
- 这些在 PowerPoint 里是普通文本 run,可以继续编辑改字号;**不是公式对象**(OMML)。
  真 OMML 公式对象虽然更"正规",但 LibreOffice 无头转换时渲染不可靠,会毁掉导出链路 —— 所以这里刻意不用
- 数学符号直接用 Unicode:`⊙`(U+2299 逐元素积)、`σ`(U+03C3)、`*`(卷积)、`⊙` 已验证在微软雅黑下正常显示
- **上下标不缩字号**:`baseline` 只移基线。`w_1` 这种单字符没问题;`ŷ^{(ensemble)}` 这类长上标会明显偏大,尽量控制在 1~2 字符

---

## 配色与字体(已按出版规范固化)

**这两个不用你操心,`diagram_kit` 已经写死,每次生成自动套用。**

### 配色:Okabe-Ito 色盲友好色板

来源:Okabe, M. & Ito, K. (2008). *Color Universal Design (CUD): How to Make
Figures and Presentations That Are Friendly to Colorblind People.*
<https://jfly.uni-koeln.de/color/> —— 科学出版最广泛推荐的色盲友好配色。
约 8% 的男性有色觉障碍,用这套能保证他们也能区分。

| style 名 | 色值 | | style 名 | 色值 |
|---|---|---|---|---|
| `blue` | `#0072B2` | | `vermillion` | `#D55E00` |
| `orange` | `#E69F00` | | `purple` | `#CC79A7` |
| `green` | `#009E73` | | `yellow` | `#F0E442` |
| `sky` | `#56B4E9` | | `gray` | `#666666` |

`filled=False` 时用的浅色填充,是从上表原色**自动向白色混合 82%** 派生的,所以永远和描边同色系。

**用色纪律**:
- **一张图不超过 3 种主色**;同层级/同类型的元素用同一种颜色
- 要强调的那一个用 `vermillion` 或 `orange`(暖色前进感),其余用 `blue`/`sky`/`gray`
- 非主流程的补充框(约束条件、说明)用 `gray` + `filled=False`
- `filled=False`(次级框)的文字色**不跟 theme 走**:实心/浅底主题用同色系深色,白底主题用近黑。否则 presentation 的白字会落在浅灰底上,直接看不见
- **不要自己引入色板外的颜色**。要加,从 `OKABE_ITO` 字典里取

### 字体

| 用途 | 字体 |
|---|---|
| 拉丁字符(英文、数字) | **Arial** |
| 中文 | **微软雅黑** |

与 `publication.mplstyle` 里 `font.sans-serif: Arial, Helvetica, DejaVu Sans` 保持一致。中文必须显式写 east-asian 字体名,否则部分环境会回落成宋体 —— `diagram_kit._set_text` 已处理。

> 自检:导出后若看到中文是宋体或数字变衬线,说明字体没生效,检查 `rPr` 里是否同时有 `a:latin` 和 `a:ea`。

### 风格预设(theme)

`Diagram(..., theme=...)` 四选一。**配色永远是 Okabe-Ito,theme 只决定怎么用色**;
几何(画布尺寸、字号、内边距)不由 theme 决定,那是调用时传参控制的。

| theme | 填充 | 文字 | 字重 | 适合 | 默认字号 |
|---|---|---|---|---|---|
| `presentation` | 实心饱和色 | 白色 | 粗 | 答辩 / 汇报 PPT(画布 10in) | 13 |
| **`tinted`** (默认) | 原色向白混 82% | 同色系深色 | 常规 | 论文主图,层级靠色块区分 | 8 |
| `academic` | 白 | 近黑 | 常规 | 最素净的期刊插图风 | 8 |
| `mono` | 白(次级框浅灰) | 黑 | 常规 | 黑白印刷 / 不依赖颜色 | 8 |

后三个是**按印刷尺寸**设的默认(8pt / 0.04in 内边距 / 1pt 描边 / 1pt 箭头),
`presentation` 是按 PPT 尺寸设的。`box()` 的 `size` / `bold` / `line_w` / `pad` /
`bg` / `text_color` 都能逐个覆盖 theme 默认。

想直接看效果,跑下面这条 —— 同一版式渲染 4 种风格并拼成一张对比图:

```powershell
python "$HOME\.claude\skills\ppt-diagram\scripts\style_gallery.py"
```

**怎么选**:
- 不确定选哪个 → 默认 `tinted`(层级清楚又不像 PPT)
- 追求最素净、最像期刊插图 → `academic`
- 论文要黑白印刷,或审稿要求「不依赖颜色」 → `mono`
- 同一张图还要上答辩 PPT → `presentation`(画布改 10in、字号 13)

**对比风格时几何要固定**:同一版式只换 theme、画布与字号不变,否则比的是尺寸不是风格。

---

## 图形质量要求

这一节是**验收标准**。下面的要求来自期刊插图规范与学术示意图实践(来源见文末),
不是审美偏好 —— 每一条都有具体理由,违反哪条都会让图"看起来业余"。

### 箭头

| 要求 | 为什么 |
|---|---|
| **连到边框**,不脱框、不插进框里 | 悬空的箭头读者不知道它从哪来;插进框里则像把框划开了。用 `connect()` / `bus()` / `port()` 画,这条自动成立 |
| 带箭头的线段要有**最小杆长**(约 0.14 英寸) | 否则只看得见箭尖看不见杆,退化成"孤立的小三角"。`audit()` 会替你查 |
| 相连两框之间要留够间距(≥ 0.2 英寸) | 太窄就挪框,别硬塞 —— 挤出来的箭头必然短 |
| **走正交**,避免斜箭头 | 斜率各不相同的斜线是"图显乱"的头号来源。`audit()` 会查 |
| **一条箭头只表达一种关系** | 同一根线一会儿表示"导致"一会儿表示"传给",读者无法解码 |
| 不用**双向箭头** | 语义含糊;要互指就画两条 |
| 全图箭头**样式统一** | 实线=主流程、虚线=反馈/可选路径。别出现第三种含义的实线 |
| 标注放在**线旁,不压线** | 文字压在线上既难读又难看 |

### 线宽与形状

- **只有两档线宽**:**框线粗、连接线细**(细线约为粗线的 1/2)。
  `diagram_kit` 的 theme 已按此设好,一般不用动
- 避免过粗的描边和过大的箭头头部
- **不用投影、不用渐变**。`_strip_style()` 已经帮你删掉主题投影 —— 自己写
  代码时注意别把它们加回来
- **同形同意**:同一形状/颜色始终代表同一类概念。判断框用菱形,处理用圆角矩形

### 布局

- 自上而下 或 从左到右,顺着读图视线
- **疏密适中,不留大空白**;高宽比协调
- **同层级元素等宽等距** —— 用 `grid()` 或手算固定间距,别凭感觉摆
- 各框大小尽量一致,并且**装得下里面的字**
- 避免"糖葫芦"式一条直线排到底;3~5 个阶段以上就分层或分叉
- 一幅图只讲一个主题;元素别堆太多

### 文字与内容(「内容充实」的实质)

- **框内写具体内容**:公式、张量名、算子、具体方法名。
  写"数据输入与预处理"这种泛称,等于什么都没说 —— 换成
  `y_{1:T} 观测序列`、`θ ← argmax Q(θ)` 才叫图里有信息
- 同图**字号统一**;缩放后仍可读(见「画布尺寸」)
- 能用正文/图注说明的别塞进图里,图中文字以少为好
- 中英文与公式混排时,外文字母的正斜体、大小写与正文保持一致

### 自查

画完跑一遍 `d.audit()`,它会报出**盒子重叠、箭头杆过短、斜箭头、内容出画布**。

但 `audit()` **查不出"箭头有没有连到边框"** —— 用 `connect()` 画的话那是结构上
成立的,不需要查;手算坐标画的话它也查不出来。所以真正该做的是:**别手算坐标**。

```python
issues = d.audit()
for i in issues:
    print("⚠️", i)
```

> 调研来源:
> [arXiv 图形生成规范(箭头-边框连接规则)](https://arxiv.org/pdf/2609.01006) ·
> [AlterLab 科学示意图 best practices](https://github.com/alterlab-ieu/alterlab-academic-skills) ·
> [科技论文插图的构思设计及要求](https://jdxb.bjtu.edu.cn/CN/PDF/380)

---

## 版面建议

实测好用的默认值:

- **画布宽度 = 论文版心** —— 最重要的一条,见上一节。默认 **5.77 英寸**(A4+3.17cm)
- **字号 7~9pt** —— 期刊硬性区间。正文框 8pt、箭头标注 7.5pt 实测好用,**绝不低于 6pt**
- **盒子高 0.5~0.55 英寸** —— 8pt 撑三行不溢出
- **左右内边距 0.04 英寸** —— tinted/academic/mono 已默认(pptx 默认是 0.1,小图上会把字挤到换行)
- **字重常规** —— 粗体只留给面板标号(A/B/C);「实心饱和色 + 白色粗体」是 PPT 汇报风,不是期刊风
- **主流程 3~5 个阶段** —— 再多就分两行或分层
- **配色纪律见上一节**(不超过 3 种主色,同层同色)
- **画布底部留白 ~0.4 英寸** —— 不写图注(铁律),但留出空间让用户在 PPT 里自己加
- **虚线只用于反馈/可选路径** —— 与主流程区分

---

## 坑

| 现象 | 原因与解法 |
|---|---|
| 中文变成宋体/乱码 | 中文字体需显式设 east-asian。`diagram_kit` 的 `_set_text` 已处理;自己写代码时要设 `rPr` 的 `a:ea` 值 |
| 文字溢出盒子 | 减字号或加盒高。导出后务必开 PNG 看一眼,别只看 PPT |
| 形状重叠 | 用 `grid()` 算坐标,别手拍数字 |
| LibreOffice 报 `Could not find platform independent libraries` | **无害警告**,转换照常成功,忽略 |
| 导出是 144 DPI 而非 600 | 那是 LibreOffice 直接导出的 PNG。**必须走 pdftocairo**(见 export_figure.ps1),分辨率参数才生效 |
| `soffice` 转换卡住 | 加 `-env:UserInstallation=file:///...` 隔离配置(脚本已含) |
| **导出图上有投影** | `<p:style>` 里的 `a:effectRef` 引用了主题投影效果。代码里 `shadow.inherit = False` 只生成空 `<a:effectLst/>`,**PowerPoint 认这个覆盖、LibreOffice 不认**,所以两边不一致。`_strip_style()` 已整块删掉 `<p:style>`,自己写代码要注意 |
| **文字莫名换行** | pptx 形状默认左右内边距各 0.1 英寸。版心尺寸下光内边距就吃掉 0.2 英寸 ≈ 3 个中文字的宽度。tinted/academic/mono 已收到 0.04 |
| **插进 Word 字太小** | 画布没跟最终印刷尺寸对齐。见「画布尺寸」一节 |
| **上下标看起来偏大** | `baseline` 只移基线、**不缩字号**,所以 `ŷ^{(ensemble)}` 这种长上下标会顶出一大块。上下标控制在 1~2 字符 |
| **变量名莫名变下标** | `_` / `^` 默认就是上下标语法,`model_v2` 会被渲染成 model 带下标 v 再接 2。写字面下划线要转义成 `model\_v2`,脱字符用 `\^` |
| **找不到 soffice / pdftocairo** | 按 环境变量 → PATH → 常见目录 的顺序找。装在别处就设 `PPT_DIAGRAM_SOFFICE` / `PPT_DIAGRAM_PDFTOCAIRO`,别改代码。跑 `python scripts/tools.py` 看当前解析到了哪里 |
