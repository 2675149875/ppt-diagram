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

```python
import sys, os
sys.path.insert(0, os.path.expanduser("~/.claude/skills/ppt-diagram/scripts"))
from diagram_kit import Diagram

# theme 见「风格预设」:tinted(默认)/ academic / mono / presentation
# 画布宽度取论文版心 —— 这样字号就是最终印刷字号,见「画布尺寸」一节
d = Diagram(width_in=5.77, height_in=2.60)          # theme 默认 tinted

# 三阶段,网格自动等距排列
Y, H = 0.45, 0.55
d.rbox(0, Y, H, "数据输入\n与预处理", 3, style="blue")
d.rbox(1, Y, H, "模型建立\n与求解",   3, style="green")
d.rbox(2, Y, H, "结果分析\n与验证",   3, style="blue")

# 箭头 + 下方标注
cells = d.grid(3, start=0.15, end=5.62, gap=0.45)   # 间隙要留够,见下
for i in range(2):
    x1 = cells[i][0] + cells[i][1]
    x2 = cells[i+1][0]
    d.arrow(x1, Y + H/2, x2, Y + H/2)
    d.label((x1+x2)/2 - 0.225, Y + H/2 + 0.05, 0.45,
            "特征提取" if i == 0 else "误差评估", size=7)

# 虚线反馈回路 —— 用多段 arrow 拼直角,别用 elbow(走线不可控)
bx = cells[2][0] + cells[2][1] - 0.2
d.arrow(bx, Y + H, bx, Y + H + 0.35, dashed=True, head=False)
d.arrow(bx, Y + H + 0.35, cells[1][0] + cells[1][1]/2, Y + H + 0.35, dashed=True)

# 不要加 d.caption(...) —— 见上面「铁律」
```

跑完把 pptx 交给用户微调。完整可运行示例见 `scripts/example_pipeline.py`。

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
| `box(x, y, w, h, text, style, shape, size, bold, filled, align, line_w, bg, text_color, pad)` | 带文字的盒子。后六个不传就按 theme 取默认 |
| `arrow(x1, y1, x2, y2, style, width, dashed, head)` | 直线箭头 |
| `elbow(x1, y1, x2, y2, ...)` | 肘形折线(绕行关系线) |
| `label(x, y, w, text, size, color, italic, align, bold)` | 纯文字标签 |
| `caption(text, size, y)` | 图注(默认贴底居中)。**默认不用**,见上面铁律 |
| `title(text, size, y)` | 图内标题。**默认不用**,见上面铁律 |
| `grid(n, start, end, gap)` | 把宽度均分 n 列,返回 `[(x, w), ...]` |
| `rbox(col, y, h, text, total_cols, **kw)` | 按网格放盒子(省去手算坐标) |
| `save(path)` | 存文件 |

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

来源:`scientific-toolkit-skill` 的 `scientific-visualization` 模块(`assets/color_palettes.py`)。这是 Okabe & Ito (2008) 提出的、科学出版最广泛推荐的色盲友好配色 —— 约 8% 的男性有色觉障碍,用这套能保证他们也能区分。

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
