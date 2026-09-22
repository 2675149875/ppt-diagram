<#
.SYNOPSIS
    把 PPTX 示意图导出成论文可用的图片。

.DESCRIPTION
    流水线:PPTX --(LibreOffice)--> PDF(矢量) --(pdftocairo)--> 高 DPI PNG / SVG(矢量)

    为什么要绕道 PDF:PowerPoint COM 在受限/沙箱环境下常起不来
    (CO_E_SERVER_EXEC_FAILURE),而 LibreOffice 命令行无头模式稳定可靠。

.PARAMETER Pptx
    输入 .pptx 路径(必填)。

.PARAMETER OutDir
    输出目录。默认与输入的 pptx 同目录。

.PARAMETER Formats
    要产出的格式,可多选:png / svg / pdf。默认 png, svg。

.PARAMETER Dpi
    位图分辨率,默认 600(期刊通常要求 >=300)。

.PARAMETER Page
    导出第几页(从 1 开始)。默认 0 = 导出全部页;PPT 只有一页时不受影响。
    多页时文件名带编号(如 fig-1.png、fig-2.png);只有一页时不带编号。

.PARAMETER BaseName
    输出文件名(不含扩展名)。默认沿用 pptx 文件名。

.EXAMPLE
    .\export_figure.ps1 -Pptx .\fig1.pptx
    导出 fig1.png (600 DPI) 和 fig1.svg 到同目录。

.EXAMPLE
    .\export_figure.ps1 -Pptx .\fig1.pptx -Formats png,svg,pdf -Dpi 300 -OutDir .\out

.NOTES
    依赖 LibreOffice(转 PDF)与 poppler 的 pdftocairo(PDF 转图),缺任何一个
    都导不出图。两者的查找顺序都是:环境变量 → PATH → 常见安装目录

        PPT_DIAGRAM_SOFFICE       soffice 可执行文件
        PPT_DIAGRAM_PDFTOCAIRO    pdftocairo 可执行文件

    scripts/tools.py 认同一组环境变量,两边保持一致。

    本脚本刻意不 import tools.py —— 导出链不该为了找路径反过来依赖 Python。

   ⚠️ 不要并发跑。LibreOffice 无头实例共用同一个 UserInstallation 配置目录,
   同时起两个会互相踢掉,表现为"未能生成 PDF"但没有任何报错输出。

   -BaseName 只影响**输出**文件名。LibreOffice 本身总是按源文件名产出 PDF,
   脚本内部会按源名找到产物再改名,所以传任意 BaseName 都可以。

    编码要求:.ps1 必须存成 **UTF-8 with BOM**。PowerShell 5.1 对无 BOM 的
    .ps1 会按系统 ANSI 代码页读,本文件里的中文提示会全部乱码。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [string]$OutDir,
    [string[]]$Formats = @('png', 'svg'),
    [int]$Dpi = 600,
    [int]$Page = 0,
    [string]$BaseName
)

$ErrorActionPreference = 'Stop'

function Find-Tool {
    param([string[]]$Candidates, [string]$Name, [string]$EnvVar)

    # 1) 环境变量优先 —— 这是显式意图,指错了要报出来,不能静默回退到别处
    if ($EnvVar) {
        $override = [Environment]::GetEnvironmentVariable($EnvVar)
        if ($override) {
            $override = [Environment]::ExpandEnvironmentVariables($override)
            if (Test-Path -LiteralPath $override) { return $override }
            throw "环境变量 $EnvVar 指向的文件不存在: $override  -- 请改成实际路径,或清空该变量改用自动查找。"
        }
    }
    # 2) PATH
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    # 3) 候选目录
    foreach ($c in $Candidates) {
        if (-not $c) { continue }
        $expanded = [Environment]::ExpandEnvironmentVariables($c)
        if (Test-Path -LiteralPath $expanded) { return $expanded }
    }
    return $null
}

# ---- 定位依赖 ----
# 装在非标准位置时,设 PPT_DIAGRAM_SOFFICE / PPT_DIAGRAM_PDFTOCAIRO 指过去,
# 不用改这个脚本。各 skill 的 scripts/tools.py 认同一组环境变量。
$soffice = Find-Tool -Name 'soffice' -EnvVar 'PPT_DIAGRAM_SOFFICE' -Candidates @(
    "$env:ProgramFiles\LibreOffice\program\soffice.com",
    "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.com",
    "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.com",
    "C:\Program Files\LibreOffice\program\soffice.exe"
)

$pdftocairo = Find-Tool -Name 'pdftocairo' -EnvVar 'PPT_DIAGRAM_PDFTOCAIRO' -Candidates @(
    "$env:USERPROFILE\.claude\skills\shared\tools\poppler\pdftocairo.exe",
    "$env:LOCALAPPDATA\poppler\pdftocairo.exe",
    "$env:ProgramFiles\poppler\bin\pdftocairo.exe",
    "$env:ProgramData\chocolatey\bin\pdftocairo.exe"
)

if (-not $soffice) {
    throw @"
找不到 LibreOffice (soffice) —— 本脚本靠它把 PPTX 转成矢量 PDF。

  Windows : winget install TheDocumentFoundation.LibreOffice
            或到 https://www.libreoffice.org/download/ 下载
  macOS   : brew install --cask libreoffice
  Linux   : sudo apt install libreoffice

已装在非标准位置时,设环境变量 PPT_DIAGRAM_SOFFICE 指向 soffice 可执行文件。
"@
}

$Pptx = (Resolve-Path -LiteralPath $Pptx).Path
if (-not $OutDir) { $OutDir = Split-Path -Parent $Pptx }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path
if (-not $BaseName) { $BaseName = [IO.Path]::GetFileNameWithoutExtension($Pptx) }

Write-Host "输入   : $Pptx"
Write-Host "输出   : $OutDir\$BaseName.*"
Write-Host "LibreOffice : $soffice"
if ($pdftocairo) { Write-Host "pdftocairo  : $pdftocairo" }
Write-Host ""

# ---- 阶段 1:PPTX -> PDF(矢量) ----
$profile = "file:///" + (($env:TEMP + "\lo_export_profile") -replace '\\', '/')

# LibreOffice **总是按源文件名**输出 PDF,它不认识 -BaseName。
# 所以先按源名找产物,再改名成 $BaseName.pdf —— 直接去找 $BaseName.pdf 是
# 找不到的(-BaseName 一旦显式传值就会报"未能生成 PDF")。
$srcBase = [IO.Path]::GetFileNameWithoutExtension($Pptx)
$loPdf   = Join-Path $OutDir ($srcBase + ".pdf")
$pdfPath = Join-Path $OutDir ($BaseName + ".pdf")

$soArgs = @('--headless', '--norestore', "-env:UserInstallation=$profile",
            '--convert-to', 'pdf', '--outdir', $OutDir, $Pptx)
# LibreOffice 会往 stderr 打无害的 "platform independent libraries" 警告,丢弃它;
# 转换是否成功以下面的结果文件检查为准。
$ErrorActionPreference = 'Continue'
& $soffice @soArgs 2>$null | ForEach-Object { Write-Host "  $_" }
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $loPdf)) {
    # LibreOffice 偶尔输出到别处,兜底找一下
    $found = Get-ChildItem -LiteralPath $OutDir -Filter ($srcBase + ".pdf") -File -ErrorAction SilentlyContinue
    if ($found) {
        $loPdf = $found.FullName
    } else {
        throw @"
LibreOffice 未能生成 PDF,导出失败。

可能的原因:
  1. **另一个导出正在跑** —— LibreOffice 无头实例共用同一个 UserInstallation
     配置目录(见脚本里的 -env:UserInstallation),同时起两个会互相踢掉。
     等前一个跑完再试,别并发。
  2. 输入 pptx 打不开或损坏 —— 先用 PowerPoint / LibreOffice 手动打开确认。
  3. 转换进程被留下 —— 任务管理器里结束 soffice.bin 后重试。

手动复现转换过程:
  & "$soffice" --headless --norestore --convert-to pdf --outdir "$OutDir" "$Pptx"
"@
    }
}
if ($loPdf -ne $pdfPath) { Move-Item -LiteralPath $loPdf -Destination $pdfPath -Force }
$pdfKB = [math]::Round((Get-Item -LiteralPath $pdfPath).Length / 1KB, 1)
Write-Host "[1/2] PDF  (矢量)     -> $([IO.Path]::GetFileName($pdfPath))  $pdfKB KB"

# ---- 阶段 2:PDF -> PNG / SVG ----
# pdftocairo 对非 ASCII(中文)路径支持不佳,统一在 ASCII 临时目录里转换,再搬回目标位置
$tmpWork = Join-Path $env:TEMP ("pdftocairo_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path $tmpWork | Out-Null
$tmpPdf = Join-Path $tmpWork "in.pdf"
Copy-Item -LiteralPath $pdfPath -Destination $tmpPdf -Force
$made = @()
if ($Formats -contains 'png') {
    if (-not $pdftocairo) { Write-Warning "缺少 pdftocairo,跳过 PNG 导出。" }
    else {
        # Page>0 → 只导该页,用 -singlefile 直接得到 out.png
        # Page=0 → 导全部页,pdftocairo 输出 out-1.png / out-2.png …
        $pageArgs = if ($Page -gt 0) { @('-f', "$Page", '-l', "$Page", '-singlefile') } else { @() }
        $tmpBase = Join-Path $tmpWork "out"
        & $pdftocairo -png -r $Dpi @pageArgs $tmpPdf $tmpBase | Out-Null
        $produced = @(Get-ChildItem -LiteralPath $tmpWork -Filter "out*.png" -File -ErrorAction SilentlyContinue)
        if ($produced.Count -eq 0) { Write-Warning "PNG 生成失败。" }
        # 只有一页时不保留 "-1" 编号,直接叫 base.png
        $dropIndex = ($produced.Count -eq 1)
        foreach ($p in $produced) {
            $suffix = if ($dropIndex) { "" } else { $p.BaseName -replace '^out', '' }
            $dest = Join-Path $OutDir ($BaseName + $suffix + ".png")
            Copy-Item -LiteralPath $p.FullName -Destination $dest -Force
            $kb = [math]::Round((Get-Item -LiteralPath $dest).Length / 1KB, 1)
            $dims = ""
            try {
                Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
                $img = [System.Drawing.Image]::FromFile($dest)
                $dims = "  $($img.Width)x$($img.Height) px"
                $img.Dispose()
            } catch { }
            Write-Host "[2/2] PNG  ($Dpi DPI)    -> $([IO.Path]::GetFileName($dest))  $kb KB$dims"
            $made += $dest
        }
    }
}
if ($Formats -contains 'svg') {
    if (-not $pdftocairo) { Write-Warning "缺少 pdftocairo,跳过 SVG 导出。" }
    else {
        $svgPath = Join-Path $OutDir ($BaseName + ".svg")
        $tmpSvgOut = Join-Path $tmpWork "out.svg"
        # 矢量图通常只要一页;Page=0(全部页)时默认取第 1 页
        $svgPage = if ($Page -gt 0) { $Page } else { 1 }
        & $pdftocairo -svg -f $svgPage -l $svgPage $tmpPdf $tmpSvgOut | Out-Null
        if (Test-Path -LiteralPath $tmpSvgOut) {
            Copy-Item -LiteralPath $tmpSvgOut -Destination $svgPath -Force
        }
        if (Test-Path -LiteralPath $svgPath) {
            $kb = [math]::Round((Get-Item -LiteralPath $svgPath).Length / 1KB, 1)
            Write-Host "[2/2] SVG  (矢量)     -> $BaseName.svg  $kb KB"
            $made += $svgPath
        } else { Write-Warning "SVG 生成失败。" }
    }
}
if ($Formats -contains 'pdf' -and (Test-Path -LiteralPath $pdfPath)) { $made += $pdfPath }

# 清理临时工作目录
Remove-Item -LiteralPath $tmpWork -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "完成。共 $($made.Count) 个文件:"
$made | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "插入 Word 建议:"
Write-Host "  - SVG  : 插入 > 图片 > 此设备,选 .svg。Word 2016+ 支持,矢量,可右键'转换为形状'编辑。"
Write-Host "  - PNG  : 直接插入。$Dpi DPI 足够印刷;若期刊要求更高可调 -Dpi。"
