@echo off
setlocal

rem ---------------------------------------------------------------------------
rem  Drag-and-drop exporter. Drop one or more .pptx files onto this file's icon
rem  and it writes the images next to each pptx.
rem
rem  Deliberately ASCII-only: a .cmd containing non-ASCII text is a portability
rem  trap. Windows batch is parsed byte-by-byte, so the file must be saved in a
rem  codepage the console agrees with -- and anyone who opens it in an editor
rem  that saves UTF-8 by default silently turns it into a broken script.
rem  Plain ASCII sidesteps the whole problem. Keep it that way.
rem
rem  Line endings must be CRLF, not LF. With LF, cmd.exe shreds multi-line
rem  blocks like "if ... (" into separate broken lines.
rem ---------------------------------------------------------------------------

if "%~1"=="" (
    echo.
    echo   Drag and drop .pptx file^(s^) onto this file icon to export images.
    echo.
    echo   Output ^(written next to each pptx^):
    echo     - PNG at 600 DPI ^(one per slide^)
    echo     - Vector SVG ^(first slide^)
    echo.
    echo   Need other options ^(page, DPI, formats^)? Call scripts\export_figure.ps1
    echo   directly -- see the README.
    echo.
    pause
    exit /b
)

rem Absolute paths to external commands. Some machines (this one included) have
rem no C:\Windows\System32 on PATH, which makes chcp / powershell unresolvable.
set "CHCPEXE=%SystemRoot%\System32\chcp.com"
set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

rem 936 = GBK. Keeps the Chinese console output of export_figure.ps1 readable;
rem PowerShell 5.1 mangles it under codepage 65001.
if exist "%CHCPEXE%" "%CHCPEXE%" 936 >nul

for %%F in (%*) do (
    echo.
    echo ============================================================
    echo   Processing: %%~nxF
    echo ============================================================
    if exist "%PSEXE%" (
        "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\export_figure.ps1" -Pptx "%%~fF"
    ) else (
        echo   [ERROR] PowerShell not found: %PSEXE%
    )
)

echo.
echo All done. Press any key to close.
pause >nul
