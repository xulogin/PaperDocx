#requires -Version 5
<#
  把排版版 docx 用**真 Word** 导成 PDF，再把 PDF 每页转成 PNG，交付前肉眼过一遍。

      powershell -ExecutionPolicy Bypass -File _工作区\排版\渲染看图.ps1
      powershell -ExecutionPolicy Bypass -File _工作区\排版\渲染看图.ps1 -Docx 别的.docx -Dpi 110

  ★ 用 Word 不用 LibreOffice：作者最终就是在 Word 里看。LibreOffice 画 OMML 很糙
    （`\lVert` 画成 `//`、`\mathbf` 不加粗），拿它验收公式会看错。
    （想快速看版面时 LibreOffice 够用：
     soffice --headless --convert-to pdf --outdir <目录> <docx>）
  ★ 这个文件必须存成 **UTF-8 带 BOM**，否则 PowerShell 5.1 读中文会报
    「字符串缺少终止符」。

  ## 两个坑（都实测过，别再踩）

  1. **`New-Object -ComObject Word.Application` 会接管已经在跑的 Word。**
     早先导一份 59 MB 的稿子把 Word 卡住了，之后每次自动化导出都接到那个卡住的
     实例上干等，看着像「我的文档把 Word 搞死了」。所以这里先查有没有在跑的 Word。
  2. **Word 的 COM 调用放在当前进程里做会莫名卡住**，同一个文件丢进 `Start-Job`
     里 13 秒就导完了。所以 Word 那段统一在子作业里跑，还带超时，卡了就报出来。
#>
param(
  [string]$Docx = "",          # 留空 = 自动找项目里最新的 *_排版版.docx
  [int]$Dpi = 100,
  [string]$Out = "",           # 留空 = <脚本目录>\_预览
  [int]$TimeoutSec = 300
)
$ErrorActionPreference = "Stop"

# ★ 项目根**不许写死层级**，而且要**先从当前工作目录找，再从脚本位置找**。
#   顺序不能反：skill 装好之后脚本在 ~/.claude/plugins/cache/... 里，
#   从那儿往上找永远找不到用户的项目（实测报「往上找不到项目根」）。
$root = $null
foreach ($start in @((Get-Location).Path, $PSScriptRoot)) {
  $d = $start
  while ($d -and -not $root) {
    if ((Get-ChildItem $d -Filter "*_排版版.docx" -ErrorAction SilentlyContinue) -or
        (Get-ChildItem $d -Filter "*_英文版.md"  -ErrorAction SilentlyContinue)) { $root = $d }
    $d = Split-Path -Parent $d
  }
  if ($root) { break }
}
if (-not $root) {
  throw "往上找不到项目根（判据：某层目录里有 *_排版版.docx 或 *_英文版.md）。" +
        "⇒ 在项目目录里跑，或者用 -Docx 给绝对路径。"
}

if (-not $Docx) {
  $cand = Get-ChildItem $root -Filter "*_排版版.docx" | Sort-Object Name
  if (-not $cand) { throw "$root 里没有 *_排版版.docx，先跑 出Word.py" }
  $src = $cand[-1].FullName            # 字典序最大 = 版本号最大
} elseif ([System.IO.Path]::IsPathRooted($Docx)) { $src = $Docx }
else { $src = Join-Path $root $Docx }
if (-not (Test-Path $src)) { throw "找不到 $src" }

# ★ 预览图必须落在**项目自己的目录**里，不能落在脚本目录。
#   skill 装好之后脚本在插件缓存里，所有项目共用一份 —— 2026-10-19 实测：
#   先渲染 demo，再渲染另一个项目，两边的 p001.png 互相覆盖，
#   连 PDF 都混在一个文件夹里，看图时拿错了文件还不自知。
if (-not $Out) { $Out = Join-Path $root "_预览" }
$outDir = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$pdf = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($src) + ".pdf")

$在跑 = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($在跑.Count) {
  Write-Warning ("已经有 {0} 个 Word 在跑，这次会接管它们。迟迟不出结果就先把 Word 关干净" -f $在跑.Count +
                 "（确认里面没有你正在编辑的文档）再跑。")
}

Write-Host "Word 导出 PDF（超时 $TimeoutSec 秒）..."
$t0 = Get-Date
$job = Start-Job -ScriptBlock {
  param($s, $o)
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false; $w.DisplayAlerts = 0
  # ReadOnly + 不加进最近文件：不动原文件，也不污染作者的「最近使用」
  $d = $w.Documents.Open($s, $false, $true)
  $d.ExportAsFixedFormat($o, 17)       # 17 = wdExportFormatPDF
  $d.Close($false); $w.Quit()
} -ArgumentList $src, $pdf
if (-not (Wait-Job $job -Timeout $TimeoutSec)) {
  Stop-Job $job; Remove-Job $job -Force
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  throw "Word 超过 $TimeoutSec 秒还没导完。先确认没有别的 Word 在跑；还卡就是图太大了，调小 排版配置.图['送入dpi']。"
}
Receive-Job $job | Out-Null; Remove-Job $job -Force
Write-Host ("  {0:N0} 秒  {1:N1} MB  ->  {2}" -f ((Get-Date) - $t0).TotalSeconds,
            ((Get-Item $pdf).Length / 1MB), $pdf)

Write-Host "PDF -> PNG ..."
$py = @"
import sys, pathlib, fitz
pdf, out, dpi = sys.argv[1], pathlib.Path(sys.argv[2]), int(sys.argv[3])
for old in out.glob('p*.png'):
    old.unlink()
d = fitz.open(pdf)
for i, page in enumerate(d, 1):
    page.get_pixmap(dpi=dpi).save(str(out / ('p%03d.png' % i)))
print('  %d 页 -> %s' % (len(d), out))
"@
$tmp = Join-Path $env:TEMP ("topng_{0}.py" -f [guid]::NewGuid().ToString('N'))
[System.IO.File]::WriteAllText($tmp, $py, [System.Text.UTF8Encoding]::new($false))
& py -3.12 -X utf8 $tmp $pdf $outDir $Dpi
Remove-Item $tmp -Force
