<#
.SYNOPSIS
    用 Word COM 把 HTML 巡检报告转成 .docx
.PARAMETER Input
    输入 HTML 路径 (必需)
.PARAMETER Output
    输出 .docx 路径; 留空则与输入同目录同名 + .docx
.EXAMPLE
    .\html_to_docx.ps1 -Input .\report_172.28.1.150_2026-05-25.html
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $InputPath,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $InputPath)) { throw "找不到 HTML: $InputPath" }
$abs = (Resolve-Path -LiteralPath $InputPath).Path
if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::ChangeExtension($abs, '.docx')
}
# 转绝对路径
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path (Get-Location) $OutputPath
}

Write-Host ''
Write-Host "  ─── HTML → DOCX 转换 ──────────────────────────" -ForegroundColor DarkCyan
Write-Host "    Input  : $abs"           -ForegroundColor Gray
Write-Host "    Output : $OutputPath"    -ForegroundColor Gray
Write-Host "  ────────────────────────────────────────────"   -ForegroundColor DarkCyan
Write-Host ''

$t0 = Get-Date
$word = New-Object -ComObject Word.Application
$word.Visible       = $false
$word.DisplayAlerts = 0    # wdAlertsNone
$word.ScreenUpdating = $false

try {
    # Word 看到 <!DOCTYPE html> 时, 即便 OpenFormat 强制 HTML 也可能按 XML 解析报 DTD 错
    # 兜底: 复制成 .htm 后缀临时文件, Word 看扩展名走 HTML 路径
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(),
            ("vcinspect_" + [Guid]::NewGuid().ToString('N') + ".htm"))
    Copy-Item -LiteralPath $abs -Destination $tmp -Force

    Write-Host "  [1/3] 打开 HTML (临时副本 .htm) ..." -ForegroundColor Cyan
    $doc = $word.Documents.Open(
        $tmp,           # FileName (.htm 后缀避免 XML 解析)
        $false,         # ConfirmConversions
        $true,          # ReadOnly
        $false,         # AddToRecentFiles
        '',             # PasswordDocument
        '',             # PasswordTemplate
        $false,         # Revert
        '',             # WritePasswordDocument
        '',             # WritePasswordTemplate
        7               # Format = wdOpenFormatWebPages (7)
    )

    Write-Host "  [2/3] 另存为 .docx (wdFormatDocumentDefault = 16) ..." -ForegroundColor Cyan
    # 16 = wdFormatDocumentDefault (.docx)
    $doc.SaveAs([ref]$OutputPath, [ref]16)
    $doc.Close($false)

    Write-Host "  [3/3] 清理 Word 进程 ..." -ForegroundColor Cyan
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

if (Test-Path -LiteralPath $OutputPath) {
    $size = (Get-Item -LiteralPath $OutputPath).Length
    $cost = [int]((Get-Date) - $t0).TotalSeconds
    Write-Host ''
    Write-Host "  ✓ 转换完成 — 耗时 ${cost}s — 大小 $('{0:N1} KB' -f ($size/1KB))" -ForegroundColor Green
    Write-Host "    $OutputPath" -ForegroundColor White
    Write-Host ''
} else {
    throw "转换失败,未生成 $OutputPath"
}
