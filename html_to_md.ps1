<#
.SYNOPSIS
    把 vcenter_inspect 生成的 HTML 报告转成 Markdown
.EXAMPLE
    .\html_to_md.ps1 -InputPath .\report_172.28.1.150_2026-05-25.html
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string] $InputPath,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if (-not (Test-Path -LiteralPath $InputPath)) { throw "找不到 HTML: $InputPath" }
$abs = (Resolve-Path -LiteralPath $InputPath).Path
if (-not $OutputPath) { $OutputPath = [System.IO.Path]::ChangeExtension($abs, '.md') }
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath = Join-Path (Get-Location) $OutputPath }

$html = [System.IO.File]::ReadAllText($abs, [System.Text.Encoding]::UTF8)

# ============================================================================
#  正则全部抽到变量, 避免 PS 5.1 parser 对 method-invocation 内的 string literal 误判
# ============================================================================
$RX_TITLE      = [regex]'<title>([^<]+)</title>'
$RX_BANNER     = [regex]'<div class="banner">([\s\S]*?)</div>\s*<div class="cards">'
$RX_H1         = [regex]'<h1>([\s\S]*?)</h1>'
$RX_SUB        = [regex]'<div class="sub">([\s\S]*?)</div>'
$RX_META_DL    = [regex]'<dl class="meta">([\s\S]*?)</dl>'
$RX_DL         = [regex]'<div><dt>([\s\S]*?)</dt>\s*<dd>([\s\S]*?)</dd></div>'
$RX_CARDS_BLK  = [regex]'<div class="cards">([\s\S]*?)</div>\s*<section'
$RX_CARD       = [regex]'<div class="card">[\s\S]*?<div class="num">([^<]+)</div><div class="lbl">([^<]+)</div><div class="sub">([\s\S]*?)</div></div>'
$RX_SECTION    = [regex]'<section id="([^"]+)">([\s\S]*?)</section>'
$RX_H2         = [regex]'<h2>([\s\S]*?)</h2>'
$RX_H2_NUM     = [regex]'<span class="num">(\d+)</span>'
$RX_H2_LBL     = [regex]'<span class="lbl">([\s\S]*?)</span>'
$RX_INFOGRID   = [regex]'<div class="info-grid"[^>]*>([\s\S]*?)\r?\n\s*</div>'
$RX_TABLE      = [regex]'<table>[\s\S]*?</table>'
$RX_TBODY      = [regex]'<tbody>([\s\S]*?)</tbody>'
$RX_TH         = [regex]'<th[^>]*>([\s\S]*?)</th>'
$RX_TR         = [regex]'<tr>([\s\S]*?)</tr>'
$RX_TD         = [regex]'<td[^>]*>([\s\S]*?)</td>'
$RX_FINDINGS   = [regex]'<div class="findings">([\s\S]*?)</div>\s*\Z'
$RX_FINDCOL    = [regex]'<div class="find-col (short|mid|long)">([\s\S]*?)</div>(?=\s*<div class="find-col|\s*\Z)'
$RX_FIND_H3    = [regex]'<h3>([\s\S]*?)</h3>'
$RX_FIND_EMPTY = [regex]"<p class='empty'>([\s\S]*?)</p>"
$RX_LI         = [regex]'<li>([\s\S]*?)</li>'
$RX_AREA       = [regex]"<span class='area'>\[([^\]]+)\]</span>"
$RX_B          = [regex]'<b>([\s\S]*?)</b>'
$RX_MUTED_SPAN = [regex]"<span class='muted'>([\s\S]*?)</span>"
$RX_DISCLAIMER = [regex]'<div class="disclaimer">([\s\S]*?)</div>\s*\Z'
$RX_NOTE       = [regex]'<p class="muted"[^>]*>([\s\S]*?)</p>'
$RX_FOOTER     = [regex]'<footer>([\s\S]*?)</footer>'

# StripTags 内部用
$RX_BADGE      = [regex]"<span class='badge badge-(\w+)'>([^<]*)</span>"
$RX_MUTED_DQ   = [regex]'<span class="muted">([^<]*)</span>'
$RX_BTAG       = [regex]'<b>([^<]*)</b>'
$RX_BAR        = [regex]'<div class="bar[^"]*"><i style="width:([^"]+)"></i></div>'
$RX_ANY_TAG    = [regex]'<[^>]+>'
$RX_MULTI_WS   = [regex]'\s+'

# ============================================================================
#  辅助函数
# ============================================================================
function HtmlDecode([string]$s) {
    if (-not $s) { return '' }
    return [System.Net.WebUtility]::HtmlDecode($s)
}

function StripTags([string]$s) {
    if (-not $s) { return '' }

    # badge → 文字标记
    $s = $RX_BADGE.Replace($s, {
        param($m)
        $kind = $m.Groups[1].Value
        $txt  = $m.Groups[2].Value.Trim()
        switch ($kind) {
            'red'   { "**[!] $txt**" }
            'amber' { "**[~] $txt**" }
            'green' { "**$txt**" }
            'blue'  { "*$txt*" }
            default { "[$txt]" }
        }
    })

    # <span class="muted">X</span>
    $s = $RX_MUTED_DQ.Replace($s, { param($m) '_' + $m.Groups[1].Value.Trim() + '_' })

    # <b>X</b>
    $s = $RX_BTAG.Replace($s, '**$1**')

    # 进度条
    $s = $RX_BAR.Replace($s, { param($m) '[' + $m.Groups[1].Value + ' 进度]' })

    # 其他 tag 全部去掉
    $s = $RX_ANY_TAG.Replace($s, '')
    $s = HtmlDecode $s
    $s = $RX_MULTI_WS.Replace($s, ' ')
    return $s.Trim()
}

function MdEscape([string]$s) {
    if ($null -eq $s) { return '' }
    return $s -replace '\|', '\|'
}

function Convert-InfoGrid([string]$block) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('| 项 | 值 |')
    [void]$sb.AppendLine('|---|---|')
    foreach ($m in $RX_DL.Matches($block)) {
        $k = MdEscape (StripTags $m.Groups[1].Value)
        $v = MdEscape (StripTags $m.Groups[2].Value)
        if (-not $v) { $v = '—' }
        [void]$sb.AppendLine("| $k | $v |")
    }
    return $sb.ToString()
}

function Convert-Table([string]$block) {
    $headers = New-Object System.Collections.Generic.List[string]
    foreach ($m in $RX_TH.Matches($block)) { [void]$headers.Add((StripTags $m.Groups[1].Value)) }
    if ($headers.Count -eq 0) { return '' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('| ' + ($headers -join ' | ') + ' |')
    $sepCells = @()
    foreach ($h in $headers) { $sepCells += '---' }
    [void]$sb.AppendLine('| ' + ($sepCells -join ' | ') + ' |')

    $bodyM = $RX_TBODY.Match($block)
    $tbody = if ($bodyM.Success) { $bodyM.Groups[1].Value } else { $block }

    foreach ($tr in $RX_TR.Matches($tbody)) {
        $cells = New-Object System.Collections.Generic.List[string]
        foreach ($td in $RX_TD.Matches($tr.Groups[1].Value)) {
            $c = MdEscape (StripTags $td.Groups[1].Value)
            if (-not $c) { $c = ' ' }
            [void]$cells.Add($c)
        }
        if ($cells.Count -gt 0) {
            [void]$sb.AppendLine('| ' + ($cells -join ' | ') + ' |')
        }
    }
    return $sb.ToString()
}

function Convert-Cards([string]$block) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('| 指标 | 数值 | 备注 |')
    [void]$sb.AppendLine('|---|---|---|')
    foreach ($m in $RX_CARD.Matches($block)) {
        $num = MdEscape (StripTags $m.Groups[1].Value)
        $lbl = MdEscape (StripTags $m.Groups[2].Value)
        $sub = MdEscape (StripTags $m.Groups[3].Value)
        if (-not $sub) { $sub = ' ' }
        [void]$sb.AppendLine("| $lbl | **$num** | $sub |")
    }
    return $sb.ToString()
}

function Convert-Findings([string]$block) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in $RX_FINDCOL.Matches($block)) {
        $body  = $m.Groups[2].Value
        $h3m   = $RX_FIND_H3.Match($body)
        $title = if ($h3m.Success) { StripTags $h3m.Groups[1].Value } else { '' }
        [void]$sb.AppendLine('### ' + $title)
        [void]$sb.AppendLine('')

        $emptyM = $RX_FIND_EMPTY.Match($body)
        if ($emptyM.Success) {
            [void]$sb.AppendLine('> _' + (StripTags $emptyM.Groups[1].Value) + '_')
            [void]$sb.AppendLine('')
            continue
        }
        foreach ($li in $RX_LI.Matches($body)) {
            $content = $li.Groups[1].Value
            $areaM   = $RX_AREA.Match($content)
            $titM    = $RX_B.Match($content)
            $detM    = $RX_MUTED_SPAN.Match($content)
            $area    = if ($areaM.Success) { StripTags $areaM.Groups[1].Value } else { '' }
            $titT    = if ($titM.Success)  { StripTags $titM.Groups[1].Value }  else { '' }
            $det     = if ($detM.Success)  { StripTags $detM.Groups[1].Value }  else { '' }
            $line = "- **[$area]** $titT"
            if ($det) { $line += "  `n  _" + $det + '_' }
            [void]$sb.AppendLine($line)
        }
        [void]$sb.AppendLine('')
    }
    return $sb.ToString()
}

# ============================================================================
#  主转换
# ============================================================================
$out = New-Object System.Text.StringBuilder

# ---- Title ----
$tm = $RX_TITLE.Match($html)
$mainTitle = if ($tm.Success) { HtmlDecode $tm.Groups[1].Value } else { 'vCenter 巡检报告' }
[void]$out.AppendLine("# $mainTitle")
[void]$out.AppendLine('')

# ---- Banner ----
$bm = $RX_BANNER.Match($html)
if ($bm.Success) {
    $banner = $bm.Groups[1].Value
    $h1m = $RX_H1.Match($banner)
    $h1  = if ($h1m.Success) { StripTags $h1m.Groups[1].Value } else { '' }
    $subm = $RX_SUB.Match($banner)
    $sub = if ($subm.Success) { StripTags $subm.Groups[1].Value } else { '' }
    if ($h1)  { [void]$out.AppendLine("> $h1") }
    if ($sub) { [void]$out.AppendLine("> $sub") }
    [void]$out.AppendLine('')

    $metaM = $RX_META_DL.Match($banner)
    if ($metaM.Success) {
        [void]$out.AppendLine('| 项 | 值 |')
        [void]$out.AppendLine('|---|---|')
        foreach ($mm in $RX_DL.Matches($metaM.Groups[1].Value)) {
            $k = MdEscape (StripTags $mm.Groups[1].Value)
            $v = MdEscape (StripTags $mm.Groups[2].Value)
            [void]$out.AppendLine("| $k | $v |")
        }
        [void]$out.AppendLine('')
    }
}

# ---- Summary cards ----
$cm = $RX_CARDS_BLK.Match($html)
if ($cm.Success) {
    [void]$out.AppendLine('## 关键指标')
    [void]$out.AppendLine('')
    [void]$out.Append((Convert-Cards $cm.Groups[1].Value))
    [void]$out.AppendLine('')
}

# ---- Sections ----
foreach ($sm in $RX_SECTION.Matches($html)) {
    $secId   = $sm.Groups[1].Value
    $secBody = $sm.Groups[2].Value

    $h2m = $RX_H2.Match($secBody)
    if ($h2m.Success) {
        $h2raw = $h2m.Groups[1].Value
        $numM = $RX_H2_NUM.Match($h2raw)
        $lblM = $RX_H2_LBL.Match($h2raw)
        $num = if ($numM.Success) { $numM.Groups[1].Value } else { '' }
        $lbl = if ($lblM.Success) { $lblM.Groups[1].Value } else { '' }
        $titleText = $h2raw
        $titleText = $RX_H2_NUM.Replace($titleText, '')
        $titleText = $RX_H2_LBL.Replace($titleText, '')
        $titleText = StripTags $titleText
        $header = if ($num) { "## $num. $titleText" } else { "## $titleText" }
        [void]$out.AppendLine($header)
        if ($lbl) {
            [void]$out.AppendLine('')
            [void]$out.AppendLine('`endpoint: ' + $lbl + '`')
        }
        [void]$out.AppendLine('')
    }

    # 按出现顺序收集所有 block
    $tokens = New-Object System.Collections.Generic.List[psobject]
    foreach ($x in $RX_INFOGRID.Matches($secBody))   { [void]$tokens.Add([pscustomobject]@{ Pos=$x.Index; Kind='infogrid';   Text=$x.Groups[1].Value }) }
    foreach ($x in $RX_TABLE.Matches($secBody))      { [void]$tokens.Add([pscustomobject]@{ Pos=$x.Index; Kind='table';      Text=$x.Value }) }
    foreach ($x in $RX_FINDINGS.Matches($secBody))   { [void]$tokens.Add([pscustomobject]@{ Pos=$x.Index; Kind='findings';   Text=$x.Groups[1].Value }) }
    foreach ($x in $RX_DISCLAIMER.Matches($secBody)) { [void]$tokens.Add([pscustomobject]@{ Pos=$x.Index; Kind='disclaimer'; Text=$x.Groups[1].Value }) }
    foreach ($x in $RX_NOTE.Matches($secBody))       { [void]$tokens.Add([pscustomobject]@{ Pos=$x.Index; Kind='note';       Text=$x.Groups[1].Value }) }

    foreach ($t in ($tokens | Sort-Object Pos)) {
        switch ($t.Kind) {
            'infogrid' {
                [void]$out.Append((Convert-InfoGrid $t.Text))
                [void]$out.AppendLine('')
            }
            'table' {
                [void]$out.Append((Convert-Table $t.Text))
                [void]$out.AppendLine('')
            }
            'findings' {
                [void]$out.Append((Convert-Findings $t.Text))
            }
            'disclaimer' {
                [void]$out.AppendLine('> **关于本报告**')
                [void]$out.AppendLine('>')
                $clean = StripTags $t.Text
                $lines = $clean -split '。'
                foreach ($ln in $lines) {
                    $tn = $ln.Trim()
                    if ($tn) { [void]$out.AppendLine("> $tn" + '。') }
                }
                [void]$out.AppendLine('')
            }
            'note' {
                [void]$out.AppendLine('> _' + (StripTags $t.Text) + '_')
                [void]$out.AppendLine('')
            }
        }
    }
}

# ---- Footer ----
$fm = $RX_FOOTER.Match($html)
if ($fm.Success) {
    [void]$out.AppendLine('---')
    [void]$out.AppendLine('')
    [void]$out.AppendLine('_' + (StripTags $fm.Groups[1].Value) + '_')
}

# ---- 写出 ----
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutputPath, $out.ToString(), $utf8NoBom)

$size = (Get-Item -LiteralPath $OutputPath).Length
Write-Host ''
Write-Host ('  Markdown 已生成 — 大小 ' + ('{0:N1} KB' -f ($size/1KB))) -ForegroundColor Green
Write-Host "    $OutputPath" -ForegroundColor White
Write-Host ''
