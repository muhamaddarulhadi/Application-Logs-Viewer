<#
.SYNOPSIS
    Collects Error and Warning events from Windows Event Viewer > Windows Logs > Application
    for a date range and builds a modern, interactive HTML report with real .xlsx download.

.EXAMPLE
    .\Get-AppEventReport.ps1
    (prompts for start and end date)

.EXAMPLE
    .\Get-AppEventReport.ps1 -StartDate 2026-10-01 -EndDate 2026-10-13 -IncludeCritical

.LINK
    https://github.com/muhamaddarulhadi/Application-Logs-Viewer
#>
[CmdletBinding()]
param(
    [datetime]$StartDate,
    [datetime]$EndDate,
    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'EventReports'),
    [switch]$IncludeCritical,
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'

# ---------- 1. Date range input ----------
function Read-DateInput($prompt, $default) {
    while ($true) {
        $text = Read-Host "$prompt [default: $($default.ToString('yyyy-MM-dd'))]"
        if ([string]::IsNullOrWhiteSpace($text)) { return $default }
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($text, [ref]$parsed)) { return $parsed }
        Write-Host "  Invalid date. Try format yyyy-MM-dd (e.g. 2026-10-13)." -ForegroundColor Yellow
    }
}

if (-not $PSBoundParameters.ContainsKey('StartDate')) { $StartDate = Read-DateInput 'Start date' (Get-Date).Date.AddDays(-7) }
if (-not $PSBoundParameters.ContainsKey('EndDate'))   { $EndDate   = Read-DateInput 'End date'   (Get-Date).Date }

# A date-only end date means "through the end of that day"
if ($EndDate.TimeOfDay.TotalSeconds -eq 0) { $EndDate = $EndDate.Date.AddDays(1).AddSeconds(-1) }
if ($EndDate -lt $StartDate) { throw "End date must be after start date." }

# ---------- 2. Query the Application log ----------
$levels = @(2, 3)                      # 2 = Error, 3 = Warning
if ($IncludeCritical) { $levels = @(1, 2, 3) }

Write-Host "Reading Application log from $StartDate to $EndDate ..." -ForegroundColor Cyan

$filter = @{
    LogName   = 'Application'
    Level     = $levels
    StartTime = $StartDate
    EndTime   = $EndDate
}

$events = Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue

if (-not $events) {
    Write-Host "No errors or warnings found in that date range." -ForegroundColor Green
    return
}

Write-Host "Found $(@($events).Count) events. Processing ..." -ForegroundColor Cyan

# ---------- 3. Shape the data ----------
$rows = foreach ($e in $events) {
    $msg = $e.Message
    if ([string]::IsNullOrWhiteSpace($msg)) {
        # Provider message DLL missing: fall back to raw event data
        $props = @($e.Properties | ForEach-Object { $_.Value }) -join ' | '
        $msg = if ($props) { "(Description unavailable) Raw data: $props" } else { '(No description available)' }
    }
    [pscustomobject]@{
        Date      = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        Level     = $e.LevelDisplayName
        Source    = $e.ProviderName
        EventID   = $e.Id
        Category  = $e.TaskDisplayName
        Computer  = $e.MachineName
        RecordId  = $e.RecordId
        Message   = $msg.Trim()
    }
}
$rows = @($rows)

# ---------- 4. Output files ----------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseName = "AppEvents_$stamp"
$htmlPath = Join-Path $OutputFolder "$baseName.html"
$csvPath  = Join-Path $OutputFolder "$baseName.csv"
$xlsxPath = Join-Path $OutputFolder "$baseName.xlsx"

$rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# ---------- Native .xlsx writer (no Excel or modules required) ----------
function New-XlsxFile {
    param([object[]]$Rows, [string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    $cols    = 'Date','Level','Source','EventID','Category','Computer','RecordId','Message'
    $letters = 'A','B','C','D','E','F','G','H'
    $numeric = 'EventID','RecordId'
    $lvlStyle = @{ Error = 3; Warning = 4; Critical = 5 }

    function ConvertTo-XmlText([string]$t) {
        [System.Security.SecurityElement]::Escape(($t -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''))
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<row r="1" ht="22" customHeight="1">')
    for ($i = 0; $i -lt $cols.Count; $i++) {
        [void]$sb.Append("<c r=`"$($letters[$i])1`" s=`"1`" t=`"inlineStr`"><is><t>$($cols[$i])</t></is></c>")
    }
    [void]$sb.Append('</row>')

    $n = 2
    foreach ($row in $Rows) {
        [void]$sb.Append("<row r=`"$n`">")
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $c = $cols[$i]; $ref = "$($letters[$i])$n"; $val = $row.$c
            if ($numeric -contains $c) {
                [void]$sb.Append("<c r=`"$ref`"><v>$([int64]$val)</v></c>")
            } else {
                $style = 0
                if ($c -eq 'Level')   { $style = if ($lvlStyle.ContainsKey([string]$val)) { $lvlStyle[[string]$val] } else { 0 } }
                if ($c -eq 'Message') { $style = 2 }
                $text = [string]$val
                if ($text.Length -gt 32000) { $text = $text.Substring(0, 32000) }
                [void]$sb.Append("<c r=`"$ref`" s=`"$style`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $text)</t></is></c>")
            }
        }
        [void]$sb.Append('</row>')
        $n++
    }
    $last = $n - 1

    $sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="11" customWidth="1"/><col min="3" max="3" width="28" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/><col min="6" max="6" width="22" customWidth="1"/><col min="7" max="7" width="12" customWidth="1"/><col min="8" max="8" width="100" customWidth="1"/></cols><sheetData>' + $sb.ToString() + '</sheetData><autoFilter ref="A1:H' + $last + '"/></worksheet>'

    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4F46E5"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFE4EA"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEDE9FE"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    $parts = [ordered]@{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Errors and Warnings" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/styles.xml'              = $styles
        'xl/worksheets/sheet1.xml'   = $sheet
    }

    if (Test-Path $Path) { Remove-Item $Path -Force }
    $fs  = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $entry  = $zip.CreateEntry($name)
            $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($parts[$name])
            $writer.Dispose()
        }
    } finally { $zip.Dispose(); $fs.Dispose() }
}

# Always build a real .xlsx next to the HTML report
$xlsxMade = $false
try {
    New-XlsxFile -Rows $rows -Path $xlsxPath
    $xlsxMade = $true
} catch { Write-Host "XLSX export failed: $($_.Exception.Message)" -ForegroundColor Yellow }

# ---------- 5. Build the HTML report ----------
$meta = [pscustomobject]@{
    Computer  = $env:COMPUTERNAME
    From      = $StartDate.ToString('yyyy-MM-dd HH:mm:ss')
    To        = $EndDate.ToString('yyyy-MM-dd HH:mm:ss')
    Generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
}

# Escape "</" so event text can never close the <script> block
$dataJson = ($rows | ConvertTo-Json -Depth 3 -Compress)
if ($rows.Count -eq 1) { $dataJson = "[$dataJson]" }
$dataJson = $dataJson -replace '</', '<\/'
$metaJson = ($meta | ConvertTo-Json -Compress) -replace '</', '<\/'

$template = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Application Event Report</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --bg:#f5f7fb; --surface:#ffffff; --surface2:#f1f4f9; --text:#0f172a; --muted:#64748b; --line:#e5e9f2;
    --brand:#6366f1; --brand2:#8b5cf6; --shadow:0 1px 2px rgba(15,23,42,.05),0 8px 24px rgba(15,23,42,.06);
    --err:#e11d48; --err-soft:#ffe4ea; --warn:#d97706; --warn-soft:#fef3c7; --crit:#7c3aed; --crit-soft:#ede9fe;
    --radius:16px;
  }
  :root[data-theme="dark"] {
    --bg:#0b1020; --surface:#131a2e; --surface2:#1a2340; --text:#e6eaf5; --muted:#8c97b2; --line:#252f4f;
    --shadow:0 1px 2px rgba(0,0,0,.3),0 12px 32px rgba(0,0,0,.35);
    --err:#fb7185; --err-soft:rgba(251,113,133,.15); --warn:#fbbf24; --warn-soft:rgba(251,191,36,.14);
    --crit:#a78bfa; --crit-soft:rgba(167,139,250,.16);
  }
  * { box-sizing:border-box; }
  html { scroll-behavior:smooth; }
  body { margin:0; font:14px/1.5 "Segoe UI Variable","Segoe UI",system-ui,-apple-system,Roboto,sans-serif; background:var(--bg); color:var(--text); -webkit-font-smoothing:antialiased; }
  button,input,select { font:inherit; color:inherit; }

  /* ---------- Hero ---------- */
  .hero { position:relative; overflow:hidden; color:#fff; padding:34px 32px 90px;
          background:radial-gradient(1200px 400px at 85% -20%,rgba(255,255,255,.22),transparent 60%),
                     linear-gradient(120deg,#4f46e5 0%,#7c3aed 55%,#c026d3 100%); }
  .hero::after { content:""; position:absolute; inset:0; background-image:radial-gradient(rgba(255,255,255,.12) 1px,transparent 1px); background-size:22px 22px; mask-image:linear-gradient(to bottom,#000,transparent 85%); pointer-events:none; }
  .hero-inner { position:relative; z-index:1; max-width:1400px; margin:0 auto; display:flex; gap:20px; align-items:flex-start; justify-content:space-between; flex-wrap:wrap; }
  .brand { display:flex; gap:16px; align-items:center; }
  .logo { width:52px; height:52px; border-radius:14px; background:rgba(255,255,255,.18); backdrop-filter:blur(8px); display:grid; place-items:center; border:1px solid rgba(255,255,255,.3); }
  .hero h1 { margin:0; font-size:26px; letter-spacing:-.02em; font-weight:700; }
  .hero p { margin:2px 0 0; opacity:.85; }
  .chips { display:flex; gap:8px; flex-wrap:wrap; margin-top:16px; }
  .chip { background:rgba(255,255,255,.16); border:1px solid rgba(255,255,255,.26); backdrop-filter:blur(8px); padding:5px 12px; border-radius:999px; font-size:12.5px; display:inline-flex; gap:6px; align-items:center; }
  .chip b { font-weight:600; }
  .iconbtn { background:rgba(255,255,255,.16); border:1px solid rgba(255,255,255,.28); color:#fff; width:42px; height:42px; border-radius:12px; cursor:pointer; display:grid; place-items:center; backdrop-filter:blur(8px); transition:.15s; }
  .iconbtn:hover { background:rgba(255,255,255,.28); transform:translateY(-1px); }

  .heroact { display:flex; gap:10px; align-items:center; }
  .ghbtn { display:inline-flex; gap:9px; align-items:center; height:42px; padding:0 16px 0 13px; border-radius:12px; color:#fff; text-decoration:none; font-weight:600; font-size:13.5px; background:rgba(255,255,255,.16); border:1px solid rgba(255,255,255,.28); backdrop-filter:blur(8px); transition:.15s; }
  .ghbtn:hover { background:rgba(255,255,255,.28); transform:translateY(-1px); }
  @media (max-width:520px) { .ghbtn span { display:none; } .ghbtn { padding:0; width:42px; justify-content:center; } }

  main { max-width:1400px; margin:-62px auto 40px; padding:0 32px; position:relative; z-index:2; }

  /* ---------- KPI cards ---------- */
  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:16px; }
  .kpi { background:var(--surface); border:1px solid var(--line); border-radius:var(--radius); padding:18px 20px; box-shadow:var(--shadow); display:flex; gap:14px; align-items:center; transition:.18s; }
  .kpi:hover { transform:translateY(-2px); }
  .kpi .ico { width:46px; height:46px; border-radius:13px; display:grid; place-items:center; flex:none; }
  .kpi .n { font-size:28px; font-weight:700; letter-spacing:-.02em; line-height:1.1; }
  .kpi .l { color:var(--muted); font-size:12.5px; font-weight:500; }
  .k-all .ico { background:linear-gradient(135deg,var(--brand),var(--brand2)); color:#fff; }
  .k-err .ico { background:var(--err-soft); color:var(--err); }
  .k-warn .ico { background:var(--warn-soft); color:var(--warn); }
  .k-crit .ico { background:var(--crit-soft); color:var(--crit); }
  .k-src .ico { background:var(--surface2); color:var(--muted); }

  /* ---------- Panels ---------- */
  .grid { display:grid; grid-template-columns:2fr 1fr 1fr; gap:16px; margin-top:16px; }
  @media (max-width:1100px) { .grid { grid-template-columns:1fr 1fr; } .grid .wide { grid-column:1/-1; } }
  @media (max-width:700px) { .grid { grid-template-columns:1fr; } main,.hero { padding-left:16px; padding-right:16px; } }
  .panel { background:var(--surface); border:1px solid var(--line); border-radius:var(--radius); padding:18px 20px; box-shadow:var(--shadow); }
  .panel h3 { margin:0 0 14px; font-size:13px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); font-weight:600; }

  .legend { display:flex; gap:14px; font-size:12px; color:var(--muted); margin-top:8px; flex-wrap:wrap; }
  .legend i { display:inline-block; width:9px; height:9px; border-radius:3px; margin-right:6px; vertical-align:middle; }
  .bar-err { fill:var(--err); } .bar-warn { fill:var(--warn); } .bar-crit { fill:var(--crit); }
  .axis { fill:var(--muted); font-size:10px; }
  .gridline { stroke:var(--line); stroke-dasharray:3 4; }

  .donutwrap { display:flex; align-items:center; gap:18px; }
  .donut-center { font-size:24px; font-weight:700; fill:var(--text); }
  .donut-sub { font-size:10px; fill:var(--muted); text-transform:uppercase; letter-spacing:.08em; }

  .topsrc { display:flex; flex-direction:column; gap:10px; }
  .src-row { cursor:pointer; }
  .src-row .top { display:flex; justify-content:space-between; font-size:12.5px; margin-bottom:4px; gap:8px; }
  .src-row .top span:first-child { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .src-row .top span:last-child { color:var(--muted); font-variant-numeric:tabular-nums; }
  .track { height:7px; background:var(--surface2); border-radius:99px; overflow:hidden; }
  .fill { height:100%; border-radius:99px; background:linear-gradient(90deg,var(--brand),var(--brand2)); transition:width .5s; }
  .src-row:hover .top span:first-child { color:var(--brand); }

  /* ---------- Toolbar ---------- */
  .toolbar { margin-top:16px; background:var(--surface); border:1px solid var(--line); border-radius:var(--radius); padding:12px; box-shadow:var(--shadow); display:flex; gap:10px; flex-wrap:wrap; align-items:center; }
  .search { position:relative; flex:1; min-width:240px; }
  .search svg { position:absolute; left:12px; top:50%; transform:translateY(-50%); color:var(--muted); }
  .search input { width:100%; padding:10px 12px 10px 38px; border-radius:10px; border:1px solid var(--line); background:var(--surface2); outline:none; transition:.15s; }
  .search input:focus { border-color:var(--brand); box-shadow:0 0 0 3px rgba(99,102,241,.2); background:var(--surface); }
  .seg { display:inline-flex; background:var(--surface2); border-radius:10px; padding:3px; gap:2px; }
  .seg button { border:0; background:transparent; padding:7px 13px; border-radius:8px; cursor:pointer; color:var(--muted); font-weight:500; transition:.15s; }
  .seg button.on { background:var(--surface); color:var(--text); box-shadow:0 1px 3px rgba(0,0,0,.12); }
  select { padding:10px 12px; border-radius:10px; border:1px solid var(--line); background:var(--surface2); outline:none; max-width:220px; cursor:pointer; }
  select:focus { border-color:var(--brand); }
  .btn { border:0; padding:10px 16px; border-radius:10px; cursor:pointer; font-weight:600; display:inline-flex; gap:8px; align-items:center; transition:.15s; }
  .btn.primary { color:#fff; background:linear-gradient(135deg,#16a34a,#15803d); box-shadow:0 4px 14px rgba(22,163,74,.35); }
  .btn.primary:hover { transform:translateY(-1px); box-shadow:0 6px 18px rgba(22,163,74,.45); }
  .btn.ghost { background:var(--surface2); color:var(--text); }
  .btn.ghost:hover { background:var(--line); }

  /* ---------- Table ---------- */
  .tablecard { margin-top:16px; background:var(--surface); border:1px solid var(--line); border-radius:var(--radius); box-shadow:var(--shadow); overflow:hidden; }
  .scroll { overflow:auto; max-height:640px; }
  table { border-collapse:separate; border-spacing:0; width:100%; }
  thead th { position:sticky; top:0; z-index:1; background:var(--surface2); text-align:left; padding:12px 16px; font-size:11.5px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); font-weight:600; cursor:pointer; white-space:nowrap; user-select:none; border-bottom:1px solid var(--line); }
  thead th:hover { color:var(--text); }
  thead th .arr { opacity:.35; margin-left:4px; }
  thead th.sorted .arr { opacity:1; color:var(--brand); }
  tbody td { padding:13px 16px; border-bottom:1px solid var(--line); vertical-align:middle; }
  tbody tr { cursor:pointer; transition:background .12s; }
  tbody tr:hover { background:var(--surface2); }
  tbody tr:last-child td { border-bottom:0; }
  .date { white-space:nowrap; font-variant-numeric:tabular-nums; }
  .date small { color:var(--muted); display:block; font-size:12px; }
  .pill { display:inline-flex; align-items:center; gap:7px; padding:4px 11px 4px 9px; border-radius:99px; font-size:12px; font-weight:600; }
  .pill::before { content:""; width:7px; height:7px; border-radius:50%; background:currentColor; }
  .Error { background:var(--err-soft); color:var(--err); }
  .Warning { background:var(--warn-soft); color:var(--warn); }
  .Critical { background:var(--crit-soft); color:var(--crit); }
  .src { font-weight:600; }
  .eid { font-family:Consolas,"Cascadia Mono",monospace; font-size:12px; background:var(--surface2); border:1px solid var(--line); padding:2px 8px; border-radius:7px; }
  .msg { color:var(--muted); max-width:480px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .empty { padding:56px 20px; text-align:center; color:var(--muted); }
  .empty svg { opacity:.5; margin-bottom:8px; }

  .pager { display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; padding:12px 16px; border-top:1px solid var(--line); color:var(--muted); font-size:13px; }
  .pager .ctl { display:flex; gap:6px; align-items:center; }
  .pager button { border:1px solid var(--line); background:var(--surface); padding:6px 12px; border-radius:8px; cursor:pointer; }
  .pager button:disabled { opacity:.4; cursor:default; }
  .pager select { padding:6px 8px; }

  /* ---------- Drawer ---------- */
  .overlay { position:fixed; inset:0; background:rgba(8,12,28,.5); backdrop-filter:blur(3px); opacity:0; pointer-events:none; transition:.2s; z-index:20; }
  .overlay.on { opacity:1; pointer-events:auto; }
  .drawer { position:fixed; top:0; right:0; height:100%; width:min(640px,100%); background:var(--surface); z-index:21; transform:translateX(102%); transition:transform .28s cubic-bezier(.2,.8,.2,1); display:flex; flex-direction:column; box-shadow:-20px 0 60px rgba(0,0,0,.25); }
  .drawer.on { transform:none; }
  .dhead { padding:20px 24px; border-bottom:1px solid var(--line); display:flex; gap:12px; align-items:flex-start; justify-content:space-between; }
  .dhead h2 { margin:8px 0 2px; font-size:18px; letter-spacing:-.01em; }
  .dhead .sub { color:var(--muted); font-size:13px; }
  .dbody { padding:22px 24px; overflow:auto; flex:1; }
  .facts { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:22px; }
  .fact { background:var(--surface2); border-radius:12px; padding:12px 14px; }
  .fact span { display:block; font-size:11px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); margin-bottom:2px; font-weight:600; }
  .fact b { font-weight:600; word-break:break-word; }
  .dbody h4 { margin:0 0 8px; font-size:12px; text-transform:uppercase; letter-spacing:.07em; color:var(--muted); display:flex; justify-content:space-between; align-items:center; }
  pre { margin:0; white-space:pre-wrap; word-break:break-word; font:13px/1.55 Consolas,"Cascadia Mono",monospace; background:var(--surface2); border:1px solid var(--line); border-radius:12px; padding:16px; }
  .dfoot { padding:14px 24px; border-top:1px solid var(--line); display:flex; justify-content:space-between; gap:10px; }
  .copy { border:0; background:transparent; color:var(--brand); cursor:pointer; font-weight:600; font-size:12px; text-transform:none; letter-spacing:0; }

  footer { text-align:center; color:var(--muted); font-size:12px; padding:0 0 30px; }
  footer a { color:var(--brand); text-decoration:none; font-weight:600; }
  footer a:hover { text-decoration:underline; }
  @media print { .hero { background:#4f46e5; } .toolbar,.pager,.iconbtn,.ghbtn,.drawer,.overlay { display:none; } .scroll { max-height:none; } }
</style>
</head>
<body>

<header class="hero">
  <div class="hero-inner">
    <div>
      <div class="brand">
        <div class="logo">
          <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/></svg>
        </div>
        <div>
          <h1>Application Event Report</h1>
          <p>Windows Logs &rsaquo; Application &middot; Errors &amp; Warnings</p>
        </div>
      </div>
      <div class="chips" id="chips"></div>
    </div>
    <div class="heroact">
      <a class="ghbtn" href="https://github.com/muhamaddarulhadi/Application-Logs-Viewer" target="_blank" rel="noopener noreferrer" title="View on GitHub"><svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12"/></svg><span>Application-Logs-Viewer</span></a>
      <button class="iconbtn" id="theme" title="Toggle dark / light mode" aria-label="Toggle theme"></button>
    </div>
  </div>
</header>

<main>
  <section class="kpis" id="kpis"></section>

  <section class="grid">
    <div class="panel wide">
      <h3>Events per day</h3>
      <div id="trend"></div>
      <div class="legend"><span><i style="background:var(--err)"></i>Error</span><span><i style="background:var(--warn)"></i>Warning</span><span id="critleg" hidden><i style="background:var(--crit)"></i>Critical</span></div>
    </div>
    <div class="panel">
      <h3>Severity split</h3>
      <div class="donutwrap" id="donut"></div>
    </div>
    <div class="panel">
      <h3>Top sources <span style="text-transform:none;letter-spacing:0;font-weight:400">(click to filter)</span></h3>
      <div class="topsrc" id="topsrc"></div>
    </div>
  </section>

  <section class="toolbar">
    <div class="search">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      <input type="search" id="q" placeholder="Search source, event ID, message...">
    </div>
    <div class="seg" id="seg"></div>
    <select id="src"><option value="">All sources</option></select>
    <button class="btn ghost" id="csv">CSV</button>
    <button class="btn primary" id="xls">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
      Download Excel (.xlsx)
    </button>
  </section>

  <section class="tablecard">
    <div class="scroll">
      <table>
        <thead><tr id="head"></tr></thead>
        <tbody id="body"></tbody>
      </table>
      <div class="empty" id="empty" hidden>
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5M8.5 11h5"/></svg>
        <div>No events match your filters.</div>
      </div>
    </div>
    <div class="pager">
      <span id="count"></span>
      <div class="ctl">
        <span>Rows</span>
        <select id="size"><option>25</option><option selected>50</option><option>100</option><option>250</option></select>
        <button id="prev">&lsaquo; Prev</button>
        <span id="pageinfo"></span>
        <button id="next">Next &rsaquo;</button>
      </div>
    </div>
  </section>
</main>
<footer><span id="foot"></span> &middot; <a href="https://github.com/muhamaddarulhadi/Application-Logs-Viewer" target="_blank" rel="noopener noreferrer">Application-Logs-Viewer</a> by <a href="https://github.com/muhamaddarulhadi" target="_blank" rel="noopener noreferrer">muhamaddarulhadi</a></footer>

<div class="overlay" id="overlay"></div>
<aside class="drawer" id="drawer" aria-hidden="true">
  <div class="dhead">
    <div>
      <span id="dpill"></span>
      <h2 id="dtitle"></h2>
      <div class="sub" id="dsub"></div>
    </div>
    <button class="iconbtn" id="dclose" style="background:var(--surface2);border-color:var(--line);color:var(--text)" aria-label="Close">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>
    </button>
  </div>
  <div class="dbody">
    <div class="facts" id="facts"></div>
    <h4>General info <button class="copy" id="copy">Copy</button></h4>
    <pre id="dmsg"></pre>
  </div>
  <div class="dfoot">
    <button class="btn ghost" id="dprev">&lsaquo; Previous</button>
    <button class="btn ghost" id="dnext">Next &rsaquo;</button>
  </div>
</aside>

<script id="data" type="application/json">__DATA__</script>
<script id="metadata" type="application/json">__META__</script>
<script>
const data = JSON.parse(document.getElementById('data').textContent);
const meta = JSON.parse(document.getElementById('metadata').textContent);
const $ = id => document.getElementById(id);
const el = (tag, props = {}, ...kids) => { const n = Object.assign(document.createElement(tag), props); kids.forEach(k => n.append(k)); return n; };
const NS = 'http://www.w3.org/2000/svg';
const svg = (tag, attrs = {}, text) => { const n = document.createElementNS(NS, tag); for (const k in attrs) n.setAttribute(k, attrs[k]); if (text !== undefined) n.textContent = text; return n; };
const lvlClass = { Error:'err', Warning:'warn', Critical:'crit' };
const fmt = n => n.toLocaleString();

/* ---------- theme ---------- */
const sun = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
const moon = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';
function setTheme(t) { document.documentElement.dataset.theme = t; $('theme').innerHTML = t === 'dark' ? sun : moon; try { localStorage.setItem('evtheme', t); } catch (e) {} }
let saved = null; try { saved = localStorage.getItem('evtheme'); } catch (e) {}
setTheme(saved || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
$('theme').onclick = () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');

/* ---------- hero chips & footer ---------- */
[['Computer', meta.Computer], ['From', meta.From], ['To', meta.To], ['Generated', meta.Generated]]
  .forEach(([k, v]) => $('chips').append(el('span', {className:'chip'}, el('span', {textContent:k + ':'}), el('b', {textContent:v}))));
$('foot').textContent = `Generated by Get-AppEventReport.ps1 on ${meta.Generated}`;

/* ---------- KPIs ---------- */
const countBy = k => data.filter(d => d.Level === k).length;
const ic = {
  all:'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h16M4 12h16M4 18h10"/></svg>',
  err:'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg>',
  warn:'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0zM12 9v4M12 17h.01"/></svg>',
  crit:'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m13 2-9 12h7l-1 8 9-12h-7z"/></svg>',
  src:'<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>'
};
const kpis = [['k-all','all','Total events',data.length],['k-err','err','Errors',countBy('Error')],['k-warn','warn','Warnings',countBy('Warning')]];
if (countBy('Critical')) kpis.push(['k-crit','crit','Critical',countBy('Critical')]);
kpis.push(['k-src','src','Unique sources',new Set(data.map(d => d.Source)).size]);
kpis.forEach(([cls, icon, label, n]) => {
  const ico = el('div', {className:'ico'}); ico.innerHTML = ic[icon];
  $('kpis').append(el('div', {className:'kpi ' + cls}, ico, el('div', {}, el('div', {className:'n', textContent:fmt(n)}), el('div', {className:'l', textContent:label}))));
});

/* ---------- trend chart (stacked bars per day) ---------- */
(function trend() {
  const byDay = {};
  data.forEach(d => { const k = d.Date.slice(0, 10); (byDay[k] ||= {Error:0, Warning:0, Critical:0}); byDay[k][d.Level] = (byDay[k][d.Level] || 0) + 1; });
  const days = Object.keys(byDay).sort();
  const W = 640, H = 200, pl = 34, pr = 8, pt = 10, pb = 28;
  const max = Math.max(...days.map(k => byDay[k].Error + byDay[k].Warning + byDay[k].Critical));
  const nice = Math.max(1, Math.ceil(max / 4) * 4);
  const s = svg('svg', {viewBox:`0 0 ${W} ${H}`, width:'100%', role:'img'});
  for (let i = 0; i <= 4; i++) {
    const y = pt + (H - pt - pb) * (1 - i / 4);
    s.append(svg('line', {x1:pl, x2:W - pr, y1:y, y2:y, class:'gridline'}), svg('text', {x:pl - 6, y:y + 3, 'text-anchor':'end', class:'axis'}, Math.round(nice * i / 4)));
  }
  const slot = (W - pl - pr) / days.length, bw = Math.min(34, slot * 0.68);
  const step = Math.ceil(days.length / 10);
  days.forEach((k, i) => {
    const x = pl + slot * i + (slot - bw) / 2;
    let y = H - pb;
    [['Error','bar-err'],['Warning','bar-warn'],['Critical','bar-crit']].forEach(([lv, cls]) => {
      const v = byDay[k][lv]; if (!v) return;
      const h = (H - pt - pb) * v / nice; y -= h;
      const r = svg('rect', {x, y, width:bw, height:h, rx:3, class:cls});
      r.append(svg('title', {}, `${k}: ${v} ${lv}${v > 1 ? 's' : ''}`));
      s.append(r);
    });
    if (i % step === 0) s.append(svg('text', {x:x + bw / 2, y:H - 9, 'text-anchor':'middle', class:'axis'}, k.slice(5)));
  });
  $('trend').append(s);
  if (countBy('Critical')) $('critleg').hidden = false;
})();

/* ---------- donut ---------- */
(function donut() {
  const parts = ['Error','Warning','Critical'].map(l => [l, countBy(l)]).filter(p => p[1]);
  const R = 52, C = 2 * Math.PI * R, total = data.length;
  const s = svg('svg', {viewBox:'0 0 140 140', width:140, height:140});
  s.append(svg('circle', {cx:70, cy:70, r:R, fill:'none', stroke:'var(--surface2)', 'stroke-width':16}));
  let off = 0;
  parts.forEach(([l, n]) => {
    const len = C * n / total;
    const c = svg('circle', {cx:70, cy:70, r:R, fill:'none', stroke:`var(--${lvlClass[l]})`, 'stroke-width':16, 'stroke-dasharray':`${Math.max(len - 2, 0)} ${C}`, 'stroke-dashoffset':-off, transform:'rotate(-90 70 70)', 'stroke-linecap':'butt'});
    c.append(svg('title', {}, `${l}: ${n}`)); s.append(c); off += len;
  });
  s.append(svg('text', {x:70, y:72, 'text-anchor':'middle', class:'donut-center'}, fmt(total)), svg('text', {x:70, y:88, 'text-anchor':'middle', class:'donut-sub'}, 'events'));
  const legend = el('div', {className:'topsrc', style:'flex:1;gap:8px'});
  parts.forEach(([l, n]) => legend.append(el('div', {style:'display:flex;justify-content:space-between;gap:10px'},
    el('span', {className:'pill ' + l, textContent:l}), el('b', {textContent:`${fmt(n)}  (${Math.round(100 * n / total)}%)`, style:'font-variant-numeric:tabular-nums'}))));
  $('donut').append(s, legend);
})();

/* ---------- top sources ---------- */
(function topSources() {
  const m = {}; data.forEach(d => m[d.Source] = (m[d.Source] || 0) + 1);
  const top = Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 6), max = top[0][1];
  top.forEach(([name, n]) => {
    const row = el('div', {className:'src-row', title:'Filter by ' + name},
      el('div', {className:'top'}, el('span', {textContent:name}), el('span', {textContent:fmt(n)})),
      el('div', {className:'track'}, el('div', {className:'fill', style:`width:${100 * n / max}%`})));
    row.onclick = () => { $('src').value = name; page = 0; render(); $('q').scrollIntoView({block:'center'}); };
    $('topsrc').append(row);
  });
})();

/* ---------- filters ---------- */
let levelFilter = '', page = 0, sortKey = 'Date', sortDir = -1, current = [], openIdx = -1;
[['', 'All'], ...['Error','Warning','Critical'].filter(l => countBy(l)).map(l => [l, l + 's'])].forEach(([v, label]) => {
  const b = el('button', {textContent:label, className:v === '' ? 'on' : ''});
  b.onclick = () => { levelFilter = v; [...$('seg').children].forEach(x => x.classList.toggle('on', x === b)); page = 0; render(); };
  $('seg').append(b);
});
[...new Set(data.map(d => d.Source))].sort().forEach(v => $('src').append(new Option(v, v)));
['q', 'src', 'size'].forEach(id => $(id).addEventListener('input', () => { page = 0; render(); }));

const cols = [['Date','Date / Time'],['Level','Level'],['Source','Source'],['EventID','Event ID'],['Category','Category'],['Message','General info']];
cols.forEach(([k, label]) => {
  const th = el('th', {}, label, el('span', {className:'arr', textContent:'\u2195'}));
  th.dataset.k = k;
  th.onclick = () => { sortDir = sortKey === k ? -sortDir : 1; sortKey = k; render(); };
  $('head').append(th);
});

function filtered() {
  const q = $('q').value.trim().toLowerCase(), sr = $('src').value;
  return data.filter(d => (!levelFilter || d.Level === levelFilter) && (!sr || d.Source === sr) &&
    (!q || (d.Source + ' ' + d.EventID + ' ' + d.Message + ' ' + d.Category).toLowerCase().includes(q)))
    .sort((a, b) => { const x = a[sortKey], y = b[sortKey]; return (typeof x === 'number' ? x - y : String(x).localeCompare(String(y))) * sortDir; });
}

function render() {
  current = filtered();
  const size = +$('size').value, pages = Math.max(1, Math.ceil(current.length / size));
  page = Math.min(page, pages - 1);
  const slice = current.slice(page * size, page * size + size), body = $('body');
  body.replaceChildren();
  slice.forEach((d, i) => {
    const tr = el('tr');
    const dt = el('td', {className:'date'}, d.Date.slice(0, 10), el('small', {textContent:d.Date.slice(11)}));
    tr.append(dt,
      el('td', {}, el('span', {className:'pill ' + d.Level, textContent:d.Level})),
      el('td', {className:'src', textContent:d.Source}),
      el('td', {}, el('span', {className:'eid', textContent:d.EventID})),
      el('td', {textContent:d.Category || '-', style:'color:var(--muted)'}),
      el('td', {className:'msg', textContent:d.Message.split('\n')[0]}));
    tr.onclick = () => openDrawer(page * size + i);
    body.append(tr);
  });
  [...$('head').children].forEach(th => { th.classList.toggle('sorted', th.dataset.k === sortKey); th.lastChild.textContent = th.dataset.k === sortKey ? (sortDir > 0 ? '\u2191' : '\u2193') : '\u2195'; });
  $('empty').hidden = current.length > 0;
  $('count').textContent = `Showing ${current.length ? page * size + 1 : 0}\u2013${Math.min(current.length, (page + 1) * size)} of ${fmt(current.length)} filtered (${fmt(data.length)} total)`;
  $('pageinfo').textContent = `${page + 1} / ${pages}`;
  $('prev').disabled = page === 0; $('next').disabled = page >= pages - 1;
}
$('prev').onclick = () => { page--; render(); };
$('next').onclick = () => { page++; render(); };

/* ---------- detail drawer ---------- */
function openDrawer(idx) {
  if (idx < 0 || idx >= current.length) return;
  openIdx = idx; const d = current[idx];
  $('dpill').replaceChildren(el('span', {className:'pill ' + d.Level, textContent:d.Level}));
  $('dtitle').textContent = d.Source;
  $('dsub').textContent = `Event ID ${d.EventID}  \u00b7  ${d.Date}`;
  $('facts').replaceChildren();
  [['Date / Time', d.Date], ['Level', d.Level], ['Source', d.Source], ['Event ID', d.EventID], ['Category', d.Category || '-'], ['Computer', d.Computer], ['Record ID', d.RecordId], ['Log', 'Application']]
    .forEach(([k, v]) => $('facts').append(el('div', {className:'fact'}, el('span', {textContent:k}), el('b', {textContent:v}))));
  $('dmsg').textContent = d.Message;
  $('copy').textContent = 'Copy';
  $('dprev').disabled = idx === 0; $('dnext').disabled = idx === current.length - 1;
  $('drawer').classList.add('on'); $('overlay').classList.add('on'); $('drawer').setAttribute('aria-hidden', 'false');
}
function closeDrawer() { $('drawer').classList.remove('on'); $('overlay').classList.remove('on'); $('drawer').setAttribute('aria-hidden', 'true'); }
$('dclose').onclick = closeDrawer; $('overlay').onclick = closeDrawer;
$('dprev').onclick = () => openDrawer(openIdx - 1); $('dnext').onclick = () => openDrawer(openIdx + 1);
document.addEventListener('keydown', e => {
  if (!$('drawer').classList.contains('on')) return;
  if (e.key === 'Escape') closeDrawer();
  if (e.key === 'ArrowLeft') openDrawer(openIdx - 1);
  if (e.key === 'ArrowRight') openDrawer(openIdx + 1);
});
$('copy').onclick = () => {
  const t = $('dmsg').textContent;
  const done = () => $('copy').textContent = 'Copied!';
  if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(done, () => {});
  else { const ta = el('textarea', {value:t, style:'position:fixed;opacity:0'}); document.body.append(ta); ta.select(); try { document.execCommand('copy'); done(); } catch (e) {} ta.remove(); }
};

/* ---------- exports (use the currently filtered rows) ---------- */
const allCols = ['Date','Level','Source','EventID','Category','Computer','RecordId','Message'];
const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function download(name, mime, content) {
  const url = URL.createObjectURL(new Blob([content], {type:mime}));
  const a = el('a', {href:url, download:name}); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
const stamp = () => new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
/* ---- real .xlsx writer (no libraries, works offline) ---- */
const crcTable = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = crcTable[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
function zipStore(files) {
  const enc = new TextEncoder(), chunks = [], central = []; let offset = 0;
  files.forEach(f => {
    const name = enc.encode(f.name), data = enc.encode(f.text), crc = crc32(data), size = data.length;
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
    lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true); lh.setUint32(14, crc, true); lh.setUint32(18, size, true);
    lh.setUint32(22, size, true); lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    chunks.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(10, 0, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, size, true); ch.setUint32(24, size, true); ch.setUint16(28, name.length, true);
    ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + size;
  });
  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...chunks, ...central, new Uint8Array(end.buffer)], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
const xmlEsc = s => String(s ?? '').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function buildXlsx(rows) {
  const L = 'ABCDEFGH', numeric = ['EventID', 'RecordId'], lvlStyle = {Error:3, Warning:4, Critical:5};
  let sd = '<row r="1" ht="22" customHeight="1">' + allCols.map((c, i) => `<c r="${L[i]}1" s="1" t="inlineStr"><is><t>${c}</t></is></c>`).join('') + '</row>';
  rows.forEach((d, r) => {
    const n = r + 2;
    sd += `<row r="${n}">` + allCols.map((c, i) => {
      const ref = L[i] + n;
      if (numeric.includes(c)) return `<c r="${ref}"><v>${Number(d[c]) || 0}</v></c>`;
      const st = c === 'Level' ? (lvlStyle[d.Level] || 0) : c === 'Message' ? 2 : 0;
      return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(String(d[c] ?? '').slice(0, 32000))}</t></is></c>`;
    }).join('') + '</row>';
  });
  const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const sheet = H + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="11" customWidth="1"/><col min="3" max="3" width="28" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/><col min="6" max="6" width="22" customWidth="1"/><col min="7" max="7" width="12" customWidth="1"/><col min="8" max="8" width="100" customWidth="1"/></cols><sheetData>' + sd + '</sheetData><autoFilter ref="A1:H' + (rows.length + 1) + '"/></worksheet>';
  const fill = c => `<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor indexed="64"/></patternFill></fill>`;
  const xf = (fillId, wrap, extra = '') => `<xf numFmtId="0" fontId="${extra ? 1 : 0}" fillId="${fillId}" borderId="0" xfId="0"${extra ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''} applyAlignment="1"><alignment vertical="${extra ? 'center' : 'top'}"${wrap ? ' wrapText="1"' : ''}/></xf>`;
  const styles = H + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' + fill('FF4F46E5') + fill('FFFFE4EA') + fill('FFFEF3C7') + fill('FFEDE9FE') + '</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6">' + xf(0, false) + xf(2, false, 'hdr') + xf(0, true) + xf(3, false) + xf(4, false) + xf(5, false) + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const R = 'http://schemas.openxmlformats.org/package/2006/relationships', O = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  return zipStore([
    {name:'[Content_Types].xml', text:H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'},
    {name:'_rels/.rels', text:H + `<Relationships xmlns="${R}"><Relationship Id="rId1" Type="${O}/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
    {name:'xl/workbook.xml', text:H + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${O}"><sheets><sheet name="Errors and Warnings" sheetId="1" r:id="rId1"/></sheets></workbook>`},
    {name:'xl/_rels/workbook.xml.rels', text:H + `<Relationships xmlns="${R}"><Relationship Id="rId1" Type="${O}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${O}/styles" Target="styles.xml"/></Relationships>`},
    {name:'xl/styles.xml', text:styles},
    {name:'xl/worksheets/sheet1.xml', text:sheet}
  ]);
}
$('xls').onclick = () => download(`AppEvents_${stamp()}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buildXlsx(current));
$('csv').onclick = () => {
  const q = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const csv = [allCols.map(q).join(',')].concat(current.map(d => allCols.map(c => q(d[c])).join(','))).join('\r\n');
  download(`AppEvents_${stamp()}.csv`, 'text/csv;charset=utf-8', '\ufeff' + csv);
};

render();
</script>
</body>
</html>
'@

$html = $template.Replace('__DATA__', $dataJson).Replace('__META__', $metaJson)
[System.IO.File]::WriteAllText($htmlPath, $html, (New-Object System.Text.UTF8Encoding($true)))

# ---------- 6. Done ----------
Write-Host ""
Write-Host "Report created:" -ForegroundColor Green
Write-Host "  HTML : $htmlPath"
Write-Host "  CSV  : $csvPath"
if ($xlsxMade) { Write-Host "  XLSX : $xlsxPath" }
Write-Host ""

if ($OpenReport -or (-not $PSBoundParameters.ContainsKey('StartDate'))) {
    Start-Process $htmlPath
}
