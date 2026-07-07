# =====================================================================
#  UM6P Patent Dashboard — one-click data refresh
#  Run this after replacing any of the source data files.
#  It (1) checks the required files are present, (2) recomputes all
#  metrics into dashboard_data.js, and (3) builds a standalone
#  single-file dashboard (UM6P_Dashboard.html) you can open or share.
# =====================================================================
$ErrorActionPreference = 'Stop'
$base = Split-Path $PSScriptRoot -Parent

function Line { param($c='Gray',$t) Write-Host $t -ForegroundColor $c }

Line Cyan  "`n============================================================"
Line Cyan  "  UM6P Patent Dashboard - refreshing from source data"
Line Cyan  "============================================================`n"

# ---- 1. Check the data files are present (exact names matter) ----
$required = @(
    'File 3 - Master Dashboard + Dataset - UM6P - Morocco-only.xlsx'   # core inventory (426 patents)
)
$recommended = @(
    'File 4 - UM6P Patents Data - Raw - Morocco-only.xlsx',
    'File 5 - Patent Renewal Data - MAScIR - Morocco-only.xlsx',
    'File 6 - Patent Family Data - Escapenet - International.xlsx',
    'File 7 - Citation Data - Orbit - International.xlsx',
    'File 8 - PCT - WIPO - International.xls',
    'morocco_assignees_tech_patent_counts.csv',
    'tunisia_assignees_tech_patent_counts.csv',
    'south_africa_assignees_tech_patent_counts.csv'
)

$missingCore = @($required | Where-Object { -not (Test-Path (Join-Path $base $_)) })
if ($missingCore.Count) {
    Line Red "ERROR - the core data file is missing. Cannot continue:"
    $missingCore | ForEach-Object { Line Red "   x  $_" }
    Line Yellow "`nMake sure the file name matches EXACTLY (including spaces and dashes)."
    exit 1
}
$missingRec = @($recommended | Where-Object { -not (Test-Path (Join-Path $base $_)) })
if ($missingRec.Count) {
    Line Yellow "Note - these files are not present; the sections that use them will be blank or reduced:"
    $missingRec | ForEach-Object { Line Yellow "   -  $_" }
    Line Gray ""
}

# ---- 2. Recompute all metrics (writes dashboard_data.js) ----
Line White "Reading the data files and recomputing metrics..."
try {
    & (Join-Path $PSScriptRoot 'build_dashboard_data.ps1')
} catch {
    Line Red "`nERROR while reading the Excel files:"
    Line Red "   $($_.Exception.Message)"
    Line Yellow "`nMost common cause: the Microsoft Access Database Engine is not installed."
    Line Yellow "Fix: install 'Microsoft Access Database Engine 2016 Redistributable' (64-bit),"
    Line Yellow "or make sure Microsoft Office / Excel is installed on this machine."
    exit 1
}

# ---- 3. Build the standalone single-file dashboard ----
Line White "`nBuilding the standalone dashboard file..."
$htmlPath = Join-Path $base 'um6p_dashboard_v3.html'
$dataPath = Join-Path $base 'dashboard_data.js'
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$data = Get-Content $dataPath -Raw -Encoding UTF8
$needle = '<script src="dashboard_data.js"></script>'
if ($html.Contains($needle)) {
    $out = $html.Replace($needle, "<script>`r`n$data`r`n</script>")
    $outPath = Join-Path $base 'UM6P_Dashboard.html'
    Set-Content -Path $outPath -Value $out -Encoding UTF8
    Line Green "   OK  -> UM6P_Dashboard.html  (open this file in a web browser)"
} else {
    Line Yellow "   Could not inline the data automatically; open um6p_dashboard_v3.html instead."
}

Line Green "`nDone. Open UM6P_Dashboard.html to view the updated dashboard.`n"
