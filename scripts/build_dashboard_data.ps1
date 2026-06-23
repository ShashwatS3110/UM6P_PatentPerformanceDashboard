# Extract UM6P dashboard metrics from Files 3-8 into dashboard_data.js (OLEDB)
$ErrorActionPreference = 'Stop'
$base = Split-Path $PSScriptRoot -Parent

function Esc-Json([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace '\\','\\' -replace '"','\"' -replace "`r",'' -replace "`n",' ' -replace "`t",' ').Trim()
}

function Short-Title([string]$t, [int]$max = 72) {
    $t = (Esc-Json $t)
    if ($t.Length -le $max) { return $t }
    return $t.Substring(0, $max - 1) + [char]0x2026
}

function Theme-Short([string]$t) {
    if ($t -match 'Environmental') { return 'Environmental Tech' }
    if ($t -match 'Green Agriculture') { return 'Sustainable Agriculture' }
    if ($t -match 'Mineral Processing') { return 'Mineral Processing' }
    if ($t -match 'Circular Chemistry') { return 'Circular Chemistry' }
    if ($t -match 'Sustainable Energy') { return "Energy & Storage" }
    return (Esc-Json $t)
}

function Lab-Bucket([string]$lab) {
    $lab = $lab.Trim()
    if ($lab -match 'MAScIR') { return 'MAScIR' }
    if ($lab -match 'OCP|CERPHOS') { return 'OCP/CERPHOS' }
    return 'UM6P TTO'
}

function Parse-Date([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s) -or $s -eq '/') { return $null }
    $s = $s.Trim()
    foreach ($fmt in @('dd/MM/yyyy','dd-MM-yyyy','yyyy-MM-dd','M/d/yyyy','d/M/yyyy')) {
        try { return [datetime]::ParseExact($s, $fmt, $null) } catch {}
    }
    try { return [datetime]::Parse($s) } catch { return $null }
}

function Open-Excel([string]$path, [string]$ext = 'Xml') {
    $cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$path;Extended Properties='Excel 12.0 $ext;HDR=NO;IMEX=1';"
    $conn = New-Object System.Data.OleDb.OleDbConnection $cs
    $conn.Open()
    return $conn
}

function Get-SheetName([System.Data.OleDb.OleDbConnection]$conn, [string]$pattern) {
    $tables = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    ($tables | Where-Object { $_.TABLE_NAME -like "*$pattern*" -and $_.TABLE_NAME -notlike "*Print*" } | Select-Object -First 1).TABLE_NAME
}

function Query-Excel([System.Data.OleDb.OleDbConnection]$conn, [string]$sheetPattern, [string]$where = '') {
    $sn = Get-SheetName $conn $sheetPattern
    if (-not $sn) { return @() }
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT * FROM [$sn] $where"
    $da = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt.Rows
}

function Cell($row, [int]$i) {
    if ($null -eq $row -or $i -ge $row.Table.Columns.Count) { return '' }
    $v = $row[$i]
    if ($null -eq $v) { return '' }
    return "$v".Trim()
}

# ---- File 3 ----
$conn3 = Open-Excel (Join-Path $base 'File 3 - Master Dashboard + Dataset - UM6P - Morocco-only.xlsx')
$invRows = Query-Excel $conn3 'Inventaire Complet'
$paysRows = Query-Excel $conn3 'PAYS DASHBOARD'

$patents = [System.Collections.Generic.List[object]]::new()
foreach ($row in $invRows) {
    $id = Cell $row 0
    if ($id -notmatch '^[A-Z]{3}-\d+') { continue }
    $theme = Cell $row 1
    $title = Cell $row 3
    $lab = Cell $row 5
    $dateStr = Cell $row 8
    $yearRaw = Cell $row 9
    $statut = Cell $row 11
    if (-not $statut) { $statut = Cell $row 10 }
    $geo = Cell $row 12
    $partners = Cell $row 15
    $year = 0
    if ($yearRaw -match '(\d{4})') { [void][int]::TryParse($Matches[1], [ref]$year) }
    if ($year -eq 0) { $d = Parse-Date $dateStr; if ($d) { $year = $d.Year } }
    $patents.Add([pscustomobject]@{
        id=$id; theme=$theme; title=$title; lab=$lab; year=$year; date=$dateStr
        statut=$statut; geo=$geo; partners=$partners; bucket=(Lab-Bucket $lab)
    })
}

$countries = [System.Collections.Generic.List[object]]::new()
foreach ($row in $paysRows) {
    $rank = Cell $row 0
    $name = Cell $row 1
    $cnt = Cell $row 2
    if (-not $name) { $name = Cell $row 2; $cnt = Cell $row 3 }
    if ($rank -match '^\d+$' -and $name -and $cnt -match '^\d+$') {
        $countries.Add([pscustomobject]@{ country=$name; count=[int]$cnt })
    } elseif ($name -and $cnt -match '^\d+$' -and $name -notmatch 'PAYS|Rang|R.partition') {
        if ($countries | Where-Object { $_.country -eq $name }) { continue }
        $countries.Add([pscustomobject]@{ country=$name; count=[int]$cnt })
    }
}
$conn3.Close()

# ---- File 5 ----
$conn5 = Open-Excel (Join-Path $base 'File 5 - Patent Renewal Data - MAScIR - Morocco-only.xlsx')
$renRows = Query-Excel $conn5 'Portefeuille'
$renewals = [System.Collections.Generic.List[object]]::new()
foreach ($row in $renRows) {
    $title = Cell $row 2
    if (-not $title) { continue }
    $filed = Parse-Date (Cell $row 5)
    $granted = Parse-Date (Cell $row 6)
    $status = Cell $row 8
    $years = $null
    if ($filed -and $granted) { $years = [math]::Round(($granted - $filed).TotalDays / 365.25, 1) }
    $active = ($status -notmatch 'DECHU|D.chu')
    $renewals.Add([pscustomobject]@{
        ref=(Cell $row 0); title=$title
        filed= if ($filed) { $filed.ToString('yyyy-MM-dd') } else { '' }
        granted= if ($granted) { $granted.ToString('yyyy-MM-dd') } else { '' }
        years=$years; active=$active; status=$status
    })
}
$conn5.Close()

# ---- File 6 ----
$conn6 = Open-Excel (Join-Path $base 'File 6 - Patent Family Data - Escapenet - International.xlsx')
$famRows = Query-Excel $conn6 'sultat'
$families = [System.Collections.Generic.List[object]]::new()
$first = $true
foreach ($row in $famRows) {
    if ($first) { $first = $false; continue }
    $title = Cell $row 1
    if (-not $title) { continue }
    $cib = Cell $row 6
    $cpc = Cell $row 7
    $prio = Cell $row 5
    $pubs = Cell $row 4
    $ipcCodes = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($block in @($cib, $cpc)) {
        foreach ($m in [regex]::Matches($block, '[A-HY]\d{2}[A-Z]\d{2}/\d{2}')) {
            [void]$ipcCodes.Add($m.Value.Substring(0, 4))
        }
    }
    $countrySet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($m in [regex]::Matches($cpc, '\(([A-Z]{2}(?:,[A-Z]{2})*)\)')) {
        foreach ($c in $m.Groups[1].Value.Split(',')) { [void]$countrySet.Add($c.Trim()) }
    }
    foreach ($code in @('EP','US','JP','MA','WO','FR')) {
        if ($pubs -match "\b$code\b" -or $cpc -match "\b$code\b") { [void]$countrySet.Add($code) }
    }
    $families.Add([pscustomobject]@{
        title=$title; prio=$prio; ipcCount=$ipcCodes.Count
        countries=@($countrySet); countryCount=$countrySet.Count
        hasEP=$countrySet.Contains('EP'); hasUS=$countrySet.Contains('US'); hasJP=$countrySet.Contains('JP')
    })
}
$conn6.Close()

# ---- File 7 ----
$conn7 = Open-Excel (Join-Path $base 'File 7 - Citation Data - Orbit - International.xlsx')
$citeRows = Query-Excel $conn7 'SHEET'
$citations = [System.Collections.Generic.List[object]]::new()
$first = $true
foreach ($row in $citeRows) {
    if ($first) { $first = $false; continue }
    $title = Cell $row 6
    if (-not $title) { continue }
    $title = ($title -replace '^\([A-Z0-9]+\)\s*','').Trim()
    $prio = Cell $row 5
    $citeRaw = Cell $row 10
    $cite = 0
    if ($citeRaw -match '^\d+$') { $cite = [int]$citeRaw }
    $citations.Add([pscustomobject]@{ title=$title; prio=$prio; citations=$cite })
}
$conn7.Close()

# ---- File 8 (.xls) ----
$pctCount8 = 0
try {
    $conn8 = Open-Excel (Join-Path $base 'File 8 - PCT - WIPO - International.xls') 'HTML'
    $tables = $conn8.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Tables, $null)
    foreach ($t in $tables.TABLE_NAME) {
        if ($t -match 'Print') { continue }
        $cmd = $conn8.CreateCommand()
        $cmd.CommandText = "SELECT * FROM [$t]"
        $da = New-Object System.Data.OleDb.OleDbDataAdapter $cmd
        $dt = New-Object System.Data.DataTable
        [void]$da.Fill($dt)
        $pctCount8 += ($dt.Rows | Where-Object { "$($_[5])" -match 'PCT|WO' -or "$($_[3])" -match 'WO' }).Count
    }
    $conn8.Close()
} catch { $pctCount8 = 0 }

# ---- File 4 ----
$conn4 = Open-Excel (Join-Path $base 'File 4 - UM6P Patents Data - Raw - Morocco-only.xlsx')
$ttoRows = Query-Excel $conn4 'UM6P Patents filed by TTO'
$ocpPatents = [System.Collections.Generic.List[object]]::new()
foreach ($row in $ttoRows) {
    $title = Cell $row 5
    if (-not $title) { $title = Cell $row 6 }
    $date = Cell $row 1
    $lab = Cell $row 6
    $partners = Cell $row 8
    if (-not $partners) { $partners = Cell $row 9 }
    $y = 0; if ($date -match '(\d{4})') { $y = [int]$Matches[1] }
    if ($title) {
        if ($partners -match 'OCP') { $ocpPatents.Add([pscustomobject]@{ title=$title; year=$y; source='TTO' }) }
    }
}
foreach ($pat in @('CBS- OCP','MSN-OCP')) {
    $rows = Query-Excel $conn4 $pat
    foreach ($row in $rows) {
        $title = Cell $row 5
        if (-not $title) { $title = Cell $row 4 }
        $date = Cell $row 8
        $y = 0; if ($date -match '(\d{4})') { $y = [int]$Matches[1] }
        if ($title) { $ocpPatents.Add([pscustomobject]@{ title=$title; year=$y; source=$pat }) }
    }
}
$conn4.Close()

# ---- Aggregations ----
$totalPatents = $patents.Count
$datedPatents = @($patents | Where-Object { $_.year -gt 0 })
$filingsUndated = $totalPatents - $datedPatents.Count
$yearNums = @($datedPatents | ForEach-Object { $_.year })
$yearMin = if ($yearNums.Count) { ($yearNums | Measure-Object -Minimum).Minimum } else { 2015 }
$yearMax = if ($yearNums.Count) { ($yearNums | Measure-Object -Maximum).Maximum } else { 2025 }
if ($yearMin -gt 2010) { $yearMin = 2010 }
if ($yearMax -lt 2025) { $yearMax = 2025 }
$years = $yearMin..$yearMax
$filingsByYear = @{}; $filingsMascir = @{}; $filingsTto = @{}; $filingsOther = @{}
foreach ($y in $years) { $filingsByYear[$y]=0; $filingsMascir[$y]=0; $filingsTto[$y]=0; $filingsOther[$y]=0 }
foreach ($p in $patents) {
    if ($p.year -ge $yearMin -and $p.year -le $yearMax) {
        $filingsByYear[$p.year]++
        switch ($p.bucket) {
            'MAScIR' { $filingsMascir[$p.year]++ }
            'UM6P TTO' { $filingsTto[$p.year]++ }
            default { $filingsOther[$p.year]++ }
        }
    }
}

$themeKeys = @(
    "Environmental Tech & Digital Mining",
    "Green Agriculture & Biotechnology",
    "Mineral Processing & Strategic Synthesis",
    "Circular Chemistry & Waste Valorization",
    "Sustainable Energy Storage & Management"
)
$themeMap = @{}
foreach ($k in $themeKeys) { $themeMap[$k] = 0 }
$otherCount = 0
foreach ($p in $patents) {
    if ($themeMap.ContainsKey($p.theme)) { $themeMap[$p.theme]++ } else { $otherCount++ }
}

$pctCount = ($patents | Where-Object { $_.statut -match 'PCT' -or $_.geo -match 'PCT|WIPO|WO' }).Count
if ($pctCount8 -gt $pctCount) { $pctCount = $pctCount8 }
$grantedCount = ($patents | Where-Object { $_.statut -match 'Maintenu|D.livr|Délivré|Granted|Delivered' }).Count
$pctRate = if ($totalPatents) { [math]::Round(100 * $pctCount / $totalPatents, 0) } else { 0 }
$grantRate = if ($totalPatents) { [math]::Round(100 * $grantedCount / $totalPatents, 0) } else { 0 }

$grantLags = @($renewals | Where-Object { $null -ne $_.years } | ForEach-Object { $_.years })
$avgGrantLag = if ($grantLags.Count) { [math]::Round(($grantLags | Measure-Object -Average).Average, 1) } else { 0 }
$renewalTotal = $renewals.Count
$renewalActiveCount = @($renewals | Where-Object { $_.active }).Count
$renewalLapsedCount = $renewalTotal - $renewalActiveCount
$renewalRate = if ($renewalTotal) { [math]::Round(100 * $renewalActiveCount / $renewalTotal, 0) } else { 0 }

$avgCites = if ($citations.Count) { [math]::Round(($citations.citations | Measure-Object -Average).Average, 1) } else { 0 }
$pctCited = if ($citations.Count) { [math]::Round(100 * ($citations | Where-Object { $_.citations -gt 0 }).Count / $citations.Count, 0) } else { 0 }
$avgIpc = if ($families.Count) { [math]::Round(($families.ipcCount | Measure-Object -Average).Average, 1) } else { 0 }
$triadicShare = 0
if ($families.Count) {
    $tri = ($families | Where-Object { $_.hasEP -and $_.hasUS -and $_.hasJP }).Count
    $triadicShare = [math]::Round(100 * $tri / $families.Count, 1)
}

$v0 = [math]::Max(1, $filingsByYear[2022])
$v1 = [math]::Max(1, $filingsByYear[2025])
$cagr = [math]::Round((100 * ([math]::Pow($v1 / $v0, 1 / 3) - 1)), 0)

$strategyTargets = @{
    "Environmental Tech & Digital Mining" = 20
    "Green Agriculture & Biotechnology" = 22
    "Mineral Processing & Strategic Synthesis" = 18
    "Circular Chemistry & Waste Valorization" = 20
    "Sustainable Energy Storage & Management" = 15
}
$alignGap = @{}
$alignShares = @{}
foreach ($k in $themeKeys) {
    $share = [math]::Round(100 * $themeMap[$k] / [math]::Max(1,$totalPatents), 0)
    $alignShares[$k] = $share
    $alignGap[$k] = $share - $strategyTargets[$k]
}
$avgAbsGap = if ($alignGap.Count) { ($alignGap.Values | ForEach-Object { [math]::Abs([double]$_) } | Measure-Object -Average).Average } else { 0 }

function Patent-Market($p) {
    $geo = if ($p.geo) { $p.geo.Trim() } else { '' }
    if ($geo) {
        if ($geo -match '^Maroc$') { return 'Morocco' }
        if ($geo -match 'PCT|WIPO') { return 'WIPO/PCT' }
        if ($geo -match ';') { return 'Multi-country' }
        return $geo
    }
    if ($p.statut -match 'PCT|WIPO') { return 'WIPO/PCT' }
    return 'Morocco'
}
$countryPatent = @{}
foreach ($p in $patents) {
    $m = Patent-Market $p
    if (-not $countryPatent[$m]) { $countryPatent[$m] = 0 }
    $countryPatent[$m]++
}

$partnerYears = 2019..2025
$partnerOcp=@{}; $partnerOther=@{}; $partnerAcad=@{}; $partnerSolo=@{}
foreach ($y in $partnerYears) { $partnerOcp[$y]=0; $partnerOther[$y]=0; $partnerAcad[$y]=0; $partnerSolo[$y]=0 }
foreach ($p in $patents) {
    if ($p.year -lt 2019 -or $p.year -gt 2025) { continue }
    $co = $p.partners
    if ($co -match 'OCP') { $partnerOcp[$p.year]++ }
    elseif ($co -match 'Universit|Hassan|Cadi|Ibn|ENSA|ESITH|CNESTEN') { $partnerAcad[$p.year]++ }
    elseif ($co -and $co.Trim()) { $partnerOther[$p.year]++ }
    else { $partnerSolo[$p.year]++ }
}
$withPartner = @($patents | Where-Object { $_.partners -and $_.partners.Trim() })
$ocpCoTotal = @($withPartner | Where-Object { $_.partners -match 'OCP' }).Count
$coTotal = $withPartner.Count
$ocpShare = if ($coTotal) { [math]::Round(100 * $ocpCoTotal / $coTotal, 0) } else { 0 }
$hhi = 0
if ($coTotal) {
    $groups = $withPartner | ForEach-Object {
        if ($_.partners -match 'OCP') { 'OCP' } else { 'Other partner' }
    } | Group-Object
    foreach ($g in $groups) { $hhi += [math]::Pow(100 * $g.Count / $coTotal, 2) }
    $hhi = [math]::Round($hhi / 10000, 2)
}

# Match citations to themes by title prefix
$citeThemeAvg = @{}
foreach ($k in $themeKeys) { $citeThemeAvg[$k] = [System.Collections.Generic.List[double]]::new() }
foreach ($c in $citations) {
    $matched = $false
    foreach ($p in $patents) {
        if ($p.title.Length -ge 12 -and $c.title.Length -ge 12 -and $p.title.Substring(0,12).ToUpper() -eq $c.title.Substring(0,12).ToUpper()) {
            [void]$citeThemeAvg[$p.theme].Add($c.citations); $matched = $true; break
        }
    }
    if (-not $matched) { [void]$citeThemeAvg["Environmental Tech & Digital Mining"].Add($c.citations) }
}
$citeVals = @($themeKeys | ForEach-Object { if ($citeThemeAvg[$_].Count) { [math]::Round(($citeThemeAvg[$_] | Measure-Object -Average).Average, 1) } else { 0 } })

function Dive([string]$key,[string]$title,[string]$sub,[array]$cols,[array]$rows,[bool]$illus) {
    return @{ key=$key; title=$title; sub=$sub; cols=$cols; rows=$rows; illustrative=$illus }
}

$dives = @{}
$dives['netFilings'] = Dive 'netFilings' 'Most Recent Patent Filings' '10 most recently filed applications (File 3)' @('Patent title','Year','Lab','Theme') @(
    @($patents | Sort-Object year, date -Descending | Select-Object -First 10 | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.year)", (Esc-Json $_.lab), (Theme-Short $_.theme))
    })
) $false

$dives['timeToGrant'] = Dive 'timeToGrant' 'Slowest Patents to Grant' 'MAScIR renewal portfolio - filing to grant lag (File 5)' @('Patent title','Filed','Granted','Years') @(
    @($renewals | Where-Object { $_.years } | Sort-Object years -Descending | Select-Object -First 5 | ForEach-Object {
        ,@((Short-Title $_.title), $_.filed, $_.granted, "$($_.years) yrs")
    })
) $false

$dives['forwardCitations'] = Dive 'forwardCitations' 'Most-Cited Patents' 'Orbit international citation export (File 7)' @('Patent title','Citations','Priority date') @(
    @($citations | Sort-Object citations -Descending | Select-Object -First 10 | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.citations)", $_.prio)
    })
) $false

$dives['techBreadth'] = Dive 'techBreadth' 'Broadest Patents by Technology Scope' 'Distinct IPC classes per family (File 6)' @('Patent title','IPC classes','Jurisdictions') @(
    @($families | Sort-Object ipcCount -Descending | Select-Object -First 5 | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.ipcCount) classes", ($_.countries -join ', '))
    })
) $false

$dives['globalFootprint'] = Dive 'globalFootprint' 'Most Globally Protected Inventions' 'Top families by jurisdiction count (File 6)' @('Patent title','Coverage','# Jurisdictions') @(
    @($families | Sort-Object countryCount -Descending | Select-Object -First 5 | ForEach-Object {
        ,@((Short-Title $_.title), ($_.countries -join ', '), "$($_.countryCount)")
    })
) $false

$dives['triadicList'] = Dive 'triadicList' 'Top Triadic Patent Candidates' 'Multi-jurisdiction families that could complete EU+US+JP (Files 6/8)' @('Patent title','Current coverage','Note') @(
    @($families | Where-Object { $_.countryCount -ge 2 } | Sort-Object countryCount -Descending | Select-Object -First 5 | ForEach-Object {
        $note = if ($_.hasEP -and $_.hasUS -and -not $_.hasJP) { 'Add Japan to complete triad' }
                elseif ($_.hasEP -or $_.hasUS) { 'Extend to EU+US+JP' }
                else { 'File PCT before priority deadline' }
        ,@((Short-Title $_.title), ($_.countries -join ', '), $note)
    })
) $false

$dives['renewal'] = Dive 'renewal' 'Longest-Active Patents' 'Oldest MAScIR patents still renewed (File 5)' @('Patent title','Filed','Status') @(
    @($renewals | Where-Object { $_.active -and $_.filed } | Sort-Object filed | Select-Object -First 5 | ForEach-Object {
        $yrs = [math]::Round(((Get-Date) - [datetime]::Parse($_.filed)).TotalDays / 365.25, 0)
        ,@((Short-Title $_.title), $_.filed, "Active ~$yrs yrs")
    })
) $false

$leadTheme = ($themeMap.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
$dives['themeTreemap'] = Dive 'themeTreemap' "Top Patents - $(Theme-Short $leadTheme)" 'Largest theme in File 3 inventory' @('Patent title','Year','Status') @(
    @($patents | Where-Object { $_.theme -eq $leadTheme } | Sort-Object year -Descending | Select-Object -First 10 | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.year)", (Esc-Json $_.statut))
    })
) $false

$dives['citeByTheme'] = Dive 'citeByTheme' 'Most-Cited Patents (Orbit sample)' 'File 7 citation-ranked list' @('Patent title','Citations','Priority') @(
    @($citations | Sort-Object citations -Descending | Select-Object -First 10 | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.citations)", $_.prio)
    })
) $false

$ocpTop = @($ocpPatents | Sort-Object year -Descending | Select-Object -First 5)
if (-not $ocpTop.Count) {
    $ocpTop = @($patents | Where-Object { $_.partners -match 'OCP' -or $_.lab -match 'OCP' } | Sort-Object year -Descending | Select-Object -First 5)
}
$dives['partnerTrend'] = Dive 'partnerTrend' 'Top OCP-Linked Patents' 'Co-owned or OCP-partner filings (Files 3/4)' @('Patent title','Year','Source') @(
    @($ocpTop | ForEach-Object {
        ,@((Short-Title $_.title), "$($_.year)", (Esc-Json $_.source))
    })
) $false

$dives['commercialization'] = Dive 'commercialization' 'Actively Licensed Patents' 'Awaiting TTO licensing registry' @('Patent title','Pathway','Partner') @() $true
$dives['dealsList'] = Dive 'dealsList' 'Recent Licensing Deals' 'Awaiting TTO deal log' @('Patent title','Year','Type','Sector') @() $true
$dives['commByTheme'] = Dive 'commByTheme' 'Commercialization by Theme' 'Awaiting TTO licensing data' @('Patent title','Theme','Pathway') @() $true
$dives['origIndex'] = Dive 'origIndex' 'Most Original Patents' 'Requires PATSTAT backward citations' @('Patent title','Score','Note') @() $true
$dives['spinoffsList'] = Dive 'spinoffsList' 'UM6P Spin-off Companies' 'Awaiting TTO venture registry' @('Company','Founded','Area','Sector') @() $true

$treemapLabels = @(
    "Environmental Tech & Digital Mining",
    'Sustainable Agriculture',
    'Mineral Processing',
    'Circular Chemistry',
    "Energy & Storage"
)
$treemapKeys = @(
    "Environmental Tech & Digital Mining",
    "Green Agriculture & Biotechnology",
    "Mineral Processing & Strategic Synthesis",
    "Circular Chemistry & Waste Valorization",
    "Sustainable Energy Storage & Management"
)
$treemap = @()
for ($i = 0; $i -lt $treemapLabels.Count; $i++) {
    $k = $treemapKeys[$i]
    $treemap += @{ label = $treemapLabels[$i]; value = [math]::Round(100 * $themeMap[$k] / [math]::Max(1, $totalPatents), 0) }
}
$treemapSum = ($treemap | ForEach-Object { $_.value } | Measure-Object -Sum).Sum
if ($treemapSum -ne 100 -and $treemap.Count) {
    $maxI = 0
    for ($i = 1; $i -lt $treemap.Count; $i++) { if ($treemap[$i].value -gt $treemap[$maxI].value) { $maxI = $i } }
    $adj = $treemap[$maxI].value - ($treemapSum - 100)
    $treemap[$maxI] = @{ label = $treemap[$maxI].label; value = $adj }
}

$alignLabels = @("Env. Tech & Mining", 'Sust. Agriculture', 'Mineral Processing', 'Circular Chemistry', "Energy & Storage")
$alignVals = @($themeKeys | ForEach-Object { $alignGap[$_] })
$alignShareVals = @($themeKeys | ForEach-Object { $alignShares[$_] })
$strategyTargetVals = @($themeKeys | ForEach-Object { $strategyTargets[$_] })
$countryTop = @($countryPatent.GetEnumerator() | ForEach-Object { [pscustomobject]@{ country = $_.Key; count = $_.Value } } | Sort-Object count -Descending)
if (-not $countryTop.Count) {
    # Fallback if geo parsing yields nothing
    $countryTop = @(
        @{ country='Morocco'; count=$totalPatents }
    )
}

$peakYear = 2015; $peakCount = 0
foreach ($y in $years) {
    $t = $filingsByYear[$y]
    if ($t -gt $peakCount) { $peakCount = $t; $peakYear = $y }
}

$countryTotal = 0
foreach ($c in $countryTop) { $countryTotal += $c.count }

function Score-Radar([double]$um, [double]$mor, [double]$top, [bool]$lower) {
    if ($lower) {
        $span = $mor - $top
        if ($span -le 0) { return 50 }
        return [math]::Max(0, [math]::Min(90, [math]::Round(100 * ($mor - $um) / $span)))
    }
    $span = $top - $mor
    if ($span -le 0) { return 50 }
    return [math]::Max(0, [math]::Min(90, [math]::Round(100 * ($um - $mor) / $span)))
}

$radarScores = @(
    (Score-Radar $cagr 14 35 $false),
    (Score-Radar $avgAbsGap 18 6 $true),
    (Score-Radar $pctRate 45 85 $false),
    (Score-Radar $avgCites 0.7 1.8 $false),
    (Score-Radar 17 8 30 $false),
    (Score-Radar $ocpShare 75 35 $true)
)
$radarBenchMorocco = @(
    (Score-Radar 14 14 35 $false),
    (Score-Radar 12 18 6 $true),
    (Score-Radar 77 45 85 $false),
    (Score-Radar 0.9 0.7 1.8 $false),
    (Score-Radar 8 8 30 $false),
    (Score-Radar 55 75 35 $true)
)
$radarBenchAfrica = @(
    (Score-Radar 11 14 35 $false),
    (Score-Radar 14 18 6 $true),
    (Score-Radar 45 45 85 $false),
    (Score-Radar 0.7 0.7 1.8 $false),
    (Score-Radar 6 8 30 $false),
    (Score-Radar 60 75 35 $true)
)
$radarBenchTop = @(
    (Score-Radar 35 14 35 $false),
    (Score-Radar 6 18 6 $true),
    (Score-Radar 85 45 85 $false),
    (Score-Radar 1.8 0.7 1.8 $false),
    (Score-Radar 30 8 30 $false),
    (Score-Radar 35 75 35 $true)
)
$radarActuals = @(
    "CAGR +$cagr% ($($filingsByYear[2025]) filings in 2025)",
    "Avg |gap| $([math]::Round($avgAbsGap,0)) pts vs declared theme weights",
    "$pctRate% PCT/international (File 3)",
    "$avgCites avg forward cites ($pctCited% cited, File 7)",
    "17% on live path (illustrative / TTO)",
    "$ocpShare% of co-filings with OCP (HHI $hhi)"
)

$payload = [ordered]@{
    generated = (Get-Date).ToString('yyyy-MM-dd')
    totalPatents = $totalPatents
    filingsUndated = $filingsUndated
    filingsDated = $datedPatents.Count
    filings2025 = $filingsByYear[2025]
    filings2024 = $filingsByYear[2024]
    cagr = $cagr
    pctRate = $pctRate
    grantRate = $grantRate
    grantedCount = $grantedCount
    avgGrantLag = $avgGrantLag
    renewalRate = $renewalRate
    renewalTotal = $renewalTotal
    renewalActiveCount = $renewalActiveCount
    renewalLapsedCount = $renewalLapsedCount
    renewalScope = 'MAScIR renewal portfolio (File 5 only)'
    avgAlignGap = [math]::Round($avgAbsGap, 1)
    strategyTargetVals = $strategyTargetVals
    alignShareVals = $alignShareVals
    avgCites = $avgCites
    pctCited = $pctCited
    avgIpc = $avgIpc
    triadicShare = $triadicShare
    ocpShare = $ocpShare
    hhi = $hhi
    filingsYears = @($years)
    yearMin = $yearMin
    yearMax = $yearMax
    filingsMascir = @($years | ForEach-Object { $filingsMascir[$_] })
    filingsTto = @($years | ForEach-Object { $filingsTto[$_] })
    filingsOther = @($years | ForEach-Object { $filingsOther[$_] })
    treemap = $treemap
    alignLabels = $alignLabels
    alignVals = $alignVals
    citeThemeLabels = $alignLabels
    citeThemeVals = $citeVals
    countries = @($countryTop | ForEach-Object { @{ country=$_.country; count=$_.count } })
    partnerYears = @($partnerYears)
    partnerOcp = @($partnerYears | ForEach-Object { $partnerOcp[$_] })
    partnerOther = @($partnerYears | ForEach-Object { $partnerOther[$_] })
    partnerAcad = @($partnerYears | ForEach-Object { $partnerAcad[$_] })
    partnerSolo = @($partnerYears | ForEach-Object { $partnerSolo[$_] })
    peakYear = $peakYear
    peakCount = $peakCount
    countryTotal = $countryTotal
    countryMetric = 'Unique applications by primary country/route from File 3 geo (empty geo defaults to Morocco)'
    radarScores = $radarScores
    radarActuals = $radarActuals
    radarBench = @{
        morocco = $radarBenchMorocco
        africa = $radarBenchAfrica
        top = $radarBenchTop
    }
    bench = @{
        cagr=$cagr; timeToGrant=$avgGrantLag; fwdCit=$avgCites; techBreadth=$avgIpc
        goingIntl=$pctRate; topTier=$triadicShare; keptAlive=$renewalRate
    }
    deepdives = $dives
}

$json = ($payload | ConvertTo-Json -Depth 20 -Compress)
$outPath = Join-Path $base 'dashboard_data.js'
"/* Auto-generated from Files 3-8 - run scripts/build_dashboard_data.ps1 to refresh */`nvar DASH_DATA = $json;" | Set-Content -Path $outPath -Encoding UTF8
Write-Output "OK: $outPath | patents=$totalPatents | 2025=$($filingsByYear[2025]) | cite=$avgCites | PCT=$pctRate%"
