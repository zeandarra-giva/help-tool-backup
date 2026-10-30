<#
  Exports every table (and saved select query) in an Access database to ONE
  self-contained HTML file: styles.css is inlined and every table lives in the
  same file, so the snapshot can be shared and opened offline on its own.
  Read-only: nothing in the database changes, and no forms or VBA run, so
  corrupted modules don't get in the way.

  Usage:
    powershell -ExecutionPolicy Bypass -File export-to-html.ps1 -OutDir ".\snapshots" [-DbPath "X:\db.accdb"] [-DbName "HELP+.accdb"] [-NoOpen]

  Without -DbPath, the database is located by file name (see Find-Database)
  and the result is remembered in last-db-path.txt for the next run.

  Exit codes: 0 = ok, 1 = error, 2 = Access driver not registered for this
  PowerShell bitness (export.bat then retries with 32-bit PowerShell).
#>
param(
    [Parameter(Mandatory)] [string] $OutDir,
    [string] $DbPath,
    [string] $DbName = 'HELP+.accdb',
    [switch] $SkipQueries,
    [switch] $NoOpen
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$cacheFile = Join-Path $PSScriptRoot 'last-db-path.txt'

function Find-Newest([string] $root, [int] $depth = -1) {
    if (-not (Test-Path -LiteralPath $root)) { return $null }
    $params = @{ LiteralPath = $root; Filter = $DbName; Recurse = $true; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($depth -ge 0) { $params.Depth = $depth }
    Get-ChildItem @params | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}

# Search order, cheapest first:
#   1. last remembered location
#   2. the script folder and each parent folder above it
#   3. Desktop, Documents, Downloads, OneDrive (a few levels deep)
#   4. every local and mapped network drive (slow, last resort)
function Find-Database {
    if (Test-Path -LiteralPath $cacheFile) {
        $cached = (Get-Content -LiteralPath $cacheFile -TotalCount 1).Trim()
        if ($cached -and (Test-Path -LiteralPath $cached)) { return $cached }
    }

    $dir = $PSScriptRoot
    while ($dir) {
        $candidate = Join-Path $dir $DbName
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $dir = Split-Path -Parent $dir
    }

    $userFolders = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('MyDocuments'),
        (Join-Path $env:USERPROFILE 'Downloads'),
        $env:OneDrive
    ) | Where-Object { $_ } | Select-Object -Unique
    foreach ($folder in $userFolders) {
        Write-Host "Searching $folder ..."
        $found = Find-Newest $folder 4
        if ($found) { return $found }
    }

    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -and (Test-Path -LiteralPath $_.Root) }
    foreach ($drive in $drives) {
        Write-Host "Searching $($drive.Root) (this can take a few minutes) ..."
        $found = Find-Newest $drive.Root
        if ($found) { return $found }
    }
    return $null
}

function Enc([object] $v) { [System.Net.WebUtility]::HtmlEncode([string]$v) }

function Format-Cell([object] $v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return '' }
    if ($v -is [byte[]]) { return '<span class="muted">[binary]</span>' }
    if ($v -is [bool]) { if ($v) { return 'Yes' } else { return 'No' } }
    if ($v -is [datetime]) {
        # Access stores time-only values on 1899-12-30
        if ($v.Date -eq [datetime]'1899-12-30') { return $v.ToString('HH:mm') }
        if ($v.TimeOfDay.Ticks -eq 0) { return $v.ToString('yyyy-MM-dd') }
        return $v.ToString('yyyy-MM-dd HH:mm')
    }
    return (Enc $v) -replace "`r?`n", '<br>'
}

# ---------------------------------------------------------------- locate ---

if ($DbPath) {
    if (-not (Test-Path -LiteralPath $DbPath)) {
        Write-Host "Database not found: $DbPath"
        exit 1
    }
} else {
    Write-Host "Looking for $DbName ..."
    $DbPath = Find-Database
    if (-not $DbPath) {
        Write-Host "Could not find $DbName on this PC or its mapped drives."
        Write-Host "Drag the database file onto export.bat to point to it directly."
        exit 1
    }
}
$DbPath = (Resolve-Path -LiteralPath $DbPath).ProviderPath
Write-Host "Using database: $DbPath"
[System.IO.File]::WriteAllText($cacheFile, $DbPath, $utf8)

# ------------------------------------------------------------------ read ---
# Fills $entries: one object per table/query with its DataTable (or Error).

$csb = New-Object System.Data.OleDb.OleDbConnectionStringBuilder
$csb['Provider'] = 'Microsoft.ACE.OLEDB.12.0'
$csb['Data Source'] = $DbPath
$csb['Mode'] = 'Read'
$conn = New-Object System.Data.OleDb.OleDbConnection($csb.ConnectionString)

try {
    $conn.Open()
} catch {
    $msg = $_.Exception.Message
    if ($msg -match 'not registered') {
        Write-Host "Access driver not available in this PowerShell ($([IntPtr]::Size * 8)-bit)."
        exit 2
    }
    Write-Host "Could not open database: $msg"
    exit 1
}

$types = @('TABLE', 'LINK', 'PASS-THROUGH')
if (-not $SkipQueries) { $types += 'VIEW' }

$objects = $conn.GetSchema('Tables').Rows |
    Where-Object { $types -contains $_.TABLE_TYPE -and $_.TABLE_NAME -notlike 'MSys*' -and $_.TABLE_NAME -notlike '~*' } |
    Sort-Object @{ Expression = { $_.TABLE_TYPE -eq 'VIEW' } }, TABLE_NAME

$entries = @()
foreach ($obj in $objects) {
    $name = $obj.TABLE_NAME
    $kind = if ($obj.TABLE_TYPE -eq 'VIEW') { 'Query' } else { 'Table' }
    Write-Host "Reading $kind $name ..."
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT * FROM [$name]"
        $dt = New-Object System.Data.DataTable
        [void](New-Object System.Data.OleDb.OleDbDataAdapter($cmd)).Fill($dt)
        $entries += [pscustomobject]@{ Name = $name; Kind = $kind; Data = $dt; Error = $null }
    } catch {
        # Queries that reference forms or VBA functions can't run outside Access
        $entries += [pscustomobject]@{ Name = $name; Kind = $kind; Data = $null; Error = $_.Exception.Message }
    }
}
$conn.Close()

# ---------------------------------------------------------------- render ---
# One page, no JavaScript: each table is a <section> shown via CSS :target
# when its link is clicked; the index shows when no table is selected.

$dbTitle = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
$cssPath = Join-Path $PSScriptRoot 'styles.css'
$css = if (Test-Path -LiteralPath $cssPath) { [System.IO.File]::ReadAllText($cssPath) } else { '' }

$usedIds = @{}
foreach ($e in $entries) {
    $base = ($e.Name -replace '[^\w\-]+', '-').Trim('-').ToLower()
    if (-not $base) { $base = 'object' }
    $id = "t-$base"
    $n = 2
    while ($usedIds.ContainsKey($id)) { $id = "t-$base-$n"; $n++ }
    $usedIds[$id] = $true
    $e | Add-Member -NotePropertyName Id -NotePropertyValue $id
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Enc $dbTitle) - snapshot $stamp</title>
<style>
$css
</style>
</head>
<body>
<header class="bar"><a href="#" class="brand">$(Enc $dbTitle)</a><span class="muted">Read-only snapshot &middot; $stamp</span></header>

"@)

[void]$sb.Append("<main class=`"home`">`n<h1>$(Enc $dbTitle)</h1>`n")
foreach ($group in @(@('Table', 'Tables'), @('Query', 'Queries'))) {
    $items = @($entries | Where-Object Kind -eq $group[0])
    if ($items.Count -eq 0) { continue }
    [void]$sb.Append("<h2>$($group[1])</h2>`n<ul class=`"cards`">`n")
    foreach ($e in $items) {
        if ($e.Data) {
            [void]$sb.Append("<li><a class=`"card`" href=`"#$($e.Id)`"><span class=`"card-name`">$(Enc $e.Name)</span><span class=`"muted`">$($e.Data.Rows.Count) rows</span></a></li>`n")
        } else {
            [void]$sb.Append("<li><div class=`"card card-off`" title=`"$(Enc $e.Error)`"><span class=`"card-name`">$(Enc $e.Name)</span><span class=`"muted`">Can't run outside Access</span></div></li>`n")
        }
    }
    [void]$sb.Append("</ul>`n")
}
[void]$sb.Append("</main>`n`n")

foreach ($e in ($entries | Where-Object Data)) {
    $dt = $e.Data
    [void]$sb.Append("<section class=`"sheet`" id=`"$($e.Id)`">`n<a href=`"#`" class=`"back`">&larr; All tables</a>`n")
    [void]$sb.Append("<h1>$(Enc $e.Name) <span class=`"tag`">$($e.Kind)</span></h1>`n<p class=`"muted`">$($dt.Rows.Count) rows</p>`n")

    if ($dt.Rows.Count -eq 0) {
        [void]$sb.Append("<p class=`"empty`">No records.</p>`n")
    } else {
        [void]$sb.Append("<div class=`"table-wrap`"><table>`n<thead><tr>")
        foreach ($col in $dt.Columns) { [void]$sb.Append("<th>$(Enc $col.ColumnName)</th>") }
        [void]$sb.Append("</tr></thead>`n<tbody>`n")
        foreach ($row in $dt.Rows) {
            [void]$sb.Append('<tr>')
            foreach ($col in $dt.Columns) {
                $v = $row[$col]
                $cls = if ($v -is [ValueType] -and $v -isnot [bool] -and $v -isnot [datetime]) { ' class="num"' } else { '' }
                [void]$sb.Append("<td$cls>$(Format-Cell $v)</td>")
            }
            [void]$sb.Append("</tr>`n")
        }
        [void]$sb.Append("</tbody>`n</table></div>`n")
    }
    [void]$sb.Append("</section>`n`n")
}
[void]$sb.Append("</body>`n</html>`n")

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$safeTitle = $dbTitle -replace '[\\/:*?"<>|]', '_'
$outFile = Join-Path $OutDir "$safeTitle snapshot $(Get-Date -Format 'yyyy-MM-dd HHmm').html"
[System.IO.File]::WriteAllText($outFile, $sb.ToString(), $utf8)

$sizeKb = [math]::Ceiling((Get-Item -LiteralPath $outFile).Length / 1KB)
Write-Host ""
Write-Host "Done: $(@($entries | Where-Object Data).Count) tables/queries in one file ($sizeKb KB)"
Write-Host "  $outFile"
Write-Host "Share this single file - it opens offline in any browser."
if (-not $NoOpen) { Invoke-Item -LiteralPath $outFile }
exit 0
