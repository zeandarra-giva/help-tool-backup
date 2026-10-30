<#
  Exports every table (and saved select query) in an Access database to ONE
  self-contained HTML file with search, tabs and a record view. template.html,
  styles.css, app.js and the data are all inlined, so the snapshot can be
  shared and opened offline on its own. Tabs are configured in categories.psd1.
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
try { Add-Type -AssemblyName System.Web } catch { }

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

function Format-Text([object] $v) {
    if ($null -eq $v -or $v -is [System.DBNull]) { return '' }
    if ($v -is [byte[]]) { return '[binary]' }
    if ($v -is [bool]) { if ($v) { return 'Yes' } else { return 'No' } }
    if ($v -is [datetime]) {
        # Access stores time-only values on 1899-12-30
        if ($v.Date -eq [datetime]'1899-12-30') { return $v.ToString('HH:mm') }
        if ($v.TimeOfDay.Ticks -eq 0) { return $v.ToString('yyyy-MM-dd') }
        return $v.ToString('yyyy-MM-dd HH:mm')
    }
    return [string]$v
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
# One self-contained file: template.html with styles.css, app.js and the data
# (as JSON) inlined. Tables are grouped into tabs using categories.psd1.

function Test-AnyLike([string] $name, $patterns) {
    foreach ($p in @($patterns)) { if ($name -like $p) { return $true } }
    return $false
}

$dbTitle = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'

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

# Assign each table to the first tab whose Match pattern fits its name.
# Leftover tables go to "Other Tables"; queries go to "Queries".
$configPath = Join-Path $PSScriptRoot 'categories.psd1'
$config = if (Test-Path -LiteralPath $configPath) { Import-PowerShellDataFile -LiteralPath $configPath } else { @{ Categories = @() } }

$ready = @($entries | Where-Object Data)
$tables = @($ready | Where-Object Kind -eq 'Table')
$assigned = @{}
$categories = @()
foreach ($c in @($config.Categories)) {
    $tabs = @()
    foreach ($t in @($c.Tabs)) {
        $hits = @($tables | Where-Object { -not $assigned.ContainsKey($_.Id) -and (Test-AnyLike $_.Name $t.Match) })
        foreach ($h in $hits) {
            $assigned[$h.Id] = $true
            $label = if ($hits.Count -eq 1 -and $t.Label) { $t.Label } else { $h.Name }
            $tabs += [pscustomobject]@{ Id = $h.Id; Label = $label }
        }
    }
    if ($tabs.Count) {
        $categories += [pscustomobject]@{ Name = $c.Name; Sidebar = ($c.Placement -eq 'Sidebar'); Tabs = $tabs }
    }
}
$leftover = @($tables | Where-Object { -not $assigned.ContainsKey($_.Id) })
if ($leftover.Count) {
    $categories += [pscustomobject]@{ Name = 'Other Tables'; Sidebar = $false; Tabs = @($leftover | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Label = $_.Name } }) }
}
$queries = @($ready | Where-Object Kind -eq 'Query')
if ($queries.Count) {
    $categories += [pscustomobject]@{ Name = 'Queries'; Sidebar = $false; Tabs = @($queries | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Label = $_.Name } }) }
}

Write-Host ""
foreach ($c in $categories) {
    $where = if ($c.Sidebar) { 'sidebar' } else { 'top tab' }
    Write-Host "[$($c.Name)] ($where): $(($c.Tabs | ForEach-Object Label) -join ', ')"
}

# JSON is built by hand: ConvertTo-Json is slow on large tables in PS 5.1.
# JavaScriptStringEncode also escapes < and >, so data can't close the <script>.
function J([object] $s) { '"' + [System.Web.HttpUtility]::JavaScriptStringEncode([string]$s) + '"' }

$numTypes = @([int16], [int32], [int64], [byte], [decimal], [double], [single])
$json = New-Object System.Text.StringBuilder
[void]$json.Append('{"title":' + (J $dbTitle) + ',"stamp":' + (J $stamp) + ',"categories":[')
[void]$json.Append((($categories | ForEach-Object {
    $tabsJson = ($_.Tabs | ForEach-Object { '{"id":' + (J $_.Id) + ',"label":' + (J $_.Label) + '}' }) -join ','
    '{"name":' + (J $_.Name) + ',"sidebar":' + $_.Sidebar.ToString().ToLower() + ',"tabs":[' + $tabsJson + ']}'
}) -join ','))
[void]$json.Append('],"unavailable":[' + ((@($entries | Where-Object { -not $_.Data }) | ForEach-Object { J $_.Name }) -join ',') + '],"tables":{')

$first = $true
foreach ($e in $ready) {
    $dt = $e.Data
    if (-not $first) { [void]$json.Append(',') }
    $first = $false
    $cols = ($dt.Columns | ForEach-Object { J $_.ColumnName }) -join ','
    $num = ($dt.Columns | ForEach-Object { ($numTypes -contains $_.DataType).ToString().ToLower() }) -join ','
    [void]$json.Append((J $e.Id) + ':{"name":' + (J $e.Name) + ',"kind":' + (J $e.Kind) + ',"cols":[' + $cols + '],"num":[' + $num + '],"rows":[')
    $firstRow = $true
    foreach ($row in $dt.Rows) {
        if (-not $firstRow) { [void]$json.Append(',') }
        $firstRow = $false
        [void]$json.Append('[' + (($row.ItemArray | ForEach-Object { J (Format-Text $_) }) -join ',') + ']')
    }
    [void]$json.Append(']}')
}
[void]$json.Append('}}')

function Read-Asset([string] $file) {
    $path = Join-Path $PSScriptRoot $file
    if (-not (Test-Path -LiteralPath $path)) { Write-Host "Missing $file next to the script."; exit 1 }
    [System.IO.File]::ReadAllText($path)
}

# Plain string Replace (not -replace) so $ and \ in the content stay literal
$html = (Read-Asset 'template.html').
    Replace('{{TITLE}}', (Enc "$dbTitle - snapshot $stamp")).
    Replace('{{CSS}}', (Read-Asset 'styles.css')).
    Replace('{{JS}}', (Read-Asset 'app.js')).
    Replace('{{DATA}}', $json.ToString())

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$safeTitle = $dbTitle -replace '[\\/:*?"<>|]', '_'
$outFile = Join-Path $OutDir "$safeTitle snapshot $(Get-Date -Format 'yyyy-MM-dd HHmm').html"
[System.IO.File]::WriteAllText($outFile, $html, $utf8)

$sizeKb = [math]::Ceiling((Get-Item -LiteralPath $outFile).Length / 1KB)
Write-Host ""
Write-Host "Done: $($ready.Count) tables/queries in one file ($sizeKb KB)"
Write-Host "  $outFile"
Write-Host "Share this single file - it opens offline in any browser."
if (-not $NoOpen) { Invoke-Item -LiteralPath $outFile }
exit 0
