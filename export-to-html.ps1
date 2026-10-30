<#
  Exports every table (and saved select query) in an Access database to static
  HTML pages styled by styles.css. Read-only: nothing in the database changes,
  and no forms or VBA run, so corrupted modules don't get in the way.

  Usage:
    powershell -ExecutionPolicy Bypass -File export-to-html.ps1 -OutDir ".\html" [-DbPath "X:\db.accdb"] [-DbName "HELP+.accdb"]

  Without -DbPath, the database is located by file name (see Find-Database)
  and the result is remembered in last-db-path.txt for the next run.

  Exit codes: 0 = ok, 1 = error, 2 = Access driver not registered for this
  PowerShell bitness (export.bat then retries with 32-bit PowerShell).
#>
param(
    [Parameter(Mandatory)] [string] $OutDir,
    [string] $DbPath,
    [string] $DbName = 'HELP+.accdb',
    [switch] $SkipQueries
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
    $s = [System.Net.WebUtility]::HtmlEncode([string]$v)
    return $s -replace "`r?`n", '<br>'
}

function Write-Page([string] $file, [string] $title, [string] $body) {
    $t = [System.Net.WebUtility]::HtmlEncode($title)
    $html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$t</title>
<link rel="stylesheet" href="styles.css">
</head>
<body>
$body
</body>
</html>
"@
    [System.IO.File]::WriteAllText((Join-Path $OutDir $file), $html, $utf8)
}

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

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$css = Join-Path $PSScriptRoot 'styles.css'
if (Test-Path -LiteralPath $css) { Copy-Item -LiteralPath $css -Destination $OutDir -Force }

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

$dbName = [System.IO.Path]::GetFileNameWithoutExtension($DbPath)
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
$types = @('TABLE', 'LINK', 'PASS-THROUGH')
if (-not $SkipQueries) { $types += 'VIEW' }

$objects = $conn.GetSchema('Tables').Rows |
    Where-Object { $types -contains $_.TABLE_TYPE -and $_.TABLE_NAME -notlike 'MSys*' -and $_.TABLE_NAME -notlike '~*' } |
    Sort-Object @{ Expression = { $_.TABLE_TYPE -eq 'VIEW' } }, TABLE_NAME

$entries = @()
$usedFiles = @{}

foreach ($obj in $objects) {
    $name = $obj.TABLE_NAME
    $kind = if ($obj.TABLE_TYPE -eq 'VIEW') { 'Query' } else { 'Table' }

    $base = ($name -replace '[^\w\-]+', '_').Trim('_')
    if (-not $base) { $base = 'object' }
    $file = "$base.html"
    $n = 2
    while ($usedFiles.ContainsKey($file.ToLower())) { $file = "$base-$n.html"; $n++ }
    $usedFiles[$file.ToLower()] = $true

    Write-Host "Exporting $kind $name ..."
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT * FROM [$name]"
        $dt = New-Object System.Data.DataTable
        [void](New-Object System.Data.OleDb.OleDbDataAdapter($cmd)).Fill($dt)
    } catch {
        # Queries that reference forms or VBA functions can't run outside Access
        $entries += [pscustomobject]@{ Name = $name; Kind = $kind; File = $null; Rows = 0; Error = $_.Exception.Message }
        continue
    }

    $sb = New-Object System.Text.StringBuilder
    $enc = [System.Net.WebUtility]::HtmlEncode($name)
    [void]$sb.Append("<header class=`"bar`"><a href=`"index.html`">&larr; $([System.Net.WebUtility]::HtmlEncode($dbName))</a>")
    [void]$sb.Append("<span class=`"muted`">Snapshot $stamp</span></header>`n<main>`n")
    [void]$sb.Append("<h1>$enc <span class=`"tag`">$kind</span></h1>`n<p class=`"muted`">$($dt.Rows.Count) rows</p>`n")

    if ($dt.Rows.Count -eq 0) {
        [void]$sb.Append("<p class=`"empty`">No records.</p>`n")
    } else {
        [void]$sb.Append("<div class=`"table-wrap`"><table>`n<thead><tr>")
        foreach ($col in $dt.Columns) {
            [void]$sb.Append("<th>$([System.Net.WebUtility]::HtmlEncode($col.ColumnName))</th>")
        }
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
    [void]$sb.Append('</main>')

    Write-Page $file "$name - $dbName" $sb.ToString()
    $entries += [pscustomobject]@{ Name = $name; Kind = $kind; File = $file; Rows = $dt.Rows.Count; Error = $null }
}

$conn.Close()

$idx = New-Object System.Text.StringBuilder
[void]$idx.Append("<header class=`"bar`"><strong>$([System.Net.WebUtility]::HtmlEncode($dbName))</strong>")
[void]$idx.Append("<span class=`"muted`">Read-only snapshot &middot; $stamp</span></header>`n<main>`n")
[void]$idx.Append("<h1>$([System.Net.WebUtility]::HtmlEncode($dbName))</h1>`n")

foreach ($group in @('Table', 'Query')) {
    $items = @($entries | Where-Object Kind -eq $group)
    if ($items.Count -eq 0) { continue }
    [void]$idx.Append("<h2>$($group)s</h2>`n<ul class=`"cards`">`n")
    foreach ($e in $items) {
        $en = [System.Net.WebUtility]::HtmlEncode($e.Name)
        if ($e.File) {
            [void]$idx.Append("<li><a class=`"card`" href=`"$($e.File)`"><span class=`"card-name`">$en</span><span class=`"muted`">$($e.Rows) rows</span></a></li>`n")
        } else {
            $err = [System.Net.WebUtility]::HtmlEncode($e.Error)
            [void]$idx.Append("<li><div class=`"card card-off`" title=`"$err`"><span class=`"card-name`">$en</span><span class=`"muted`">Can't run outside Access</span></div></li>`n")
        }
    }
    [void]$idx.Append("</ul>`n")
}
[void]$idx.Append('</main>')

Write-Page 'index.html' $dbName $idx.ToString()
Write-Host "Done: $(($entries | Where-Object File).Count) pages written to $OutDir"
exit 0
