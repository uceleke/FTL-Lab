<# 

.DESCRIPTION
  - One-pass scan of files with access-denied tracking (noisy "0 B" folders get flagged Partial=$true).
  - End-of-run prompt to export CSV/HTML to a path you pick (or skip).
  - Outputs:
      1) Top-level folder sizes (+ Partial flag)
      2) Rollups to an arbitrary depth (+ Partial flag)
      3) Heaviest directories (any depth)
      4) Largest files
      5) Size by extension (file type)
      6) Access issues table (paths where traversal failed)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [ValidateScript({ Test-Path -LiteralPath $_ })]
  [string]$Path,

  [int]$Depth = 2,

  [int]$Top   = 20,

  [string[]]$ExcludePaths = @(),

  # If you pre-provide ReportPath, it will be shown as the default during the export prompt.
  [string]$ReportPath = $(Join-Path $env:USERPROFILE ("Desktop\StorageReport_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

  # Set -NoPromptToExport to skip prompting (and also skip exporting).
  [switch]$NoPromptToExport
)
function HtmlEnc {
  param([string]$s)
  try {
    return [System.Net.WebUtility]::HtmlEncode($s)
  } catch {
    try {
      Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue | Out-Null
      if ([type]::GetType('System.Web.HttpUtility')) {
        return [System.Web.HttpUtility]::HtmlEncode($s)
      } else {
        return $s
      }
    } catch { return $s }
  }
}


# ---------- Helpers ----------
# Binary units
$KB = [double]1024
$MB = $KB * 1024
$GB = $MB * 1024
$TB = $GB * 1024

function Convert-Size {
  param([long]$Bytes)
  $b = [double]$Bytes
  if     ($b -lt $KB) { return ("{0:N0} B"  -f $b) }
  elseif ($b -lt $MB) { return ("{0:N2} KB" -f ($b / $KB)) }
  elseif ($b -lt $GB) { return ("{0:N2} MB" -f ($b / $MB)) }
  elseif ($b -lt $TB) { return ("{0:N2} GB" -f ($b / $GB)) }  # <- divide by $GB
  else                { return ("{0:N2} TB" -f ($b / $TB)) }
}

function New-Row {
  param(
    [string]$Path,
    [long]$Bytes,
    [bool]$Partial=$false
  )
  [pscustomobject]@{
    Path    = $Path
    Bytes   = [int64]$Bytes
    Size    = Convert-Size $Bytes
    Partial = $Partial
  }
}

# Normalize root
$root = (Resolve-Path -LiteralPath $Path).Path
while ($root.EndsWith('\') -or $root.EndsWith('/')) { $root = $root.TrimEnd('\','/') }

Write-Host "Scanning: $root" -ForegroundColor Cyan

# Aggregates
$dirSize        = @{}
$topLevelSize   = @{}
$levelRollup    = @{}   # key: "$depth||path"
$extSize        = @{}
$largestFiles   = New-Object System.Collections.Generic.List[object]

# Exclude regexes
$excludeRegexes = @()
foreach ($ex in $ExcludePaths) {
  try   { $excludeRegexes += [regex]::new($ex, 'IgnoreCase') }
  catch { $excludeRegexes += [regex]::new([regex]::Escape($ex), 'IgnoreCase') }
}

# Capture non-terminating errors from recursion to detect access-denied nodes
$gciErrors = @()
$allFiles = Get-ChildItem -LiteralPath $root -Recurse -Force -File -ErrorAction SilentlyContinue -ErrorVariable +gciErrors

# Optional exclude
if ($excludeRegexes.Count -gt 0) {
  $allFiles = $allFiles | Where-Object {
    $full = $_.FullName
    $include = $true
    foreach ($rx in $excludeRegexes) {
      if ($rx.IsMatch($full)) { $include = $false; break }
    }
    $include
  }
}

# Build an access-issues list from the collected errors
$accessErrors = $gciErrors | Where-Object {
  ($_.Exception -is [System.UnauthorizedAccessException]) -or
  ($_.FullyQualifiedErrorId -like '*UnauthorizedAccessException*') -or
  ($_.Exception.Message -match 'denied')
}

# Try to extract meaningful paths from each error
$accessIssueItems = @()
foreach ($e in $accessErrors) {
  $p = $null
  if ($e.TargetObject) {
    $p = [string]$e.TargetObject
  } else {
    # Heuristic parse of message if TargetObject is null
    if ($e.Exception.Message -match '(: |Access to the path .)(?<pp>\\\\.*?)(?: is denied|$)') {
      $p = $matches['pp']
    } else {
      $p = $e.Exception.Message
    }
  }
  if (-not [string]::IsNullOrWhiteSpace($p)) {
    $accessIssueItems += [pscustomobject]@{
      Path    = $p
      Message = $e.Exception.Message
      Type    = $e.Exception.GetType().FullName
    }
  }
}
# Unique & normalized
$accessIssues = $accessIssueItems | Sort-Object Path -Unique

# Helper: does a rollup path intersect any denied subtree?
function Test-Partial {
  param([string]$SomePath)
  foreach ($ai in $accessIssues) {
    $ap = [string]$ai.Path
    if ([string]::IsNullOrWhiteSpace($ap)) { continue }
    if ($ap.StartsWith($SomePath, [System.StringComparison]::OrdinalIgnoreCase)) {
      return $true
    }
  }
  return $false
}

# Single pass aggregate with progress (PS5/7-safe)
$idx = 0
$total = ($allFiles | Measure-Object).Count
$sep = [IO.Path]::DirectorySeparatorChar
if ($Depth -lt 1) { $Depth = 1 }

foreach ($f in $allFiles) {
  $idx++
  if ($idx -eq 1 -or ($idx % 1000 -eq 0)) {
    $percent = 100
    if ($total -gt 0) { $percent = ($idx / $total) * 100 }
    Write-Progress -Activity "Analyzing files..." -Status "$idx / $total" -PercentComplete $percent
  }

  $size = [int64]$f.Length
  $dir  = $f.DirectoryName

  if (-not $dirSize.ContainsKey($dir)) { $dirSize[$dir] = 0L }
  $dirSize[$dir] += $size

  if ($dir.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
    $rel = $dir.Substring($root.Length).TrimStart('\','/')
    $first = if ([string]::IsNullOrEmpty($rel)) { '' } else { ($rel -split '[\\/]')[0] }
    if ([string]::IsNullOrEmpty($first)) {
      $level1 = $root
    } else {
      $level1 = (Join-Path $root $first)
    }
    if (-not $topLevelSize.ContainsKey($level1)) { $topLevelSize[$level1] = 0L }
    $topLevelSize[$level1] += $size

    $parts = @()
    if (-not [string]::IsNullOrEmpty($rel)) { $parts = $rel -split '[\\/]'}
    for ($d = 1; $d -le $Depth; $d++) {
      if (($parts.Count -eq 0) -or ($d -gt $parts.Count)) {
        $rollPath = $root
      } else {
        $rollPath = Join-Path $root ( ($parts[0..($d-1)] -join $sep) )
      }
      $key = "$d||$rollPath"
      if (-not $levelRollup.ContainsKey($key)) { $levelRollup[$key] = 0L }
      $levelRollup[$key] += $size
    }
  }

  $ext = [IO.Path]::GetExtension($f.Name)
  if ($null -eq $ext) { $ext = '' }
  $ext = $ext.ToLowerInvariant()
  if ([string]::IsNullOrWhiteSpace($ext)) { $ext = '<noext>' }
  if (-not $extSize.ContainsKey($ext)) { $extSize[$ext] = 0L }
  $extSize[$ext] += $size

  $largestFiles.Add([pscustomobject]@{
    Path  = $f.FullName
    Bytes = $size
    Size  = Convert-Size $size
  })
}

Write-Progress -Activity "Analyzing files..." -Completed
# Sanity check: total bytes vs. formatted
$__totalBytes = ($largestFiles | Measure-Object Bytes -Sum).Sum
if ($__totalBytes -is [int64] -and $__totalBytes -gt 0) {
  Write-Host ("`n== Totals (sanity check) ==") -ForegroundColor Yellow
  Write-Host ("Files scanned: {0:N0}" -f ($largestFiles.Count))
  Write-Host ("Total bytes:   {0:N0}" -f $__totalBytes)
  Write-Host ("As size:       {0}" -f (Convert-Size $__totalBytes))
}

# ---------- Build result tables ----------
# Top-level (with Partial flag)
$topLevelTable = $topLevelSize.GetEnumerator() | ForEach-Object {
  $p = $_.Key
  $b = $_.Value
  New-Row -Path $p -Bytes $b -Partial (Test-Partial $p)
} | Sort-Object Bytes -Descending

# Rollups to Depth (with Partial flag)
$rollupTable = foreach ($k in $levelRollup.Keys) {
  $d,$p = $k -split '\|\|',2
  [pscustomobject]@{
    Depth   = [int]$d
    Path    = $p
    Bytes   = [int64]$levelRollup[$k]
    Size    = Convert-Size $levelRollup[$k]
    Partial = (Test-Partial $p)
  }
} 

$rollupAtDepth = $rollupTable | Where-Object { $_.Depth -eq $Depth } | Sort-Object Bytes -Descending

# Heaviest directories (any depth)
$heaviestDirs = $dirSize.GetEnumerator() | ForEach-Object {
  $p = $_.Key; $b = $_.Value
  New-Row -Path $p -Bytes $b -Partial (Test-Partial $p)
} | Sort-Object Bytes -Descending

# Largest files
$largestFiles = $largestFiles | Sort-Object Bytes -Descending

# By extension
$byExt = $extSize.GetEnumerator() |
  ForEach-Object {
    [pscustomobject]@{
      Extension = $_.Key
      Bytes     = [int64]$_.Value
      Size      = Convert-Size $_.Value
    }
  } | Sort-Object Bytes -Descending

  # ---------- Recompute Size strings for consistency ----------
$topLevelTable  | ForEach-Object { $_.Size = Convert-Size $_.Bytes }
$rollupAtDepth  | ForEach-Object { $_.Size = Convert-Size $_.Bytes }
$heaviestDirs   | ForEach-Object { $_.Size = Convert-Size $_.Bytes }
$largestFiles   | ForEach-Object { $_.Size = Convert-Size $_.Bytes }
$byExt          | ForEach-Object { if ($_.PSObject.Properties.Match('Bytes')) { $_.Size = Convert-Size $_.Bytes } }

# ---------- Console summary ----------
Write-Host ""
Write-Host "== Top-level folders (level 1) ==" -ForegroundColor Yellow
$topLevelTable | Select-Object -First $Top | Format-Table -AutoSize

Write-Host ""
Write-Host ("== Rollup at depth {0} ==" -f $Depth) -ForegroundColor Yellow
$rollupAtDepth | Select-Object -First $Top | Format-Table -Property Path,Size,Partial -AutoSize

Write-Host ""
Write-Host "== Heaviest directories (any depth) ==" -ForegroundColor Yellow
$heaviestDirs | Select-Object -First $Top | Format-Table -Property Path,Size,Partial -AutoSize

Write-Host ""
Write-Host "== Largest files ==" -ForegroundColor Yellow
$largestFiles | Select-Object -First $Top | Format-Table -Property Path,Size -AutoSize

Write-Host ""
Write-Host "== By file extension ==" -ForegroundColor Yellow
$byExt | Select-Object -First $Top | Format-Table -Property Extension,Size -AutoSize

if ($accessIssues.Count -gt 0) {
  Write-Host ""
  Write-Host ("== Access issues detected == ({0} unique paths)" -f $accessIssues.Count) -ForegroundColor DarkYellow
  $accessIssues | Select-Object -First ([Math]::Min($Top,50)) | Format-Table -Property Path,Type -AutoSize
  Write-Host "Rows with Partial = True include one or more inaccessible subpaths."
}

# ---------- Export prompt ----------
if (-not $NoPromptToExport) {
  Write-Host ""
  $ans = Read-Host ("Export CSV + HTML report? (Y/n) [default: Y]")
  if ([string]::IsNullOrWhiteSpace($ans) -or $ans -match '^(y|yes)$') {
    $chosen = Read-Host ("Enter export folder or press Enter for default:`n  " + $ReportPath)
    if (-not [string]::IsNullOrWhiteSpace($chosen)) { $ReportPath = $chosen }
    if (-not (Test-Path -LiteralPath $ReportPath)) {
      $null = New-Item -ItemType Directory -Path $ReportPath -Force
    }

    $csvTop1        = Join-Path $ReportPath "TopLevel.csv"
    $csvRollupDepth = Join-Path $ReportPath ("Rollup_Level{0}.csv" -f $Depth)
    $csvDirs        = Join-Path $ReportPath "HeaviestDirectories.csv"
    $csvFiles       = Join-Path $ReportPath "LargestFiles.csv"
    $csvExt         = Join-Path $ReportPath "ByExtension.csv"
    $csvAccess      = Join-Path $ReportPath "AccessIssues.csv"
    $htmlReport     = Join-Path $ReportPath "StorageReport.html"

    $topLevelTable        | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvTop1
    $rollupAtDepth        | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvRollupDepth
    $heaviestDirs         | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvDirs
    $largestFiles         | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvFiles
    $byExt                | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvExt
    $accessIssues         | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvAccess

    # Quick HTML summary
    $sb = New-Object -TypeName System.Text.StringBuilder
    $null = $sb.AppendLine("<html><head><meta charset='utf-8'><title>Storage Report</title>")
    $null = $sb.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;margin:20px} table{border-collapse:collapse;margin:12px 0;width:100%} th,td{border:1px solid #ddd;padding:6px;text-align:left} th{background:#f3f3f3}</style>")
    $null = $sb.AppendLine("</head><body>")
    $null = $sb.AppendLine("<h1>Storage Report</h1><p><b>Root:</b> $root<br><b>Generated:</b> $(Get-Date)</p>")
    $null = $sb.AppendLine("<p><i>Note:</i> Rows with <b>Partial = True</b> include one or more inaccessible subpaths; totals reflect only accessible files.</p>")

    function Add-TableHtml {
      param([string]$Title, [object[]]$Rows, [string[]]$Columns, [int]$Max = 30)
      $null = $sb.AppendLine("<h2>$Title</h2><table><tr>")
      foreach($c in $Columns){ $null = $sb.AppendLine("<th>$c</th>") }
      $null = $sb.AppendLine("</tr>")
      $i = 0
      foreach($r in $Rows){
        if ($i -ge $Max) { break }
        $null = $sb.AppendLine("<tr>")
        foreach($c in $Columns){
          $v = HtmlEnc ($r.$c)
          $null = $sb.AppendLine("<td>$v</td>")
        }
        $null = $sb.AppendLine("</tr>")
        $i++
      }
      $null = $sb.AppendLine("</table>")
    }

    Add-TableHtml -Title "Top-level folders (level 1) — top $Top" -Rows $topLevelTable -Columns @('Path','Size','Bytes','Partial') -Max $Top
    Add-TableHtml -Title "Rollup at depth $Depth — top $Top" -Rows $rollupAtDepth -Columns @('Path','Size','Bytes','Partial') -Max $Top
    Add-TableHtml -Title "Heaviest directories — top $Top" -Rows $heaviestDirs -Columns @('Path','Size','Bytes','Partial') -Max $Top
    Add-TableHtml -Title "Largest files — top $Top" -Rows $largestFiles -Columns @('Path','Size','Bytes') -Max $Top
    Add-TableHtml -Title "By extension — top $Top" -Rows $byExt -Columns @('Extension','Size','Bytes') -Max $Top

    if ($accessIssues.Count -gt 0) {
      Add-TableHtml -Title "Access issues (first $Top shown)" -Rows $accessIssues -Columns @('Path','Type') -Max $Top
    }

    $null = $sb.AppendLine("</body></html>")
    [IO.File]::WriteAllText($htmlReport, $sb.ToString(), [Text.UTF8Encoding]::new($true))

    Write-Host ""
    Write-Host "Saved CSVs and HTML report to: $ReportPath" -ForegroundColor Green
    Write-Host "Open report: $htmlReport"
  } else {
    Write-Host "Export skipped."
  }
}
# ---------- Examples ----------
<#
.\Get-ShareUsage.ps1 -Path \\filesrv\Shares\Finance -Depth 2 -Top 30
.\Get-ShareUsage.ps1 -Path D:\Shares -ExcludePaths '$RECYCLE.BIN','System Volume Information','\\node_modules\\' -Depth 3 -Top 50

#>
