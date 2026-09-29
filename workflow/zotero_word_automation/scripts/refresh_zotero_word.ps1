param(
  [Parameter(Mandatory=$true)][string]$DocPath,
  [Parameter(Mandatory=$true)][string]$LogPath,
  [ValidateRange(120,3600)][int]$WaitSeconds = 120,
  [ValidateRange(135,7200)][int]$MaxWaitSeconds = 300,
  [string]$PayloadPath
)

$ErrorActionPreference = "Stop"
if (Test-Path -LiteralPath $LogPath) { throw "Use a new log path: $LogPath" }
if ($MaxWaitSeconds -lt ($WaitSeconds + 15)) { throw 'MaxWaitSeconds must allow at least 15 seconds beyond WaitSeconds.' }
if (-not (Test-Path -LiteralPath $DocPath -PathType Leaf)) { throw "Word file not found: $DocPath" }
$placeholders = @()
if ($PayloadPath) {
  $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $placeholders = @($payload.groups | ForEach-Object { $_.placeholder }) + @($payload.bibliography.placeholder)
}

function Add-LogLine([string]$message) {
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
  Add-Content -LiteralPath $LogPath -Encoding UTF8 -Value "$timestamp`t$message"
}

function Get-RefreshSnapshot($doc) {
  $parts = @()
  $ready = $true
  $items = 0
  $bibs = 0
  foreach ($field in $doc.Fields) {
    $code = [string]$field.Code.Text
    if ($code -notmatch 'ZOTERO_(ITEM|BIBL)') { continue }
    $display = [string]$field.Result.Text
    $parts += $code + "`n" + $display
    if ([string]::IsNullOrWhiteSpace($display) -or $display -match '\[(?:[^\]]*placeholder[^\]]*|(?:TG\d+|CL\d+)[\w.\-]*|[\w.\-]+_G\d+)\]' -or $display -match 'missing item|no longer exists|citation error') { $ready = $false }
    foreach ($placeholder in $placeholders) {
      if ($placeholder -and $display.Contains($placeholder)) { $ready = $false }
    }
    if ($code -match 'ZOTERO_ITEM') {
      $items++
      try {
        $obj = $code.Substring($code.IndexOf('{')).Trim() | ConvertFrom-Json
        if (-not $obj.properties.formattedCitation -or -not $obj.properties.plainCitation) { $ready = $false }
      } catch { $ready = $false }
    } else { $bibs++ }
  }
  if ($items -eq 0 -or $bibs -gt 1) { $ready = $false }
  return [pscustomobject]@{Ready=$ready; Snapshot=($parts -join "`n"); Items=$items; Bibliographies=$bibs}
}

Add-LogLine "refresh_start doc=$DocPath wait=$WaitSeconds"
$backupPath = "$DocPath.before_refresh_$(Get-Date -Format yyyyMMdd_HHmmss)_$([guid]::NewGuid().ToString('N').Substring(0,8)).docx"
Copy-Item -LiteralPath $DocPath -Destination $backupPath -ErrorAction Stop
Add-LogLine "backup=$backupPath"
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = 0
$completed = $false

try {
  $doc = $word.Documents.Open($DocPath, $false, $false)
  if ($doc.ReadOnly) { throw 'Word opened read-only.' }
  $doc.Activate()
  $word.Activate()
  Add-LogLine "opened fields=$($doc.Fields.Count)"
  Start-Sleep -Seconds 2
  Add-LogLine "before_ZoteroRefresh"
  $word.Run("ZoteroRefresh")
  Add-LogLine "after_ZoteroRefresh_return fields=$($doc.Fields.Count)"
  $clock = [Diagnostics.Stopwatch]::StartNew()
  $lastSnapshot = $null
  $stableSince = 0
  $candidateReady = $false
  while ($clock.Elapsed.TotalSeconds -lt $MaxWaitSeconds) {
    Start-Sleep -Seconds 5
    $state = Get-RefreshSnapshot $doc
    $elapsed = $clock.Elapsed.TotalSeconds
    if ($state.Snapshot -cne $lastSnapshot) { $stableSince = $elapsed; $lastSnapshot = $state.Snapshot }
    Add-LogLine "waited=$([int]$elapsed)s ready=$($state.Ready) items=$($state.Items) bibliography=$($state.Bibliographies)"
    if ($elapsed -ge $WaitSeconds -and $state.Ready -and ($elapsed - $stableSince) -ge 15) {
      $candidateReady = $true
      break
    }
  }
  if (-not $candidateReady) { throw 'Refresh observations did not become ready/stable before timeout.' }
  Add-LogLine 'display_stable_after_minimum_wait; observational evidence only, not a Zotero completion signal'
  Add-LogLine "before_save fields=$($doc.Fields.Count)"
  $doc.Save()
  if (-not $doc.Saved) { throw 'Word did not confirm saved state.' }
  Add-LogLine "saved"
  $doc.Close($true)
  Add-LogLine "closed"
  $completed = $true
} catch {
  Add-LogLine "FAILED: $($_.Exception.Message); Word left open; inspect before saving or closing."
  throw
} finally {
  if ($completed) { $word.Quit() }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
