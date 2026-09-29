param(
  [Parameter(Mandatory=$true)][string]$PayloadPath,
  [Parameter(Mandatory=$true)][string]$LogPath,
  [switch]$ReplaceExistingZoteroFields,
  [switch]$RemoveExistingReferencesHeading,
  [string]$ZoteroWordDllPath
)

$ErrorActionPreference = "Stop"
$wdCollapseEnd = 0
$wdCollapseStart = 1

function Add-LogLine([string]$message) {
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
  Add-Content -LiteralPath $LogPath -Encoding UTF8 -Value "$timestamp`t$message"
}

function Resolve-ZoteroWordDllPath([string]$explicitPath) {
  $dllRelativePath = "integration\word-for-windows\libzoteroWinWordIntegration.dll"

  if (-not [string]::IsNullOrWhiteSpace($explicitPath)) {
    $expandedPath = [Environment]::ExpandEnvironmentVariables($explicitPath)
    if (Test-Path -LiteralPath $expandedPath -PathType Leaf) {
      return (Get-Item -LiteralPath $expandedPath).FullName
    }
    throw "Specified Zotero Word DLL does not exist: $expandedPath"
  }

  $installRoots = @()
  if ($env:ProgramFiles) {
    $installRoots += (Join-Path $env:ProgramFiles "Zotero")
  }
  $programFilesX86 = ${env:ProgramFiles(x86)}
  if ($programFilesX86) {
    $installRoots += (Join-Path $programFilesX86 "Zotero")
  }
  if ($env:LOCALAPPDATA) {
    $installRoots += (Join-Path $env:LOCALAPPDATA "Programs\Zotero")
    $installRoots += (Join-Path $env:LOCALAPPDATA "Zotero")
  }

  Get-Process -Name zotero -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      if ($_.Path) {
        $installRoots += (Split-Path -Parent $_.Path)
      }
    } catch {}
  }

  $zoteroCommand = Get-Command zotero.exe -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($null -ne $zoteroCommand -and $zoteroCommand.Source) {
    $installRoots += (Split-Path -Parent $zoteroCommand.Source)
  }

  $uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zotero",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Zotero",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zotero"
  )
  foreach ($key in $uninstallKeys) {
    $entry = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    if ($null -eq $entry) {
      continue
    }
    if ($entry.InstallLocation) {
      $installRoots += [string]$entry.InstallLocation
    }
    if ($entry.DisplayIcon) {
      $displayIcon = ([string]$entry.DisplayIcon -replace ',\d+$', '').Trim('"')
      if ($displayIcon) {
        $installRoots += (Split-Path -Parent $displayIcon)
      }
    }
  }

  $checkedPaths = @()
  foreach ($root in ($installRoots | Where-Object { $_ } | Select-Object -Unique)) {
    $candidate = Join-Path ([Environment]::ExpandEnvironmentVariables([string]$root)) $dllRelativePath
    $checkedPaths += $candidate
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
      return (Get-Item -LiteralPath $candidate).FullName
    }
  }

  $checkedSummary = if ($checkedPaths.Count -gt 0) { $checkedPaths -join "; " } else { "no candidate paths were available" }
  throw "Zotero Word integration DLL was not found. Checked: $checkedSummary. Pass -ZoteroWordDllPath with the verified full DLL path; do not guess."
}

function Check-Status([UInt16]$status, [string]$label) {
  if ($status -ne 0) {
    $errPtr = [ZoteroWordDll]::getError()
    $err = if ($errPtr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::PtrToStringUni($errPtr) } else { "unknown error" }
    [ZoteroWordDll]::clearError()
    throw "$label status=$status error=$err"
  }
  Add-LogLine "$label status=0"
}

function Find-TargetRange($doc, [string]$searchText, [string]$position, [string]$occurrence, [string]$selectionReason) {
  if ([string]::IsNullOrWhiteSpace($searchText)) { throw 'Empty anchor_text' }
  if ($position -notin @('before', 'after')) { throw "Invalid position: $position" }
  if ([string]::IsNullOrWhiteSpace($occurrence)) { $occurrence = 'unique' }
  if ($occurrence -notin @('unique', 'first', 'last', 'after_last_in_document') -and $occurrence -notmatch '^[1-9]\d*$') {
    throw "Invalid anchor_occurrence: $occurrence"
  }
  $targetMatches = @()
  $scan = $doc.Content.Duplicate
  $scan.Find.ClearFormatting()
  $scan.Find.Text = $searchText
  $scan.Find.MatchCase = $false
  $scan.Find.MatchWholeWord = $false
  $scan.Find.MatchWildcards = $false
  $scan.Find.Forward = $true
  $scan.Find.Wrap = 0

  while ($scan.Find.Execute()) {
    $targetMatches += $scan.Duplicate
    $scan.Start = $scan.End
    $scan.End = $doc.Content.End
  }

  if ($targetMatches.Count -eq 0) {
    throw "Target text not found: $searchText"
  }
  if ($targetMatches.Count -gt 1 -and ($occurrence -eq 'unique' -or [string]::IsNullOrWhiteSpace($selectionReason))) {
    throw "Ambiguous anchor ($($targetMatches.Count) matches): $searchText. Supply explicit occurrence and anchor_selection_reason from the verified plan."
  }

  $target = $null
  if ($occurrence -eq "last" -or $occurrence -eq "after_last_in_document") {
    $target = $targetMatches[$targetMatches.Count - 1].Duplicate
  } elseif ($occurrence -match '^\d+$') {
    $idx = [int]$occurrence
    if ($idx -lt 1 -or $idx -gt $targetMatches.Count) {
      throw "Occurrence $occurrence out of range for target text count $($targetMatches.Count): $searchText"
    }
    $target = $targetMatches[$idx - 1].Duplicate
  } else {
    if ($targetMatches.Count -gt 1) {
      Add-LogLine "target_explicit_selection count=$($targetMatches.Count) occurrence=$occurrence reason=$selectionReason"
    }
    $target = $targetMatches[0].Duplicate
  }

  if ($position -eq "before") {
    $target.Collapse($wdCollapseStart)
  } else {
    $target.Collapse($wdCollapseEnd)
  }
  return $target
}

if (Test-Path -LiteralPath $LogPath) { throw "Use a new log path: $LogPath" }

$payload = (Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8) | ConvertFrom-Json
$docPath = [IO.Path]::GetFullPath([string]$payload.doc_path)
if (-not (Test-Path -LiteralPath $docPath -PathType Leaf)) { throw "Word file not found: $docPath" }
if ($payload.doc_sha256 -and (Get-FileHash -LiteralPath $docPath -Algorithm SHA256).Hash -ne $payload.doc_sha256) {
  throw 'Word file changed since payload generation; inspect current fields and regenerate the task payload.'
}
foreach ($group in $payload.groups) {
  if ($group.action -and $group.action -ne 'add') { throw 'This script is add-only; use project-specific replace/delete steps.' }
}
if ($RemoveExistingReferencesHeading -and -not $ReplaceExistingZoteroFields) {
  throw 'References heading removal is only supported with an explicitly authorized full rebuild.'
}

$resolvedZoteroWordDllPath = Resolve-ZoteroWordDllPath $ZoteroWordDllPath
Add-LogLine "zotero_word_dll=$resolvedZoteroWordDllPath"
$dllLiteral = $resolvedZoteroWordDllPath.Replace('"', '""')

$source = @"
using System;
using System.Runtime.InteropServices;

public static class ZoteroWordDll {
  const string DllPath = @"$dllLiteral";

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern void clearError();

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern IntPtr getError();

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern UInt16 getDocument([MarshalAs(UnmanagedType.LPWStr)] string documentName, out IntPtr returnValue);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall)]
  public static extern void freeDocument(IntPtr doc);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern UInt16 setDocumentData(IntPtr doc, [MarshalAs(UnmanagedType.LPWStr)] string documentData);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern UInt16 insertField(IntPtr doc, [MarshalAs(UnmanagedType.LPWStr)] string fieldType, UInt16 noteType, out IntPtr returnValue);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern UInt16 setCode(IntPtr field, [MarshalAs(UnmanagedType.LPWStr)] string code);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall, CharSet=CharSet.Unicode)]
  public static extern UInt16 setText(IntPtr field, [MarshalAs(UnmanagedType.LPWStr)] string text, [MarshalAs(UnmanagedType.I1)] bool isRich);

  [DllImport(DllPath, CallingConvention=CallingConvention.StdCall)]
  public static extern UInt16 complete(IntPtr doc);
}
"@

Add-Type -TypeDefinition $source

Add-LogLine "insert_start doc=$docPath payload=$PayloadPath"
$backupPath = "$docPath.before_insert_$(Get-Date -Format yyyyMMdd_HHmmss)_$([guid]::NewGuid().ToString('N').Substring(0,8)).docx"
Copy-Item -LiteralPath $docPath -Destination $backupPath -ErrorAction Stop
Add-LogLine "backup=$backupPath"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$docPtr = [IntPtr]::Zero
$doc = $null
$completed = $false

try {
  $doc = $word.Documents.Open($docPath, $false, $false)
  if ($doc.ReadOnly) { throw 'Word opened read-only; stop before writing.' }
  Add-LogLine "opened paragraphs=$($doc.Paragraphs.Count) fields_before=$($doc.Fields.Count)"

  $existingIds = @{}
  $existingItems = 0
  $existingBibs = 0
  foreach ($field in $doc.Fields) {
    $code = $field.Code.Text
    if ($code -match 'ZOTERO_ITEM CSL_CITATION') {
      $existingItems++
      $obj = $code.Substring($code.IndexOf('{')).Trim() | ConvertFrom-Json
      if ($existingIds.ContainsKey([string]$obj.citationID)) { throw 'Duplicate citationID already in Word.' }
      $existingIds[[string]$obj.citationID] = $true
    }
    if ($code -match 'ZOTERO_BIBL') { $existingBibs++ }
  }
  if ($existingBibs -gt 1) { throw 'Multiple existing bibliography fields; resolve before adding citations.' }
  $targets = @{}
  $newIds = @{}
  foreach ($group in $payload.groups) {
    $id = [string]$group.group_id
    if (-not $id -or $targets.ContainsKey($id)) { throw 'Missing or duplicate group_id' }
    $code = [string]$group.field_code
    $obj = $code.Substring($code.IndexOf('{')).Trim() | ConvertFrom-Json
    $cid = [string]$obj.citationID
    if (-not $cid -or $newIds.ContainsKey($cid) -or (-not $ReplaceExistingZoteroFields -and $existingIds.ContainsKey($cid))) {
      throw "Duplicate or existing citationID: $cid. Inspect saved fields; do not blindly rerun."
    }
    $newIds[$cid] = $true
    # Verify every target before the first field mutation; live ranges track later edits.
    $targets[$id] = Find-TargetRange $doc ([string]$group.anchor_text) ([string]$group.position) ([string]$group.anchor_occurrence) ([string]$group.anchor_selection_reason)
  }

  if ($ReplaceExistingZoteroFields) {
    for ($i = $doc.Fields.Count; $i -ge 1; $i--) {
      $field = $doc.Fields.Item($i)
      if ($field.Type -eq 81 -and $field.Code.Text -like "*ZOTERO*") {
        $field.Select()
        $word.Selection.Delete()
      }
    }
    Add-LogLine "removed_existing_zotero_fields fields_now=$($doc.Fields.Count)"
  }

  if ($RemoveExistingReferencesHeading) {
    for ($i = $doc.Paragraphs.Count; $i -ge 1; $i--) {
      $text = ($doc.Paragraphs.Item($i).Range.Text -replace "`r", "" -replace "`a", "").Trim()
      if ($text -eq "References") {
        $doc.Paragraphs.Item($i).Range.Delete()
      }
    }
    Add-LogLine "removed_existing_references_heading_if_any"
  }

  $status = [ZoteroWordDll]::getDocument($null, [ref]$docPtr)
  Check-Status $status "getDocument"
  if ($existingItems -eq 0 -or $ReplaceExistingZoteroFields) {
    Check-Status ([ZoteroWordDll]::setDocumentData($docPtr, [string]$payload.doc_prefs)) "setDocumentData"
  } else {
    Add-LogLine 'preserved_existing_document_preferences; style compatibility must be checked by the execution plan'
  }

  foreach ($group in $payload.groups) {
    $range = $targets[[string]$group.group_id]
    $range.Select()
    $fieldPtr = [IntPtr]::Zero
    Check-Status ([ZoteroWordDll]::insertField($docPtr, "Field", 0, [ref]$fieldPtr)) ("insertField " + [string]$group.group_id)
    $dllCode = ([string]$group.field_code) -replace "^ZOTERO_", ""
    Check-Status ([ZoteroWordDll]::setCode($fieldPtr, $dllCode)) ("setCode " + [string]$group.group_id)
    Check-Status ([ZoteroWordDll]::setText($fieldPtr, [string]$group.placeholder, $false)) ("setText " + [string]$group.group_id)
    Add-LogLine ("inserted group={0} citations={1}" -f $group.group_id, (($group.citation_ids -join ";")))
  }

  if ($null -ne $payload.bibliography -and ($existingBibs -eq 0 -or $ReplaceExistingZoteroFields)) {
    $selection = $word.Selection
    $null = $selection.EndKey(6)
    $selection.TypeParagraph()
    $selection.TypeText([string]$payload.bibliography.heading)
    $selection.TypeParagraph()
    $bibPtr = [IntPtr]::Zero
    Check-Status ([ZoteroWordDll]::insertField($docPtr, "Field", 0, [ref]$bibPtr)) "insertField bibliography"
    $bibCode = ([string]$payload.bibliography.field_code) -replace "^ZOTERO_", ""
    Check-Status ([ZoteroWordDll]::setCode($bibPtr, $bibCode)) "setCode bibliography"
    Check-Status ([ZoteroWordDll]::setText($bibPtr, [string]$payload.bibliography.placeholder, $false)) "setText bibliography"
  } elseif ($existingBibs -eq 1) {
    Add-LogLine 'reused_existing_bibliography'
  }

  Check-Status ([ZoteroWordDll]::complete($docPtr)) "complete"
  $docPtr = [IntPtr]::Zero
  $doc.Save()
  Add-LogLine "saved_after_insert fields_after=$($doc.Fields.Count)"
  $doc.Close($true)
  Add-LogLine "closed"
  $completed = $true
} catch {
  Add-LogLine "FAILED: $($_.Exception.Message); Word left open for inspection; backup=$backupPath"
  $word.Visible = $true
  throw
} finally {
  if ($docPtr -ne [IntPtr]::Zero) {
    [ZoteroWordDll]::freeDocument($docPtr)
  }
  if ($null -ne $doc) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
  }
  if ($completed) { $word.Quit() }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
