<#
.SYNOPSIS
Hardened Toolkit - System Cleaner, Duplicate Finder, Uninstaller & More (GUI + Console)
.DESCRIPTION
Runs a dark-themed GUI by default. Use -Console for the text menu.
.VERSION
2.4.0 (2026-10-22)
.AUTHOR
SecuritySuite
.SAFETY
  * SCAN anywhere. DELETE always skips Windows, Program Files, ProgramData,
    drive roots, the user profile root, the Users folder, other user
    profiles and this tool's log folder.
  * Deletions bypass the Recycle Bin (irreversible). Links are removed, never followed.
  * Duplicate deletion re-hashes the keeper and every candidate immediately
    before removal; changed files are skipped.
  * Scan results are exported to your Desktop before any delete prompt.
  * Deletes are logged to %LOCALAPPDATA%\HardenedToolkit\deleted.log
.NOTES
  Requires Windows PowerShell 5.1 (built into Windows 10/11).
  The script requests administrator rights because cleaning system folders
  and installing to Program Files requires them.
#>

param(
  [switch]$Console,
  [switch]$Gui,
  [switch]$Uninstall,
  [switch]$Quiet,
  [switch]$Hidden,
  [switch]$SmokeTest
)

$ErrorActionPreference = 'Stop'
$AppName      = "HardenedToolkit"
$Version      = "2.4.0"
$DisplayTitle = "Hardened Toolkit v$Version"
$ShortcutName = "Hardened Toolkit.lnk"
$IconName     = "HardenedToolkit.ico"

$script:ErrorLog  = Join-Path $env:LOCALAPPDATA "$AppName\errors.log"
$script:DeleteLog = Join-Path $env:LOCALAPPDATA "$AppName\deleted.log"
$script:LogDir    = Join-Path $env:LOCALAPPDATA $AppName
$script:UiMode    = 'console'
$script:LogBox    = $null
$script:Busy      = $false
$script:MainForm  = $null

# =========================================================
# UI abstraction (works in both GUI and console mode)
# =========================================================
function Write-Ui {
  param([string]$Text, [string]$Color = 'Gray')
  $Text = [string]$Text
  if ($script:UiMode -eq 'gui' -and $script:LogBox -and -not $script:LogBox.IsDisposed) {
    $script:LogBox.AppendText((($Text -replace "`r?`n", "`r`n")) + "`r`n")
    $script:LogBox.SelectionStart = $script:LogBox.TextLength
    $script:LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
  } else {
    Write-Host $Text -ForegroundColor $Color
  }
}

function Write-Table {
  param($Data)
  if ($null -eq $Data) { Write-Ui "(none)"; return }
  $arr = @($Data)
  if ($arr.Count -eq 0) { Write-Ui "(none)"; return }
  $text = ($arr | Format-Table -AutoSize | Out-String -Width 180).TrimEnd()
  Write-Ui $text
}

function Confirm-Action {
  param([string]$Message)
  if ($script:UiMode -eq 'gui') {
    $r = [System.Windows.Forms.MessageBox]::Show($Message, $DisplayTitle,
         [System.Windows.Forms.MessageBoxButtons]::YesNo,
         [System.Windows.Forms.MessageBoxIcon]::Warning)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
  }
  Write-Host "`n$Message" -ForegroundColor Cyan
  $resp = (Read-Host "Type Y or YES to proceed (anything else cancels)").Trim().ToUpperInvariant()
  return ($resp -eq 'Y' -or $resp -eq 'YES')
}

function Show-Message {
  param([string]$Message, [string]$Icon = 'Information')
  if ($script:UiMode -ne 'gui') { return }
  $ico = [System.Windows.Forms.MessageBoxIcon]::Information
  switch ($Icon) {
    'Error'       { $ico = [System.Windows.Forms.MessageBoxIcon]::Error }
    'Warning'     { $ico = [System.Windows.Forms.MessageBoxIcon]::Warning }
    'Question'    { $ico = [System.Windows.Forms.MessageBoxIcon]::Question }
    'Information' { $ico = [System.Windows.Forms.MessageBoxIcon]::Information }
  }
  [void][System.Windows.Forms.MessageBox]::Show($Message, $DisplayTitle,
    [System.Windows.Forms.MessageBoxButtons]::OK, $ico)
}

function Select-FromList {
  param([string]$Title, [string[]]$Items)
  if (-not $Items -or $Items.Count -eq 0) { return -1 }

  if ($script:UiMode -eq 'gui') {
    $all = @($Items)
    $map = New-Object System.Collections.ArrayList
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(680, 560)
    $dlg.MinimumSize = New-Object System.Drawing.Size(520, 400)
    $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(28,28,30)
    $dlg.ForeColor = [System.Drawing.Color]::White

    $lb = New-Object System.Windows.Forms.ListBox
    $lb.Dock = 'Fill'
    $lb.BackColor = [System.Drawing.Color]::FromArgb(18,18,20)
    $lb.ForeColor = [System.Drawing.Color]::Gainsboro
    $lb.BorderStyle = 'None'
    $lb.IntegralHeight = $false
    $dlg.Controls.Add($lb)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Dock = 'Top'
    $txtFilter.Height = 28
    $txtFilter.BackColor = [System.Drawing.Color]::FromArgb(37,37,40)
    $txtFilter.ForeColor = [System.Drawing.Color]::White
    $txtFilter.BorderStyle = 'FixedSingle'
    $dlg.Controls.Add($txtFilter)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Select"
    $ok.Dock = 'Bottom'
    $ok.Height = 40
    $ok.FlatStyle = 'Flat'
    $ok.BackColor = [System.Drawing.Color]::FromArgb(45,45,48)
    $ok.ForeColor = [System.Drawing.Color]::White
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($ok)
    $dlg.AcceptButton = $ok

    $refresh = {
      $f = $txtFilter.Text.Trim()
      $lb.BeginUpdate()
      try {
        $lb.Items.Clear()
        $map.Clear()
        for ($i = 0; $i -lt $all.Count; $i++) {
          if ($f -eq '' -or $all[$i].IndexOf($f, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            [void]$lb.Items.Add($all[$i])
            [void]$map.Add($i)
          }
        }
      } finally { $lb.EndUpdate() }
      if ($lb.Items.Count -gt 0) { $lb.SelectedIndex = 0 }
    }
    $txtFilter.Add_TextChanged($refresh.GetNewClosure())
    $lb.Add_DoubleClick({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK }.GetNewClosure())
    & $refresh

    $result = -1
    $owner = $script:MainForm
    if ($owner -and -not $owner.IsDisposed) {
      if ($dlg.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { $result = $lb.SelectedIndex }
    } else {
      if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $result = $lb.SelectedIndex }
    }
    if ($result -ge 0 -and $result -lt $map.Count) { return [int]$map[$result] }
    return -1
  }

  $filter = (Read-Host "Filter list (Enter for all)").Trim()
  $shown = New-Object System.Collections.ArrayList
  for ($i = 0; $i -lt $Items.Count; $i++) {
    if ($filter -eq '' -or $Items[$i].IndexOf($filter, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
      [void]$shown.Add($i)
    }
  }
  if ($shown.Count -eq 0) { Write-Host "No matches." -ForegroundColor Yellow; return -1 }
  for ($i = 0; $i -lt $shown.Count; $i++) { Write-Host "$($i+1). $($Items[$shown[$i]])" }
  $n = Read-Host "Enter number (Enter to cancel)"
  $idx = 0
  if ([int]::TryParse($n, [ref]$idx) -and $idx -ge 1 -and $idx -le $shown.Count) { return [int]$shown[$idx - 1] }
  return -1
}

# =========================================================
# Logging / error handling
# =========================================================
function Write-ErrorLog {
  param([string]$Name, $ErrorRecord)
  try {
    $dir = Split-Path $script:ErrorLog -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $entry = "[{0}] {1}: {2}`r`n{3}`r`n" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Name,
             $ErrorRecord.Exception.Message, $ErrorRecord.ScriptStackTrace
    Add-Content -Path $script:ErrorLog -Value $entry -Encoding UTF8
  } catch { }
}

function Write-DeleteLog {
  param([string]$Path, [string]$Result)
  try {
    $dir = Split-Path $script:DeleteLog -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -Path $script:DeleteLog -Value ("{0}`t{1}`t{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Result, $Path) -Encoding UTF8
  } catch { }
}

function Invoke-SafeAction {
  param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Action)
  try {
    & $Action
  } catch {
    Write-Ui "[!] '$Name' failed: $($_.Exception.Message)" 'Red'
    Write-ErrorLog -Name $Name -ErrorRecord $_
    Write-Ui "   Details logged to: $script:ErrorLog" 'DarkGray'
    Show-Message "'$Name' failed:`r`n$($_.Exception.Message)" 'Error'
  }
}

function Pause-OnError {
  param([string]$Message)
  Write-Host "`n[!] $Message" -ForegroundColor Red
  Read-Host "Press Enter to exit" | Out-Null
}

# =========================================================
# Elevation
# =========================================================
function Require-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
  if ($isAdmin) { return }

  Write-Host "`n[!]  Administrator privileges required - relaunching elevated..." -ForegroundColor Yellow
  if (-not $PSCommandPath -or -not (Test-Path $PSCommandPath)) {
    Pause-OnError "Cannot relaunch elevated because the script path is unavailable."
    exit 1
  }
  $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
  if ($Console)   { $argList += '-Console' }
  if ($Gui)       { $argList += '-Gui' }
  if ($Uninstall) { $argList += '-Uninstall' }
  if ($Quiet)     { $argList += '-Quiet' }
  if ($Hidden)    { $argList += '-Hidden' }
  try {
    $exe = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -LiteralPath $exe)) { $exe = 'powershell.exe' }
    $sp = @{ FilePath = $exe; ArgumentList = $argList; Verb = 'RunAs' }
    if ($Hidden) { $sp.WindowStyle = 'Hidden' }
    Start-Process @sp
    exit 0
  } catch {
    Pause-OnError "Failed to relaunch elevated: $($_.Exception.Message)"
    exit 1
  }
}

# =========================================================
# Path safety
# =========================================================
function Get-NormalizedPath {
  param([string]$Path)
  if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
  try {
    $p = $Path -replace '^\\\\\?\\UNC\\', '\\'
    $p = $p -replace '^\\\\\?\\', ''
    $full = [System.IO.Path]::GetFullPath($p)
    $root = [System.IO.Path]::GetPathRoot($full)
    if ([string]::IsNullOrEmpty($root)) { return $full.TrimEnd('\') }
    $rest = $full.Substring($root.Length)
    $parts = @()
    foreach ($seg in ($rest -split '\\')) {
      if ($seg -eq '') { continue }
      $clean = $seg.TrimEnd([char[]]@(' ', '.'))
      if ($clean -ne '') { $parts += $clean }
    }
    if ($parts.Count -eq 0) { return $root }
    return ((Join-Path $root ($parts -join '\'))).TrimEnd('\')
  } catch { return $null }
}

$usersRoot = $null
if ($env:USERPROFILE) { $usersRoot = Split-Path $env:USERPROFILE -Parent }
$sensitiveRoots = @(
  $env:SystemRoot,
  ${env:ProgramFiles},
  ${env:ProgramFiles(x86)},
  $env:ProgramData,
  (Join-Path $env:LOCALAPPDATA $AppName),
  (Join-Path $env:SystemDrive 'Recovery'),
  (Join-Path $env:SystemDrive 'Boot'),
  (Join-Path $env:SystemDrive '$Recycle.Bin'),
  (Join-Path $env:SystemDrive 'System Volume Information'),
  (Join-Path $env:SystemDrive 'PerfLogs'),
  (Join-Path $env:SystemDrive 'Windows.old'),
  (Join-Path $env:SystemDrive 'Config.Msi'),
  (Join-Path $env:SystemDrive 'Documents and Settings'),
  (Join-Path $env:SystemDrive 'Users\Public'),
  (Join-Path $env:SystemDrive 'Users\Default'),
  (Join-Path $env:SystemDrive 'Users\Default User')
)
if ($usersRoot -and (Test-Path -LiteralPath $usersRoot -PathType Container)) {
  $sensitiveRoots += @(Get-ChildItem -LiteralPath $usersRoot -Directory -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -and -not $_.FullName.Equals($env:USERPROFILE, [System.StringComparison]::OrdinalIgnoreCase) } |
    ForEach-Object { $_.FullName })
}
$protectedRoots = @($sensitiveRoots |
  Where-Object { $_ } | ForEach-Object { Get-NormalizedPath $_ } | Where-Object { $_ })

$protectedExact = @($env:USERPROFILE, $usersRoot, "$env:SystemDrive\", $PSCommandPath) |
  Where-Object { $_ } | ForEach-Object { Get-NormalizedPath $_ } | Where-Object { $_ }

function Test-PathIsProtected {
  param([string]$Path)
  $norm = Get-NormalizedPath $Path
  if (-not $norm) { return $true }
  try {
    $normRoot = Get-NormalizedPath ([System.IO.Path]::GetPathRoot($norm))
    if ($normRoot -and $normRoot.Equals($norm, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  } catch { return $true }
  foreach ($exact in $protectedExact) {
    if ($exact -and $norm.Equals($exact, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  }
  foreach ($root in $protectedRoots) {
    if ($root -and ($norm.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -or
                    $norm.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase))) { return $true }
  }
  return $false
}

# Walks a folder without following reparse points. Returns files and directories
# that are safe to inspect; links themselves are never included.
function Get-TreeSnapshot {
  param([Parameter(Mandatory)][string]$Root)
  $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
  $dirs  = New-Object 'System.Collections.Generic.List[System.IO.DirectoryInfo]'
  $stack = New-Object 'System.Collections.Generic.Stack[string]'
  $start = Get-NormalizedPath $Root
  if (-not $start -or -not (Test-Path -LiteralPath $start -PathType Container)) {
    return [pscustomobject]@{ Files = $files; Directories = $dirs }
  }
  try {
    if (([System.IO.File]::GetAttributes($start) -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      return [pscustomobject]@{ Files = $files; Directories = $dirs }
    }
  } catch {
    return [pscustomobject]@{ Files = $files; Directories = $dirs }
  }
  $stack.Push($start)
  while ($stack.Count -gt 0) {
    $dir = $stack.Pop()
    try {
      foreach ($p in [System.IO.Directory]::EnumerateFileSystemEntries($dir, '*', [System.IO.SearchOption]::TopDirectoryOnly)) {
        try {
          $attr = [System.IO.File]::GetAttributes($p)
          $isDir  = (($attr -band [System.IO.FileAttributes]::Directory) -ne 0)
          $isLink = (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
          if ($isDir) {
            if (-not $isLink) {
              $dirs.Add((New-Object System.IO.DirectoryInfo($p)))
              $stack.Push($p)
            }
          } elseif (-not $isLink) {
            $files.Add((New-Object System.IO.FileInfo($p)))
          }
        } catch { }
      }
    } catch { }
  }
  return [pscustomobject]@{ Files = $files; Directories = $dirs }
}

# Single guarded delete. Directories are removed only if empty; links are removed, never followed.
function Remove-SafeItem {
  param([Parameter(Mandatory)][string]$Path, [switch]$AllowProtected)
  if (-not $AllowProtected -and (Test-PathIsProtected $Path)) {
    Write-DeleteLog $Path 'SKIP-PROTECTED'
    return $false
  }
  try {
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.PSIsContainer) {
      [System.IO.Directory]::Delete($item.FullName, $false)
    } else {
      Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
    }
    Write-DeleteLog $Path 'DELETED'
    return $true
  } catch {
    Write-DeleteLog $Path "FAILED: $($_.Exception.Message)"
    return $false
  }
}

# =========================================================
# Folder picking / export
# =========================================================
function Get-TargetFolder {
  param([string]$PromptMessage)
  $path = $null

  if ($script:UiMode -eq 'gui') {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $PromptMessage
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $path = $dlg.SelectedPath
  } else {
    Write-Host "`n$PromptMessage" -ForegroundColor Cyan
    Write-Host "  Type path, or type B to browse, or Enter to cancel"
    $path = Read-Host "Folder path"
    if ([string]::IsNullOrWhiteSpace($path)) { return $null }
    $path = $path.Trim().Trim('"')
    if ($path -eq 'B' -or $path -eq 'b') {
      try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $PromptMessage
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $path = $dlg.SelectedPath } else { return $null }
      } catch {
        Write-Host "[!]  Browser unavailable - type path manually" -ForegroundColor Yellow
        $path = Read-Host "Folder path"
        if (-not $path) { return $null }
      }
    }
  }

  $norm = Get-NormalizedPath $path
  if (-not $norm -or -not (Test-Path -LiteralPath $norm -PathType Container)) {
    Write-Ui "[!] Path not found: $path" 'Red'
    return $null
  }

  $drv = [System.IO.Path]::GetPathRoot($norm).TrimEnd('\')
  if ($drv -eq $env:SystemDrive -and $norm.TrimEnd('\') -eq $drv) {
    if (-not (Confirm-Action "[!]  Scanning the entire $drv drive can take MANY MINUTES. Proceed anyway?")) { return $null }
  }
  return $norm
}

function Export-Results {
  param([string]$Label, [array]$Lines)
  $dir = [Environment]::GetFolderPath("Desktop")
  if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) {
    $dir = [Environment]::GetFolderPath("MyDocuments")
  }
  if (-not $dir -or -not (Test-Path -LiteralPath $dir -PathType Container)) {
    $dir = [System.IO.Path]::GetTempPath()
  }
  $out = Join-Path $dir "$Label-$(Get-Date -Format 'yyyyMMdd-HHmmss-fff').txt"
  $text = @($Lines | ForEach-Object { [string]$_ })
  [System.IO.File]::WriteAllLines($out, [string[]]$text, [System.Text.UTF8Encoding]::new($true))
  Write-Ui "[FILE] List saved: $out" 'Green'
}

function Get-FolderSize {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return 0L }
  $snap = Get-TreeSnapshot $Path
  $sum = 0L
  foreach ($f in $snap.Files) {
    try { $sum += [int64]$f.Length } catch { }
  }
  return $sum
}

# =========================================================
# Installer / uninstaller
# =========================================================
function Get-StartMenuProgramsPath {
  $base = [Environment]::GetFolderPath("Programs")
  if (-not $base) { $base = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs" }
  return $base
}

function Get-ProgramFilesPath {
  $p = $env:ProgramW6432
  if (-not $p) { $p = [Environment]::GetFolderPath('ProgramFiles') }
  if (-not $p) { $p = $env:ProgramFiles }
  if (-not $p) { throw "The Program Files folder could not be located." }
  return $p
}

function Get-ShortcutDirs {
  $dirs = @()
  $desktop = [Environment]::GetFolderPath('Desktop')
  if ($desktop) { $dirs += $desktop }
  $sm = Get-StartMenuProgramsPath
  if ($sm) { $dirs += $sm }
  return @($dirs | Select-Object -Unique)
}

function Get-LegacyShortcutPaths {
  $out = @()
  foreach ($dir in Get-ShortcutDirs) {
    if (Test-Path -LiteralPath $dir -PathType Container) {
      $out += @(Get-ChildItem -LiteralPath $dir -Filter "Hardened Toolkit v*.lnk" -File -Force -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty FullName)
    }
  }
  return @($out)
}

function New-ToolkitShortcut {
  param(
    [Parameter(Mandatory)][string]$LnkPath,
    [Parameter(Mandatory)][string]$TargetPath,
    [Parameter(Mandatory)][string]$Arguments,
    [Parameter(Mandatory)][string]$WorkingDirectory,
    [Parameter(Mandatory)][string]$Description,
    [string]$IconPath
  )
  $shell = New-Object -ComObject WScript.Shell
  $sc = $shell.CreateShortcut($LnkPath)
  $sc.TargetPath = $TargetPath
  $sc.Arguments = $Arguments
  $sc.WorkingDirectory = $WorkingDirectory
  $sc.Description = $Description
  if ($IconPath -and (Test-Path -LiteralPath $IconPath)) {
    $sc.IconLocation = "$IconPath,0"
  } else {
    $sc.IconLocation = "$TargetPath,0"
  }
  $sc.Save()
  try {
    $bytes = [System.IO.File]::ReadAllBytes($LnkPath)
    if ($bytes.Length -gt 0x15) {
      $bytes[0x15] = $bytes[0x15] -bor 0x20   # "Run as administrator" flag
      [System.IO.File]::WriteAllBytes($LnkPath, $bytes)
    }
    return $true
  } catch {
    Write-Ui "[!] Shortcut created, but the auto-admin flag could not be set: $($_.Exception.Message)" 'Yellow'
    return $false
  }
}

function Install-Toolkit {
  Write-Ui "`n=== Install $DisplayTitle ===" 'Cyan'
  if (-not $PSCommandPath) { throw "Script path unavailable; run this from a saved .ps1 file." }
  $programFiles = Get-ProgramFilesPath
  $targetDir = Join-Path $programFiles $AppName
  Write-Ui "Install folder: $targetDir"
  if (-not (Confirm-Action "Copy this tool to Program Files and create Desktop + Start Menu shortcuts?")) { return }

  if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
  }
  $destFile = Join-Path $targetDir "$AppName.ps1"
  $srcNorm = Get-NormalizedPath $PSCommandPath
  $dstNorm = Get-NormalizedPath $destFile
  if (-not $srcNorm -or -not $dstNorm -or -not $srcNorm.Equals($dstNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $PSCommandPath -Destination $destFile -Force
  }
  if (-not (Test-Path -LiteralPath $destFile)) { throw "The script could not be copied to $destFile" }
  Write-Ui "[OK] Installed to $destFile" 'Green'

  $iconDest = Join-Path $targetDir $IconName
  $sourceIcon = Join-Path (Split-Path -Parent $PSCommandPath) $IconName
  $srcIconNorm = Get-NormalizedPath $sourceIcon
  $dstIconNorm = Get-NormalizedPath $iconDest
  if ((Test-Path -LiteralPath $sourceIcon) -and $srcIconNorm -and $dstIconNorm -and
      -not $srcIconNorm.Equals($dstIconNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
    Copy-Item -LiteralPath $sourceIcon -Destination $iconDest -Force
  }
  $iconForShortcut = $null
  if (Test-Path -LiteralPath $iconDest) {
    $iconForShortcut = $iconDest
    Write-Ui "[OK] Icon installed to $iconDest" 'Green'
  } else {
    Write-Ui "[!] $IconName was not found next to the script; shortcuts will use the PowerShell icon." 'Yellow'
  }

  $psExe = Join-Path $PSHOME 'powershell.exe'
  if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
  $runArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$destFile`" -Hidden"

  $desktop = [Environment]::GetFolderPath('Desktop')
  $smDir = Get-StartMenuProgramsPath
  $created = @()
  if ($desktop -and (Test-Path -LiteralPath $desktop -PathType Container)) {
    $lnkPath = Join-Path $desktop $ShortcutName
    if (New-ToolkitShortcut -LnkPath $lnkPath -TargetPath $psExe -Arguments $runArgs -WorkingDirectory $targetDir -Description "Launch $DisplayTitle" -IconPath $iconForShortcut) {
      Write-Ui "[OK] Desktop shortcut created (runs as Admin)" 'Green'
    }
    $created += $lnkPath
  }
  if ($smDir) {
    if (-not (Test-Path -LiteralPath $smDir -PathType Container)) {
      New-Item -ItemType Directory -Path $smDir -Force | Out-Null
    }
    $menuLnk = Join-Path $smDir $ShortcutName
    if ($created.Count -gt 0 -and (Test-Path -LiteralPath $created[0])) {
      Copy-Item -LiteralPath $created[0] -Destination $menuLnk -Force
    } else {
      [void](New-ToolkitShortcut -LnkPath $menuLnk -TargetPath $psExe -Arguments $runArgs -WorkingDirectory $targetDir -Description "Launch $DisplayTitle" -IconPath $iconForShortcut)
    }
    Write-Ui "[OK] Start Menu shortcut created" 'Green'
  }

  foreach ($old in Get-LegacyShortcutPaths) {
    try { Remove-Item -LiteralPath $old -Force -ErrorAction Stop; Write-Ui "[i] Removed old shortcut: $old" 'DarkGray' } catch { }
  }

  $regPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppName"
  if (-not (Test-Path -LiteralPath $regPath)) { New-Item -Path $regPath -Force | Out-Null }
  $uninstallCmd = "`"$psExe`" -NoProfile -ExecutionPolicy Bypass -File `"$destFile`" -Uninstall"
  $quietCmd = "$uninstallCmd -Quiet"
  $props = @{
    DisplayName          = $DisplayTitle
    DisplayVersion       = $Version
    Publisher            = "SecuritySuite"
    InstallLocation      = $targetDir
    DisplayIcon          = $(if ($iconForShortcut) { "$iconForShortcut,0" } else { "$psExe,0" })
    UninstallString      = $uninstallCmd
    QuietUninstallString = $quietCmd
    InstallDate          = (Get-Date -Format 'yyyyMMdd')
  }
  foreach ($k in $props.Keys) {
    New-ItemProperty -Path $regPath -Name $k -Value $props[$k] -PropertyType String -Force | Out-Null
  }
  $sizeKb = [int][math]::Ceiling((Get-FolderSize $targetDir) / 1KB)
  New-ItemProperty -Path $regPath -Name 'EstimatedSize' -Value $sizeKb -PropertyType DWord -Force | Out-Null
  New-ItemProperty -Path $regPath -Name 'NoModify' -Value 1 -PropertyType DWord -Force | Out-Null
  New-ItemProperty -Path $regPath -Name 'NoRepair' -Value 1 -PropertyType DWord -Force | Out-Null
  Write-Ui "`n[OK] INSTALL COMPLETE - use the Desktop shortcut from now on!" 'Green'
}

function Uninstall-Toolkit {
  Write-Ui "`n=== Uninstall $DisplayTitle ===" 'Cyan'
  if (-not $Quiet) {
    if (-not (Confirm-Action "Remove program files, shortcuts and the registry entry? Logs in $env:LOCALAPPDATA\$AppName are kept.")) { return }
  }

  $programFiles = Get-ProgramFilesPath
  $targetDir = Join-Path $programFiles $AppName
  $targets = New-Object System.Collections.ArrayList
  foreach ($dir in Get-ShortcutDirs) {
    [void]$targets.Add((Join-Path $dir $ShortcutName))
  }
  foreach ($old in Get-LegacyShortcutPaths) { [void]$targets.Add($old) }
  [void]$targets.Add($targetDir)
  [void]$targets.Add("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppName")

  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($t in $targets) {
    if (-not $t) { continue }
    if (-not $seen.Add($t)) { continue }
    try {
      if (Test-Path -LiteralPath $t) {
        Remove-Item -LiteralPath $t -Force -Recurse -ErrorAction Stop
        Write-Ui "[OK] Removed: $t" 'Green'
      } else {
        Write-Ui "[i] Not present: $t" 'DarkGray'
      }
    } catch {
      Write-Ui "[!] Could not remove $t : $($_.Exception.Message)" 'Yellow'
    }
  }
  Write-Ui "`n[OK] Uninstall finished. Logs kept in: $(Join-Path $env:LOCALAPPDATA $AppName)" 'Green'
}

# =========================================================
# Scans / cleaners
# =========================================================
function Get-DuplicateGroups {
  param([Parameter(Mandatory)][string]$Root)
  $snap = Get-TreeSnapshot $Root
  $files = @($snap.Files | Where-Object { try { $_.Length -gt 0 } catch { $false } })
  $sizeGroups = @($files | Group-Object Length | Where-Object { $_.Count -gt 1 })

  $groups = New-Object System.Collections.ArrayList
  $wasted = 0L
  foreach ($grp in $sizeGroups) {
    $hashGroups = @($grp.Group | Group-Object {
      try { (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $null }
    } | Where-Object { $_.Name -and $_.Count -gt 1 })

    foreach ($hg in $hashGroups) {
      $sorted = @($hg.Group | Sort-Object FullName)
      $keep = $sorted[0]
      $dups = @($sorted | Select-Object -Skip 1)
      if ($dups.Count -gt 0) {
        $dupSizes = @($dups | ForEach-Object { try { [int64]$_.Length } catch { 0L } })
        $wasted += [int64](($dupSizes | Measure-Object -Sum).Sum)
        [void]$groups.Add([pscustomobject]@{
          Hash = [string]$hg.Name
          Keep = [string]$keep.FullName
          Dups = @($dups | ForEach-Object { [string]$_.FullName })
        })
      }
    }
  }
  return [pscustomobject]@{ Groups = @($groups); Wasted = $wasted }
}

# Re-hashes the keeper and every candidate immediately before deleting anything.
function Remove-DuplicateGroups {
  param([array]$Groups)
  $removed = 0
  $skipped = 0
  foreach ($grp in @($Groups)) {
    $dups = @($grp.Dups)
    $keepItem = Get-Item -LiteralPath $grp.Keep -Force -ErrorAction SilentlyContinue
    if (-not $keepItem -or $keepItem.PSIsContainer -or $keepItem.LinkType) { $skipped += $dups.Count; continue }
    $keepHash = $null
    $keepLen = -1L
    try {
      $keepLen = [int64]$keepItem.Length
      $keepHash = (Get-FileHash -LiteralPath $grp.Keep -Algorithm SHA256 -ErrorAction Stop).Hash
    } catch { }
    if (-not $keepHash -or $keepHash -ne $grp.Hash) { $skipped += $dups.Count; continue }
    foreach ($f in $dups) {
      $cur = Get-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
      if (-not $cur -or $cur.PSIsContainer -or $cur.LinkType) { $skipped++; continue }
      $curLen = -1L
      try { $curLen = [int64]$cur.Length } catch { $skipped++; continue }
      if ($curLen -ne $keepLen) { $skipped++; continue }
      $curHash = $null
      try { $curHash = (Get-FileHash -LiteralPath $f -Algorithm SHA256 -ErrorAction Stop).Hash } catch { }
      if (-not $curHash -or $curHash -ne $grp.Hash) { $skipped++; continue }
      if (Remove-SafeItem -Path $f) { $removed++ } else { $skipped++ }
    }
  }
  return [pscustomobject]@{ Removed = $removed; Skipped = $skipped }
}

function Find-Duplicates {
  $root = Get-TargetFolder "Select folder to scan for duplicates"
  if (-not $root) { return }
  Write-Ui "`nScanning $root - please wait..." 'Cyan'

  $scan = Get-DuplicateGroups $root
  $dupGroups = @($scan.Groups)
  $wasted = [int64]$scan.Wasted

  if ($dupGroups.Count -eq 0) { Write-Ui "[OK] No duplicates found" 'Green'; Show-Message "No duplicates found."; return }

  $lines = foreach ($g in $dupGroups) {
    "KEEP: $($g.Keep)"
    foreach ($d in $g.Dups) { "  DUP:  $d" }
    ""
  }
  Export-Results "Duplicate-Files" @($lines)
  $dupCount = [int](($dupGroups | ForEach-Object { @($_.Dups).Count } | Measure-Object -Sum).Sum)
  Write-Ui "Found $($dupGroups.Count) duplicate groups ($dupCount extra copies) - wasted space: $([math]::Round($wasted/1MB,2)) MB" 'Cyan'

  if (Confirm-Action "Delete $dupCount duplicate copies (one copy per group is kept)? This cannot be undone.") {
    $res = Remove-DuplicateGroups $dupGroups
    Write-Ui "Removed: $($res.Removed) | Skipped (protected, changed or failed): $($res.Skipped)" 'Green'
  }
}

function Find-EmptyFiles {
  $root = Get-TargetFolder "Select folder to scan for 0-byte files"
  if (-not $root) { return }
  Write-Ui "`nScanning $root - please wait..." 'Cyan'

  $snap = Get-TreeSnapshot $root
  $empty = @($snap.Files |
             Where-Object { try { $_.Length -eq 0 } catch { $false } } |
             Select-Object -ExpandProperty FullName)
  if ($empty.Count -eq 0) { Write-Ui "[OK] No empty files found" 'Green'; Show-Message "No empty files found."; return }

  Export-Results "Empty-Files" $empty
  Write-Ui "Found $($empty.Count) empty files" 'Cyan'

  if (Confirm-Action "Delete $($empty.Count) empty files? This cannot be undone.") {
    $removed = 0; $skipped = 0
    foreach ($f in $empty) {
      $cur = Get-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
      if (-not $cur -or $cur.PSIsContainer -or $cur.LinkType -or $cur.Length -ne 0) { $skipped++; continue }
      if (Remove-SafeItem -Path $f) { $removed++ } else { $skipped++ }
    }
    Write-Ui "Removed: $removed | Skipped (protected, changed or failed): $skipped" 'Green'
  }
}

function Find-EmptyFolders {
  $root = Get-TargetFolder "Select folder to scan for empty folders"
  if (-not $root) { return }
  Write-Ui "`nScanning $root - please wait..." 'Cyan'

  # Deepest first; a folder is "empty" if everything in it is already an empty folder.
  $snap = Get-TreeSnapshot $root
  $dirs = @($snap.Directories | Sort-Object { $_.FullName.Length } -Descending)
  $emptySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $empty = New-Object System.Collections.ArrayList

  foreach ($d in $dirs) {
    if (Test-PathIsProtected $d.FullName) { continue }
    $children = @(Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue)
    $allEmpty = $true
    foreach ($c in $children) {
      if ($c.LinkType) { $allEmpty = $false; break }
      if (-not $emptySet.Contains($c.FullName)) { $allEmpty = $false; break }
    }
    if ($allEmpty) { [void]$emptySet.Add($d.FullName); [void]$empty.Add($d.FullName) }
  }

  if ($empty.Count -eq 0) { Write-Ui "[OK] No empty folders found" 'Green'; Show-Message "No empty folders found."; return }

  Export-Results "Empty-Folders" @($empty)
  Write-Ui "Found $($empty.Count) empty folders" 'Cyan'

  if (Confirm-Action "Delete $($empty.Count) empty folders? This cannot be undone.") {
    $removed = 0; $skipped = 0
    foreach ($f in $empty) {   # already deepest-first
      $cur = Get-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue
      if (-not $cur -or -not $cur.PSIsContainer -or $cur.LinkType) { $skipped++; continue }
      if (Get-ChildItem -LiteralPath $f -Force -ErrorAction SilentlyContinue) { $skipped++; continue }
      if (Remove-SafeItem -Path $f) { $removed++ } else { $skipped++ }
    }
    Write-Ui "Removed: $removed | Skipped (protected, not empty or failed): $skipped" 'Green'
  }
}

function Find-LargestFiles {
  $root = Get-TargetFolder "Select folder to list largest files"
  if (-not $root) { return }
  Write-Ui "`nScanning $root - please wait..." 'Cyan'

  $snap = Get-TreeSnapshot $root
  $large = @($snap.Files | Sort-Object Length -Descending | Select-Object -First 20)
  if ($large.Count -eq 0) { Write-Ui "No files found." 'Yellow'; return }

  $lines = $large | ForEach-Object { "{0} MB - {1}" -f ([math]::Round($_.Length/1MB,2)), $_.FullName }
  Export-Results "Largest-Files" @($lines)
  Write-Ui "`nTop 20 largest files:" 'Cyan'
  foreach ($l in $lines) { Write-Ui "  $l" }
}

function Get-JunkFolders {
  @(
    @{ n = 'User Temp';            p = $env:TEMP },
    @{ n = 'Windows Temp';         p = "$env:SystemRoot\Temp" },
    @{ n = 'Prefetch';             p = "$env:SystemRoot\Prefetch" },
    @{ n = 'Windows Update Cache'; p = "$env:SystemRoot\SoftwareDistribution\Download" }
  ) | Where-Object { $_.p }
}

# Walks a folder without following links; links are removed themselves.
function Clear-FolderContents {
  param([string]$Path)
  $items = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
  foreach ($i in $items) {
    if (-not $i.LinkType -and $i.PSIsContainer) { Clear-FolderContents $i.FullName }
    [void](Remove-SafeItem -Path $i.FullName -AllowProtected)
  }
}

function Get-JunkEstimate {
  Write-Ui "Estimate - nothing is deleted" 'Cyan'
  $total = 0L
  foreach ($f in Get-JunkFolders) {
    $s = Get-FolderSize $f.p
    $total += $s
    Write-Ui ("  {0,-22} {1,10} MB" -f $f.n, [math]::Round($s/1MB,2))
  }
  Write-Ui "Approx reclaimable: $([math]::Round($total/1MB,2)) MB" 'Green'
}

function Clean-SafeJunk {
  Write-Ui "`n=== Safe Junk Cleaner ===" 'Cyan'
  if (-not (Confirm-Action "Delete temp files, Prefetch and the Windows Update download cache? Files in use are skipped.")) { return }
  $totalFreed = 0L
  foreach ($f in Get-JunkFolders) {
    if (-not (Test-Path -LiteralPath $f.p)) { continue }
    $before = Get-FolderSize $f.p
    Write-Ui "Cleaning: $($f.n)" 'Cyan'
    Clear-FolderContents $f.p
    $after = Get-FolderSize $f.p
    $freed = [math]::Max(0, $before - $after)
    $totalFreed += $freed
    Write-Ui "  Freed: $([math]::Round($freed/1MB,2)) MB" 'Green'
  }
  Write-Ui "[OK] Total freed: $([math]::Round($totalFreed/1MB,2)) MB" 'Green'
}

function Clear-Bin {
  if (Confirm-Action "Empty the Recycle Bin? This cannot be undone.") {
    try {
      Clear-RecycleBin -Force -ErrorAction Stop
      Write-Ui "[OK] Recycle Bin emptied" 'Green'
    } catch {
      Write-Ui "[!]  Could not empty Recycle Bin (it may already be empty): $($_.Exception.Message)" 'Yellow'
    }
  }
}

# =========================================================
# System views
# =========================================================
function Show-DiskSpace {
  $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop |
    Select-Object DeviceID, VolumeName,
      @{n='Free(GB)'; e={[math]::Round($_.FreeSpace/1GB,1)}},
      @{n='Size(GB)'; e={[math]::Round($_.Size/1GB,1)}},
      @{n='Free%';    e={ if ($_.Size) { [math]::Round(100*$_.FreeSpace/$_.Size,0) } else { 0 } }}
  Write-Table $disks
}

function Show-TopProcesses {
  $p = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 15 Name, Id,
       @{n='RAM(MB)'; e={[math]::Round($_.WorkingSet64/1MB,1)}}
  Write-Table $p
}

function Show-StartupPrograms {
  Write-Ui "Startup entries:" 'Cyan'
  $runKeys = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
  )
  foreach ($rk in $runKeys) {
    if (Test-Path $rk) {
      Write-Ui "`n[$rk]" 'Yellow'
      $props = Get-ItemProperty $rk -ErrorAction SilentlyContinue
      if ($props) {
        $props.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } |
          ForEach-Object { Write-Ui "  $($_.Name) = $($_.Value)" }
      }
    }
  }
  foreach ($sf in @([Environment]::GetFolderPath('Startup'), [Environment]::GetFolderPath('CommonStartup'))) {
    if ($sf -and (Test-Path $sf)) {
      Write-Ui "`n[$sf]" 'Yellow'
      Get-ChildItem $sf -Force -ErrorAction SilentlyContinue | ForEach-Object { Write-Ui "  $($_.Name)" }
    }
  }
}

function Get-InstalledPrograms {
  $keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
          'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
          'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
  Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | Sort-Object DisplayName
}

function Show-InstalledPrograms {
  Write-Table (Get-InstalledPrograms | Select-Object DisplayName, DisplayVersion, Publisher)
}

function Uninstall-Program {
  $apps = @(Get-InstalledPrograms | Where-Object { $_.UninstallString })
  if ($apps.Count -eq 0) { Write-Ui "No uninstallable programs found." 'Yellow'; return }
  $names = $apps | ForEach-Object { if ($_.DisplayVersion) { "$($_.DisplayName)  ($($_.DisplayVersion))" } else { $_.DisplayName } }
  $idx = Select-FromList "Select a program to uninstall" @($names)
  if ($idx -lt 0) { Write-Ui "Cancelled." 'DarkGray'; return }
  $prog = $apps[$idx]
  if (Confirm-Action "Launch the uninstaller for: $($prog.DisplayName)?") {
    Write-Ui "Launching uninstaller for $($prog.DisplayName)..." 'Cyan'
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $prog.UninstallString -Wait
    Write-Ui "Uninstaller finished." 'Green'
  }
}

function Start-DiskCleanup {
  Start-Process cleanmgr.exe -ErrorAction Stop
}

# =========================================================
# GUI
# =========================================================
function Invoke-GuiAction {
  param([string]$Name, [scriptblock]$Action)
  if ($script:Busy) {
    Write-Ui "Another operation is still running. Please wait." 'Yellow'
    return
  }
  $script:Busy = $true
  $script:Sidebar.Enabled = $false
  $script:Progress.Visible = $true
  $script:StatusLabel.Text = "Status: Running - $Name"
  $script:StatusLabel.ForeColor = [System.Drawing.Color]::Gold
  Write-Ui "`n>> $Name" 'Cyan'
  [System.Windows.Forms.Application]::DoEvents()
  try {
    Invoke-SafeAction -Name $Name -Action $Action
  } finally {
    $script:Progress.Visible = $false
    $script:Sidebar.Enabled = $true
    $script:StatusLabel.Text = "Status: Ready"
    $script:StatusLabel.ForeColor = [System.Drawing.Color]::LightGreen
    $script:Busy = $false
    [System.Windows.Forms.Application]::DoEvents()
  }
}

function Start-GuiMode {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()
  $script:UiMode = 'gui'
  $script:ExitTo = $null

  $bg    = [System.Drawing.Color]::FromArgb(28,28,30)
  $panel = [System.Drawing.Color]::FromArgb(37,37,40)
  $dark  = [System.Drawing.Color]::FromArgb(20,20,22)
  $btnBg = [System.Drawing.Color]::FromArgb(45,45,48)

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $DisplayTitle
  $form.StartPosition = 'CenterScreen'
  $form.Size = New-Object System.Drawing.Size(1120, 740)
  $form.MinimumSize = New-Object System.Drawing.Size(900, 600)
  $form.BackColor = $bg
  $form.ForeColor = [System.Drawing.Color]::White
  $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $script:MainForm = $form
  $script:Busy = $false
  $form.Add_FormClosing({
    param($sender, $e)
    if ($script:Busy) {
      $e.Cancel = $true
      try {
        [void][System.Windows.Forms.MessageBox]::Show(
          "An operation is still running. Please wait for it to finish before closing.",
          $DisplayTitle, [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Information)
      } catch { }
    }
  })

  # Fill control first so Top/Left docks carve space out of it
  $main = New-Object System.Windows.Forms.Panel
  $main.Dock = 'Fill'
  $main.Padding = New-Object System.Windows.Forms.Padding(14)
  $form.Controls.Add($main)

  $logBox = New-Object System.Windows.Forms.TextBox
  $logBox.Multiline = $true
  $logBox.ReadOnly = $true
  $logBox.ScrollBars = 'Vertical'
  $logBox.Dock = 'Fill'
  $logBox.BackColor = [System.Drawing.Color]::FromArgb(18,18,20)
  $logBox.ForeColor = [System.Drawing.Color]::Gainsboro
  $logBox.BorderStyle = 'FixedSingle'
  $logBox.Font = New-Object System.Drawing.Font("Consolas", 10)
  $main.Controls.Add($logBox)
  $script:LogBox = $logBox

  $header = New-Object System.Windows.Forms.Panel
  $header.Dock = 'Top'
  $header.Height = 84
  $header.BackColor = $dark
  $form.Controls.Add($header)

  $title = New-Object System.Windows.Forms.Label
  $title.Text = "Hardened Toolkit"
  $title.AutoSize = $true
  $title.Location = New-Object System.Drawing.Point(20, 12)
  $title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
  $header.Controls.Add($title)

  $subtitle = New-Object System.Windows.Forms.Label
  $subtitle.Text = "Cleaner | Duplicate Finder | Scanner | Uninstaller   (v$Version)"
  $subtitle.AutoSize = $true
  $subtitle.Location = New-Object System.Drawing.Point(22, 48)
  $subtitle.ForeColor = [System.Drawing.Color]::Gainsboro
  $header.Controls.Add($subtitle)

  $status = New-Object System.Windows.Forms.Label
  $status.Text = "Status: Ready"
  $status.AutoSize = $false
  $status.Size = New-Object System.Drawing.Size(360, 24)
  $status.Anchor = 'Top,Right'
  $status.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 380), 18)
  $status.TextAlign = 'MiddleRight'
  $status.ForeColor = [System.Drawing.Color]::LightGreen
  $header.Controls.Add($status)
  $script:StatusLabel = $status

  $progress = New-Object System.Windows.Forms.ProgressBar
  $progress.Style = 'Marquee'
  $progress.MarqueeAnimationSpeed = 25
  $progress.Visible = $false
  $progress.Size = New-Object System.Drawing.Size(360, 12)
  $progress.Anchor = 'Top,Right'
  $progress.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 380), 50)
  $header.Controls.Add($progress)
  $script:Progress = $progress

  $sidebar = New-Object System.Windows.Forms.FlowLayoutPanel
  $sidebar.Dock = 'Left'
  $sidebar.Width = 270
  $sidebar.BackColor = $panel
  $sidebar.FlowDirection = 'TopDown'
  $sidebar.WrapContents = $false
  $sidebar.AutoScroll = $true
  $sidebar.Padding = New-Object System.Windows.Forms.Padding(14, 10, 8, 10)
  $form.Controls.Add($sidebar)
  $script:Sidebar = $sidebar

  function Add-SectionLabel([string]$text) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text.ToUpper()
    $l.AutoSize = $false
    $l.Size = New-Object System.Drawing.Size(230, 26)
    $l.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 2)
    $l.ForeColor = [System.Drawing.Color]::FromArgb(140,150,175)
    $l.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $l.TextAlign = 'BottomLeft'
    $sidebar.Controls.Add($l)
  }

  function Add-ToolButton([string]$text, [string]$actionName, [scriptblock]$action) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Size = New-Object System.Drawing.Size(230, 36)
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
    $b.BackColor = $btnBg
    $b.ForeColor = [System.Drawing.Color]::White
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70,70,74)
    $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(60,64,84)
    $b.TextAlign = 'MiddleLeft'
    $b.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $n = $actionName; $a = $action
    $b.Add_Click({ Invoke-GuiAction -Name $n -Action $a }.GetNewClosure())
    $sidebar.Controls.Add($b)
  }

  Add-SectionLabel "Overview"
  Add-ToolButton "Disk space"                "Disk space overview"        { Show-DiskSpace }
  Add-ToolButton "Estimate junk"             "Estimate reclaimable junk"  { Get-JunkEstimate }
  Add-ToolButton "Top memory processes"      "Top memory processes"       { Show-TopProcesses }

  Add-SectionLabel "Clean"
  Add-ToolButton "Clean safe junk"           "Clean safe junk"            { Clean-SafeJunk }
  Add-ToolButton "Empty Recycle Bin"         "Empty Recycle Bin"          { Clear-Bin }
  Add-ToolButton "Windows Disk Cleanup"      "Windows Disk Cleanup"       { Start-DiskCleanup }

  Add-SectionLabel "Find"
  Add-ToolButton "Duplicate files"           "Find duplicate files"       { Find-Duplicates }
  Add-ToolButton "Empty files"               "Find empty files"           { Find-EmptyFiles }
  Add-ToolButton "Empty folders"             "Find empty folders"         { Find-EmptyFolders }
  Add-ToolButton "Largest 20 files"          "Find largest files"         { Find-LargestFiles }

  Add-SectionLabel "Programs"
  Add-ToolButton "Startup programs"          "List startup programs"      { Show-StartupPrograms }
  Add-ToolButton "Installed programs"        "List installed programs"    { Show-InstalledPrograms }
  Add-ToolButton "Uninstall a program"       "Uninstall a program"        { Uninstall-Program }

  Add-SectionLabel "Toolkit"
  Add-ToolButton "Install + shortcuts"       "Install toolkit"            { Install-Toolkit }
  Add-ToolButton "Uninstall toolkit"         "Uninstall toolkit"          { Uninstall-Toolkit }

  Add-SectionLabel "Session"
  Add-ToolButton "Switch to console"         "Switch to console"          { $script:ExitTo = 'console'; $form.Close() }
  Add-ToolButton "Clear log"                 "Clear log"                  { $script:LogBox.Clear() }
  Add-ToolButton "Exit"                      "Exit"                       { $form.Close() }

  $form.Add_Shown({
    Write-Ui "Welcome to $DisplayTitle. Scan results are saved to your Desktop; deletions are logged to:"
    Write-Ui "  $script:DeleteLog"
    Write-Ui "Deleted files bypass the Recycle Bin. Every delete asks for confirmation first.`r`n"
    if ($SmokeTest) {
      $script:SmokeTimer = New-Object System.Windows.Forms.Timer
      $script:SmokeTimer.Interval = 1500
      $script:SmokeTimer.Add_Tick({
        $script:SmokeTimer.Stop()
        if ($script:MainForm) { $script:MainForm.Close() }
      })
      $script:SmokeTimer.Start()
      Write-Ui "Smoke test: GUI opened; closing automatically." 'Yellow'
    }
  })

  [void]$form.ShowDialog()
  $script:LogBox = $null
  $script:MainForm = $null
  $script:SmokeTimer = $null
  $script:UiMode = 'console'
  return $script:ExitTo
}

# =========================================================
# Console
# =========================================================
function Start-ConsoleMode {
  $script:UiMode = 'console'
  while ($true) {
    Clear-Host
    Write-Host "`n$DisplayTitle" -ForegroundColor Cyan
    Write-Host "========================================"
    Write-Host " 1. Show disk space overview"
    Write-Host " 2. Estimate reclaimable junk"
    Write-Host " 3. Clean safe junk (temp, cache, prefetch)"
    Write-Host " 4. Empty Recycle Bin"
    Write-Host " 5. Find duplicate files"
    Write-Host " 6. Find empty (0-byte) files"
    Write-Host " 7. Find empty folders"
    Write-Host " 8. Find largest 20 files"
    Write-Host " 9. Top memory processes"
    Write-Host "10. List startup programs"
    Write-Host "11. List installed programs"
    Write-Host "12. Uninstall a program"
    Write-Host "13. Launch Windows Disk Cleanup"
    Write-Host "14. Install & create Desktop shortcut"
    Write-Host "15. Uninstall this toolkit"
    Write-Host " 0. Exit"
    Write-Host "========================================"

    $sel = Read-Host "Enter option number"
    switch ($sel) {
      '0'  { Write-Host "Done - goodbye!" -ForegroundColor Green; return }
      '1'  { Invoke-SafeAction "Disk space overview"       { Show-DiskSpace } }
      '2'  { Invoke-SafeAction "Estimate reclaimable junk" { Get-JunkEstimate } }
      '3'  { Invoke-SafeAction "Clean safe junk"           { Clean-SafeJunk } }
      '4'  { Invoke-SafeAction "Empty Recycle Bin"         { Clear-Bin } }
      '5'  { Invoke-SafeAction "Find duplicate files"      { Find-Duplicates } }
      '6'  { Invoke-SafeAction "Find empty files"          { Find-EmptyFiles } }
      '7'  { Invoke-SafeAction "Find empty folders"        { Find-EmptyFolders } }
      '8'  { Invoke-SafeAction "Find largest files"        { Find-LargestFiles } }
      '9'  { Invoke-SafeAction "Top memory processes"      { Show-TopProcesses } }
      '10' { Invoke-SafeAction "List startup programs"     { Show-StartupPrograms } }
      '11' { Invoke-SafeAction "List installed programs"   { Show-InstalledPrograms } }
      '12' { Invoke-SafeAction "Uninstall a program"       { Uninstall-Program } }
      '13' { Invoke-SafeAction "Windows Disk Cleanup"      { Start-DiskCleanup } }
      '14' { Invoke-SafeAction "Install toolkit"           { Install-Toolkit } }
      '15' { Invoke-SafeAction "Uninstall toolkit"         { Uninstall-Toolkit } }
      default { Write-Host "Invalid option" -ForegroundColor Red }
    }
    Read-Host "`nPress Enter to return to menu" | Out-Null
  }
}

# =========================================================
# Entry point
# =========================================================
if ($MyInvocation.InvocationName -ne '.') {
try {
  if ($SmokeTest) { $Console = $false }
  Require-Admin

  if ($Uninstall) {
    Uninstall-Toolkit
    if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }
    exit 0
  }

  if ($Console) {
    Start-ConsoleMode
  } else {
    $next = $null
    try {
      $next = Start-GuiMode
    } catch {
      $script:UiMode = 'console'
      Write-ErrorLog -Name "GUI" -ErrorRecord $_
      if ($Hidden) {
        try {
          Add-Type -AssemblyName System.Windows.Forms
          [void][System.Windows.Forms.MessageBox]::Show(
            "The graphical interface failed to start.`r`n`r`n$($_.Exception.Message)`r`n`r`nSee the log in $script:LogDir",
            $DisplayTitle, [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        } catch { }
        exit 1
      }
      Write-Host "GUI failed ($($_.Exception.Message)) - falling back to console mode." -ForegroundColor Yellow
      Write-Host "See the log in $script:LogDir" -ForegroundColor Yellow
      Start-Sleep -Seconds 2
      $next = 'console'
    }
    if ($next -eq 'console') { Start-ConsoleMode }
  }
}
catch {
  Write-ErrorLog -Name "Fatal" -ErrorRecord $_
  if ($Hidden) {
    try {
      Add-Type -AssemblyName System.Windows.Forms
      [void][System.Windows.Forms.MessageBox]::Show(
        "A fatal error occurred.`r`n`r`n$($_.Exception.Message)`r`n`r`nSee the log in $script:LogDir",
        $DisplayTitle, [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    } catch { }
  } else {
    Write-Host "`nFatal error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "See the log in $script:LogDir" -ForegroundColor Red
    Read-Host "Press Enter to exit" | Out-Null
  }
  exit 1
}
}