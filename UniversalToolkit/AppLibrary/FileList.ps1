#Requires -Version 5.1
<#
.SYNOPSIS
    Directory File & Folder Name Lister
.DESCRIPTION
    Scans a directory and saves all file/folder names to a .TXT or .RTF file.
    Features:
    - GUI folder browse dialogs
    - Optional recursive subfolder search
    - Files grouped by extension (.EXE, .TXT, .PDF, etc.) with counts
    - Default save location = Desktop
    - Auto-opens output file when done
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
#  Helper: Show folder browser dialog
# ============================================================
function Show-FolderBrowser {
    param(
        [string]$Description = "Select a folder",
        [bool]$ShowNewFolderButton = $true
    )
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $Description
    $dlg.ShowNewFolderButton = $ShowNewFolderButton
    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return $null
}

# ============================================================
#  Helper: Escape string for RTF
# ============================================================
function ConvertTo-RtfSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text -replace '\\', '\\'
    $Text = $Text -replace '{', '\{'
    $Text = $Text -replace '}', '\}'
    return $Text
}

# ============================================================
#  MAIN SCRIPT
# ============================================================

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "     DIRECTORY FILE & FOLDER NAME LISTER  (PowerShell)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# --- Get Desktop path ---
$desktopDir = [Environment]::GetFolderPath("Desktop")
if (-not $desktopDir) { $desktopDir = Join-Path $env:USERPROFILE "Desktop" }

# ============================================================
#  STEP 1: Browse for SOURCE directory
# ============================================================
Write-Host "Step 1 of 4: Select the SOURCE folder to scan..." -ForegroundColor Yellow
Write-Host ""

$sourceDir = Show-FolderBrowser -Description "Select the folder to scan for files and folders"

if (-not $sourceDir) {
    Write-Host "ERROR: No folder was selected." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Source folder: " -NoNewline
Write-Host $sourceDir -ForegroundColor Green
Write-Host ""

# ============================================================
#  STEP 2: Subfolder search option
# ============================================================
Write-Host "Step 2 of 4: Include subfolders?" -ForegroundColor Yellow
Write-Host "  [1] Top-level only (default) - scan just this folder"
Write-Host "  [2] Recursive - search ALL subfolders too"
Write-Host ""

do {
    $subChoice = Read-Host "Enter your choice [1/2] (press Enter = top-level only)"
    if ([string]::IsNullOrWhiteSpace($subChoice)) { $subChoice = "1" }
} while ($subChoice -notin @("1", "2"))

$recursive = ($subChoice -eq "2")
$modeLabel = if ($recursive) { "Recursive (includes all subfolders)" } else { "Top-level only" }

Write-Host "Mode: " -NoNewline
Write-Host $modeLabel -ForegroundColor Green
Write-Host ""

# ============================================================
#  STEP 3: Choose OUTPUT directory
# ============================================================
Write-Host "Step 3 of 4: Choose where to save the output file:" -ForegroundColor Yellow
Write-Host "  [1] Save to DESKTOP (default)  -  $desktopDir"
Write-Host "  [2] Save in the SOURCE folder   -  $sourceDir"
Write-Host "  [3] Browse for a different folder"
Write-Host ""

do {
    $outChoice = Read-Host "Enter your choice [1/2/3] (press Enter = Desktop)"
    if ([string]::IsNullOrWhiteSpace($outChoice)) { $outChoice = "1" }
} while ($outChoice -notin @("1", "2", "3"))

$outputDir = switch ($outChoice) {
    "1" { $desktopDir }
    "2" { $sourceDir }
    "3" {
        Write-Host ""
        $browsed = Show-FolderBrowser -Description "Select where to save the output file"
        if (-not $browsed) {
            Write-Host "ERROR: No output folder was selected." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        $browsed
    }
}

Write-Host "Output folder: " -NoNewline
Write-Host $outputDir -ForegroundColor Green
Write-Host ""

# ============================================================
#  STEP 4: Choose output format
# ============================================================
Write-Host "Step 4 of 4: Choose output format:" -ForegroundColor Yellow
Write-Host "  [1] Plain Text (.TXT)"
Write-Host "  [2] Rich Text Format (.RTF)"
Write-Host ""

do {
    $fmtChoice = Read-Host "Enter your choice (1 or 2)"
} while ($fmtChoice -notin @("1", "2"))

$outputExt = if ($fmtChoice -eq "1") { "txt" } else { "rtf" }

# ============================================================
#  Set output filename with timestamp
# ============================================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $outputDir "FileList_$timestamp.$outputExt"

# ============================================================
#  COLLECT DATA
# ============================================================
Write-Host "Scanning... " -NoNewline

$getFolderParams = @{
    Path    = $sourceDir
    Directory = $true
}
$getFileParams = @{
    Path    = $sourceDir
    File    = $true
}
if ($recursive) {
    $getFolderParams["Recurse"] = $true
    $getFileParams["Recurse"] = $true
}

# Get folders
$folders = Get-ChildItem @getFolderParams | Sort-Object FullName
$folderCount = $folders.Count

# Get files
$files = Get-ChildItem @getFileParams | Sort-Object FullName
$fileCount = $files.Count

# Group files by extension
$filesByExt = $files | Group-Object {
    $ext = $_.Extension
    if ([string]::IsNullOrEmpty($ext)) { "(NONE)" } else { $ext }
} | Sort-Object Name

Write-Host "Done" -ForegroundColor Green
Write-Host "  Folders: $folderCount"
Write-Host "  Files:   $fileCount"
Write-Host ""

# ============================================================
#  Helper: Get relative path
# ============================================================
function Get-RelativePath {
    param([string]$FullPath, [string]$BasePath)
    $baseUri = New-Object System.Uri($BasePath.TrimEnd('\') + '\')
    $fullUri = New-Object System.Uri($FullPath)
    $rel = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($fullUri).ToString())
    return $rel.Replace('/', '\')
}

# ============================================================
#  WRITE OUTPUT
# ============================================================
Write-Host "Generating $outputExt file... " -NoNewline

if ($outputExt -eq "txt") {
    # --- TXT OUTPUT ---
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("============================================================")
    [void]$sb.AppendLine(" DIRECTORY LISTING")
    [void]$sb.AppendLine("============================================================")
    [void]$sb.AppendLine(" Source:    $sourceDir")
    [void]$sb.AppendLine(" Mode:      $modeLabel")
    [void]$sb.AppendLine(" Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine(" Folders:   $folderCount")
    [void]$sb.AppendLine(" Files:     $fileCount")
    [void]$sb.AppendLine("============================================================")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("------------------------------------------------------------")
    [void]$sb.AppendLine(" FOLDERS ($folderCount)")
    [void]$sb.AppendLine("------------------------------------------------------------")

    if ($folderCount -gt 0) {
        foreach ($f in $folders) {
            $display = if ($recursive) { Get-RelativePath $f.FullName $sourceDir } else { $f.Name }
            [void]$sb.AppendLine($display)
        }
    } else {
        [void]$sb.AppendLine("(none)")
    }

    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("------------------------------------------------------------")
    [void]$sb.AppendLine(" FILES BY EXTENSION ($fileCount total)")
    [void]$sb.AppendLine("------------------------------------------------------------")

    if ($fileCount -gt 0) {
        foreach ($group in $filesByExt) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("[$($group.Name) FILES]  ($($group.Count))")
            [void]$sb.AppendLine("----------------------------------------")
            foreach ($file in $group.Group | Sort-Object FullName) {
                $display = if ($recursive) { Get-RelativePath $file.FullName $sourceDir } else { $file.Name }
                [void]$sb.AppendLine($display)
            }
        }
    } else {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("(no files found)")
    }

    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("============================================================")
    [void]$sb.AppendLine(" END OF LIST")
    [void]$sb.AppendLine("============================================================")

    $sb.ToString() | Out-File -FilePath $outputFile -Encoding UTF8

} else {
    # --- RTF OUTPUT ---
    $rtfSource = ConvertTo-RtfSafe $sourceDir
    $rtfMode = ConvertTo-RtfSafe $modeLabel
    $rtfDate = ConvertTo-RtfSafe (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

    $rtfLines = @()
    $rtfLines += '{\rtf1\ansi\deff0'
    $rtfLines += '{\fonttbl{\f0 Arial;}{\f1 Courier New;}}'
    $rtfLines += '{\colortbl ;\red0\green0\blue128;\red128\green0\blue0;\red0\green100\blue0;\red128\green0\blue128;}'
    $rtfLines += '\f0\fs24'
    $rtfLines += '\b\fs32 DIRECTORY LISTING\b0\fs24\par'
    $rtfLines += '\par'
    $rtfLines += "\b Source:\b0  $rtfSource\par"
    $rtfLines += "\b Mode:\b0  $rtfMode\par"
    $rtfLines += "\b Generated:\b0  $rtfDate\par"
    $rtfLines += "\b Folders:\b0  $folderCount\par"
    $rtfLines += "\b Files:\b0  $fileCount\par"
    $rtfLines += '\par'
    $rtfLines += '\b\fs28\cf1 FOLDERS \b0\i (' + $folderCount + ')\i0\fs24\par'
    $rtfLines += '\f1'

    if ($folderCount -gt 0) {
        foreach ($f in $folders) {
            $display = if ($recursive) { Get-RelativePath $f.FullName $sourceDir } else { $f.Name }
            $rtfLines += (ConvertTo-RtfSafe $display) + '\par'
        }
    } else {
        $rtfLines += '\i (none)\i0\par'
    }

    $rtfLines += '\f0\par'
    $rtfLines += '\b\fs28\cf2 FILES BY EXTENSION \b0\i (' + $fileCount + ' total)\i0\fs24\par'

    if ($fileCount -gt 0) {
        foreach ($group in $filesByExt) {
            $extName = ConvertTo-RtfSafe $group.Name
            $rtfLines += '\f0\par'
            $rtfLines += '\b\fs26\cf3 ' + $extName + ' FILES \b0\i (' + $group.Count + ')\i0\fs24\par'
            $rtfLines += '\f1'
            foreach ($file in $group.Group | Sort-Object FullName) {
                $display = if ($recursive) { Get-RelativePath $file.FullName $sourceDir } else { $file.Name }
                $rtfLines += (ConvertTo-RtfSafe $display) + '\par'
            }
        }
    } else {
        $rtfLines += '\f1\par\i (no files found)\i0\par'
    }

    $rtfLines += '\f0\par'
    $rtfLines += '\cf4\i --- End of list ---\i0\cf0\par'
    $rtfLines += '}'

    $rtfLines | Out-File -FilePath $outputFile -Encoding ASCII
}

Write-Host "Done" -ForegroundColor Green

# ============================================================
#  DONE
# ============================================================
Clear-Host
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " SUCCESS!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Scanned:     $sourceDir"
Write-Host " Mode:        $modeLabel"
Write-Host " Folders:     $folderCount"
Write-Host " Files:       $fileCount"
Write-Host " Saved to:    $outputFile"
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to open the file"

# Open the file
Invoke-Item $outputFile
