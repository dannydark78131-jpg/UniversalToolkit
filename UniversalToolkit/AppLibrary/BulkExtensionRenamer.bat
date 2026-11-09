@echo off
setlocal

:: Run the PowerShell block embedded below this batch section.
:: The script re-reads itself and executes everything after the marker line,
:: which safely handles spaces / parentheses / apostrophes in this file's path.
set "SCRIPT_PATH=%~f0"

cls
echo ==================================================
echo       BULK FILE EXTENSION RENAMER (TESTED)
echo ==================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Sta -Command "$m='#---PS'+'-START---#'; $code = [System.IO.File]::ReadAllText($env:SCRIPT_PATH); $code = $code.Substring($code.IndexOf($m)); Invoke-Command -ScriptBlock ([scriptblock]::Create($code))"

echo.
echo Press any key to exit...
pause >nul
exit /b

#---PS-START---#
Add-Type -AssemblyName System.Windows.Forms

Write-Host "--- Configuration ---" -ForegroundColor Cyan
$oldExt = Read-Host "Enter current extension to target (e.g., txt)"
$newExt = Read-Host "Enter new extension to apply (e.g., md)"

# Clean up input (remove leading periods if the user typed them)
$oldExt = $oldExt.Trim().TrimStart('.')
$newExt = $newExt.Trim().TrimStart('.')

if ([string]::IsNullOrEmpty($oldExt) -or [string]::IsNullOrEmpty($newExt)) {
    Write-Host "`n[ERROR] Both extensions must be specified." -ForegroundColor Red
    return
}

# Open the Windows folder picker dialog
Write-Host "`nOpening folder picker..." -ForegroundColor Yellow
$browser = New-Object System.Windows.Forms.FolderBrowserDialog
$browser.Description = "Select the folder containing files to rename"
$browser.ShowNewFolderButton = $false

$result = $browser.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "`n[CANCELLED] No folder selected." -ForegroundColor Yellow
    return
}

$targetFolder = $browser.SelectedPath
Write-Host "Target Folder: $targetFolder" -ForegroundColor Green

# Process files and generate a log
$logPath = Join-Path -Path $targetFolder -ChildPath "renamer_log.txt"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

$logContent = @"
==================================================
BULK RENAME LOG
==================================================
Date/Time : $timestamp
Folder    : $targetFolder
Rule      : *.$oldExt --> *.$newExt
--------------------------------------------------
RESULTS:
"@

Write-Host "`nScanning for *.$oldExt files..." -ForegroundColor Cyan
$files = @(Get-ChildItem -Path $targetFolder -Filter "*.$oldExt" -File)

if ($files.Count -eq 0) {
    Write-Host "No files found matching *.$oldExt in this directory." -ForegroundColor Yellow
    $logContent += "`r`nNo files found matching criteria."
} else {
    $successCount = 0
    $failCount = 0

    foreach ($file in $files) {
        $newName = [System.IO.Path]::ChangeExtension($file.Name, $newExt)
        $oldPath = $file.FullName

        try {
            Rename-Item -LiteralPath $oldPath -NewName $newName -ErrorAction Stop
            Write-Host " [OK] $($file.Name) -> $newName" -ForegroundColor Green
            $logContent += "`r`n[SUCCESS] $($file.Name) -> $newName"
            $successCount++
        } catch {
            Write-Host " [ERROR] $($file.Name) - $_" -ForegroundColor Red
            $logContent += "`r`n[FAILED]  $($file.Name) - $_"
            $failCount++
        }
    }
    Write-Host "`nFinished! Successfully renamed: $successCount | Failed: $failCount" -ForegroundColor Cyan
}

$logContent += "`r`n=================================================="
$logContent | Out-File -FilePath $logPath -Encoding utf8

Write-Host "`nLog saved to: renamer_log.txt" -ForegroundColor Gray