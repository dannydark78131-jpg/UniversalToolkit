#Requires -Version 5.1
# Universal Toolkit Launcher
# One hub screen for a library of PowerShell / batch tools: launch, elevate
# where needed, install desktop + Start Menu shortcuts, and uninstall.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$script:AppName    = 'Universal Toolkit'
$script:AppVersion = '1.1'

# ---------- paths ----------
$script:BaseDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($script:BaseDir)) {
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        $script:BaseDir = Split-Path -Parent $proc.MainModule.FileName
    } catch {
        $script:BaseDir = (Get-Location).Path
    }
}
$script:LibDir      = Join-Path $script:BaseDir 'AppLibrary'
$script:IconPath    = Join-Path $script:BaseDir 'AppIcon.ico'
$script:InstallDir  = Join-Path $env:LOCALAPPDATA 'Programs\UniversalToolkit'
$script:ShortcutName = 'Universal Toolkit.lnk'
$script:IsInstalled = $false

# ---------- app catalog ----------
# Kind: Ps1 = launch a .ps1 with powershell ; Bat = launch a .bat via cmd
# Admin: $true -> request elevation (RunAs) when launching
$script:Catalog = @(
    [PSCustomObject]@{
        Id='HardenedToolkit'; Name='Hardened Toolkit'; Version='2.4.0'
        File='HardenedToolkit.ps1'; Kind='Ps1'; Admin=$true; Icon=$null
        Quote='Cleaner, duplicate finder, scanner and uninstaller in one.'
        Tags='Cleanup, Duplicates, System'
    }
    [PSCustomObject]@{
        Id='ComponentOptimizer'; Name='Component Optimizer'; Version='2.0'
        File='ComponentDeduplicator.ps1'; Kind='Ps1'; Admin=$false; Icon='ComponentDeduplicator.ico'
        Quote='Optimize, deduplicate and relocate application components.'
        Tags='Optimize, Duplicates, WinSxS'
    }
    [PSCustomObject]@{
        Id='DeviceCleanup'; Name='Device Cleanup'; Version=''
        File='DeviceCleanup.ps1'; Kind='Ps1'; Admin=$true; Icon=$null
        Quote='Removes ghost and unused devices from Windows.'
        Tags='Devices, System'
    }
    [PSCustomObject]@{
        Id='ExecutablePerformanceManager'; Name='Executable Performance Manager'; Version=''
        File='ExecutablePerformanceManager.ps1'; Kind='Ps1'; Admin=$true; Icon=$null
        Quote='Tunes per-app priority, QoS, GPU and compatibility settings.'
        Tags='Performance, Tuning'
    }
    [PSCustomObject]@{
        Id='DGsFileFinder'; Name='DGs File Finder'; Version='2.0.0'
        File='DGsFileSearch.ps1'; Kind='Ps1'; Admin=$false; Icon=$null
        Quote='Search ZIP archives and folders, filter, and extract files.'
        Tags='Search, ZIP, Extract'
    }
    [PSCustomObject]@{
        Id='FileList'; Name='File and Folder Lister'; Version=''
        File='FileList.bat'; Kind='Bat'; Admin=$false; Icon=$null
        Quote='Lists every file and folder in a directory to TXT or RTF.'
        Tags='Listing, Report'
    }
    [PSCustomObject]@{
        Id='BulkExtensionRenamer'; Name='Bulk Extension Renamer'; Version=''
        File='BulkExtensionRenamer.bat'; Kind='Bat'; Admin=$false; Icon=$null
        Quote='Renames files from one extension to another, with a log.'
        Tags='Rename, Bulk'
    }
)

# ---------- helpers ----------
function Get-DesktopPath {
    $p = [Environment]::GetFolderPath('Desktop')
    if ([string]::IsNullOrWhiteSpace($p)) {
        $shell = New-Object -ComObject WScript.Shell
        try { $p = $shell.SpecialFolders.Item('Desktop') }
        finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    }
    return $p
}
function Get-ProgramsPath { return [Environment]::GetFolderPath('Programs') }

function Get-ShortcutPaths {
    [PSCustomObject]@{
        Desktop   = Join-Path (Get-DesktopPath) $script:ShortcutName
        StartDir  = Join-Path (Get-ProgramsPath) $script:AppName
        StartMain = Join-Path (Get-ProgramsPath) (Join-Path $script:AppName $script:ShortcutName)
    }
}

function New-Shortcut {
    param(
        [string]$LinkPath,
        [string]$TargetPath,
        [string]$Arguments = '',
        [string]$WorkingDirectory = '',
        [string]$IconLocation = '',
        [string]$Description = ''
    )
    $parent = Split-Path -Parent $LinkPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $shell = New-Object -ComObject WScript.Shell
    try {
        $sc = $shell.CreateShortcut($LinkPath)
        $sc.TargetPath = $TargetPath
        if ($Arguments) { $sc.Arguments = $Arguments }
        if ($WorkingDirectory) { $sc.WorkingDirectory = $WorkingDirectory }
        if ($IconLocation -and (Test-Path -LiteralPath $IconLocation)) { $sc.IconLocation = "$IconLocation,0" }
        if ($Description) { $sc.Description = $Description }
        $sc.WindowStyle = 1
        $sc.Save()
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
}

function Resolve-AppPath {
    param($App)
    $p = Join-Path $script:LibDir $App.File
    if (Test-Path -LiteralPath $p) { return $p }
    $p2 = Join-Path $script:InstallDir (Join-Path 'AppLibrary' $App.File)
    if (Test-Path -LiteralPath $p2) { return $p2 }
    return $null
}

function Resolve-AppIcon {
    param($App)
    if ($App.Icon) {
        $p = Join-Path $script:LibDir $App.Icon
        if (Test-Path -LiteralPath $p) { return $p }
        $p = Join-Path $script:InstallDir (Join-Path 'AppLibrary' $App.Icon)
        if (Test-Path -LiteralPath $p) { return $p }
    }
    if (Test-Path -LiteralPath $script:IconPath) { return $script:IconPath }
    return $null
}

function Get-LaunchSpec {
    param($App)
    $path = Resolve-AppPath $App
    if (-not $path) { return $null }
    $wd = Split-Path -Parent $path
    if ($App.Kind -eq 'Bat') {
        return [PSCustomObject]@{ Target = 'cmd.exe'; Arguments = "/c `"$path`""; WorkingDirectory = $wd }
    }
    # -Sta is mandatory for the WinForms and WPF scripts in this library.
    return [PSCustomObject]@{
        Target = 'powershell.exe'
        Arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File `"$path`""
        WorkingDirectory = $wd
    }
}

function Start-App {
    param($App)
    $spec = Get-LaunchSpec $App
    if (-not $spec) {
        [System.Windows.Forms.MessageBox]::Show("Could not find '$($App.File)' in AppLibrary.", 'Missing app', 'OK', 'Error') | Out-Null
        return
    }
    $params = @{
        FilePath         = $spec.Target
        ArgumentList     = $spec.Arguments
        WorkingDirectory = $spec.WorkingDirectory
    }
    if ($App.Admin) {
        # Elevate at the launcher level and pass -Sta to the elevated child so
        # the app's own self-elevation is a no-op and WPF apps keep their STA thread.
        $params['Verb'] = 'RunAs'
        $params['ArgumentList'] = $spec.Arguments
    }
    try {
        $proc = Start-Process @params -PassThru
        Start-Sleep -Milliseconds 900
        if ($proc -and $proc.HasExited -and $proc.ExitCode -ne 0 -and -not $App.Admin) {
            Set-Status "$($App.Name) exited immediately (code $($proc.ExitCode))." ([System.Drawing.Color]::FromArgb(255, 248, 113, 113))
        } else {
            Set-Status "Launched $($App.Name)$(if ($App.Admin) { ' (elevated)' })" ([System.Drawing.Color]::FromArgb(255, 74, 222, 128))
        }
    } catch {
        if ($_.Exception.Message -match 'canceled|cancelled|1223|The operation was canceled') {
            Set-Status "Elevation cancelled for $($App.Name)." ([System.Drawing.Color]::FromArgb(255, 250, 204, 21))
        } else {
            Set-Status "Failed to launch $($App.Name): $($_.Exception.Message)" ([System.Drawing.Color]::FromArgb(255, 248, 113, 113))
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Launch failed", 'OK', 'Error') | Out-Null
        }
    }
}

function Set-Status {
    param([string]$Text, [System.Drawing.Color]$Color)
    if ($script:lblStatus) {
        $script:lblStatus.Text = $Text
        if ($Color) { $script:lblStatus.ForeColor = $Color }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------- install / uninstall ----------
function Copy-Payload {
    if ($script:BaseDir.TrimEnd('\') -ieq $script:InstallDir.TrimEnd('\')) { return }
    New-Item -ItemType Directory -Path $script:InstallDir -Force | Out-Null
    $libDest = Join-Path $script:InstallDir 'AppLibrary'
    New-Item -ItemType Directory -Path $libDest -Force | Out-Null
    foreach ($item in Get-ChildItem -LiteralPath $script:LibDir -File) {
        Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $libDest $item.Name) -Force
    }
    foreach ($name in @('UniversalLauncher.exe', 'UniversalLauncher.exe.config', 'UniversalLauncher.ps1', 'AppIcon.ico', 'README.txt')) {
        $src = Join-Path $script:BaseDir $name
        if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination (Join-Path $script:InstallDir $name) -Force }
    }
}

function Get-InstallerTarget {
    $exe = Join-Path $script:InstallDir 'UniversalLauncher.exe'
    if (Test-Path -LiteralPath $exe) { return [PSCustomObject]@{ Target = $exe; Arguments = '' } }
    $ps1 = Join-Path $script:InstallDir 'UniversalLauncher.ps1'
    if (Test-Path -LiteralPath $ps1) {
        return [PSCustomObject]@{ Target = 'powershell.exe'; Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ps1`"" }
    }
    return $null
}

function Update-InstallState {
    $script:IsInstalled = (Test-Path -LiteralPath (Join-Path $script:InstallDir 'UniversalLauncher.exe')) -or
                          (Test-Path -LiteralPath (Join-Path $script:InstallDir 'UniversalLauncher.ps1')) -or
                          ($script:BaseDir.TrimEnd('\') -ieq $script:InstallDir.TrimEnd('\'))
    if ($script:IsInstalled) {
        $script:btnInstall.Text = 'Repair / update install'
        Set-Status "Installed at $($script:InstallDir)" ([System.Drawing.Color]::FromArgb(255, 74, 222, 128))
    } else {
        $script:btnInstall.Text = 'Install shortcuts'
        Set-Status 'Portable mode. Install to add desktop and Start Menu shortcuts.' ([System.Drawing.Color]::FromArgb(255, 148, 163, 184))
    }
    $script:header.Invalidate()
}

function Invoke-Install {
    try {
        Set-Busy $true
        Set-Status 'Installing...' ([System.Drawing.Color]::FromArgb(255, 96, 165, 250))
        Copy-Payload
        $launcher = Get-InstallerTarget
        if (-not $launcher) { throw 'Could not find the launcher in the install folder.' }
        $icon = Join-Path $script:InstallDir 'AppIcon.ico'
        if (-not (Test-Path -LiteralPath $icon)) { $icon = $launcher.Target }
        $paths = Get-ShortcutPaths
        $created = New-Object System.Collections.ArrayList
        if ($script:chkDesktop.Checked) {
            New-Shortcut -LinkPath $paths.Desktop -TargetPath $launcher.Target -Arguments $launcher.Arguments `
                -WorkingDirectory $script:InstallDir -IconLocation $icon -Description 'Open the Universal Toolkit launcher'
            [void]$created.Add('Desktop shortcut')
        }
        if ($script:chkStartMenu.Checked) {
            New-Shortcut -LinkPath $paths.StartMain -TargetPath $launcher.Target -Arguments $launcher.Arguments `
                -WorkingDirectory $script:InstallDir -IconLocation $icon -Description 'Open the Universal Toolkit launcher'
            foreach ($app in $script:Catalog) {
                $spec = Get-LaunchSpec $app
                if (-not $spec) { continue }
                $appIcon = Resolve-AppIcon $app
                $lnk = Join-Path $paths.StartDir ("{0}.lnk" -f $app.Name)
                New-Shortcut -LinkPath $lnk -TargetPath $spec.Target -Arguments $spec.Arguments `
                    -WorkingDirectory $spec.WorkingDirectory -IconLocation $appIcon -Description $app.Quote
            }
            [void]$created.Add('Start Menu shortcuts')
        }
        Update-InstallState
        $extra = ''
        if ($created.Count -gt 0) { $extra = "`n`nCreated: " + ($created -join ', ') }
        [System.Windows.Forms.MessageBox]::Show(
            "Universal Toolkit is installed.`n`nLocation: $($script:InstallDir)$extra",
            'Install complete', 'OK', 'Information') | Out-Null
    } catch {
        Set-Status "Install failed: $($_.Exception.Message)" ([System.Drawing.Color]::FromArgb(255, 248, 113, 113))
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Install failed', 'OK', 'Error') | Out-Null
    } finally { Set-Busy $false }
}

function Invoke-Uninstall {
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Remove Universal Toolkit shortcuts and the installed copy?`n`nYour apps and files are never modified during rename/cleanup.",
        'Uninstall', 'YesNo', 'Warning')
    if ($answer -ne 'Yes') { return }
    try {
        Set-Busy $true
        $paths = Get-ShortcutPaths
        if (Test-Path -LiteralPath $paths.Desktop) { Remove-Item -LiteralPath $paths.Desktop -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $paths.StartDir) { Remove-Item -LiteralPath $paths.StartDir -Recurse -Force -ErrorAction SilentlyContinue }
        if ($script:BaseDir.TrimEnd('\') -ieq $script:InstallDir.TrimEnd('\')) {
            $safe = $script:InstallDir.Replace("'", "''")
            $cmd = "Start-Sleep -Seconds 2; Remove-Item -LiteralPath '$safe' -Recurse -Force -ErrorAction SilentlyContinue"
            Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-WindowStyle', 'Hidden', '-Command', $cmd) -WindowStyle Hidden
            [System.Windows.Forms.MessageBox]::Show('Shortcuts removed. The install folder will be deleted after this window closes.', 'Uninstalled', 'OK', 'Information') | Out-Null
            $form.Close(); return
        }
        if (Test-Path -LiteralPath $script:InstallDir) { Remove-Item -LiteralPath $script:InstallDir -Recurse -Force -ErrorAction SilentlyContinue }
        Update-InstallState
        Set-Status 'Shortcuts and installed copy removed.' ([System.Drawing.Color]::FromArgb(255, 74, 222, 128))
        [System.Windows.Forms.MessageBox]::Show('Universal Toolkit has been removed.', 'Uninstalled', 'OK', 'Information') | Out-Null
    } catch {
        Set-Status "Uninstall failed: $($_.Exception.Message)" ([System.Drawing.Color]::FromArgb(255, 248, 113, 113))
    } finally { Set-Busy $false }
}

# ---------- form scaffolding ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Universal Toolkit - Launcher'
$form.ClientSize = New-Object System.Drawing.Size(900, 620)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(820, 560)
$form.BackColor = [System.Drawing.Color]::FromArgb(255, 11, 18, 32)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.KeyPreview = $true

$script:formIcon = $null
$script:headerIconBmp = $null
if (Test-Path -LiteralPath $script:IconPath) {
    try {
        $script:formIcon = New-Object System.Drawing.Icon -ArgumentList $script:IconPath
        $form.Icon = $script:formIcon
        $hi = New-Object System.Drawing.Icon -ArgumentList $script:IconPath, 96, 96
        $script:headerIconBmp = $hi.ToBitmap()
        $hi.Dispose()
    } catch { }
}

$titleFont    = New-Object System.Drawing.Font('Segoe UI', 24, [System.Drawing.FontStyle]::Bold)
$subtitleFont = New-Object System.Drawing.Font('Segoe UI', 10.5)
$versionFont  = New-Object System.Drawing.Font('Segoe UI', 9)
$whiteBrush   = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::White)
$softBrush    = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(255, 224, 231, 255))
$verBrush     = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(255, 199, 210, 254))

$script:header = New-Object System.Windows.Forms.Panel
$script:header.Dock = 'Top'
$script:header.Height = 140
$script:header.BackColor = [System.Drawing.Color]::FromArgb(255, 14, 165, 233)
$script:header.Add_Paint({
        param($sender, $e)
        $rect = $sender.ClientRectangle
        if ($rect.Width -le 0 -or $rect.Height -le 0) { return }
        $grad = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            $rect,
            [System.Drawing.Color]::FromArgb(255, 14, 165, 233),
            [System.Drawing.Color]::FromArgb(255, 124, 58, 237),
            [System.Single]0)
        try { $e.Graphics.FillRectangle($grad, $rect) } finally { $grad.Dispose() }

        $shineRect = [System.Drawing.Rectangle]::new(0, 0, $rect.Width, [int]($rect.Height * 0.6))
        $shine = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            $shineRect,
            [System.Drawing.Color]::FromArgb(45, 255, 255, 255),
            [System.Drawing.Color]::FromArgb(0, 255, 255, 255),
            [System.Single]90)
        try { $e.Graphics.FillRectangle($shine, $shineRect) } finally { $shine.Dispose() }

        if ($script:headerIconBmp) {
            $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $e.Graphics.DrawImage($script:headerIconBmp, 26, 26, 88, 88)
        }
        $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
        $e.Graphics.DrawString($script:AppName, $titleFont, $whiteBrush, 130, 22)
        $e.Graphics.DrawString('Your toolkit, one window', $subtitleFont, $softBrush, 133, 72)
        $state = if ($script:IsInstalled) { 'Installed' } else { 'Portable' }
        $count = $script:Catalog.Count
        $e.Graphics.DrawString("v$($script:AppVersion)   |   $state   |   $count apps   |   runs locally", $versionFont, $verBrush, 135, 98)
        $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(90, 255, 255, 255), [System.Single]1)
        try { $e.Graphics.DrawLine($pen, 0, $rect.Height - 1, $rect.Width, $rect.Height - 1) } finally { $pen.Dispose() }
    })

$body = New-Object System.Windows.Forms.FlowLayoutPanel
$body.Dock = 'Fill'
$body.BackColor = [System.Drawing.Color]::FromArgb(255, 11, 18, 32)
$body.AutoScroll = $true
$body.Padding = New-Object System.Windows.Forms.Padding(14)
$body.WrapContents = $true
$body.FlowDirection = 'LeftToRight'

$footer = New-Object System.Windows.Forms.Panel
$footer.Dock = 'Bottom'
$footer.Height = 56
$footer.BackColor = [System.Drawing.Color]::FromArgb(255, 15, 23, 42)
$footer.Add_Paint({
        param($sender, $e)
        $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(255, 30, 41, 59), [System.Single]1)
        try { $e.Graphics.DrawLine($pen, 0, 0, $sender.Width, 0) } finally { $pen.Dispose() }
    })

$script:btnLaunchSelected = New-Object System.Windows.Forms.Button
$script:btnLaunchSelected.Text = 'Launch selected'
$script:btnLaunchSelected.Location = New-Object System.Drawing.Point(14, 12)
$script:btnLaunchSelected.Size = New-Object System.Drawing.Size(150, 32)
$script:btnLaunchSelected.FlatStyle = 'Flat'
$script:btnLaunchSelected.FlatAppearance.BorderSize = 0
$script:btnLaunchSelected.BackColor = [System.Drawing.Color]::FromArgb(255, 34, 197, 94)
$script:btnLaunchSelected.ForeColor = [System.Drawing.Color]::White
$script:btnLaunchSelected.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$script:btnLaunchSelected.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:btnLaunchSelected.UseVisualStyleBackColor = $false
$script:btnLaunchSelected.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 22, 163, 74)

$script:btnInstall = New-Object System.Windows.Forms.Button
$script:btnInstall.Text = 'Install shortcuts'
$script:btnInstall.Location = New-Object System.Drawing.Point(174, 12)
$script:btnInstall.Size = New-Object System.Drawing.Size(160, 32)
$script:btnInstall.FlatStyle = 'Flat'
$script:btnInstall.FlatAppearance.BorderSize = 0
$script:btnInstall.BackColor = [System.Drawing.Color]::FromArgb(255, 59, 130, 246)
$script:btnInstall.ForeColor = [System.Drawing.Color]::White
$script:btnInstall.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$script:btnInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$script:btnInstall.UseVisualStyleBackColor = $false
$script:btnInstall.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 37, 99, 235)

$script:chkDesktop = New-Object System.Windows.Forms.CheckBox
$script:chkDesktop.Text = 'Desktop shortcut'
$script:chkDesktop.Location = New-Object System.Drawing.Point(346, 18)
$script:chkDesktop.AutoSize = $true
$script:chkDesktop.Checked = $true
$script:chkDesktop.ForeColor = [System.Drawing.Color]::FromArgb(255, 203, 213, 225)
$script:chkDesktop.BackColor = [System.Drawing.Color]::Transparent

$script:chkStartMenu = New-Object System.Windows.Forms.CheckBox
$script:chkStartMenu.Text = 'Start Menu group'
$script:chkStartMenu.Location = New-Object System.Drawing.Point(346, 40)
$script:chkStartMenu.AutoSize = $true
$script:chkStartMenu.Checked = $true
$script:chkStartMenu.ForeColor = [System.Drawing.Color]::FromArgb(255, 203, 213, 225)
$script:chkStartMenu.BackColor = [System.Drawing.Color]::Transparent

$script:lnkUninstall = New-Object System.Windows.Forms.LinkLabel
$script:lnkUninstall.Text = 'Uninstall'
$script:lnkUninstall.Location = New-Object System.Drawing.Point(520, 20)
$script:lnkUninstall.AutoSize = $true
$script:lnkUninstall.LinkColor = [System.Drawing.Color]::FromArgb(255, 248, 113, 113)
$script:lnkUninstall.ActiveLinkColor = [System.Drawing.Color]::FromArgb(255, 252, 165, 165)
$script:lnkUninstall.VisitedLinkColor = [System.Drawing.Color]::FromArgb(255, 248, 113, 113)
$script:lnkUninstall.BackColor = [System.Drawing.Color]::Transparent

$script:lnkOpenFolder = New-Object System.Windows.Forms.LinkLabel
$script:lnkOpenFolder.Text = 'Open app folder'
$script:lnkOpenFolder.Location = New-Object System.Drawing.Point(590, 20)
$script:lnkOpenFolder.AutoSize = $true
$script:lnkOpenFolder.LinkColor = [System.Drawing.Color]::FromArgb(255, 96, 165, 250)
$script:lnkOpenFolder.ActiveLinkColor = [System.Drawing.Color]::FromArgb(255, 147, 197, 253)
$script:lnkOpenFolder.VisitedLinkColor = [System.Drawing.Color]::FromArgb(255, 96, 165, 250)
$script:lnkOpenFolder.BackColor = [System.Drawing.Color]::Transparent

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Location = New-Object System.Drawing.Point(676, 20)
$script:lblStatus.Size = New-Object System.Drawing.Size(210, 20)
$script:lblStatus.TextAlign = 'MiddleRight'
$script:lblStatus.AutoSize = $false
$script:lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$script:lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 148, 163, 184)
$script:lblStatus.BackColor = [System.Drawing.Color]::Transparent
$script:lblStatus.Anchor = 'Top,Right'

$footer.Controls.AddRange(@($script:btnLaunchSelected, $script:btnInstall, $script:chkDesktop, $script:chkStartMenu, $script:lnkUninstall, $script:lnkOpenFolder, $script:lblStatus))

function Set-Busy {
    param([bool]$Busy)
    foreach ($c in @($script:btnInstall, $script:lnkUninstall, $script:chkDesktop, $script:chkStartMenu)) {
        if ($c) { $c.Enabled = -not $Busy }
    }
    $form.Cursor = if ($Busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------- app cards ----------
$script:SelectedApp = $null
$script:CardPanels = New-Object System.Collections.ArrayList
$cardWidth = 420
$cardHeight = 96
$gap = 14
$colCount = 2

function New-AppCard {
    param($App)

    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size($cardWidth, $cardHeight)
    $card.Margin = New-Object System.Windows.Forms.Padding($gap, $gap, 0, 0)
    $card.BackColor = [System.Drawing.Color]::FromArgb(255, 17, 28, 46)
    $card.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Tag = $App

    $accent = [System.Drawing.Color]::FromArgb(255, 96, 165, 250)
    $muted  = [System.Drawing.Color]::FromArgb(255, 148, 163, 184)
    $white  = [System.Drawing.Color]::FromArgb(255, 241, 245, 249)

    $card.Add_Paint({
            param($sender, $e)
            $isSel = ($script:SelectedApp -ne $null) -and ($script:SelectedApp.Id -eq $sender.Tag.Id)
            $border = if ($isSel) { [System.Drawing.Color]::FromArgb(255, 34, 197, 94) } else { [System.Drawing.Color]::FromArgb(255, 30, 41, 59) }
            $pen = [System.Drawing.Pen]::new($border, [System.Single]$(if ($isSel) { 2 } else { 1 }))
            try { $e.Graphics.DrawRectangle($pen, 0, 0, $sender.Width - 1, $sender.Height - 1) } finally { $pen.Dispose() }
            if ($isSel) {
                $bar = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 34, 197, 94))
                try { $e.Graphics.FillRectangle($bar, 0, 0, 4, $sender.Height) } finally { $bar.Dispose() }
            }
        })

    $appIcon = Resolve-AppIcon $App
    $pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(52, 52)
    $pic.Location = New-Object System.Drawing.Point(16, 22)
    $pic.SizeMode = 'Zoom'
    $pic.BackColor = [System.Drawing.Color]::Transparent
    if ($appIcon) {
        try { $pic.Image = [System.Drawing.Image]::FromFile($appIcon) } catch { }
    }
    $card.Controls.Add($pic)

    $name = New-Object System.Windows.Forms.Label
    $name.Text = $App.Name
    $name.Location = New-Object System.Drawing.Point(80, 14)
    $name.Size = New-Object System.Drawing.Size(320, 22)
    $name.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $name.ForeColor = $white
    $name.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($name)

    $desc = New-Object System.Windows.Forms.Label
    $desc.Text = $App.Quote
    $desc.Location = New-Object System.Drawing.Point(80, 38)
    $desc.Size = New-Object System.Drawing.Size(328, 34)
    $desc.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $desc.ForeColor = $muted
    $desc.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($desc)

    $badgeText = if ($App.Admin) { 'Admin' } else { 'Standard' }
    $badge = New-Object System.Windows.Forms.Label
    $badge.Text = $badgeText
    $badge.AutoSize = $true
    $badge.Location = New-Object System.Drawing.Point(80, 72)
    $badge.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    $badge.ForeColor = if ($App.Admin) { [System.Drawing.Color]::FromArgb(255, 250, 204, 21) } else { [System.Drawing.Color]::FromArgb(255, 134, 239, 172) }
    $badge.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($badge)

    $tags = New-Object System.Windows.Forms.Label
    $tags.Text = $App.Tags
    $tags.AutoSize = $true
    $tags.Location = New-Object System.Drawing.Point(134, 72)
    $tags.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $tags.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 116, 139)
    $tags.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($tags)

    $launch = New-Object System.Windows.Forms.Button
    $launch.Text = 'Open'
    $launch.Size = New-Object System.Drawing.Size(72, 26)
    $launch.Location = New-Object System.Drawing.Point(332, 62)
    $launch.FlatStyle = 'Flat'
    $launch.FlatAppearance.BorderSize = 0
    $launch.BackColor = [System.Drawing.Color]::FromArgb(255, 34, 197, 94)
    $launch.ForeColor = [System.Drawing.Color]::White
    $launch.Font = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Bold)
    $launch.Cursor = [System.Windows.Forms.Cursors]::Hand
    $launch.UseVisualStyleBackColor = $false
    $launch.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 22, 163, 74)
    $launch.Add_Click({ param($s, $e) $script:SelectedApp = $App; Refresh-Cards; Start-App $App })
    $card.Controls.Add($launch)

    $select = {
        param($s, $e)
        $script:SelectedApp = $App
        Refresh-Cards
        Set-Status $App.Quote ([System.Drawing.Color]::FromArgb(255, 148, 163, 184))
    }
    $card.Add_Click($select)
    $name.Add_Click($select)
    $desc.Add_Click($select)
    $pic.Add_Click($select)

    [void]$script:CardPanels.Add($card)
    return $card
}

function Refresh-Cards {
    foreach ($c in $script:CardPanels) { $c.Invalidate() }
    if ($script:SelectedApp) {
        $script:btnLaunchSelected.Text = "Launch $($script:SelectedApp.Name)"
        $script:btnLaunchSelected.Enabled = $true
    } else {
        $script:btnLaunchSelected.Text = 'Launch selected'
        $script:btnLaunchSelected.Enabled = $false
    }
}

function Build-Cards {
    $body.SuspendLayout()
    $body.Controls.Clear()
    $script:CardPanels.Clear()
    foreach ($app in $script:Catalog) {
        $card = New-AppCard $app
        $body.Controls.Add($card)
    }
    $body.ResumeLayout($true)
}

# ---------- wiring ----------
$script:btnLaunchSelected.Add_Click({ if ($script:SelectedApp) { Start-App $script:SelectedApp } })
$script:btnInstall.Add_Click({ Invoke-Install })
$script:lnkUninstall.Add_LinkClicked({ Invoke-Uninstall })
$script:lnkOpenFolder.Add_LinkClicked({ Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:BaseDir)`"" })
$form.Add_KeyDown({ param($sender, $e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() } })
$form.Add_FormClosed({
        foreach ($obj in @($titleFont, $subtitleFont, $versionFont, $whiteBrush, $softBrush, $verBrush, $script:formIcon)) {
            if ($obj) { try { $obj.Dispose() } catch { } }
        }
        if ($script:headerIconBmp) { try { $script:headerIconBmp.Dispose() } catch { } }
    })

# docking: fill first, then footer, then header
$form.Controls.Add($body)
$form.Controls.Add($footer)
$form.Controls.Add($script:header)

Build-Cards
Refresh-Cards
Update-InstallState
Set-Status 'Select an app and press Open, or use Launch selected.' ([System.Drawing.Color]::FromArgb(255, 148, 163, 184))
[void]$form.ShowDialog()
