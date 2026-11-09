Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

if (-not ('ComponentDedupNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ComponentDedupNative {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
}

$consoleHandle = [ComponentDedupNative]::GetConsoleWindow()
if ($consoleHandle -ne [IntPtr]::Zero) {
    [void][ComponentDedupNative]::ShowWindow($consoleHandle, 0)
}

$script:ScriptPath = $PSCommandPath
if (-not $script:ScriptPath) { $script:ScriptPath = $MyInvocation.MyCommand.Path }
$script:Busy = $false
$script:InTick = $false
$script:CancelRequested = $false
$script:Job = $null
$script:LastScanInput = $null
$script:LastScanScope = $null
$script:GridRecords = [System.Collections.Generic.List[object]]::new()
$script:LogWriter = $null
$script:CurrentLogPath = $null
$script:FileLogCounter = 0

$script:UIFont = [System.Drawing.Font]::new('Segoe UI', [float]9)
$script:BoldFont = [System.Drawing.Font]::new('Segoe UI', [float]9, [System.Drawing.FontStyle]::Bold)
$script:SmallFont = [System.Drawing.Font]::new('Segoe UI', [float]8)
$script:HeaderFont = [System.Drawing.Font]::new('Segoe UI', [float]20, [System.Drawing.FontStyle]::Bold)
$script:SubFont = [System.Drawing.Font]::new('Segoe UI', [float]9)
$script:TextColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
$script:MutedColor = [System.Drawing.Color]::FromArgb(108, 122, 137)
$script:Accent = [System.Drawing.Color]::FromArgb(30, 132, 73)
$script:AccentLight = [System.Drawing.Color]::FromArgb(39, 174, 96)
$script:HeaderDark = [System.Drawing.Color]::FromArgb(15, 76, 44)
$script:PanelColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$script:Danger = [System.Drawing.Color]::FromArgb(192, 57, 43)

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height, [System.Drawing.Font]$Font, $ForeColor)
    $control = New-Object System.Windows.Forms.Label
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    if ($Font) { $control.Font = $Font } else { $control.Font = $script:UIFont }
    if ($null -ne $ForeColor) { $control.ForeColor = $ForeColor } else { $control.ForeColor = $script:TextColor }
    $control.AutoEllipsis = $true
    return $control
}

function New-TextBox {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height, [string]$Text = '', [System.Drawing.Font]$Font)
    $control = New-Object System.Windows.Forms.TextBox
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.Text = $Text
    if ($Font) { $control.Font = $Font } else { $control.Font = $script:UIFont }
    $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    return $control
}

function New-CheckBox {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height, [bool]$Checked = $false)
    $control = New-Object System.Windows.Forms.CheckBox
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.Font = $script:UIFont
    $control.ForeColor = $script:TextColor
    $control.Checked = $Checked
    return $control
}

function New-Radio {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height, [bool]$Checked = $false)
    $control = New-Object System.Windows.Forms.RadioButton
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.Font = $script:UIFont
    $control.ForeColor = $script:TextColor
    $control.Checked = $Checked
    return $control
}

function New-Group {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    $control = New-Object System.Windows.Forms.GroupBox
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.Font = $script:BoldFont
    $control.ForeColor = $script:TextColor
    $control.BackColor = [System.Drawing.Color]::White
    return $control
}

function New-Combo {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height)
    $control = New-Object System.Windows.Forms.ComboBox
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.Font = $script:UIFont
    $control.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    return $control
}

function New-FlatButton {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width, [int]$Height, [System.Drawing.Color]$BackColor, $ForeColor, [System.Drawing.Font]$Font)
    $control = New-Object System.Windows.Forms.Button
    $control.Text = $Text
    $control.Location = [System.Drawing.Point]::new($X, $Y)
    $control.Size = [System.Drawing.Size]::new($Width, $Height)
    $control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $control.FlatAppearance.BorderSize = 0
    $control.FlatAppearance.MouseOverBackColor = [System.Windows.Forms.ControlPaint]::Light($BackColor, [float]0.12)
    $control.FlatAppearance.MouseDownBackColor = [System.Windows.Forms.ControlPaint]::Dark($BackColor, [float]0.08)
    $control.BackColor = $BackColor
    if ($null -ne $ForeColor) { $control.ForeColor = $ForeColor } else { $control.ForeColor = [System.Drawing.Color]::White }
    if ($Font) { $control.Font = $Font } else { $control.Font = $script:UIFont }
    $control.Cursor = [System.Windows.Forms.Cursors]::Hand
    $control.UseVisualStyleBackColor = $false
    return $control
}

function Get-CleanPath {
    param([string]$Path)
    if ($null -eq $Path) { return '' }
    return $Path.Trim().Trim([char]'"')
}

function Get-SettingsPath {
    return Join-Path (Join-Path $env:APPDATA 'ComponentDeduplicator') 'settings.json'
}

function Get-DefaultSettings {
    return [ordered]@{
        InputDir = ''
        OutputDir = ''
        Action = 'Move'
        Scope = 'All'
        HashAlgorithm = 'SHA256'
        MatchMode = 'Exact content (hash)'
        Extensions = @('.exe', '.dll', '.msi', '.msix', '.appx', '.bat', '.cmd', '.ps1', '.vbs', '.reg', '.ocx', '.cpl', '.config', '.manifest')
        CustomExtensions = ''
        IncludeSubfolders = $true
        IncludeHidden = $false
        PreserveStructure = $false
        DryRun = $false
        ConfirmActions = $true
        LogToFile = $false
        LogPath = (Join-Path (Join-Path $env:APPDATA 'ComponentDeduplicator') 'deduplicator.log')
        OpenOutputAfter = $false
        VerboseLog = $false
        MaxScanFiles = 100000
    }
}

function Load-Settings {
    $defaults = Get-DefaultSettings
    $path = Get-SettingsPath
    if (Test-Path -LiteralPath $path) {
        try {
            $loaded = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($key in @($defaults.Keys)) {
                $property = $loaded.PSObject.Properties[$key]
                if ($property) { $defaults[$key] = $property.Value }
            }
        } catch { }
    }
    return [pscustomobject]$defaults
}

function Save-Settings {
    if (-not $script:TxtIn) { return }
    if ($script:LogWriter -and $script:CurrentLogPath -ne (Get-CleanPath $script:TxtLogPath.Text)) {
        try { $script:LogWriter.Dispose() } catch { }
        $script:LogWriter = $null
        $script:CurrentLogPath = $null
    }
    if ($script:ChkLogToFile -and -not $script:ChkLogToFile.Checked -and $script:LogWriter) {
        try { $script:LogWriter.Dispose() } catch { }
        $script:LogWriter = $null
        $script:CurrentLogPath = $null
    }
    $settings = [ordered]@{
        InputDir = Get-CleanPath $script:TxtIn.Text
        OutputDir = Get-CleanPath $script:TxtOut.Text
        Action = Get-SelectedAction
        Scope = Get-SelectedScope
        HashAlgorithm = [string]$script:CmbHash.SelectedItem
        MatchMode = [string]$script:CmbMatch.SelectedItem
        Extensions = @(Get-CheckedExtensions)
        CustomExtensions = $script:TxtCustomExt.Text.Trim()
        IncludeSubfolders = [bool]$script:ChkRecurse.Checked
        IncludeHidden = [bool]$script:ChkHidden.Checked
        PreserveStructure = [bool]$script:ChkPreserve.Checked
        DryRun = [bool]$script:ChkDryRun.Checked
        ConfirmActions = [bool]$script:ChkConfirm.Checked
        LogToFile = [bool]$script:ChkLogToFile.Checked
        LogPath = Get-CleanPath $script:TxtLogPath.Text
        OpenOutputAfter = [bool]$script:ChkOpenOutput.Checked
        VerboseLog = [bool]$script:ChkVerbose.Checked
        MaxScanFiles = [int]$script:NumMaxFiles.Value
    }
    try {
        $path = Get-SettingsPath
        $folder = [System.IO.Path]::GetDirectoryName($path)
        if ($folder -and -not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        $settings | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $path -Encoding UTF8
    } catch { }
    $script:Settings = [pscustomobject]$settings
}

function Get-CheckedExtensions {
    $result = @()
    foreach ($item in $script:ChkExtensions.CheckedItems) { $result += [string]$item }
    return $result
}

function Get-SelectedExtensions {
    $result = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $script:ChkExtensions.CheckedItems) { $result.Add(([string]$item).ToLowerInvariant()) }
    $custom = $script:TxtCustomExt.Text.Trim()
    if ($custom) {
        foreach ($part in ($custom -split '[,;\s]+')) {
            $value = $part.Trim()
            if (-not $value) { continue }
            if (-not $value.StartsWith('.')) { $value = '.' + $value }
            $result.Add($value.ToLowerInvariant())
        }
    }
    return @($result | Sort-Object -Unique)
}

function Get-SelectedAction {
    if ($script:RadCopy.Checked) { return 'Copy' }
    if ($script:RadRecycle.Checked) { return 'Recycle' }
    return 'Move'
}

function Get-SelectedScope {
    if ($script:RadScopeAll.Checked) { return 'All' }
    return 'Duplicates'
}

function Set-ComboSelection {
    param($Combo, [string]$Value)
    if (-not $Combo) { return }
    $index = $Combo.Items.IndexOf($Value)
    if ($index -ge 0) { $Combo.SelectedIndex = $index }
    elseif ($Combo.Items.Count -gt 0) { $Combo.SelectedIndex = 0 }
}

function Write-FileLog {
    param([string]$Line)
    try {
        if (-not $script:ChkLogToFile -or -not $script:ChkLogToFile.Checked) { return }
        if (-not $script:TxtLogPath) { return }
        $path = Get-CleanPath $script:TxtLogPath.Text
        if ([string]::IsNullOrWhiteSpace($path)) { return }
        if ($script:LogWriter -and $script:CurrentLogPath -ne $path) {
            try { $script:LogWriter.Dispose() } catch { }
            $script:LogWriter = $null
            $script:CurrentLogPath = $null
        }
        if (-not $script:LogWriter) {
            $folder = [System.IO.Path]::GetDirectoryName($path)
            if ($folder -and -not (Test-Path -LiteralPath $folder)) {
                New-Item -ItemType Directory -Path $folder -Force | Out-Null
            }
            $encoding = [System.Text.UTF8Encoding]::new($false)
            $script:LogWriter = [System.IO.StreamWriter]::new($path, $true, $encoding)
            $script:LogWriter.AutoFlush = $false
            $script:CurrentLogPath = $path
        }
        $script:LogWriter.WriteLine($Line)
        $script:FileLogCounter++
        if ($script:FileLogCounter -ge 50) {
            $script:LogWriter.Flush()
            $script:FileLogCounter = 0
        }
    } catch { }
}

function Write-Log {
    param([string]$Message, [string]$Kind = 'Info')
    try {
        $color = [System.Drawing.Color]::LightGreen
        switch ($Kind) {
            'Success' { $color = [System.Drawing.Color]::FromArgb(0, 255, 160) }
            'Warning' { $color = [System.Drawing.Color]::Gold }
            'Error' { $color = [System.Drawing.Color]::Tomato }
            'Dry' { $color = [System.Drawing.Color]::Cyan }
            'Muted' { $color = [System.Drawing.Color]::Silver }
        }
        $stamp = Get-Date -Format 'HH:mm:ss'
        $line = "[$stamp] $Message"
        if ($script:LogBox) {
            if ($script:LogBox.TextLength -gt 250000) { $script:LogBox.Clear() }
            $script:LogBox.SelectionStart = $script:LogBox.TextLength
            $script:LogBox.SelectionLength = 0
            $script:LogBox.SelectionColor = $color
            $script:LogBox.AppendText($line + "`r`n")
            $script:LogBox.SelectionColor = [System.Drawing.Color]::LightGreen
            $script:LogBox.ScrollToCaret()
        }
        if ($script:ChkLogToFile -and $script:ChkLogToFile.Checked) {
            Write-FileLog $line
        }
    } catch { }
}

function Write-InstallLog {
    param([string]$Message, [string]$Kind = 'Info')
    try {
        $color = [System.Drawing.Color]::LightGreen
        switch ($Kind) {
            'Success' { $color = [System.Drawing.Color]::FromArgb(0, 255, 160) }
            'Warning' { $color = [System.Drawing.Color]::Gold }
            'Error' { $color = [System.Drawing.Color]::Tomato }
        }
        if ($script:InstallLog) {
            $stamp = Get-Date -Format 'HH:mm:ss'
            $script:InstallLog.SelectionStart = $script:InstallLog.TextLength
            $script:InstallLog.SelectionLength = 0
            $script:InstallLog.SelectionColor = $color
            $script:InstallLog.AppendText("[$stamp] $Message`r`n")
            $script:InstallLog.SelectionColor = [System.Drawing.Color]::LightGreen
            $script:InstallLog.ScrollToCaret()
        }
    } catch { }
}

function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return ('{0} B' -f $Bytes)
}

function Get-RoundedPath {
    param([float]$X, [float]$Y, [float]$Width, [float]$Height, [float]$Radius)
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $diameter = [float]($Radius * 2)
    if ($diameter -le 0) { $diameter = 1 }
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc(($X + $Width - $diameter), $Y, $diameter, $diameter, 270, 90)
    $path.AddArc(($X + $Width - $diameter), ($Y + $Height - $diameter), $diameter, $diameter, 0, 90)
    $path.AddArc($X, ($Y + $Height - $diameter), $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-AppIconBitmap {
    param([int]$Size)
    if ($Size -lt 16) { $Size = 16 }
    $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $scale = $Size / 256.0

    $shadowPath = Get-RoundedPath (12 * $scale) (14 * $scale) (232 * $scale) (232 * $scale) (50 * $scale)
    $shadowBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(50, 0, 0, 0))
    $graphics.FillPath($shadowBrush, $shadowPath)
    $shadowBrush.Dispose()
    $shadowPath.Dispose()

    $backgroundRect = [System.Drawing.RectangleF]::new((8 * $scale), (8 * $scale), (240 * $scale), (240 * $scale))
    $backgroundPath = Get-RoundedPath (8 * $scale) (8 * $scale) (240 * $scale) (240 * $scale) (48 * $scale)
    $backgroundBrush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($backgroundRect, [System.Drawing.Color]::FromArgb(255, 15, 76, 44), [System.Drawing.Color]::FromArgb(255, 46, 204, 113), [float]45)
    $graphics.FillPath($backgroundBrush, $backgroundPath)
    $backgroundBrush.Dispose()
    $backgroundPath.Dispose()

    $highlightBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(45, 255, 255, 255))
    $graphics.FillEllipse($highlightBrush, (20 * $scale), (14 * $scale), (150 * $scale), (90 * $scale))
    $highlightBrush.Dispose()

    $backPath = Get-RoundedPath (58 * $scale) (52 * $scale) (120 * $scale) (150 * $scale) (12 * $scale)
    $backBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(110, 255, 255, 255))
    $graphics.FillPath($backBrush, $backPath)
    $backBrush.Dispose()
    $backPath.Dispose()

    $frontX = 82 * $scale
    $frontY = 72 * $scale
    $frontW = 116 * $scale
    $frontH = 150 * $scale
    $frontPath = Get-RoundedPath $frontX $frontY $frontW $frontH (12 * $scale)
    $whiteBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $graphics.FillPath($whiteBrush, $frontPath)
    $whiteBrush.Dispose()
    $borderPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(40, 90, 70), [float][Math]::Max(1, (2 * $scale)))
    $graphics.DrawPath($borderPen, $frontPath)
    $borderPen.Dispose()
    $frontPath.Dispose()

    $fold = 30 * $scale
    $foldBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(235, 228, 238))
    $foldPoints = [System.Drawing.PointF[]]@(
        [System.Drawing.PointF]::new(($frontX + $frontW - $fold), $frontY),
        [System.Drawing.PointF]::new(($frontX + $frontW), $frontY),
        [System.Drawing.PointF]::new(($frontX + $frontW), ($frontY + $fold))
    )
    $graphics.FillPolygon($foldBrush, $foldPoints)
    $foldBrush.Dispose()
    $foldPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(180, 200, 215), [float][Math]::Max(1, (1.5 * $scale)))
    $graphics.DrawLine($foldPen, ($frontX + $frontW - $fold), $frontY, ($frontX + $frontW), ($frontY + $fold))
    $foldPen.Dispose()

    $linePen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(155, 180, 205), [float][Math]::Max(1, (4 * $scale)))
    $linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $graphics.DrawLine($linePen, ($frontX + (16 * $scale)), (132 * $scale), ($frontX + (86 * $scale)), (132 * $scale))
    $graphics.DrawLine($linePen, ($frontX + (16 * $scale)), (150 * $scale), ($frontX + (72 * $scale)), (150 * $scale))
    $graphics.DrawLine($linePen, ($frontX + (16 * $scale)), (168 * $scale), ($frontX + (56 * $scale)), (168 * $scale))
    $linePen.Dispose()

    $badgeX = 150 * $scale
    $badgeY = 150 * $scale
    $badgeW = 88 * $scale
    $badgeH = 88 * $scale
    $badgeBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 15, 76, 44))
    $graphics.FillEllipse($badgeBrush, $badgeX, $badgeY, $badgeW, $badgeH)
    $badgeBrush.Dispose()
    $ringPen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, [float][Math]::Max(1, (3 * $scale)))
    $graphics.DrawEllipse($ringPen, $badgeX, $badgeY, $badgeW, $badgeH)
    $ringPen.Dispose()
    $checkPen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, [float][Math]::Max(1.5, (10 * $scale)))
    $checkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $checkPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $checkPoints = [System.Drawing.PointF[]]@(
        [System.Drawing.PointF]::new((170 * $scale), (196 * $scale)),
        [System.Drawing.PointF]::new((187 * $scale), (213 * $scale)),
        [System.Drawing.PointF]::new((220 * $scale), (173 * $scale))
    )
    $graphics.DrawLines($checkPen, $checkPoints)
    $checkPen.Dispose()

    $graphics.Dispose()
    return $bitmap
}

function Convert-BitmapToIconFrame {
    param([System.Drawing.Bitmap]$Bitmap)
    $size = $Bitmap.Width
    $rectangle = [System.Drawing.Rectangle]::new(0, 0, $size, $size)
    $locked = $null
    $bytes = $null
    $stride = $size * 4
    try {
        $locked = $Bitmap.LockBits($rectangle, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $stride = $locked.Stride
        $bytes = [byte[]]::new($stride * $size)
        [System.Runtime.InteropServices.Marshal]::Copy($locked.Scan0, $bytes, 0, $bytes.Length)
    } finally {
        if ($locked) { $Bitmap.UnlockBits($locked) }
    }

    $stream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.BinaryWriter]::new($stream)
    [void]$writer.Write([uint32]40)
    [void]$writer.Write([int32]$size)
    [void]$writer.Write([int32]($size * 2))
    [void]$writer.Write([uint16]1)
    [void]$writer.Write([uint16]32)
    [void]$writer.Write([uint32]0)
    [void]$writer.Write([uint32]($size * $size * 4))
    [void]$writer.Write([int32]0)
    [void]$writer.Write([int32]0)
    [void]$writer.Write([uint32]0)
    [void]$writer.Write([uint32]0)
    [void]$writer.Write([uint32]0)
    for ($y = $size - 1; $y -ge 0; $y--) {
        [void]$writer.Write($bytes, ($y * $stride), ($size * 4))
    }
    $maskStride = [int]([Math]::Ceiling($size / 32.0) * 4)
    $mask = [byte[]]::new($maskStride * $size)
    [void]$writer.Write($mask)
    $writer.Flush()
    $result = $stream.ToArray()
    $writer.Dispose()
    $stream.Dispose()
    return ,$result
}

function New-AppIcon {
    param([string]$Path)
    $folder = [System.IO.Path]::GetDirectoryName($Path)
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $sizes = @(16, 24, 32, 48, 64, 128, 256)
    $frames = @()
    foreach ($size in $sizes) {
        $bitmap = New-AppIconBitmap $size
        $frames += ,(Convert-BitmapToIconFrame $bitmap)
        $bitmap.Dispose()
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $writer = [System.IO.BinaryWriter]::new($stream)
    [void]$writer.Write([uint16]0)
    [void]$writer.Write([uint16]1)
    [void]$writer.Write([uint16]$sizes.Count)
    $offset = 6 + (16 * $sizes.Count)
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $size = $sizes[$i]
        $dimension = if ($size -ge 256) { [byte]0 } else { [byte]$size }
        [void]$writer.Write($dimension)
        [void]$writer.Write($dimension)
        [void]$writer.Write([byte]0)
        [void]$writer.Write([byte]0)
        [void]$writer.Write([uint16]1)
        [void]$writer.Write([uint16]32)
        $frame = $frames[$i]
        [void]$writer.Write([uint32]$frame.Length)
        [void]$writer.Write([uint32]$offset)
        $offset += $frame.Length
    }
    foreach ($frame in $frames) { [void]$writer.Write($frame) }
    $writer.Flush()
    $writer.Dispose()
    $stream.Dispose()
}

function Get-RelativePath {
    param([string]$BasePath, [string]$TargetPath)
    $base = [System.IO.Path]::GetFullPath($BasePath)
    if (-not $base.EndsWith('\')) { $base += '\' }
    $target = [System.IO.Path]::GetFullPath($TargetPath)
    $baseUri = [System.Uri]::new($base)
    $targetUri = [System.Uri]::new($target)
    $relative = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
    $relative = $relative.Replace('/', '\')
    if ($relative.StartsWith('..')) { return [System.IO.Path]::GetFileName($TargetPath) }
    return $relative
}

function Get-NonConflictingPath {
    param([string]$Directory, [string]$FileName)
    $destination = Join-Path $Directory $FileName
    if (-not (Test-Path -LiteralPath $destination)) { return $destination }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)
    for ($i = 1; $i -lt 10000; $i++) {
        $destination = Join-Path $Directory ('{0}_{1}{2}' -f $baseName, $i, $extension)
        if (-not (Test-Path -LiteralPath $destination)) { return $destination }
    }
    return Join-Path $Directory ('{0}_{1}{2}' -f $baseName, [guid]::NewGuid().ToString('N').Substring(0, 8), $extension)
}

function Get-InstallDir {
    return Join-Path $env:LOCALAPPDATA 'ComponentDeduplicator'
}

function Get-DesktopShortcutPath {
    return Join-Path ([Environment]::GetFolderPath('Desktop')) 'Component Optimizer.lnk'
}

function Get-StartMenuShortcutPath {
    return Join-Path ([Environment]::GetFolderPath('Programs')) 'Component Optimizer.lnk'
}

function Install-Application {
    $installDir = Get-InstallDir
    if (-not (Test-Path -LiteralPath $installDir)) {
        New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    }
    $destinationScript = Join-Path $installDir 'ComponentDeduplicator.ps1'
    $source = $script:ScriptPath
    if ($source -and (Test-Path -LiteralPath $source)) {
        $copy = $true
        try {
            if ([System.IO.Path]::GetFullPath($source) -eq [System.IO.Path]::GetFullPath($destinationScript)) { $copy = $false }
        } catch { }
        if ($copy) { Copy-Item -LiteralPath $source -Destination $destinationScript -Force -ErrorAction Stop }
    } elseif (-not (Test-Path -LiteralPath $destinationScript)) {
        throw 'Cannot locate the script file to install.'
    }
    New-AppIcon -Path (Join-Path $installDir 'ComponentDeduplicator.ico')
    return $installDir
}

function New-Shortcut {
    param([string]$ShortcutPath)
    $installDir = Get-InstallDir
    $scriptPath = Join-Path $installDir 'ComponentDeduplicator.ps1'
    $iconPath = Join-Path $installDir 'ComponentDeduplicator.ico'
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw 'The application is not installed yet.' }
    if (-not (Test-Path -LiteralPath $iconPath)) { New-AppIcon -Path $iconPath }
    $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powerShellPath)) {
        $powerShellPath = (Get-Command powershell.exe -ErrorAction Stop).Source
    }
    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $powerShellPath
        $shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
        $shortcut.WorkingDirectory = $installDir
        $shortcut.IconLocation = "$iconPath,0"
        $shortcut.Description = 'Optimize, deduplicate and relocate application components'
        $shortcut.WindowStyle = 7
        $shortcut.Save()
    } finally {
        if ($shortcut) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut) }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
}

function New-DesktopShortcut {
    $path = Get-DesktopShortcutPath
    New-Shortcut -ShortcutPath $path
    return $path
}

function New-StartMenuShortcut {
    $path = Get-StartMenuShortcutPath
    New-Shortcut -ShortcutPath $path
    return $path
}

function Uninstall-Application {
    foreach ($path in @((Get-DesktopShortcutPath), (Get-StartMenuShortcutPath))) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    }
    $installDir = Get-InstallDir
    if (-not (Test-Path -LiteralPath $installDir)) { return 'Shortcuts removed. Install folder was already absent.' }
    $current = $script:ScriptPath
    $runningFromInstall = $false
    if ($current) {
        try {
            $runningFromInstall = [System.IO.Path]::GetFullPath($current).StartsWith(([System.IO.Path]::GetFullPath($installDir) + '\'), [System.StringComparison]::OrdinalIgnoreCase)
        } catch { }
    }
    if ($runningFromInstall) {
        $command = '/c ping 127.0.0.1 -n 3 >nul & rmdir /s /q "' + $installDir + '"'
        Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList $command -WindowStyle Hidden
        return 'Shortcuts removed. The install folder will be deleted after this window closes.'
    }
    Remove-Item -LiteralPath $installDir -Recurse -Force -ErrorAction Stop
    return 'Uninstalled successfully.'
}

function Set-Busy {
    param([bool]$Busy)
    $script:Busy = $Busy
    if ($script:BtnScan) { $script:BtnScan.Enabled = -not $Busy }
    if ($script:BtnCancel) { $script:BtnCancel.Enabled = $Busy }
    if ($script:BtnNew) { $script:BtnNew.Enabled = -not $Busy }
    if ($script:BtnCheckAll) { $script:BtnCheckAll.Enabled = -not $Busy }
    if ($script:BtnCheckNone) { $script:BtnCheckNone.Enabled = -not $Busy }
    if ($script:GroupFolders) { $script:GroupFolders.Enabled = -not $Busy }
    if ($script:GroupOptions) { $script:GroupOptions.Enabled = -not $Busy }
    if ($script:Grid) { $script:Grid.Enabled = -not $Busy }
    if ($script:BtnClean) { $script:BtnClean.Enabled = (-not $Busy) -and ($script:GridRecords.Count -gt 0) }
    Update-ActionUi
}

function Update-ActionUi {
    if (-not $script:TxtOut) { return }
    $allow = (-not $script:RadRecycle.Checked) -and (-not $script:Busy)
    $script:TxtOut.Enabled = $allow
    $script:BtnOut.Enabled = $allow
    $script:ChkPreserve.Enabled = $allow
}

function Set-GridHeaders {
    param([bool]$AllFiles)
    if (-not $script:Grid) { return }
    $script:Grid.Columns['Master'].HeaderText = if ($AllFiles) { 'Source folder' } else { 'Original (kept)' }
    $script:Grid.Columns['Duplicate'].HeaderText = if ($AllFiles) { 'File to process' } else { 'Duplicate to process' }
}

function Update-GridSummary {
    $total = $script:GridRecords.Count
    $selected = 0
    $bytes = [long]0
    foreach ($record in $script:GridRecords) {
        if ($record.Process) {
            $selected++
            $bytes += [long]$record.Size
        }
    }
    $isAll = $script:LastScanScope -eq 'All'
    $noun = if ($isAll) { 'file(s)' } else { 'duplicate(s)' }
    $sizeLabel = if ($isAll) { 'total' } else { 'reclaimable' }
    if ($script:LblStats) {
        $script:LblStats.Text = ('{0:N0} {1}, {2:N0} selected - {3} {4}' -f $total, $noun, $selected, (Format-Size $bytes), $sizeLabel)
    }
}

function Reset-ForNextOperation {
    if ($script:Busy) { return }
    $script:Grid.RowCount = 0
    $script:GridRecords = [System.Collections.Generic.List[object]]::new()
    $script:LastScanInput = $null
    $script:LastScanScope = $null
    $script:CancelRequested = $false
    $script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    $script:Progress.Value = 0
    $script:LblStatus.Text = 'Ready for another operation.'
    $script:LblStats.Text = ''
    $script:LogBox.Clear()
    Write-Log 'Ready for another operation. Choose a scope, then click Scan.' 'Info'
    Set-Busy $false
    Save-Settings
}

function Finish-Job {
    param([string]$Message, [string]$Kind = 'Info')
    $finishedJob = $script:Job
    if ($script:JobTimer) { $script:JobTimer.Stop() }
    if ($script:LogWriter) { try { $script:LogWriter.Flush() } catch { } }
    $script:Job = $null
    $script:CancelRequested = $false
    if ($script:Progress) { $script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks }
    if ($finishedJob -and $finishedJob.Type -eq 'Scan' -and $Kind -ne 'Success') {
        $script:Grid.RowCount = 0
        $script:GridRecords = [System.Collections.Generic.List[object]]::new()
        $script:LastScanInput = $null
        $script:LastScanScope = $null
        $script:LblStats.Text = ''
    }
    if ($finishedJob -and $finishedJob.Type -eq 'Clean' -and $Kind -ne 'Success') {
        $script:Grid.RowCount = 0
        $kept = [System.Collections.Generic.List[object]]::new()
        foreach ($record in $script:GridRecords) {
            if (-not $record.Removed) { $kept.Add($record) }
        }
        $script:GridRecords = $kept
        $script:Grid.RowCount = $kept.Count
        Update-GridSummary
    }
    Set-Busy $false
    if ($Message) {
        $script:LblStatus.Text = $Message
        if ($Kind -eq 'Success') {
            try { $script:LogBox.Clear() } catch { }
            Write-Log $Message 'Success'
            Write-Log 'Operation complete. Ready for another operation.' 'Info'
        } else {
            Write-Log $Message $Kind
        }
    }
}

function Update-JobStatus {
    $job = $script:Job
    if (-not $job) { return }
    if ($job.Type -eq 'Scan') {
        switch ($job.Phase) {
            'Collect' {
                $script:LblStatus.Text = ('Collecting files... {0:N0} found in {1:N0} folder(s)' -f $job.Files.Count, $job.DirectoriesScanned)
            }
            'Hash' {
                $percent = 0
                if ($job.Total -gt 0) { $percent = [int](5 + (90.0 * $job.Processed / $job.Total)) }
                $script:Progress.Value = [Math]::Min(95, [Math]::Max(5, $percent))
                $script:LblStatus.Text = ('Analyzing {0:N0} of {1:N0} file(s)...' -f $job.Processed, $job.Total)
            }
            'ListAll' {
                $percent = 0
                if ($job.Total -gt 0) { $percent = [int](5 + (92.0 * $job.Processed / $job.Total)) }
                $script:Progress.Value = [Math]::Min(97, [Math]::Max(5, $percent))
                $script:LblStatus.Text = ('Indexing {0:N0} of {1:N0} file(s)...' -f $job.Processed, $job.Total)
            }
            'Build' {
                if ($job.Keys -and $job.Keys.Count -gt 0) {
                    $percent = [int](95 + (5.0 * $job.KeyIndex / $job.Keys.Count))
                    $script:Progress.Value = [Math]::Min(100, $percent)
                } else {
                    $script:Progress.Value = 95
                }
                if ($script:Grid.RowCount -ne $script:GridRecords.Count) {
                    $script:Grid.RowCount = $script:GridRecords.Count
                }
                $script:LblStatus.Text = ('Building duplicate list... {0:N0} duplicate(s)' -f $script:GridRecords.Count)
            }
        }
    } else {
        $percent = 0
        if ($job.Records.Count -gt 0) { $percent = [int](100.0 * $job.Index / $job.Records.Count) }
        $script:Progress.Value = [Math]::Min(100, [Math]::Max(0, $percent))
        $script:LblStatus.Text = ('Processing {0:N0} of {1:N0}...' -f $job.Index, $job.Records.Count)
    }
}

function Step-ScanJob {
    param([object]$Job)
    if ($Job.Phase -eq 'Collect') {
        if ($Job.DirStack.Count -eq 0 -or $Job.LimitReached) {
            $Job.Total = $Job.Files.Count
            $Job.Processed = 0
            if ($Job.Total -eq 0) {
                Finish-Job 'No matching files were found.' 'Warning'
                return 'Finished'
            }
            Write-Log ('Collected {0:N0} candidate file(s).' -f $Job.Total) 'Info'
            if ($Job.AccessErrors -gt 0) { Write-Log ('Skipped {0:N0} inaccessible item(s).' -f $Job.AccessErrors) 'Warning' }
            if ($Job.LimitReached) { Write-Log ('Safety limit of {0:N0} files reached. Remaining files were skipped.' -f $Job.MaxFiles) 'Warning' }
            $script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
            if ($Job.ScopeAll) {
                $Job.Phase = 'ListAll'
            } else {
                $Job.Phase = 'Hash'
            }
            return 'Continue'
        }
        $directory = $Job.DirStack.Pop()
        $Job.DirectoriesScanned++
        try {
            foreach ($path in [System.IO.Directory]::EnumerateFiles($directory)) {
                if ($Job.Files.Count -ge $Job.MaxFiles) { $Job.LimitReached = $true; break }
                $extension = [System.IO.Path]::GetExtension($path)
                if (-not $Job.ExtensionSet.Contains($extension)) { continue }
                try {
                    $info = [System.IO.FileInfo]::new($path)
                    if (-not $Job.IncludeHidden) {
                        $attributes = $info.Attributes
                        if (($attributes -band [System.IO.FileAttributes]::Hidden) -ne 0) { continue }
                        if (($attributes -band [System.IO.FileAttributes]::System) -ne 0) { continue }
                    }
                    $Job.Files.Add($info)
                } catch { $Job.AccessErrors++ }
            }
        } catch { $Job.AccessErrors++ }
        if ($Job.LimitReached) { $Job.DirStack.Clear() }
        if ($Job.Recurse -and -not $Job.LimitReached) {
            try {
                foreach ($subPath in [System.IO.Directory]::EnumerateDirectories($directory)) {
                    try {
                        if ($Job.ExcludeDir -and $subPath.StartsWith($Job.ExcludeDir, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                        $subInfo = [System.IO.DirectoryInfo]::new($subPath)
                        $subAttributes = $subInfo.Attributes
                        if (-not $Job.IncludeHidden) {
                            if (($subAttributes -band [System.IO.FileAttributes]::Hidden) -ne 0) { continue }
                            if (($subAttributes -band [System.IO.FileAttributes]::System) -ne 0) { continue }
                        }
                        if (($subAttributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                        $Job.DirStack.Push($subPath)
                    } catch { $Job.AccessErrors++ }
                }
            } catch { $Job.AccessErrors++ }
        }
        return 'Continue'
    }

    if ($Job.Phase -eq 'ListAll') {
        if ($Job.Processed -ge $Job.Total) {
            $Job.Files = $null
            $script:Grid.RowCount = $script:GridRecords.Count
            $script:Grid.Invalidate()
            $script:LastScanInput = $Job.InputDir
            $script:LastScanScope = 'All'
            Set-GridHeaders $true
            Update-GridSummary
            $Job.Progress = 100
            $message = ('Scan complete. {0:N0} matching file(s), {1} total.' -f $Job.RecordCount, (Format-Size $Job.RecordBytes))
            Finish-Job $message 'Success'
            return 'Finished'
        }
        $file = $Job.Files[$Job.Processed]
        $Job.Processed++
        $record = [pscustomobject]@{
            Master = $file.DirectoryName
            Duplicate = $file.FullName
            Size = [long]$file.Length
            Match = 'All files'
            Process = $true
            Removed = $false
        }
        $script:GridRecords.Add($record)
        $Job.RecordCount++
        $Job.RecordBytes += [long]$file.Length
        return 'Continue'
    }

    if ($Job.Phase -eq 'Hash') {
        if ($Job.Processed -ge $Job.Total) {
            $Job.Phase = 'Build'
            $Job.Keys = @($Job.HashMap.Keys)
            $Job.KeyIndex = 0
            if ($Job.Files) { $Job.Files = $null }
            if ($Job.Keys.Count -eq 0) {
                Finish-Job 'No comparable files were found.' 'Warning'
                return 'Finished'
            }
            return 'Continue'
        }
        $file = $Job.Files[$Job.Processed]
        $Job.Processed++
        try {
            if ($Job.MatchByName) {
                $key = 'N|' + $file.Name.ToLowerInvariant() + '|' + [string]$file.Length
            } else {
                $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm $Job.Algorithm -ErrorAction Stop).Hash
                $key = 'H|' + $hash
            }
            if (-not $Job.HashMap.ContainsKey($key)) {
                $Job.HashMap[$key] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            }
            $Job.HashMap[$key].Add($file)
        } catch {
            $Job.HashErrors++
            if ($Job.HashErrors -le 20) { Write-Log ('Skipped unreadable file: {0}' -f $file.FullName) 'Warning' }
        }
        return 'Continue'
    }

    if ($Job.Phase -eq 'Build') {
        if ($Job.KeyIndex -ge $Job.Keys.Count) {
            $Job.HashMap = $null
            $Job.Keys = $null
            $script:Grid.RowCount = $script:GridRecords.Count
            $script:Grid.Invalidate()
            $script:LastScanInput = $Job.InputDir
            $script:LastScanScope = 'Duplicates'
            Set-GridHeaders $false
            Update-GridSummary
            $Job.Progress = 100
            $message = ('Scan complete. {0:N0} duplicate(s), {1} reclaimable.' -f $Job.RecordCount, (Format-Size $Job.RecordBytes))
            Finish-Job $message 'Success'
            return 'Finished'
        }
        $key = $Job.Keys[$Job.KeyIndex]
        $Job.KeyIndex++
        $list = $Job.HashMap[$key]
        if ($list -and $list.Count -gt 1) {
            $master = $list[0].FullName
            $matchText = if ($Job.MatchByName) { 'Name + size' } else { $Job.Algorithm }
            for ($i = 1; $i -lt $list.Count; $i++) {
                $duplicate = $list[$i]
                $record = [pscustomobject]@{
                    Master = $master
                    Duplicate = $duplicate.FullName
                    Size = [long]$duplicate.Length
                    Match = $matchText
                    Process = $true
                    Removed = $false
                }
                $script:GridRecords.Add($record)
                $Job.RecordCount++
                $Job.RecordBytes += [long]$duplicate.Length
            }
        }
        return 'Continue'
    }

    return 'Finished'
}

function Step-CleanJob {
    param([object]$Job)
    if ($Job.Index -ge $Job.Records.Count) {
        if ($Job.Action -ne 'Copy' -and -not $Job.DryRun) {
            $script:Grid.RowCount = 0
            $kept = [System.Collections.Generic.List[object]]::new()
            foreach ($record in $script:GridRecords) {
                if (-not $record.Removed) { $kept.Add($record) }
            }
            $script:GridRecords = $kept
            $script:Grid.RowCount = $kept.Count
            $script:Grid.Invalidate()
            Update-GridSummary
        }
        $summary = ''
        if ($Job.DryRun) {
            $summary = ('Dry run complete. {0:N0} file(s), {1}, would be processed.' -f $Job.Succeeded, (Format-Size $Job.ProcessedBytes))
        } else {
            $summary = ('Cleanup complete. {0:N0} succeeded, {1:N0} failed, {2} processed.' -f $Job.Succeeded, $Job.Failed, (Format-Size $Job.ProcessedBytes))
        }
        if ($Job.OpenOutput -and -not $Job.DryRun -and $Job.Action -ne 'Recycle' -and (Test-Path -LiteralPath $Job.OutputDir)) {
            try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Job.OutputDir) } catch { }
        }
        $Job.Progress = 100
        $kind = if ($Job.Failed -gt 0) { 'Warning' } else { 'Success' }
        Finish-Job $summary $kind
        return 'Finished'
    }

    $record = $Job.Records[$Job.Index]
    $Job.Index++
    $path = [string]$record.Duplicate
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $Job.Failed++
        $message = "Missing file: $path"
        if ($Job.Verbose) { Write-Log $message 'Warning' }
        Write-FileLog ("[{0}] MISSING {1}" -f (Get-Date).ToString('s'), $path)
        return 'Continue'
    }

    try {
        if ($Job.Action -eq 'Recycle') {
            if ($Job.DryRun) {
                $line = "[DRY RUN] Recycle: $path"
                if ($Job.Verbose) { Write-Log $line 'Dry' }
                Write-FileLog $line
                $Job.Succeeded++
                $Job.ProcessedBytes += [long]$record.Size
            } else {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                $record.Removed = $true
                $Job.Succeeded++
                $Job.ProcessedBytes += [long]$record.Size
                $line = "[RECYCLED] $path"
                if ($Job.Verbose) { Write-Log $line 'Success' }
                Write-FileLog $line
            }
        } else {
            $targetDir = $Job.OutputDir
            if ($Job.Preserve) {
                $relative = Get-RelativePath -BasePath $Job.InputDir -TargetPath $path
                $relativeDir = [System.IO.Path]::GetDirectoryName($relative)
                if ($relativeDir) { $targetDir = Join-Path $Job.OutputDir $relativeDir }
            }
            $fileName = [System.IO.Path]::GetFileName($path)
            $destination = Get-NonConflictingPath -Directory $targetDir -FileName $fileName
            if ($Job.DryRun) {
                $line = "[DRY RUN] {0}: {1} -> {2}" -f $Job.Action, $path, $destination
                if ($Job.Verbose) { Write-Log $line 'Dry' }
                Write-FileLog $line
                $Job.Succeeded++
                $Job.ProcessedBytes += [long]$record.Size
            } else {
                if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }
                if ($Job.Action -eq 'Move') {
                    Move-Item -LiteralPath $path -Destination $destination -Force -ErrorAction Stop
                    $record.Removed = $true
                } else {
                    Copy-Item -LiteralPath $path -Destination $destination -Force -ErrorAction Stop
                }
                $Job.Succeeded++
                $Job.ProcessedBytes += [long]$record.Size
                $line = "[{0}] {1} -> {2}" -f $Job.Action.ToUpperInvariant(), $path, $destination
                if ($Job.Verbose) { Write-Log $line 'Success' }
                Write-FileLog $line
            }
        }
    } catch {
        $Job.Failed++
        $message = "Failed: $path - $($_.Exception.Message)"
        Write-Log $message 'Error'
        Write-FileLog $message
    }
    return 'Continue'
}

function Invoke-JobTimerTick {
    if ($script:InTick) { return }
    if (-not $script:Job) {
        if ($script:JobTimer) { $script:JobTimer.Stop() }
        return
    }
    $script:InTick = $true
    $stopwatch = $null
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        while ($stopwatch.ElapsedMilliseconds -lt 45) {
            if ($script:CancelRequested) {
                Finish-Job 'Operation cancelled by user.' 'Warning'
                return
            }
            $job = $script:Job
            if (-not $job) { return }
            if ($job.Type -eq 'Scan') { $state = Step-ScanJob $job } else { $state = Step-CleanJob $job }
            if ($state -eq 'Finished') { return }
            if ($state -eq 'Wait') { break }
        }
        Update-JobStatus
    } catch {
        $message = $_.Exception.Message
        try { Write-Log ('Operation stopped: ' + $message) 'Error' } catch { }
        Finish-Job ('Stopped due to an error: ' + $message) 'Error'
    } finally {
        $script:InTick = $false
        if ($stopwatch) { $stopwatch.Stop() }
    }
}

function Start-ScanJob {
    if ($script:Busy) { return }
    $inputDir = Get-CleanPath $script:TxtIn.Text
    if ([string]::IsNullOrWhiteSpace($inputDir) -or -not (Test-Path -LiteralPath $inputDir -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show('Please choose a valid input folder.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
        return
    }
    $extensions = @(Get-SelectedExtensions)
    if ($extensions.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select at least one file type to scan.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
        return
    }
    $algorithm = [string]$script:CmbHash.SelectedItem
    if ($algorithm -notin @('MD5', 'SHA1', 'SHA256')) { $algorithm = 'SHA256' }
    $matchMode = [string]$script:CmbMatch.SelectedItem
    $scope = Get-SelectedScope
    $scopeAll = $scope -eq 'All'
    Save-Settings

    $script:LogBox.Clear()
    $script:Grid.RowCount = 0
    $script:GridRecords = [System.Collections.Generic.List[object]]::new()
    $script:LastScanInput = $null
    $script:LastScanScope = $null
    $script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $script:Progress.MarqueeAnimationSpeed = 25
    $script:Progress.Value = 0
    $script:LblStats.Text = ''

    $extensionSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($extension in $extensions) { [void]$extensionSet.Add($extension) }

    $excludeDir = $null
    $outputDir = Get-CleanPath $script:TxtOut.Text
    if ($outputDir -and (Test-Path -LiteralPath $outputDir -PathType Container)) {
        try {
            $inputFull = [System.IO.Path]::GetFullPath($inputDir).TrimEnd('\')
            $outputFull = [System.IO.Path]::GetFullPath($outputDir).TrimEnd('\')
            if ($outputFull.StartsWith(($inputFull + '\'), [System.StringComparison]::OrdinalIgnoreCase)) { $excludeDir = $outputFull }
        } catch { }
    }

    $stack = [System.Collections.Generic.Stack[string]]::new()
    $stack.Push([System.IO.Path]::GetFullPath($inputDir))

    $job = [pscustomobject]@{
        Type = 'Scan'
        Phase = 'Collect'
        InputDir = [System.IO.Path]::GetFullPath($inputDir)
        OutputDir = $outputDir
        ExcludeDir = $excludeDir
        Files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        HashMap = @{}
        Keys = $null
        KeyIndex = 0
        DirStack = $stack
        ExtensionSet = $extensionSet
        Algorithm = $algorithm
        MatchMode = $matchMode
        MatchByName = ($matchMode -like 'Same name*')
        Scope = $scope
        ScopeAll = $scopeAll
        Recurse = [bool]$script:ChkRecurse.Checked
        IncludeHidden = [bool]$script:ChkHidden.Checked
        MaxFiles = [int]$script:NumMaxFiles.Value
        Total = 0
        Processed = 0
        RecordCount = 0
        RecordBytes = [long]0
        DirectoriesScanned = 0
        AccessErrors = 0
        HashErrors = 0
        LimitReached = $false
        Progress = 0
    }

    $script:Job = $job
    Set-Busy $true
    $script:LblStatus.Text = 'Starting scan...'
    Write-Log ('Starting scan: {0}' -f $job.InputDir) 'Info'
    Write-Log ('Scope: {0} | Mode: {1}' -f $(if ($scopeAll) { 'All application components' } else { 'Duplicates only' }), $(if ($scopeAll) { 'component index' } else { $algorithm })) 'Muted'
    $script:JobTimer.Start()
}

function Start-CleanJob {
    if ($script:Busy) { return }
    if ($script:GridRecords.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('There is nothing to process. Run a scan first.', 'Component Optimizer', 'OK', 'Information') | Out-Null
        return
    }
    $inputDir = Get-CleanPath $script:TxtIn.Text
    if ([string]::IsNullOrWhiteSpace($inputDir) -or -not (Test-Path -LiteralPath $inputDir -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show('Please choose a valid input folder.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        if ($script:LastScanInput -and ([System.IO.Path]::GetFullPath($inputDir) -ne $script:LastScanInput)) {
            [System.Windows.Forms.MessageBox]::Show('The current results are from a different input folder. Please scan again.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
            return
        }
    } catch { }
    $scope = Get-SelectedScope
    if ($script:LastScanScope -and $scope -ne $script:LastScanScope) {
        [System.Windows.Forms.MessageBox]::Show('The current results use a different scope. Please scan again.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
        return
    }

    $selected = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $script:GridRecords) {
        if ($record.Process) { $selected.Add($record) }
    }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No files are checked for processing.', 'Component Optimizer', 'OK', 'Information') | Out-Null
        return
    }

    $action = Get-SelectedAction
    $outputDir = Get-CleanPath $script:TxtOut.Text
    $dryRun = [bool]$script:ChkDryRun.Checked

    if ($action -ne 'Recycle') {
        if ([string]::IsNullOrWhiteSpace($outputDir)) {
            [System.Windows.Forms.MessageBox]::Show('Please choose an output folder.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
            return
        }
        if (-not (Test-Path -LiteralPath $outputDir -PathType Container) -and -not $dryRun) {
            $answer = [System.Windows.Forms.MessageBox]::Show('The output folder does not exist. Create it now?', 'Component Optimizer', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        try {
            $inputFull = [System.IO.Path]::GetFullPath($inputDir).TrimEnd('\')
            $outputFull = [System.IO.Path]::GetFullPath($outputDir).TrimEnd('\')
            if ($inputFull -eq $outputFull) {
                [System.Windows.Forms.MessageBox]::Show('The output folder must be different from the input folder.', 'Component Optimizer', 'OK', 'Warning') | Out-Null
                return
            }
        } catch { }
    }

    if ($script:ChkConfirm.Checked) {
        $verb = switch ($action) {
            'Copy' { 'copy' }
            'Recycle' { 'send to the Recycle Bin' }
            default { 'move' }
        }
        $suffix = if ($dryRun) { ' This is a dry run; no files will be changed.' } else { '' }
        $text = ('About to {0} {1:N0} file(s).{2}' -f $verb, $selected.Count, $suffix)
        $answer = [System.Windows.Forms.MessageBox]::Show($text, 'Confirm operation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    Save-Settings
    $script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    $script:Progress.Value = 0
    $job = [pscustomobject]@{
        Type = 'Clean'
        InputDir = [System.IO.Path]::GetFullPath($inputDir)
        OutputDir = if ($outputDir) { $outputDir } else { '' }
        Action = $action
        Preserve = [bool]$script:ChkPreserve.Checked
        DryRun = $dryRun
        Verbose = [bool]$script:ChkVerbose.Checked
        OpenOutput = [bool]$script:ChkOpenOutput.Checked
        Records = $selected
        Index = 0
        Succeeded = 0
        Failed = 0
        ProcessedBytes = [long]0
        Progress = 0
    }
    $script:Job = $job
    Set-Busy $true
    $script:LblStatus.Text = 'Starting operation...'
    Write-Log ('Starting {0} operation for {1:N0} file(s){2}.' -f $action, $selected.Count, $(if ($dryRun) { ' (dry run)' } else { '' })) 'Info'
    $script:JobTimer.Start()
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Component Optimizer'
$form.ClientSize = [System.Drawing.Size]::new(1000, 770)
$form.MinimumSize = [System.Drawing.Size]::new(900, 700)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = [System.Drawing.Color]::White
$form.Font = $script:UIFont
$form.MaximizeBox = $true

$header = New-Object System.Windows.Forms.Panel
$header.Dock = [System.Windows.Forms.DockStyle]::Top
$header.Height = 88
$header.BackColor = $script:HeaderDark
$header.Add_Paint({
    param($sender, $eventArgs)
    $rectangle = $sender.ClientRectangle
    $brush = [System.Drawing.Drawing2D.LinearGradientBrush]::new($rectangle, $script:HeaderDark, $script:AccentLight, [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
    $eventArgs.Graphics.FillRectangle($brush, $rectangle)
    $brush.Dispose()
    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(70, 255, 255, 255), [float]1)
    $eventArgs.Graphics.DrawLine($pen, 0, ($sender.Height - 1), $sender.Width, ($sender.Height - 1))
    $pen.Dispose()
})
$titleLabel = New-Label 'Component Optimizer' 24 12 650 38 $script:HeaderFont ([System.Drawing.Color]::White)
$subtitleLabel = New-Label 'Optimize, deduplicate and relocate application components' 26 51 700 22 $script:SubFont ([System.Drawing.Color]::FromArgb(225, 245, 235))
$versionLabel = New-Label 'v2.0' 880 18 90 22 $script:SmallFont ([System.Drawing.Color]::FromArgb(220, 255, 255, 255))
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$header.Controls.AddRange(@($titleLabel, $subtitleLabel, $versionLabel))

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = [System.Drawing.Point]::new(12, 98)
$tabs.Size = [System.Drawing.Size]::new(976, 660)
$tabs.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabs.Font = $script:UIFont

$tabMain = New-Object System.Windows.Forms.TabPage
$tabMain.Text = 'Search & Move'
$tabMain.BackColor = [System.Drawing.Color]::White
$tabMain.Padding = [System.Windows.Forms.Padding]::new(0)

$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = 'Settings'
$tabSettings.BackColor = [System.Drawing.Color]::White
$tabSettings.Padding = [System.Windows.Forms.Padding]::new(0)

$tabInstall = New-Object System.Windows.Forms.TabPage
$tabInstall.Text = 'Install'
$tabInstall.BackColor = [System.Drawing.Color]::White
$tabInstall.Padding = [System.Windows.Forms.Padding]::new(0)

$tabs.TabPages.AddRange(@($tabMain, $tabSettings, $tabInstall))

$groupFolders = New-Group 'Folders' 14 12 948 110
$labelInput = New-Label 'Input folder:' 16 28 100 20 $script:UIFont
$script:TxtIn = New-TextBox 120 25 700 24
$script:BtnIn = New-FlatButton 'Browse...' 828 23 104 28 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$labelOutput = New-Label 'Output folder:' 16 66 100 20 $script:UIFont
$script:TxtOut = New-TextBox 120 63 700 24
$script:BtnOut = New-FlatButton 'Browse...' 828 61 104 28 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$groupFolders.Controls.AddRange(@($labelInput, $script:TxtIn, $script:BtnIn, $labelOutput, $script:TxtOut, $script:BtnOut))

$groupOptions = New-Group 'Action and scope' 14 130 948 112
$script:RadMove = New-Radio 'Move files' 18 22 110 22 $true
$script:RadCopy = New-Radio 'Copy files' 140 22 110 22
$script:RadRecycle = New-Radio 'Send to Recycle Bin' 260 22 170 22
$hintAction = New-Label 'The same action is used for every checked row below.' 450 24 470 20 $script:SmallFont $script:MutedColor
$script:RadScopeAll = New-Radio 'All application components (search and move)' 18 48 370 22 $true
$script:RadScopeDup = New-Radio 'Duplicates only' 400 48 160 22
$script:ChkRecurse = New-CheckBox 'Include subfolders' 18 76 150 22 $true
$script:ChkHidden = New-CheckBox 'Include hidden/system' 180 76 170 22
$script:ChkPreserve = New-CheckBox 'Preserve folder structure' 360 76 180 22
$script:ChkDryRun = New-CheckBox 'Dry run' 550 76 90 22
$script:ChkConfirm = New-CheckBox 'Confirm before changes' 650 76 170 22 $true
$groupOptions.Controls.AddRange(@($script:RadMove, $script:RadCopy, $script:RadRecycle, $hintAction, $script:RadScopeAll, $script:RadScopeDup, $script:ChkRecurse, $script:ChkHidden, $script:ChkPreserve, $script:ChkDryRun, $script:ChkConfirm))

$script:BtnScan = New-FlatButton 'Scan for components' 14 250 170 38 $script:Accent
$script:BtnClean = New-FlatButton 'Run action on checked' 192 250 210 38 $script:AccentLight
$script:BtnNew = New-FlatButton 'New operation' 410 250 150 38 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:BtnCheckAll = New-FlatButton 'Check all' 568 250 80 38 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor $script:SmallFont
$script:BtnCheckNone = New-FlatButton 'Check none' 654 250 90 38 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor $script:SmallFont
$script:BtnCancel = New-FlatButton 'Cancel' 752 250 200 38 $script:Danger

$script:Progress = New-Object System.Windows.Forms.ProgressBar
$script:Progress.Location = [System.Drawing.Point]::new(14, 296)
$script:Progress.Size = [System.Drawing.Size]::new(948, 16)
$script:Progress.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks

$script:LblStatus = New-Label 'Ready.' 14 316 650 20 $script:UIFont
$script:LblStats = New-Label '' 680 316 282 20 $script:UIFont
$script:LblStats.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight

$script:Grid = New-Object System.Windows.Forms.DataGridView
$script:Grid.Location = [System.Drawing.Point]::new(14, 342)
$script:Grid.Size = [System.Drawing.Size]::new(948, 190)
$script:Grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:Grid.BackgroundColor = [System.Drawing.Color]::White
$script:Grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:Grid.RowHeadersVisible = $false
$script:Grid.AllowUserToAddRows = $false
$script:Grid.AllowUserToDeleteRows = $false
$script:Grid.AllowUserToResizeRows = $false
$script:Grid.MultiSelect = $true
$script:Grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$script:Grid.ReadOnly = $true
$script:Grid.VirtualMode = $true
$script:Grid.RowCount = 0
$script:Grid.EnableHeadersVisualStyles = $false
$script:Grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
$script:Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$script:Grid.ColumnHeadersDefaultCellStyle.Font = $script:BoldFont
$script:Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
$script:Grid.ColumnHeadersHeight = 30
$script:Grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 249, 251)
$script:Grid.GridColor = [System.Drawing.Color]::FromArgb(222, 228, 234)
$script:Grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(209, 242, 222)
$script:Grid.DefaultCellStyle.SelectionForeColor = $script:TextColor

$columnProcess = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$columnProcess.Name = 'Process'
$columnProcess.HeaderText = 'Clean'
$columnProcess.Width = 50
$columnProcess.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$columnMaster = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnMaster.Name = 'Master'
$columnMaster.HeaderText = 'Original (kept)'
$columnMaster.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$columnMaster.FillWeight = 40
$columnMaster.MinimumWidth = 160
$columnMaster.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$columnDuplicate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnDuplicate.Name = 'Duplicate'
$columnDuplicate.HeaderText = 'Duplicate to process'
$columnDuplicate.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$columnDuplicate.FillWeight = 60
$columnDuplicate.MinimumWidth = 180
$columnDuplicate.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$columnSize = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnSize.Name = 'Size'
$columnSize.HeaderText = 'Size'
$columnSize.Width = 85
$columnSize.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$columnSize.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
$columnMatch = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$columnMatch.Name = 'Match'
$columnMatch.HeaderText = 'Match'
$columnMatch.Width = 100
$columnMatch.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
[void]$script:Grid.Columns.Add($columnProcess)
[void]$script:Grid.Columns.Add($columnMaster)
[void]$script:Grid.Columns.Add($columnDuplicate)
[void]$script:Grid.Columns.Add($columnSize)
[void]$script:Grid.Columns.Add($columnMatch)

$labelLog = New-Label 'Execution log:' 14 540 130 18 $script:UIFont
$script:LogBox = New-Object System.Windows.Forms.RichTextBox
$script:LogBox.Location = [System.Drawing.Point]::new(14, 558)
$script:LogBox.Size = [System.Drawing.Size]::new(948, 60)
$script:LogBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:LogBox.BackColor = [System.Drawing.Color]::FromArgb(16, 24, 32)
$script:LogBox.ForeColor = [System.Drawing.Color]::LightGreen
$script:LogBox.Font = [System.Drawing.Font]::new('Consolas', [float]9)
$script:LogBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:LogBox.ReadOnly = $true
$script:LogBox.WordWrap = $false
$script:LogBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$script:LogBox.DetectUrls = $false

$tabMain.Controls.AddRange(@($groupFolders, $groupOptions, $script:BtnScan, $script:BtnClean, $script:BtnNew, $script:BtnCheckAll, $script:BtnCheckNone, $script:BtnCancel, $script:Progress, $script:LblStatus, $script:LblStats, $script:Grid, $labelLog, $script:LogBox))

$groupTypes = New-Group 'Application component types' 14 12 470 230
$script:ChkExtensions = New-Object System.Windows.Forms.CheckedListBox
$script:ChkExtensions.Location = [System.Drawing.Point]::new(16, 26)
$script:ChkExtensions.Size = [System.Drawing.Size]::new(438, 146)
$script:ChkExtensions.CheckOnClick = $true
$script:ChkExtensions.Font = $script:UIFont
$script:ChkExtensions.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
foreach ($extension in @('.exe', '.dll', '.msi', '.msix', '.appx', '.bat', '.cmd', '.com', '.ps1', '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.reg', '.ocx', '.cpl', '.scr', '.sys', '.config', '.manifest', '.cab')) {
    [void]$script:ChkExtensions.Items.Add($extension)
}
$labelTypesHint = New-Label 'Application component and installer types only.' 16 176 438 16 $script:SmallFont $script:MutedColor
$script:BtnSelectAll = New-FlatButton 'All' 16 196 100 26 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:BtnSelectNone = New-FlatButton 'None' 124 196 100 26 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:BtnSelectDefault = New-FlatButton 'Recommended' 232 196 120 26 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$groupTypes.Controls.AddRange(@($script:ChkExtensions, $labelTypesHint, $script:BtnSelectAll, $script:BtnSelectNone, $script:BtnSelectDefault))

$groupMatch = New-Group 'Matching' 496 12 466 230
$labelHash = New-Label 'Hash algorithm:' 16 34 130 20 $script:UIFont
$script:CmbHash = New-Combo 160 31 250 24
foreach ($item in @('MD5', 'SHA1', 'SHA256')) { [void]$script:CmbHash.Items.Add($item) }
$labelMatch = New-Label 'Duplicate match mode:' 16 74 150 20 $script:UIFont
$script:CmbMatch = New-Combo 170 71 240 24
foreach ($item in @('Exact content (hash)', 'Same name and size')) { [void]$script:CmbMatch.Items.Add($item) }
$labelCustom = New-Label 'Custom extensions (comma separated, dots optional):' 16 112 380 20 $script:UIFont
$script:TxtCustomExt = New-TextBox 16 134 434 24
$labelTip = New-Label 'Hash mode verifies byte-for-byte equality. Name + size is faster but less strict.' 16 168 434 40 $script:SmallFont $script:MutedColor
$groupMatch.Controls.AddRange(@($labelHash, $script:CmbHash, $labelMatch, $script:CmbMatch, $labelCustom, $script:TxtCustomExt, $labelTip))

$groupBehavior = New-Group 'Logging and safety' 14 252 948 140
$script:ChkLogToFile = New-CheckBox 'Write log file' 16 28 120 22
$script:TxtLogPath = New-TextBox 140 25 660 24
$script:BtnLogBrowse = New-FlatButton 'Browse...' 808 23 120 28 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:ChkOpenOutput = New-CheckBox 'Open output folder when finished' 16 62 250 22
$script:ChkVerbose = New-CheckBox 'Verbose per-file log' 280 62 180 22
$labelLimit = New-Label 'Safety file limit:' 16 98 110 20 $script:UIFont
$script:NumMaxFiles = New-Object System.Windows.Forms.NumericUpDown
$script:NumMaxFiles.Location = [System.Drawing.Point]::new(140, 95)
$script:NumMaxFiles.Size = [System.Drawing.Size]::new(110, 24)
$script:NumMaxFiles.Minimum = [decimal]1000
$script:NumMaxFiles.Maximum = [decimal]500000
$script:NumMaxFiles.Increment = [decimal]10000
$script:NumMaxFiles.Value = [decimal]100000
$script:NumMaxFiles.ThousandsSeparator = $true
$script:NumMaxFiles.Font = $script:UIFont
$labelLimitTip = New-Label 'Prevents runaway memory use on huge folder trees.' 270 98 400 20 $script:SmallFont $script:MutedColor
$groupBehavior.Controls.AddRange(@($script:ChkLogToFile, $script:TxtLogPath, $script:BtnLogBrowse, $script:ChkOpenOutput, $script:ChkVerbose, $labelLimit, $script:NumMaxFiles, $labelLimitTip))

$labelSettingsHint = New-Label 'Settings are saved automatically when you scan, run an operation, or close the app.' 14 408 700 22 $script:SmallFont $script:MutedColor
$script:BtnSaveSettings = New-FlatButton 'Save settings' 14 438 180 36 $script:Accent
$tabSettings.Controls.AddRange(@($groupTypes, $groupMatch, $groupBehavior, $labelSettingsHint, $script:BtnSaveSettings))

$groupInstall = New-Group 'Installation' 14 12 948 160
$labelInstallPath = New-Label 'Install location:' 16 22 110 20 $script:UIFont
$script:TxtInstallPath = New-TextBox 128 19 620 24
$script:TxtInstallPath.ReadOnly = $true
$script:BtnInstall = New-FlatButton 'Install + Desktop Shortcut' 16 80 220 36 $script:Accent
$script:BtnDesktop = New-FlatButton 'Desktop Shortcut' 246 80 190 36 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:BtnStartMenu = New-FlatButton 'Start Menu Shortcut' 446 80 170 36 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:BtnOpenFolder = New-FlatButton 'Open Folder' 626 80 150 36 ([System.Drawing.Color]::FromArgb(236, 240, 241)) $script:TextColor
$script:ChkLaunchAfter = New-CheckBox 'Launch installed app after install' 16 126 240 22
$script:BtnUninstall = New-FlatButton 'Uninstall' 626 122 150 32 $script:Danger
$script:PicIcon = New-Object System.Windows.Forms.PictureBox
$script:PicIcon.Location = [System.Drawing.Point]::new(808, 16)
$script:PicIcon.Size = [System.Drawing.Size]::new(120, 120)
$script:PicIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$script:PicIcon.BackColor = [System.Drawing.Color]::White
$script:PicIcon.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupInstall.Controls.AddRange(@($labelInstallPath, $script:TxtInstallPath, $script:BtnInstall, $script:BtnDesktop, $script:BtnStartMenu, $script:BtnOpenFolder, $script:ChkLaunchAfter, $script:BtnUninstall, $script:PicIcon))

$script:InstallLog = New-Object System.Windows.Forms.RichTextBox
$script:InstallLog.Location = [System.Drawing.Point]::new(14, 184)
$script:InstallLog.Size = [System.Drawing.Size]::new(948, 432)
$script:InstallLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$script:InstallLog.BackColor = [System.Drawing.Color]::FromArgb(16, 24, 32)
$script:InstallLog.ForeColor = [System.Drawing.Color]::LightGreen
$script:InstallLog.Font = [System.Drawing.Font]::new('Consolas', [float]9)
$script:InstallLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:InstallLog.ReadOnly = $true
$script:InstallLog.WordWrap = $true
$script:InstallLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$tabInstall.Controls.AddRange(@($groupInstall, $script:InstallLog))

$form.Controls.Add($header)
$form.Controls.Add($tabs)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($script:RadScopeAll, 'Process every file that matches the selected extensions, not only duplicates.')
$toolTip.SetToolTip($script:ChkDryRun, 'Preview the operation without changing any files.')
$toolTip.SetToolTip($script:ChkPreserve, 'Recreate the source folder structure inside the output folder.')
$toolTip.SetToolTip($script:NumMaxFiles, 'Upper limit on the number of files held in memory during a scan.')
$toolTip.SetToolTip($script:BtnNew, 'Clear results and prepare the app for another scan.')

$script:BrowseFolder = {
    param([string]$Title)
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title
    $dialog.ShowNewFolderButton = $true
    $result = $dialog.ShowDialog()
    $path = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $path = $dialog.SelectedPath }
    $dialog.Dispose()
    return $path
}

$script:BtnIn.Add_Click({
    $path = & $script:BrowseFolder 'Select the folder to scan'
    if ($path) { $script:TxtIn.Text = $path; Save-Settings }
})
$script:BtnOut.Add_Click({
    $path = & $script:BrowseFolder 'Select the output folder'
    if ($path) { $script:TxtOut.Text = $path; Save-Settings }
})
$script:BtnLogBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Choose log file'
    $dialog.Filter = 'Log files (*.log)|*.log|All files (*.*)|*.*'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:TxtLogPath.Text = $dialog.FileName
        Save-Settings
    }
    $dialog.Dispose()
})
$script:RadRecycle.Add_CheckedChanged({ Update-ActionUi })
$script:RadMove.Add_CheckedChanged({ Update-ActionUi })
$script:RadCopy.Add_CheckedChanged({ Update-ActionUi })
$script:BtnScan.Add_Click({ Start-ScanJob })
$script:BtnClean.Add_Click({ Start-CleanJob })
$script:BtnNew.Add_Click({ Reset-ForNextOperation })
$script:BtnCancel.Add_Click({ $script:CancelRequested = $true; $script:LblStatus.Text = 'Cancelling...' })
$script:BtnCheckAll.Add_Click({
    foreach ($record in $script:GridRecords) { $record.Process = $true }
    $script:Grid.Invalidate()
    Update-GridSummary
})
$script:BtnCheckNone.Add_Click({
    foreach ($record in $script:GridRecords) { $record.Process = $false }
    $script:Grid.Invalidate()
    Update-GridSummary
})
$script:BtnSelectAll.Add_Click({
    for ($i = 0; $i -lt $script:ChkExtensions.Items.Count; $i++) { $script:ChkExtensions.SetItemChecked($i, $true) }
    Save-Settings
})
$script:BtnSelectNone.Add_Click({
    for ($i = 0; $i -lt $script:ChkExtensions.Items.Count; $i++) { $script:ChkExtensions.SetItemChecked($i, $false) }
    Save-Settings
})
$script:BtnSelectDefault.Add_Click({
    $defaults = @('.exe', '.dll', '.msi', '.msix', '.appx', '.bat', '.cmd', '.ps1', '.vbs', '.reg', '.ocx', '.cpl', '.config', '.manifest')
    for ($i = 0; $i -lt $script:ChkExtensions.Items.Count; $i++) {
        $item = [string]$script:ChkExtensions.Items[$i]
        $script:ChkExtensions.SetItemChecked($i, $defaults -contains $item)
    }
    Save-Settings
})
$script:BtnSaveSettings.Add_Click({
    Save-Settings
    Write-Log 'Settings saved.' 'Success'
})
$script:BtnInstall.Add_Click({
    try {
        $installDir = Install-Application
        $shortcut = New-DesktopShortcut
        Write-InstallLog ('Installed to {0}' -f $installDir) 'Success'
        Write-InstallLog ('Desktop shortcut created: {0}' -f $shortcut) 'Success'
        Write-InstallLog 'Icon created as ComponentDeduplicator.ico' 'Info'
        if ($script:ChkLaunchAfter.Checked) {
            $powerShellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $installedScript = Join-Path $installDir 'ComponentDeduplicator.ps1'
            Start-Process -FilePath $powerShellPath -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$installedScript`""
            Write-InstallLog 'Installed copy launched.' 'Info'
        }
    } catch {
        Write-InstallLog ('Install failed: ' + $_.Exception.Message) 'Error'
    }
})
$script:BtnDesktop.Add_Click({
    try {
        $installDir = Install-Application
        $shortcut = New-DesktopShortcut
        Write-InstallLog ('Desktop shortcut created: {0}' -f $shortcut) 'Success'
    } catch {
        Write-InstallLog ('Shortcut creation failed: ' + $_.Exception.Message) 'Error'
    }
})
$script:BtnStartMenu.Add_Click({
    try {
        $installDir = Install-Application
        $shortcut = New-StartMenuShortcut
        Write-InstallLog ('Start Menu shortcut created: {0}' -f $shortcut) 'Success'
    } catch {
        Write-InstallLog ('Shortcut creation failed: ' + $_.Exception.Message) 'Error'
    }
})
$script:BtnOpenFolder.Add_Click({
    try {
        $installDir = Get-InstallDir
        if (-not (Test-Path -LiteralPath $installDir)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $installDir)
    } catch {
        Write-InstallLog ('Could not open install folder: ' + $_.Exception.Message) 'Error'
    }
})
$script:BtnUninstall.Add_Click({
    $answer = [System.Windows.Forms.MessageBox]::Show('Remove the desktop and Start Menu shortcuts, plus the installed application files?', 'Uninstall', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        $message = Uninstall-Application
        Write-InstallLog $message 'Success'
    } catch {
        Write-InstallLog ('Uninstall failed: ' + $_.Exception.Message) 'Error'
    }
})

$script:Grid.Add_CellValueNeeded({
    param($sender, $eventArgs)
    if ($eventArgs.RowIndex -lt 0 -or $eventArgs.RowIndex -ge $script:GridRecords.Count) { return }
    $record = $script:GridRecords[$eventArgs.RowIndex]
    switch ($eventArgs.ColumnIndex) {
        0 { $eventArgs.Value = [bool]$record.Process }
        1 { $eventArgs.Value = [string]$record.Master }
        2 { $eventArgs.Value = [string]$record.Duplicate }
        3 { $eventArgs.Value = Format-Size ([long]$record.Size) }
        4 { $eventArgs.Value = [string]$record.Match }
    }
})
$script:Grid.Add_CellContentClick({
    param($sender, $eventArgs)
    if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -ne 0) { return }
    if ($eventArgs.RowIndex -ge $script:GridRecords.Count) { return }
    $record = $script:GridRecords[$eventArgs.RowIndex]
    $record.Process = -not $record.Process
    $script:Grid.InvalidateCell(0, $eventArgs.RowIndex)
    Update-GridSummary
})
$script:Grid.Add_CellDoubleClick({
    param($sender, $eventArgs)
    if ($eventArgs.RowIndex -lt 0 -or $eventArgs.ColumnIndex -eq 0) { return }
    if ($eventArgs.RowIndex -ge $script:GridRecords.Count) { return }
    $path = [string]$script:GridRecords[$eventArgs.RowIndex].Duplicate
    if ($path -and (Test-Path -LiteralPath $path)) {
        try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $path) } catch { }
    }
})

$script:JobTimer = New-Object System.Windows.Forms.Timer
$script:JobTimer.Interval = 30
$script:JobTimer.Add_Tick({ Invoke-JobTimerTick })

$script:Settings = Load-Settings
$script:TxtIn.Text = [string]$script:Settings.InputDir
$script:TxtOut.Text = [string]$script:Settings.OutputDir
switch ([string]$script:Settings.Action) {
    'Copy' { $script:RadCopy.Checked = $true }
    'Recycle' { $script:RadRecycle.Checked = $true }
    default { $script:RadMove.Checked = $true }
}
switch ([string]$script:Settings.Scope) {
    'All' { $script:RadScopeAll.Checked = $true }
    default { $script:RadScopeDup.Checked = $true }
}
Set-ComboSelection $script:CmbHash ([string]$script:Settings.HashAlgorithm)
Set-ComboSelection $script:CmbMatch ([string]$script:Settings.MatchMode)
$savedExtensions = @($script:Settings.Extensions)
for ($i = 0; $i -lt $script:ChkExtensions.Items.Count; $i++) {
    $item = [string]$script:ChkExtensions.Items[$i]
    $script:ChkExtensions.SetItemChecked($i, ($savedExtensions -contains $item))
}
$script:TxtCustomExt.Text = [string]$script:Settings.CustomExtensions
$script:ChkRecurse.Checked = [bool]$script:Settings.IncludeSubfolders
$script:ChkHidden.Checked = [bool]$script:Settings.IncludeHidden
$script:ChkPreserve.Checked = [bool]$script:Settings.PreserveStructure
$script:ChkDryRun.Checked = [bool]$script:Settings.DryRun
$script:ChkConfirm.Checked = [bool]$script:Settings.ConfirmActions
$script:ChkLogToFile.Checked = [bool]$script:Settings.LogToFile
$script:TxtLogPath.Text = [string]$script:Settings.LogPath
$script:ChkOpenOutput.Checked = [bool]$script:Settings.OpenOutputAfter
$script:ChkVerbose.Checked = [bool]$script:Settings.VerboseLog
$maxFiles = [decimal]$script:Settings.MaxScanFiles
if ($maxFiles -lt $script:NumMaxFiles.Minimum) { $maxFiles = $script:NumMaxFiles.Minimum }
if ($maxFiles -gt $script:NumMaxFiles.Maximum) { $maxFiles = $script:NumMaxFiles.Maximum }
$script:NumMaxFiles.Value = $maxFiles
$script:TxtInstallPath.Text = Get-InstallDir
Update-ActionUi
Update-GridSummary

try {
    $script:PicIcon.Image = New-AppIconBitmap 112
} catch { }

try {
    $script:TempIconPath = Join-Path ([System.IO.Path]::GetTempPath()) 'ComponentDeduplicator.ico'
    New-AppIcon -Path $script:TempIconPath
    $form.Icon = [System.Drawing.Icon]::new($script:TempIconPath)
} catch { }

$form.Add_FormClosing({
    param($sender, $eventArgs)
    if ($script:Job) {
        $answer = [System.Windows.Forms.MessageBox]::Show('An operation is still running. Close anyway?', 'Component Optimizer', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $eventArgs.Cancel = $true
            return
        }
    }
    $script:CancelRequested = $true
    try { Save-Settings } catch { }
    if ($script:JobTimer) { $script:JobTimer.Stop() }
    if ($script:LogWriter) {
        try { $script:LogWriter.Flush(); $script:LogWriter.Dispose() } catch { }
        $script:LogWriter = $null
    }
})

try {
    [System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $eventArgs)
        try { Write-Log ('UI error: ' + $eventArgs.Exception.Message) 'Error' } catch { }
    })
} catch { }

Write-Log 'Ready. Choose a scope, select your folders, then click Scan for components.' 'Info'

try {
    [System.Windows.Forms.Application]::Run($form)
} catch {
    [System.Windows.Forms.MessageBox]::Show(('The application stopped unexpectedly:' + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message), 'Component Optimizer', 'OK', 'Error') | Out-Null
}

