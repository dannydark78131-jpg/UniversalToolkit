# =====================================================================
#  DGs File Finder  (DGsFileSearch.ps1)  v2.0.0
#  Single-window GUI: scan a ZIP archive OR a folder (all subfolders),
#  list files/folders, filter, inspect results, export reports and
#  extract files (all / by type / hand-picked).
#
#  v2.0.0 highlights:
#   - Results grid with live filter, file-type grid, log tabs
#   - Include/Exclude wildcard file filters (GUI + command line)
#   - Cancel button, progress bar, saved settings
#   - No-GUI automation mode (-NoGui)
#   - Reparse-point safe scan, per-directory error isolation
#   - Zip-slip protection, long-name safe extraction
#
#  Pure ASCII, saved UTF-8 with BOM (PowerShell 5.1 safe).
# =====================================================================
[CmdletBinding()]
param(
    [string]$ZipPath = "",
    [string]$FolderPath = "",
    [string]$OutputDir = "",
    [string[]]$ExportFormat = @(),
    [switch]$Extract,
    [ValidateSet("All","ByType","Single","None")]
    [string]$ExtractMode = "All",
    [string[]]$ExtractType = @(),
    [string[]]$ExtractFile = @(),
    [string]$Include = "",
    [string]$Exclude = "",
    [switch]$Open,
    [switch]$NoGui,
    [switch]$SmokeTest
)

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"
$appName = "DGs File Finder"
$appVersion = "2.0.0"

$script:entries = $null
$script:files = $null
$script:typeGroups = $null
$script:sourceKind = ""
$script:sourcePath = ""
$script:baseName = ""
$script:zipArchive = $null
$script:lastExports = @{}
$script:logBox = $null
$script:statusLabel = $null
$script:resultsView = $null
$script:typeView = $null
$script:displayCap = 50000
$script:filterInclude = ""
$script:filterExclude = ""
$script:filteredOut = 0
$script:cancelRequested = $false
$script:tick = 0
$script:isBusy = $false
$script:configPath = Join-Path $env:APPDATA "DGsFileSearch\settings.json"
$script:reservedNames = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')

trap {
    $msg = "Unexpected error: $($_.Exception.Message)"
    Write-Host $msg -ForegroundColor Red
    if (-not $NoGui) {
        try { [System.Windows.Forms.MessageBox]::Show($msg, "$appName - Error", "OK", "Error") | Out-Null } catch { }
    }
    exit 1
}

# ---------- helpers ----------
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-SafeRelPath {
    param([string]$Rel)
    if ([string]::IsNullOrEmpty($Rel)) { return "" }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $parts = @()
    foreach ($segment in ($Rel -split '[\\/]')) {
        if ($segment -eq '' -or $segment -eq '.') { continue }
        $p = $segment
        foreach ($c in $invalid) { $p = $p.Replace([string]$c, '_') }
        $p = $p.Replace([string][char]0, '_')
        if ($p -eq '..') { $p = '_' }
        $p = $p.TrimEnd([char[]]@(' ', '.'))
        if ($p.Length -gt 120) { $p = $p.Substring(0, 120) }
        if ($p -eq '') { continue }
        if ($script:reservedNames -contains $p.ToUpperInvariant()) { $p = '_' + $p }
        $parts += $p
    }
    return ($parts -join '\')
}

function Test-UnderDir {
    param([string]$Child, [string]$Root)
    if ([string]::IsNullOrEmpty($Root)) { return $false }
    $r = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $c = [System.IO.Path]::GetFullPath($Child)
    return $c.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-InOutputTree {
    param([string]$FullPath)
    if ([string]::IsNullOrEmpty($script:outputDir)) { return $false }
    $full = [System.IO.Path]::GetFullPath($FullPath).TrimEnd('\')
    $root = [System.IO.Path]::GetFullPath($script:outputDir).TrimEnd('\')
    if ($full.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return (Test-UnderDir $FullPath $script:outputDir)
}

function Test-MaskList {
    param([string]$Masks, [string]$Name, [string]$Path)
    if ([string]::IsNullOrWhiteSpace($Masks)) { return $false }
    $norm = ([string]$Path) -replace '/','\'
    foreach ($m in ($Masks -split '[;,]')) {
        $p = $m.Trim()
        if ($p -eq '') { continue }
        if ($Name -like $p) { return $true }
        if ($Path -like $p) { return $true }
        if ($norm -like $p) { return $true }
    }
    return $false
}

function Test-FileFilter {
    param([string]$Name, [string]$Path)
    if (-not [string]::IsNullOrWhiteSpace($script:filterInclude)) {
        if (-not (Test-MaskList $script:filterInclude $Name $Path)) { return $false }
    }
    if (-not [string]::IsNullOrWhiteSpace($script:filterExclude)) {
        if (Test-MaskList $script:filterExclude $Name $Path) { return $false }
    }
    return $true
}

function HtmlEnc {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function XmlEnc {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Normalize-Formats {
    param([string[]]$Formats)
    $valid = @('CSV','JSON','TXT','HTML','XML')
    $out = New-Object 'System.Collections.Generic.List[string]'
    foreach ($f in @($Formats)) {
        foreach ($part in ([string]$f -split ',')) {
            $u = $part.Trim().ToUpperInvariant()
            if ($u -eq '') { continue }
            if ($valid -notcontains $u) { throw "Unknown export format '$part'. Use CSV, JSON, TXT, HTML or XML." }
            if (-not $out.Contains($u)) { $out.Add($u) }
        }
    }
    return $out.ToArray()
}

function Normalize-Types {
    param([string[]]$Types)
    $out = New-Object 'System.Collections.Generic.List[string]'
    foreach ($t in @($Types)) {
        foreach ($part in ([string]$t -split ',')) {
            $v = $part.Trim().ToLowerInvariant()
            if ($v -eq '') { continue }
            if ($v -ne '[no extension]' -and -not $v.StartsWith('.')) { $v = '.' + $v }
            if (-not $out.Contains($v)) { $out.Add($v) }
        }
    }
    return $out.ToArray()
}

function Log {
    param([string]$Msg)
    if ($null -ne $script:logBox) {
        $script:logBox.AppendText($Msg + "`r`n")
        $script:logBox.SelectionStart = $script:logBox.TextLength
        $script:logBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        Write-Host $Msg
    }
}

function Set-Status {
    param([string]$Text)
    if ($null -ne $script:statusLabel) {
        $script:statusLabel.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Test-Cancel {
    $script:tick++
    if ($null -ne $script:logBox -and ($script:tick % 200) -eq 0) {
        [System.Windows.Forms.Application]::DoEvents()
    }
    if ($script:cancelRequested) {
        throw [System.OperationCanceledException]::new("Cancelled by user.")
    }
}

function Close-ZipArchive {
    if ($script:zipArchive) {
        try { $script:zipArchive.Dispose() } catch { }
        $script:zipArchive = $null
    }
}

# ---------- settings ----------
function Load-Settings {
    $s = @{
        LastSource = ""
        LastOutput = ""
        LastKind = "Zip"
        Formats = @('TXT')
        OpenWhenDone = $true
        ExtractMode = "None"
        Include = ""
        Exclude = ""
    }
    try {
        if (Test-Path -LiteralPath $script:configPath) {
            $j = Get-Content -LiteralPath $script:configPath -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($k in @($s.Keys)) {
                $prop = $j.PSObject.Properties[$k]
                if ($null -ne $prop) { $s[$k] = $prop.Value }
            }
        }
    } catch { }
    if ($null -eq $s.Formats) { $s.Formats = @('TXT') }
    elseif ($s.Formats -isnot [array]) { $s.Formats = @([string]$s.Formats) }
    return $s
}

function Save-Settings {
    param($Settings)
    try {
        $dir = Split-Path -Parent $script:configPath
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $json = $Settings | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($script:configPath, $json, [System.Text.UTF8Encoding]::new($false))
    } catch { }
}

# ---------- validate + resolve source/output ----------
function Initialize-Context {
    param(
        [string]$SourcePath,
        [string]$Kind,
        [string]$OutputPath,
        [string]$IncludeMask = "",
        [string]$ExcludeMask = ""
    )
    if ([string]::IsNullOrWhiteSpace($SourcePath)) { throw "Please choose a source (zip file or folder)." }

    try { $resolved = (Resolve-Path -LiteralPath $SourcePath).ProviderPath }
    catch { throw "Source not found: $SourcePath" }

    if ($Kind -eq "Folder") {
        if (-not (Test-Path -LiteralPath $resolved -PathType Container)) { throw "Folder not found: $resolved" }
    } else {
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw "Zip file not found: $resolved" }
        $Kind = "Zip"
    }
    $script:sourcePath = $resolved
    $script:sourceKind = $Kind
    $script:filterInclude = $IncludeMask
    $script:filterExclude = $ExcludeMask

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path ([System.IO.Path]::GetTempPath()) "DGsFileSearch"
    }
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    $script:outputDir = (Resolve-Path -LiteralPath $OutputPath).ProviderPath

    if ($Kind -eq "Folder") {
        $srcFull = [System.IO.Path]::GetFullPath($script:sourcePath).TrimEnd('\')
        $outFull = [System.IO.Path]::GetFullPath($script:outputDir).TrimEnd('\')
        if ($srcFull.Equals($outFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "The output folder cannot be the same as the source folder."
        }
        if (Test-UnderDir $script:sourcePath $script:outputDir) {
            throw "The output folder cannot contain the source folder."
        }
    }

    if ($Kind -eq "Folder") {
        $bn = [System.IO.Path]::GetFileName($script:sourcePath)
    } else {
        $bn = [System.IO.Path]::GetFileNameWithoutExtension($script:sourcePath)
    }
    if ([string]::IsNullOrWhiteSpace($bn)) { $bn = Split-Path $script:sourcePath -Leaf }
    if ([string]::IsNullOrWhiteSpace($bn)) { $bn = ($script:sourcePath -replace '[:\\/]+','_').Trim('_') }
    if ([string]::IsNullOrWhiteSpace($bn)) { $bn = "scan" }
    $script:baseName = ($bn -replace '[^\w\- ]','_')
}

# ---------- core scan ----------
function Invoke-Scan {
    $script:entries = New-Object 'System.Collections.Generic.List[object]'
    $script:filteredOut = 0
    $scanErrors = 0
    Close-ZipArchive

    if ($script:sourceKind -eq "Zip") {
        try {
            $script:zipArchive = [System.IO.Compression.ZipFile]::OpenRead($script:sourcePath)
        } catch {
            throw "Could not open archive '$($script:sourcePath)': $($_.Exception.Message)"
        }
        foreach ($entry in $script:zipArchive.Entries) {
            Test-Cancel
            try {
                $isFolder = [string]::IsNullOrEmpty($entry.Name) -or $entry.FullName.EndsWith("/") -or $entry.FullName.EndsWith("\")
                $displayPath = $entry.FullName -replace '[/\\]$',''
                if ($isFolder) {
                    $ext = "[Folder]"
                } elseif ($entry.Name -match '\.') {
                    $ext = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
                    if ([string]::IsNullOrEmpty($ext)) { $ext = "[no extension]" }
                } else { $ext = "[no extension]" }

                if (-not $isFolder) {
                    if (-not (Test-FileFilter $entry.Name $displayPath)) { $script:filteredOut++; continue }
                }
                try { $mod = $entry.LastWriteTime.DateTime } catch { $mod = [datetime]'1980-01-01' }
                $script:entries.Add([PSCustomObject]@{
                    Type = $(if ($isFolder) { "Folder" } else { "File" })
                    Path = $displayPath
                    Name = $(if ($isFolder) { $displayPath } else { $entry.Name })
                    Extension = $ext
                    SizeBytes = [long]$entry.Length
                    CompressedBytes = [long]$entry.CompressedLength
                    Modified = $mod
                    FullPath = $null
                })
            } catch [System.OperationCanceledException] { throw } catch { $scanErrors++ }
        }
    } else {
        $rootLen = $script:sourcePath.TrimEnd('\').Length + 1
        $filesList = New-Object 'System.Collections.Generic.List[string]'
        $dirsList  = New-Object 'System.Collections.Generic.List[string]'
        $stack = New-Object 'System.Collections.Generic.Stack[string]'
        $stack.Push($script:sourcePath)
        while ($stack.Count -gt 0) {
            Test-Cancel
            $dir = $stack.Pop()
            try {
                foreach ($f in [System.IO.Directory]::EnumerateFileSystemEntries($dir, '*', [System.IO.SearchOption]::TopDirectoryOnly)) {
                    Test-Cancel
                    try {
                        if (Test-InOutputTree $f) { continue }
                        if ([System.IO.Directory]::Exists($f)) {
                            $dirsList.Add($f)
                            $attr = [System.IO.File]::GetAttributes($f)
                            if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                                $stack.Push($f)
                            }
                        } else {
                            $name = [System.IO.Path]::GetFileName($f)
                            $rel = $f.Substring($rootLen)
                            if (-not (Test-FileFilter $name $rel)) { $script:filteredOut++; continue }
                            $filesList.Add($f)
                        }
                    } catch [System.OperationCanceledException] { throw } catch { $scanErrors++ }
                }
            } catch [System.OperationCanceledException] { throw } catch { $scanErrors++ }
        }
        foreach ($d in $dirsList) {
            Test-Cancel
            try {
                $rel = $d.Substring($rootLen)
                $script:entries.Add([PSCustomObject]@{ Type='Folder'; Path=$rel; Name=$rel; Extension='[Folder]'
                    SizeBytes=[long]0; CompressedBytes=[long]0; Modified=[System.IO.Directory]::GetLastWriteTime($d); FullPath=$d })
            } catch [System.OperationCanceledException] { throw } catch { $scanErrors++ }
        }
        foreach ($f in $filesList) {
            Test-Cancel
            try {
                $rel = $f.Substring($rootLen)
                $name = [System.IO.Path]::GetFileName($f)
                $ext  = [System.IO.Path]::GetExtension($f).ToLowerInvariant()
                if ([string]::IsNullOrEmpty($ext)) { $ext = "[no extension]" }
                $fi = New-Object System.IO.FileInfo($f)
                $script:entries.Add([PSCustomObject]@{ Type='File'; Path=$rel; Name=$name; Extension=$ext
                    SizeBytes=[long]$fi.Length; CompressedBytes=[long]$fi.Length; Modified=$fi.LastWriteTime; FullPath=$f })
            } catch [System.OperationCanceledException] { throw } catch { $scanErrors++ }
        }
    }
    return $scanErrors
}

# ---------- exports ----------
function Write-Exports {
    param([string[]]$Formats)
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $written = @{}
    $entries = $script:entries
    $sorted = @($entries | Sort-Object Type, Path)
    $folders = @($entries | Where-Object { $_.Type -eq "Folder" })
    $files   = $script:files
    $totalSize = [long](($files | Measure-Object -Property SizeBytes -Sum).Sum)
    $typeGroups = $script:typeGroups
    $baseName = $script:baseName
    $outDir = $script:outputDir
    $hasFilter = (-not [string]::IsNullOrWhiteSpace($script:filterInclude)) -or (-not [string]::IsNullOrWhiteSpace($script:filterExclude))

    foreach ($fmt in $Formats) {
        Test-Cancel
        try {
            $outFile = Join-Path $outDir ("{0}_contents.{1}" -f $baseName, $fmt.ToLower())
            switch ($fmt) {
                'CSV' {
                    $sw = New-Object System.IO.StreamWriter($outFile, $false, $utf8NoBom)
                    $sw.WriteLine('Type,Path,Name,Extension,SizeBytes,CompressedBytes,Modified')
                    foreach ($e in $sorted) {
                        Test-Cancel
                        $line = '"{0}","{1}","{2}","{3}",{4},{5},"{6}"' -f $e.Type, ($e.Path -replace '"','""'), ($e.Name -replace '"','""'), ($e.Extension -replace '"','""'), $e.SizeBytes, $e.CompressedBytes, $e.Modified.ToString('o')
                        $sw.WriteLine($line)
                    }
                    $sw.Close(); $written.CSV = $outFile
                }
                'JSON' {
                    $cap = 20000
                    $summary = [ordered]@{
                        App = $appName; AppVersion = $appVersion; GeneratedAt = (Get-Date).ToString('o')
                        Source = $script:sourcePath; SourceKind = $script:sourceKind
                        IncludeFilter = $script:filterInclude; ExcludeFilter = $script:filterExclude
                        FilesSkippedByFilters = $script:filteredOut
                        TotalEntries = $entries.Count; FolderCount = $folders.Count; FileCount = $files.Count
                        TotalSizeBytes = $totalSize
                        FileTypes = @(foreach ($g in $typeGroups) { [ordered]@{ Extension=$g.Name; Count=$g.Count; SizeBytes=[long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum } })
                        Folders = @($folders | Sort-Object Path | Select-Object -ExpandProperty Path)
                        EntriesTruncated = ($entries.Count -gt $cap)
                        Entries = @($sorted | Select-Object -First $cap | ForEach-Object { [ordered]@{ Type=$_.Type; Path=$_.Path; Name=$_.Name; Extension=$_.Extension; SizeBytes=$_.SizeBytes; CompressedBytes=$_.CompressedBytes; Modified=$_.Modified.ToString('o') } })
                    }
                    [System.IO.File]::WriteAllText($outFile, ($summary | ConvertTo-Json -Depth 6), $utf8NoBom)
                    $written.JSON = $outFile
                }
                'TXT' {
                    $sb = New-Object System.Text.StringBuilder
                    [void]$sb.AppendLine("$appName v$appVersion - plain-text report")
                    [void]$sb.AppendLine("Source: $($script:sourcePath) ($($script:sourceKind))")
                    if ($hasFilter) {
                        [void]$sb.AppendLine("Include: $($script:filterInclude)")
                        [void]$sb.AppendLine("Exclude: $($script:filterExclude)")
                        [void]$sb.AppendLine(("Files skipped by filters: {0:N0}" -f $script:filteredOut))
                    }
                    [void]$sb.AppendLine(("Generated: {0}" -f (Get-Date)))
                    [void]$sb.AppendLine(("Entries: {0:N0}  Folders: {1:N0}  Files: {2:N0}  Total size: {3}" -f $entries.Count, $folders.Count, $files.Count, (Format-Size $totalSize)))
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("=== FILE TYPES ===")
                    foreach ($g in $typeGroups) { [void]$sb.AppendLine(("{0,-16} {1,8:N0}  {2}" -f $g.Name, $g.Count, (Format-Size ([long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum)))) }
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("=== FOLDERS ===")
                    foreach ($f in ($folders | Sort-Object Path)) { [void]$sb.AppendLine($f.Path) }
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("=== ALL ENTRIES ===")
                    foreach ($e in $sorted) {
                        Test-Cancel
                        [void]$sb.AppendLine(("{0,-7} {1,12}  {2}" -f $e.Type, (Format-Size $e.SizeBytes), $e.Path))
                    }
                    [System.IO.File]::WriteAllText($outFile, $sb.ToString(), $utf8NoBom)
                    $written.TXT = $outFile
                }
                'HTML' {
                    $sb = New-Object System.Text.StringBuilder
                    [void]$sb.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'><title>$appName - $(HtmlEnc $baseName)</title>")
                    [void]$sb.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1a2740}h1{color:#0e5a8a}table{border-collapse:collapse;width:100%;margin:12px 0}th,td{border:1px solid #cdd7e5;padding:6px 10px;text-align:left;font-size:13px}th{background:#0e5a8a;color:#fff}tr:nth-child(even){background:#f2f6fb}.tag{display:inline-block;background:#e6f0fa;color:#0e5a8a;border-radius:4px;padding:2px 8px;margin:2px}</style></head><body>")
                    [void]$sb.AppendLine("<h1>$(HtmlEnc $appName)</h1>")
                    [void]$sb.AppendLine("<p><b>Source:</b> $(HtmlEnc $script:sourcePath) ($(HtmlEnc $script:sourceKind))<br><b>Generated:</b> $(HtmlEnc (Get-Date).ToString())<br><b>Entries:</b> $($entries.Count) &nbsp; <b>Folders:</b> $($folders.Count) &nbsp; <b>Files:</b> $($files.Count) &nbsp; <b>Total size:</b> $(HtmlEnc (Format-Size $totalSize))</p>")
                    if ($hasFilter) {
                        $filterText = "Include: $($script:filterInclude) &nbsp; Exclude: $($script:filterExclude) &nbsp; Skipped: $($script:filteredOut)"
                        [void]$sb.AppendLine("<p><b>Filters:</b> $(HtmlEnc $filterText)</p>")
                    }
                    [void]$sb.AppendLine("<h2>File types</h2><table><tr><th>Extension</th><th>Count</th><th>Total size</th></tr>")
                    foreach ($g in $typeGroups) { [void]$sb.AppendLine(("<tr><td><span class='tag'>{0}</span></td><td>{1:N0}</td><td>{2}</td></tr>" -f (HtmlEnc $g.Name), $g.Count, (HtmlEnc (Format-Size ([long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum))))) }
                    [void]$sb.AppendLine("</table><h2>All entries</h2><table><tr><th>Type</th><th>Path</th><th>Size</th><th>Modified</th></tr>")
                    $cap = 50000; $n = 0
                    foreach ($e in $sorted) {
                        Test-Cancel
                        if ($n -ge $cap) { [void]$sb.AppendLine("<tr><td colspan=4>... truncated, see CSV/TXT</td></tr>"); break }
                        [void]$sb.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f (HtmlEnc $e.Type), (HtmlEnc $e.Path), (HtmlEnc (Format-Size $e.SizeBytes)), (HtmlEnc $e.Modified.ToString())))
                        $n++
                    }
                    [void]$sb.AppendLine("</table></body></html>")
                    [System.IO.File]::WriteAllText($outFile, $sb.ToString(), $utf8NoBom)
                    $written.HTML = $outFile
                }
                'XML' {
                    $sb = New-Object System.Text.StringBuilder
                    [void]$sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
                    [void]$sb.AppendLine("<Report App='$appName' Version='$appVersion' Source='$(XmlEnc $script:sourcePath)' SourceKind='$(XmlEnc $script:sourceKind)' IncludeFilter='$(XmlEnc $script:filterInclude)' ExcludeFilter='$(XmlEnc $script:filterExclude)' FilesSkippedByFilters='$($script:filteredOut)' Generated='$(Get-Date -Format o)' TotalEntries='$($entries.Count)' FolderCount='$($folders.Count)' FileCount='$($files.Count)' TotalSizeBytes='$totalSize'>")
                    [void]$sb.AppendLine("  <FileTypes>")
                    foreach ($g in $typeGroups) { [void]$sb.AppendLine(("    <Type Extension='{0}' Count='{1}' SizeBytes='{2}'/>" -f (XmlEnc $g.Name), $g.Count, [long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum)) }
                    [void]$sb.AppendLine("  </FileTypes>")
                    [void]$sb.AppendLine("  <Folders>")
                    foreach ($f in ($folders | Sort-Object Path)) { [void]$sb.AppendLine(("    <Folder Path='{0}'/>" -f (XmlEnc $f.Path))) }
                    [void]$sb.AppendLine("  </Folders>")
                    [void]$sb.AppendLine("  <Entries>")
                    $cap = 50000; $n = 0
                    foreach ($e in $sorted) {
                        Test-Cancel
                        if ($n -ge $cap) { [void]$sb.AppendLine("    <Truncated/>"); break }
                        [void]$sb.AppendLine(("    <Entry Type='{0}' Path='{1}' Extension='{2}' SizeBytes='{3}' Modified='{4}'/>" -f (XmlEnc $e.Type), (XmlEnc $e.Path), (XmlEnc $e.Extension), $e.SizeBytes, $e.Modified.ToString('o')))
                        $n++
                    }
                    [void]$sb.AppendLine("  </Entries></Report>")
                    [System.IO.File]::WriteAllText($outFile, $sb.ToString(), $utf8NoBom)
                    $written.XML = $outFile
                }
            }
            Log "  Exported $fmt -> $outFile"
        }
        catch [System.OperationCanceledException] { throw }
        catch { Log "  [WARN] Export $fmt failed: $_" }
    }
    return $written
}

# ---------- extraction ----------
function Invoke-Extract {
    param(
        [string]$Mode,
        [string[]]$SelectedTypes = @(),
        [string[]]$SelectedPaths = @()
    )
    if ($Mode -eq "None" -or $null -eq $script:files -or $script:files.Count -eq 0) { return 0 }
    $extractDir = Join-Path $script:outputDir ("{0}_extracted" -f $script:baseName)
    New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
    $wantAll = ($Mode -eq "All")
    $wantExts = @()
    $singlePaths = @{}

    if ($Mode -eq "ByType") {
        if ($null -ne $SelectedTypes -and $SelectedTypes.Count -gt 0) {
            $wantExts = @(Normalize-Types -Types $SelectedTypes)
        } elseif ($NoGui) {
            Log "  [WARN] ByType extraction needs -ExtractType when -NoGui is used. Skipped."
            return 0
        } else {
            $opts = foreach ($g in $script:typeGroups) { [PSCustomObject]@{ Extension = $g.Name; Count = $g.Count; Size = (Format-Size ([long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum)) } }
            $picked = $opts | Out-GridView -Title "$appName - Select file type(s) to extract (Ctrl+click). Cancel = skip." -PassThru
            if ($null -eq $picked) { Log "  Extraction cancelled."; return 0 }
            $wantExts = @($picked | Select-Object -ExpandProperty Extension)
        }
    }
    if ($Mode -eq "Single") {
        if ($null -ne $SelectedPaths -and $SelectedPaths.Count -gt 0) {
            foreach ($p in $SelectedPaths) {
                $key = ([string]$p) -replace '/','\'
                if ($key -ne '') { $singlePaths[$key] = $true }
            }
        } elseif ($NoGui) {
            Log "  [WARN] Single-file extraction needs -ExtractFile when -NoGui is used. Skipped."
            return 0
        } else {
            $cap = 50000
            $sp = $script:files | Select-Object -First $cap | Out-GridView -Title "$appName - Pick file(s) to extract (Ctrl+click). Cancel = none." -PassThru
            if ($null -eq $sp) { Log "  No files selected, extraction skipped."; return 0 }
            foreach ($s in @($sp)) {
                $key = ([string]$s.Path) -replace '/','\'
                if ($key -ne '') { $singlePaths[$key] = $true }
            }
            if ($script:files.Count -gt $cap) { Log "  [WARN] Single-file picker capped at $cap entries." }
        }
    }

    $matchTest = {
        param($ext, $path)
        if ($wantAll) { return $true }
        if ($wantExts -contains $ext) { return $true }
        $key = ([string]$path) -replace '/','\'
        if ($singlePaths.ContainsKey($key)) { return $true }
        return $false
    }

    $extracted = 0; $extractErrors = 0
    $openedHere = $false
    try {
        if ($script:sourceKind -eq "Zip") {
            if ($null -eq $script:zipArchive) {
                $script:zipArchive = [System.IO.Compression.ZipFile]::OpenRead($script:sourcePath)
                $openedHere = $true
            }
            foreach ($entry in $script:zipArchive.Entries) {
                Test-Cancel
                try {
                    $isDirEntry = [string]::IsNullOrEmpty($entry.Name) -or $entry.FullName.EndsWith("/") -or $entry.FullName.EndsWith("\")
                    if ($isDirEntry) {
                        if ($wantAll) {
                            $relDir = Get-SafeRelPath ($entry.FullName -replace '[/\\]$','')
                            if ($relDir -ne '') {
                                $dd = Join-Path $extractDir $relDir
                                if (Test-UnderDir $dd $extractDir) { New-Item -ItemType Directory -Path $dd -Force | Out-Null }
                            }
                        }
                        continue
                    }
                    $rel = Get-SafeRelPath ($entry.FullName -replace '[/\\]$','')
                    if ([string]::IsNullOrEmpty($rel)) { continue }
                    $ext = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
                    if ([string]::IsNullOrEmpty($ext)) { $ext = "[no extension]" }
                    $p = $entry.FullName -replace '[/\\]$',''
                    if (-not (Test-FileFilter $entry.Name $p)) { continue }
                    if (-not (& $matchTest $ext $p)) { continue }
                    $dest = Join-Path $extractDir $rel
                    if (-not (Test-UnderDir $dest $extractDir)) { $extractErrors++; continue }
                    $destDir = Split-Path $dest -Parent
                    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
                    $extracted++
                } catch [System.OperationCanceledException] { throw }
                catch { $extractErrors++ }
            }
        } else {
            foreach ($e in $script:files) {
                Test-Cancel
                try {
                    if ([string]::IsNullOrEmpty($e.FullPath)) { continue }
                    $rel = Get-SafeRelPath $e.Path
                    if ([string]::IsNullOrEmpty($rel)) { continue }
                    if (-not (& $matchTest $e.Extension $e.Path)) { continue }
                    $dest = Join-Path $extractDir $rel
                    if (-not (Test-UnderDir $dest $extractDir)) { $extractErrors++; continue }
                    $destDir = Split-Path $dest -Parent
                    if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                    Copy-Item -LiteralPath $e.FullPath -Destination $dest -Force
                    $extracted++
                } catch [System.OperationCanceledException] { throw }
                catch { $extractErrors++ }
            }
        }
    }
    finally {
        if ($openedHere) { Close-ZipArchive }
    }
    Log "  Extracted $extracted file(s) to: $extractDir"
    if ($extractErrors -gt 0) { Log "  [WARN] $extractErrors entry(ies) skipped during extraction." }
    return $extracted
}

# ---------- full run (shared by GUI and -NoGui) ----------
function Invoke-Pipeline {
    param(
        [string[]]$Formats = @(),
        [string]$Mode = "None",
        [string[]]$Types = @(),
        [string[]]$FilesToExtract = @(),
        [switch]$OpenWhenDone
    )
    $script:lastExports = @{}
    try {
        Set-Status "Scanning..."
        Log "Source : $($script:sourcePath) ($($script:sourceKind))"
        Log "Output : $($script:outputDir)"
        if (-not [string]::IsNullOrWhiteSpace($script:filterInclude)) { Log "Include: $($script:filterInclude)" }
        if (-not [string]::IsNullOrWhiteSpace($script:filterExclude)) { Log "Exclude: $($script:filterExclude)" }
        Log "Scanning (all subfolders included)..."
        $scanErrors = Invoke-Scan

        $folderList = @($script:entries | Where-Object { $_.Type -eq "Folder" })
        $script:files = @($script:entries | Where-Object { $_.Type -eq "File" })
        $totalSize = [long](($script:files | Measure-Object -Property SizeBytes -Sum).Sum)
        $script:typeGroups = @($script:files | Group-Object -Property Extension | Sort-Object Count -Descending)

        Log ""
        Log ("Total entries : {0:N0}" -f $script:entries.Count)
        Log ("  Folders     : {0:N0}" -f $folderList.Count)
        Log ("  Files       : {0:N0}" -f $script:files.Count)
        Log ("Total size    : {0} ({1:N0} bytes)" -f (Format-Size $totalSize), $totalSize)
        if ($script:filteredOut -gt 0) { Log ("Files skipped by filters: {0:N0}" -f $script:filteredOut) }
        if ($scanErrors -gt 0) { Log ("[WARN] Skipped due to errors: {0:N0}" -f $scanErrors) }
        Log ""
        Log "FILE TYPES:"
        Log ("  {0,-16} {1,8} {2,14}" -f "Extension", "Count", "Total Size")
        foreach ($g in $script:typeGroups) {
            $sum = [long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum
            Log ("  {0,-16} {1,8:N0} {2,14}" -f $g.Name, $g.Count, (Format-Size $sum))
        }
        Log ""

        if ($null -ne $script:logBox) {
            Update-ResultsGrid
            Update-TypeGrid
        }

        $fmtsChosen = @(Normalize-Formats -Formats $Formats)
        if ($fmtsChosen.Count -eq 0) { $fmtsChosen = @('TXT'); Log "No format selected, defaulting to TXT." }
        Set-Status "Writing exports..."
        Log "Writing exports..."
        $written = Write-Exports -Formats $fmtsChosen
        $script:lastExports = $written
        Log "Exports done: $($written.Keys -join ', ')"
        Log ""

        if ($Mode -ne "None") {
            Set-Status "Extracting..."
            Log "Extraction ($Mode)..."
            [void](Invoke-Extract -Mode $Mode -SelectedTypes $Types -SelectedPaths $FilesToExtract)
        }

        Set-Status "Done."
        Log ""
        Log "=== DONE ==="
        Log "Output folder: $($script:outputDir)"
        if ($OpenWhenDone) { Start-Process explorer.exe $script:outputDir }
    }
    finally {
        Close-ZipArchive
    }
}

# ---------- GUI result helpers ----------
function Apply-ResultsFilter {
    if ($null -eq $script:resultsView) { return }
    $t = $gridFilter.Text.Trim()
    if ($t -eq '') {
        $script:resultsView.RowFilter = ''
    } else {
        $safe = $t.Replace("'", "''").Replace('[', '').Replace(']', '')
        $script:resultsView.RowFilter = "Path LIKE '%$safe%' OR Name LIKE '%$safe%' OR Extension LIKE '%$safe%' OR Type LIKE '%$safe%'"
    }
    if ($null -ne $script:statusLabel) {
        $total = 0
        if ($null -ne $script:entries) { $total = $script:entries.Count }
        $script:statusLabel.Text = ("Showing {0:N0} of {1:N0} entries" -f $script:resultsView.Count, $total)
    }
}

function Update-ResultsGrid {
    if ($null -eq $grid) { return }
    $table = New-Object System.Data.DataTable "Results"
    [void]$table.Columns.Add("Type", [string])
    [void]$table.Columns.Add("Path", [string])
    [void]$table.Columns.Add("Name", [string])
    [void]$table.Columns.Add("Extension", [string])
    [void]$table.Columns.Add("SizeBytes", [long])
    [void]$table.Columns.Add("Modified", [datetime])
    $table.BeginLoadData()
    $n = 0
    foreach ($e in $script:entries) {
        if ($n -ge $script:displayCap) { break }
        $r = $table.NewRow()
        $r["Type"] = [string]$e.Type
        $r["Path"] = [string]$e.Path
        $r["Name"] = [string]$e.Name
        $r["Extension"] = [string]$e.Extension
        $r["SizeBytes"] = [long]$e.SizeBytes
        $r["Modified"] = [datetime]$e.Modified
        $table.Rows.Add($r)
        $n++
    }
    $table.EndLoadData()
    $script:resultsView = New-Object System.Data.DataView($table)
    $grid.DataSource = $script:resultsView
    if ($null -ne $grid.Columns["Type"]) { $grid.Columns["Type"].Width = 70 }
    if ($null -ne $grid.Columns["Path"]) { $grid.Columns["Path"].AutoSizeMode = "Fill" }
    if ($null -ne $grid.Columns["Name"]) { $grid.Columns["Name"].Width = 180 }
    if ($null -ne $grid.Columns["Extension"]) { $grid.Columns["Extension"].Width = 100 }
    if ($null -ne $grid.Columns["SizeBytes"]) { $grid.Columns["SizeBytes"].Width = 110; $grid.Columns["SizeBytes"].HeaderText = "Size" }
    if ($null -ne $grid.Columns["Modified"]) { $grid.Columns["Modified"].Width = 140 }
    Apply-ResultsFilter
    if ($script:entries.Count -gt $script:displayCap) {
        Log ("  [WARN] Results tab shows first {0:N0} entries (exports contain all)." -f $script:displayCap)
    }
    if ($null -ne $tabs) { $tabs.SelectedTab = $tabResults }
}

function Update-TypeGrid {
    if ($null -eq $typesGrid) { return }
    $table = New-Object System.Data.DataTable "Types"
    [void]$table.Columns.Add("Extension", [string])
    [void]$table.Columns.Add("Count", [int])
    [void]$table.Columns.Add("TotalSize", [long])
    $table.BeginLoadData()
    foreach ($g in $script:typeGroups) {
        $r = $table.NewRow()
        $r["Extension"] = [string]$g.Name
        $r["Count"] = [int]$g.Count
        $r["TotalSize"] = [long]($g.Group | Measure-Object -Property SizeBytes -Sum).Sum
        $table.Rows.Add($r)
    }
    $table.EndLoadData()
    $script:typeView = New-Object System.Data.DataView($table)
    $script:typeView.Sort = "Count DESC"
    $typesGrid.DataSource = $script:typeView
    if ($null -ne $typesGrid.Columns["Extension"]) { $typesGrid.Columns["Extension"].Width = 180 }
    if ($null -ne $typesGrid.Columns["Count"]) { $typesGrid.Columns["Count"].Width = 100 }
    if ($null -ne $typesGrid.Columns["TotalSize"]) { $typesGrid.Columns["TotalSize"].Width = 150; $typesGrid.Columns["TotalSize"].HeaderText = "Total size" }
}

function Get-CurrentSettings {
    $fmtList = @($fmts | Where-Object { $chkFmt[$_].Checked })
    $mode = "None"
    if ($rbEAll.Checked) { $mode = "All" }
    elseif ($rbEType.Checked) { $mode = "ByType" }
    elseif ($rbESingle.Checked) { $mode = "Single" }
    return [ordered]@{
        LastSource = $txtSource.Text.Trim()
        LastOutput = $txtOutput.Text.Trim()
        LastKind = $(if ($rbZip.Checked) { "Zip" } else { "Folder" })
        Formats = $fmtList
        OpenWhenDone = $chkOpen.Checked
        ExtractMode = $mode
        Include = $txtInclude.Text.Trim()
        Exclude = $txtExclude.Text.Trim()
    }
}

function Update-Layout {
    if ($null -eq $tabs -or $null -eq $panelTop -or $null -eq $status) { return }
    $h = $form.ClientSize.Height - $panelTop.Height - $status.Height
    if ($h -lt 120) { $h = 120 }
    $tabs.SetBounds(0, $panelTop.Height, $form.ClientSize.Width, $h)
}

# ---------- load assemblies ----------
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Required .NET assembly unavailable: $_", "$appName - Error", "OK", "Error") | Out-Null
    exit 1
}

# ---------- automation / no-GUI mode ----------
if ($NoGui) {
    try {
        if (-not [string]::IsNullOrWhiteSpace($ZipPath)) {
            $src = $ZipPath; $kind = "Zip"
        } elseif (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
            $src = $FolderPath; $kind = "Folder"
        } else {
            throw "Provide -ZipPath or -FolderPath when using -NoGui."
        }
        Initialize-Context -SourcePath $src -Kind $kind -OutputPath $OutputDir -IncludeMask $Include -ExcludeMask $Exclude

        $mode = "None"
        if ($Extract) { $mode = $ExtractMode }
        if ($null -ne $ExtractType -and $ExtractType.Count -gt 0) { $mode = "ByType" }
        $extractFiles = @()
        foreach ($ef in $ExtractFile) {
            foreach ($part in ([string]$ef -split ',')) {
                $x = $part.Trim()
                if ($x -ne '') { $extractFiles += $x }
            }
        }
        if ($extractFiles.Count -gt 0) { $mode = "Single" }

        Invoke-Pipeline -Formats $ExportFormat -Mode $mode -Types $ExtractType -FilesToExtract $extractFiles -OpenWhenDone:$Open
        exit 0
    }
    catch {
        if ($_.Exception -is [System.OperationCanceledException]) { Write-Host "Cancelled." }
        else { Write-Host "[ERROR] $_" -ForegroundColor Red }
        exit 1
    }
}

# ---------- high-DPI ----------
try {
    Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public static class DgsDpi { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); }' -ErrorAction Stop
    [void][DgsDpi]::SetProcessDPIAware()
} catch { }

# ---------- load saved GUI settings ----------
$settings = Load-Settings

$initZip = $ZipPath
$initFolder = $FolderPath
if ([string]::IsNullOrWhiteSpace($initZip) -and [string]::IsNullOrWhiteSpace($initFolder)) {
    $last = [string]$settings.LastSource
    if (-not [string]::IsNullOrWhiteSpace($last)) {
        if ([string]$settings.LastKind -eq "Folder") { $initFolder = $last } else { $initZip = $last }
    }
}
$initOutput = $OutputDir
if ([string]::IsNullOrWhiteSpace($initOutput)) { $initOutput = [string]$settings.LastOutput }
$initInclude = $Include
if ([string]::IsNullOrWhiteSpace($initInclude)) { $initInclude = [string]$settings.Include }
$initExclude = $Exclude
if ([string]::IsNullOrWhiteSpace($initExclude)) { $initExclude = [string]$settings.Exclude }
$savedFormats = @(@($settings.Formats) | ForEach-Object { ([string]$_).ToUpperInvariant() })
if ($savedFormats.Count -eq 0) { $savedFormats = @('TXT') }

# ---------- build form ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "$appName v$appVersion"
$form.Size = New-Object System.Drawing.Size(980, 720)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(840, 620)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AllowDrop = $true

$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock = "Top"
$panelTop.Height = 352
$form.Controls.Add($panelTop)

# --- source group ---
$grpSrc = New-Object System.Windows.Forms.GroupBox
$grpSrc.Text = "1. Source (zip file or folder - all subfolders included)"
$grpSrc.Location = New-Object System.Drawing.Point(12, 8)
$grpSrc.Size = New-Object System.Drawing.Size(920, 76)
$grpSrc.Anchor = "Top,Left,Right"
$panelTop.Controls.Add($grpSrc)

$rbZip = New-Object System.Windows.Forms.RadioButton
$rbZip.Text = "Zip archive"
$rbZip.Location = New-Object System.Drawing.Point(12, 20)
$rbZip.Size = New-Object System.Drawing.Size(100, 20)
$rbZip.Checked = $true
$grpSrc.Controls.Add($rbZip)

$rbFolder = New-Object System.Windows.Forms.RadioButton
$rbFolder.Text = "Folder"
$rbFolder.Location = New-Object System.Drawing.Point(118, 20)
$rbFolder.Size = New-Object System.Drawing.Size(80, 20)
$grpSrc.Controls.Add($rbFolder)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(12, 44)
$txtSource.Size = New-Object System.Drawing.Size(816, 22)
$txtSource.Anchor = "Top,Left,Right"
$grpSrc.Controls.Add($txtSource)

$btnSrc = New-Object System.Windows.Forms.Button
$btnSrc.Text = "Browse..."
$btnSrc.Location = New-Object System.Drawing.Point(836, 43)
$btnSrc.Size = New-Object System.Drawing.Size(72, 24)
$btnSrc.Anchor = "Top,Right"
$grpSrc.Controls.Add($btnSrc)

# --- output group ---
$grpOut = New-Object System.Windows.Forms.GroupBox
$grpOut.Text = "2. Output folder (reports + extracted files go here)"
$grpOut.Location = New-Object System.Drawing.Point(12, 92)
$grpOut.Size = New-Object System.Drawing.Size(920, 56)
$grpOut.Anchor = "Top,Left,Right"
$panelTop.Controls.Add($grpOut)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(12, 24)
$txtOutput.Size = New-Object System.Drawing.Size(816, 22)
$txtOutput.Anchor = "Top,Left,Right"
$grpOut.Controls.Add($txtOutput)

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = "Browse..."
$btnOut.Location = New-Object System.Drawing.Point(836, 23)
$btnOut.Size = New-Object System.Drawing.Size(72, 24)
$btnOut.Anchor = "Top,Right"
$grpOut.Controls.Add($btnOut)

if (-not [string]::IsNullOrWhiteSpace($initZip)) {
    $rbZip.Checked = $true; $rbFolder.Checked = $false
    $txtSource.Text = $initZip
} elseif (-not [string]::IsNullOrWhiteSpace($initFolder)) {
    $rbFolder.Checked = $true; $rbZip.Checked = $false
    $txtSource.Text = $initFolder
}
if (-not [string]::IsNullOrWhiteSpace($initOutput)) { $txtOutput.Text = $initOutput }

# --- filters group ---
$grpFilter = New-Object System.Windows.Forms.GroupBox
$grpFilter.Text = "3. File filters (wildcards separated by ; - blank = all files)"
$grpFilter.Location = New-Object System.Drawing.Point(12, 156)
$grpFilter.Size = New-Object System.Drawing.Size(920, 52)
$grpFilter.Anchor = "Top,Left,Right"
$panelTop.Controls.Add($grpFilter)

$lblInclude = New-Object System.Windows.Forms.Label
$lblInclude.Text = "Include:"
$lblInclude.Location = New-Object System.Drawing.Point(12, 24)
$lblInclude.Size = New-Object System.Drawing.Size(52, 20)
$grpFilter.Controls.Add($lblInclude)

$txtInclude = New-Object System.Windows.Forms.TextBox
$txtInclude.Location = New-Object System.Drawing.Point(68, 22)
$txtInclude.Size = New-Object System.Drawing.Size(380, 22)
$txtInclude.Text = $initInclude
$grpFilter.Controls.Add($txtInclude)

$lblExclude = New-Object System.Windows.Forms.Label
$lblExclude.Text = "Exclude:"
$lblExclude.Location = New-Object System.Drawing.Point(460, 24)
$lblExclude.Size = New-Object System.Drawing.Size(56, 20)
$grpFilter.Controls.Add($lblExclude)

$txtExclude = New-Object System.Windows.Forms.TextBox
$txtExclude.Location = New-Object System.Drawing.Point(520, 22)
$txtExclude.Size = New-Object System.Drawing.Size(388, 22)
$txtExclude.Anchor = "Top,Left,Right"
$txtExclude.Text = $initExclude
$grpFilter.Controls.Add($txtExclude)

# --- export formats ---
$grpExp = New-Object System.Windows.Forms.GroupBox
$grpExp.Text = "4. Export formats (at least one is always produced)"
$grpExp.Location = New-Object System.Drawing.Point(12, 216)
$grpExp.Size = New-Object System.Drawing.Size(920, 44)
$grpExp.Anchor = "Top,Left,Right"
$panelTop.Controls.Add($grpExp)

$fmts = @('CSV','JSON','TXT','HTML','XML')
$chkFmt = @{}
$fx = 12
foreach ($f in $fmts) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $f
    $c.Location = New-Object System.Drawing.Point($fx, 16)
    $c.Size = New-Object System.Drawing.Size(60, 20)
    $c.Checked = ($savedFormats -contains $f)
    $grpExp.Controls.Add($c)
    $chkFmt[$f] = $c
    $fx += 68
}
if (-not ($fmts | Where-Object { $chkFmt[$_].Checked })) { $chkFmt['TXT'].Checked = $true }

$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Text = "Open folder when done"
$chkOpen.Location = New-Object System.Drawing.Point(720, 16)
$chkOpen.Size = New-Object System.Drawing.Size(190, 20)
$chkOpen.Anchor = "Top,Right"
$chkOpen.Checked = [bool]$settings.OpenWhenDone
$grpExp.Controls.Add($chkOpen)

# --- extract group ---
$grpExt = New-Object System.Windows.Forms.GroupBox
$grpExt.Text = "5. Extract files after scan (preserves folder structure)"
$grpExt.Location = New-Object System.Drawing.Point(12, 268)
$grpExt.Size = New-Object System.Drawing.Size(920, 44)
$grpExt.Anchor = "Top,Left,Right"
$panelTop.Controls.Add($grpExt)

$rbENone = New-Object System.Windows.Forms.RadioButton
$rbENone.Text = "None"
$rbENone.Location = New-Object System.Drawing.Point(12, 16); $rbENone.Size = New-Object System.Drawing.Size(60,20); $rbENone.Checked = $true
$grpExt.Controls.Add($rbENone)
$rbEAll = New-Object System.Windows.Forms.RadioButton
$rbEAll.Text = "All files"
$rbEAll.Location = New-Object System.Drawing.Point(80, 16); $rbEAll.Size = New-Object System.Drawing.Size(80,20)
$grpExt.Controls.Add($rbEAll)
$rbEType = New-Object System.Windows.Forms.RadioButton
$rbEType.Text = "By file type (pick after scan)"
$rbEType.Location = New-Object System.Drawing.Point(168, 16); $rbEType.Size = New-Object System.Drawing.Size(200,20)
$grpExt.Controls.Add($rbEType)
$rbESingle = New-Object System.Windows.Forms.RadioButton
$rbESingle.Text = "Hand-pick files (Results selection or picker)"
$rbESingle.Location = New-Object System.Drawing.Point(378, 16); $rbESingle.Size = New-Object System.Drawing.Size(300,20)
$grpExt.Controls.Add($rbESingle)

switch ([string]$settings.ExtractMode) {
    'All'    { $rbEAll.Checked = $true }
    'ByType' { $rbEType.Checked = $true }
    'Single' { $rbESingle.Checked = $true }
}

# --- action buttons ---
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Scan, Export && Extract"
$btnRun.Location = New-Object System.Drawing.Point(12, 320)
$btnRun.Size = New-Object System.Drawing.Size(190, 30)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(14, 90, 138)
$btnRun.ForeColor = [System.Drawing.Color]::White
$panelTop.Controls.Add($btnRun)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(210, 320)
$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
$btnCancel.Enabled = $false
$panelTop.Controls.Add($btnCancel)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Open output folder"
$btnOpen.Location = New-Object System.Drawing.Point(308, 320)
$btnOpen.Size = New-Object System.Drawing.Size(150, 30)
$btnOpen.Enabled = $false
$panelTop.Controls.Add($btnOpen)

# --- status strip ---
$status = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = "Ready."
$lblStatus.Spring = $true
$lblStatus.TextAlign = "MiddleLeft"
$progress = New-Object System.Windows.Forms.ToolStripProgressBar
$progress.Style = "Marquee"
$progress.Width = 180
$progress.MarqueeAnimationSpeed = 30
$progress.Visible = $false
[void]$status.Items.Add($lblStatus)
[void]$status.Items.Add($progress)
$form.Controls.Add($status)
$script:statusLabel = $lblStatus

# --- tabs ---
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(0, 352)
$tabs.Size = New-Object System.Drawing.Size(964, 340)
$tabs.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($tabs)

$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = "Log"
$tabLog.Padding = New-Object System.Windows.Forms.Padding(4)
$tabs.TabPages.Add($tabLog)

$tabResults = New-Object System.Windows.Forms.TabPage
$tabResults.Text = "Results"
$tabResults.Padding = New-Object System.Windows.Forms.Padding(4)
$tabs.TabPages.Add($tabResults)

$tabTypes = New-Object System.Windows.Forms.TabPage
$tabTypes.Text = "File types"
$tabTypes.Padding = New-Object System.Windows.Forms.Padding(4)
$tabs.TabPages.Add($tabTypes)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Both"
$txtLog.WordWrap = $false
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.Dock = "Fill"
$tabLog.Controls.Add($txtLog)
$script:logBox = $txtLog

$resultsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$resultsLayout.Dock = "Fill"
$resultsLayout.ColumnCount = 1
$resultsLayout.RowCount = 2
[void]$resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
[void]$resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabResults.Controls.Add($resultsLayout)

$filterLayout = New-Object System.Windows.Forms.TableLayoutPanel
$filterLayout.Dock = "Fill"
$filterLayout.ColumnCount = 2
$filterLayout.RowCount = 1
[void]$filterLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
[void]$filterLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.TextAlign = "MiddleLeft"
$lblFilter.Dock = "Fill"
$gridFilter = New-Object System.Windows.Forms.TextBox
$gridFilter.Anchor = "Left,Right"
$gridFilter.Margin = New-Object System.Windows.Forms.Padding(3, 6, 8, 5)
$filterLayout.Controls.Add($lblFilter, 0, 0)
$filterLayout.Controls.Add($gridFilter, 1, 0)
$resultsLayout.Controls.Add($filterLayout, 0, 0)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = "None"
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = "None"
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 246, 251)
$resultsLayout.Controls.Add($grid, 0, 1)

$typesGrid = New-Object System.Windows.Forms.DataGridView
$typesGrid.Dock = "Fill"
$typesGrid.ReadOnly = $true
$typesGrid.AllowUserToAddRows = $false
$typesGrid.AllowUserToDeleteRows = $false
$typesGrid.AllowUserToResizeRows = $false
$typesGrid.RowHeadersVisible = $false
$typesGrid.SelectionMode = "FullRowSelect"
$typesGrid.MultiSelect = $false
$typesGrid.AutoSizeColumnsMode = "None"
$typesGrid.BackgroundColor = [System.Drawing.Color]::White
$typesGrid.BorderStyle = "None"
$typesGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 246, 251)
$tabTypes.Controls.Add($typesGrid)

$formatCells = {
    param($sender, $e)
    try {
        if ($e.Value -is [System.DBNull]) { return }
        $col = $sender.Columns[$e.ColumnIndex]
        if ($null -eq $col) { return }
        switch ($col.Name) {
            'SizeBytes' { $e.Value = Format-Size ([long]$e.Value); $e.FormattingApplied = $true }
            'TotalSize' { $e.Value = Format-Size ([long]$e.Value); $e.FormattingApplied = $true }
            'Modified'  { $e.Value = ([datetime]$e.Value).ToString('yyyy-MM-dd HH:mm'); $e.FormattingApplied = $true }
        }
    } catch { }
}
$grid.Add_CellFormatting($formatCells)
$typesGrid.Add_CellFormatting($formatCells)

# ---------- drag & drop ----------
$form.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})
$form.Add_DragDrop({
    param($sender, $e)
    $items = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($null -eq $items -or $items.Count -eq 0) { return }
    $p = [string]$items[0]
    if (Test-Path -LiteralPath $p -PathType Container) {
        $rbFolder.Checked = $true; $rbZip.Checked = $false
    } else {
        $rbZip.Checked = $true; $rbFolder.Checked = $false
    }
    $txtSource.Text = $p
})

# ---------- browse handlers ----------
$btnSrc.Add_Click({
    if ($rbZip.Checked) {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "$appName - Select a ZIP archive"
        $dlg.Filter = "Archives (*.zip;*.jar;*.war;*.apk;*.docx;*.xlsx)|*.zip;*.jar;*.war;*.apk;*.docx;*.xlsx|All files (*.*)|*.*"
        if ($dlg.ShowDialog() -eq "OK") { $txtSource.Text = $dlg.FileName }
    } else {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "$appName - Select a folder to scan"
        $dlg.ShowNewFolderButton = $false
        if ($dlg.ShowDialog() -eq "OK") { $txtSource.Text = $dlg.SelectedPath }
    }
})
$btnOut.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "$appName - Select output folder"
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq "OK") { $txtOutput.Text = $dlg.SelectedPath }
})
$btnOpen.Add_Click({
    if ($script:outputDir -ne "" -and (Test-Path -LiteralPath $script:outputDir)) {
        Start-Process explorer.exe $script:outputDir
    }
})
$btnCancel.Add_Click({
    $script:cancelRequested = $true
    $lblStatus.Text = "Cancelling..."
})

# ---------- grid handlers ----------
$gridFilter.Add_TextChanged({ Apply-ResultsFilter })
$typesGrid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $ext = [string]$typesGrid.Rows[$e.RowIndex].Cells["Extension"].Value
    if ($ext -ne "") {
        $tabs.SelectedTab = $tabResults
        $gridFilter.Text = $ext
        $gridFilter.Focus()
    }
})

# ---------- layout ----------
$form.Add_Shown({ Update-Layout })
$form.Add_Resize({ Update-Layout })

# ---------- RUN button ----------
$btnRun.Add_Click({
    $script:cancelRequested = $false
    $script:tick = 0
    $script:isBusy = $true
    $btnRun.Enabled = $false
    $btnCancel.Enabled = $true
    $progress.Visible = $true
    $lblStatus.Text = "Working..."
    $form.Cursor = "WaitCursor"
    $txtLog.Clear()
    $script:resultsView = $null
    $grid.DataSource = $null
    $typesGrid.DataSource = $null
    try {
        $kind = $(if ($rbZip.Checked) { "Zip" } else { "Folder" })
        Initialize-Context -SourcePath $txtSource.Text.Trim() -Kind $kind -OutputPath $txtOutput.Text.Trim() -IncludeMask $txtInclude.Text.Trim() -ExcludeMask $txtExclude.Text.Trim()

        $fmtsChosen = @($fmts | Where-Object { $chkFmt[$_].Checked })
        if ($fmtsChosen.Count -eq 0) { $fmtsChosen = @('TXT') }

        $mode = "None"
        if ($rbEAll.Checked) { $mode = "All" }
        elseif ($rbEType.Checked) { $mode = "ByType" }
        elseif ($rbESingle.Checked) { $mode = "Single" }

        $selectedPaths = @()
        if ($mode -eq "Single" -and $null -ne $grid.SelectedRows) {
            foreach ($row in $grid.SelectedRows) {
                $pv = $row.Cells["Path"].Value
                if ($null -ne $pv -and [string]$pv -ne "") { $selectedPaths += [string]$pv }
            }
        }

        Invoke-Pipeline -Formats $fmtsChosen -Mode $mode -FilesToExtract $selectedPaths -OpenWhenDone:$chkOpen.Checked

        $btnOpen.Enabled = $true
        Save-Settings (Get-CurrentSettings)
        $lblStatus.Text = ("Done. {0:N0} entries, {1} export(s)." -f $script:entries.Count, $script:lastExports.Count)
    }
    catch {
        if ($_.Exception -is [System.OperationCanceledException]) {
            Log "Cancelled."
            $lblStatus.Text = "Cancelled."
        } else {
            Log "[ERROR] $_"
            $lblStatus.Text = "Failed - see log."
            [System.Windows.Forms.MessageBox]::Show("$_", "$appName - Error", "OK", "Error") | Out-Null
        }
    }
    finally {
        $script:isBusy = $false
        $form.Cursor = "Default"
        $btnRun.Enabled = $true
        $btnCancel.Enabled = $false
        $progress.Visible = $false
        Close-ZipArchive
    }
})

# ---------- form close ----------
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:isBusy) {
        $script:cancelRequested = $true
        $e.Cancel = $true
        return
    }
    try { Save-Settings (Get-CurrentSettings) } catch { }
    Close-ZipArchive
})

# ---------- optional smoke test ----------
if ($SmokeTest) {
    $script:smokeTimer = New-Object System.Windows.Forms.Timer
    $script:smokeTimer.Interval = 1500
    $script:smokeTimer.Add_Tick({
        $script:smokeTimer.Stop()
        $form.Close()
    })
    $form.Add_Shown({
        if (-not [string]::IsNullOrWhiteSpace($txtSource.Text)) {
            $script:smokeTimer.Stop()
            foreach ($f in $fmts) { $chkFmt[$f].Checked = ($f -eq 'TXT') }
            $rbENone.Checked = $true
            $chkOpen.Checked = $false
            try { $btnRun.PerformClick() } catch { }
            $script:smokeTimer.Start()
        }
    })
    $script:smokeTimer.Start()
}

[void]$form.ShowDialog()
exit 0
