#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 Apps - Office Update Channel Analyzer

.DESCRIPTION
    Analyzes Office C2R registry settings to determine the active update
    channel, identify configuration conflicts, and provide a root cause
    analysis for channel switching issues.

    Supports two modes:
      LIVE MODE   - Run without -Path to query the local machine's registry
                    directly (requires M365 Apps to be installed).
      FILE MODE   - Provide -Path to analyze exported .txt registry files
                    from diagnostic collections (SaRA, OfficeDiag, etc.).

    Outputs a colored console report (priority table, findings, action plan)
    AND an interactive HTML report (OfficePolicies.html).

.PARAMETER Path
    (Optional) Path to the folder containing Office diagnostic files. The
    folder should contain a 'Registry' subfolder with .txt registry export
    files, or be the Registry folder itself.
    If omitted, the tool runs in live mode against the local registry.

.EXAMPLE
    .\OfficeChannelAnalyzer.ps1
    .\OfficeChannelAnalyzer.ps1 -Path "C:\Logs\CustomerMachine"
    OfficeChannelAnalyzer.exe "C:\Logs\mattphil"
    OfficeChannelAnalyzer.exe /?

.NOTES
    Reference: https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels
#>

param(
    [Parameter(Position = 0)]
    [string]$Path
)

# ─────────────────────────────────────────────────────────────────────────────
# USAGE / HELP
# ─────────────────────────────────────────────────────────────────────────────
function Show-Usage {
    $w = 72
    Write-Host ""
    Write-Host ("  " + ([string][char]0x2550) * $w) -ForegroundColor DarkCyan
    Write-Host "  Microsoft 365 Apps - Office Update Channel Analyzer" -ForegroundColor Cyan
    Write-Host ("  " + ([string][char]0x2550) * $w) -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  Analyzes Office C2R registry settings to determine the active update" -ForegroundColor White
    Write-Host "  channel, identify blocking configurations, and surface root causes" -ForegroundColor White
    Write-Host "  for channel switching failures (e.g. Semi-Annual to Monthly Enterprise)." -ForegroundColor White
    Write-Host ""
    Write-Host "  MODES:" -ForegroundColor Yellow
    Write-Host "    LIVE MODE   Run without -Path to query the local machine's registry." -ForegroundColor White
    Write-Host "    FILE MODE   Provide -Path to analyze exported registry files." -ForegroundColor White
    Write-Host ""
    Write-Host "  USAGE:" -ForegroundColor Yellow
    Write-Host "    OfficeChannelAnalyzer.exe                          (live mode)" -ForegroundColor White
    Write-Host "    OfficeChannelAnalyzer.exe <FolderPath>             (file mode)" -ForegroundColor White
    Write-Host "    .\OfficeChannelAnalyzer.ps1                        (live mode)" -ForegroundColor White
    Write-Host "    .\OfficeChannelAnalyzer.ps1 -Path <FolderPath>     (file mode)" -ForegroundColor White
    Write-Host ""
    Write-Host "  EXAMPLES:" -ForegroundColor Yellow
    Write-Host '    OfficeChannelAnalyzer.exe' -ForegroundColor Gray
    Write-Host '    OfficeChannelAnalyzer.exe "C:\Logs\CustomerMachine"' -ForegroundColor Gray
    Write-Host '    OfficeChannelAnalyzer.exe .\mattphil' -ForegroundColor Gray
    Write-Host '    OfficeChannelAnalyzer.exe .\mattphil\Registry' -ForegroundColor Gray
    Write-Host '    OfficeChannelAnalyzer.exe /?' -ForegroundColor Gray
    Write-Host ""
    Write-Host "  OUTPUT:" -ForegroundColor Yellow
    Write-Host "    - Colored console report with channel priority table" -ForegroundColor White
    Write-Host "    - Findings and root cause analysis" -ForegroundColor White
    Write-Host "    - Action plan with remediation steps" -ForegroundColor White
    Write-Host "    - OfficePolicies.html (interactive, opens in browser)" -ForegroundColor White
    Write-Host ""
    Write-Host "  CHANNEL PRIORITY ORDER (first configured value wins):" -ForegroundColor Yellow
    Write-Host "    1  Cloud Update   UpdatePath    (HKLM\...\cloud\office\...\officeupdate)" -ForegroundColor Gray
    Write-Host "    2  Cloud Update   UpdateBranch  (HKLM\...\cloud\office\...\officeupdate)" -ForegroundColor Gray
    Write-Host "    3  Policy/GPO     UpdatePath    (HKLM\...\office\16.0\...\officeupdate)" -ForegroundColor Gray
    Write-Host "    4  Policy/GPO     UpdateBranch  (HKLM\...\office\16.0\...\officeupdate)" -ForegroundColor Gray
    Write-Host "    5  ODT            UpdateUrl     (HKLM\...\ClickToRun\Configuration)" -ForegroundColor Gray
    Write-Host "    6  Unmanaged      UnmanagedUpdateUrl (HKLM\...\ClickToRun\Configuration)" -ForegroundColor Gray
    Write-Host "    7  Unmanaged      CDNBaseUrl    (HKLM\...\ClickToRun\Configuration)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  REFERENCE:" -ForegroundColor Yellow
    Write-Host "    https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host ("  " + ([string][char]0x2550) * $w) -ForegroundColor DarkCyan
    Write-Host ""
}

if ($Path -in @('/?', '-?', '--help', '-h', '-help', '/help')) {
    Show-Usage
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
# MODE DETECTION: LIVE vs FILE
# ─────────────────────────────────────────────────────────────────────────────
$liveMode = $false

if (-not $Path) {
    # No path provided — attempt live registry mode
    $c2rKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (-not (Test-Path $c2rKey)) {
        Write-Host ""
        Write-Host "  Microsoft 365 Apps (Click-to-Run) is not installed on this machine." -ForegroundColor Red
        Write-Host "  To analyze exported registry files, provide a path:" -ForegroundColor Yellow
        Write-Host "    .\OfficeChannelAnalyzer.ps1 -Path <FolderPath>" -ForegroundColor Gray
        Write-Host "  Run with /? for full usage help." -ForegroundColor Yellow
        Write-Host ""
        exit 1
    }
    $liveMode = $true
    $folderLabel  = $env:COMPUTERNAME
    $outputFolder = $PWD.Path
    Write-Host ""
    Write-Host "  LIVE MODE: Reading registry from local machine ($folderLabel)..." -ForegroundColor Cyan
}

# ─────────────────────────────────────────────────────────────────────────────
# FILE MODE: RESOLVE PATHS
# ─────────────────────────────────────────────────────────────────────────────
$regFolder = $null
if (-not $liveMode) {
    if (-not (Test-Path $Path)) {
        Write-Host "  ERROR: Path not found: '$Path'" -ForegroundColor Red
        Write-Host "  Run with /? for usage help." -ForegroundColor Yellow
        exit 1
    }

    if (Test-Path (Join-Path $Path 'Registry') -PathType Container) {
        $regFolder = Join-Path $Path 'Registry'
    } elseif ((Get-Item $Path).PSIsContainer -and (Get-ChildItem $Path -Filter '*.txt' -ErrorAction SilentlyContinue)) {
        $regFolder = $Path
    }

    if (-not $regFolder) {
        Write-Host "  ERROR: No 'Registry' subfolder or .txt registry exports found in '$Path'." -ForegroundColor Red
        Write-Host "  Ensure the folder contains Office diagnostic files with a 'Registry' subfolder." -ForegroundColor Yellow
        exit 1
    }

    $outputFolder = if ($Path -ne $regFolder) { $Path } else { Split-Path $regFolder -Parent }
    $folderLabel  = Split-Path (Resolve-Path $Path).Path -Leaf
}

# ─────────────────────────────────────────────────────────────────────────────
# CHANNEL NAME MAPPINGS
# ─────────────────────────────────────────────────────────────────────────────
$channelGuids = @{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current Channel (CC)'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current Channel Preview'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise Channel (MEC)'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise Channel (SAEC)'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise Channel Preview [DEPRECATED]'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta Channel'
}
$channelBranches = @{
    'current'           = 'Current Channel (CC)'
    'currentpreview'    = 'Current Channel Preview'
    'monthlyenterprise' = 'Monthly Enterprise Channel (MEC)'
    'semiannual'        = 'Semi-Annual Enterprise Channel (SAEC)'
    'semiannualpreview' = 'Semi-Annual Enterprise Channel Preview [DEPRECATED]'
    'betachannel'       = 'Beta Channel'
    'beta'              = 'Beta Channel'
    'insiders'          = 'Beta Channel'
}

# PS5-compatible null-coalescing helper
function Nvl { param($a, $b = '-') if ($null -ne $a -and "$a" -ne '') { $a } else { $b } }

function Resolve-ChannelName {
    param([string]$Value)
    if (-not $Value) { return $null }
    if ($Value -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
        $g = $Matches[1].ToLower()
        if ($channelGuids.ContainsKey($g)) { return $channelGuids[$g] }
        return "Unknown Channel (GUID: $g)"
    }
    $k = $channelBranches[$Value.ToLower()]
    if ($k) { return $k }
    return $Value
}

# ─────────────────────────────────────────────────────────────────────────────
# PARSE REGISTRY: FILE MODE or LIVE MODE
# ─────────────────────────────────────────────────────────────────────────────
function Parse-RegFile {
    param([string]$FilePath)
    $entries = @()
    $key = ''
    try {
        foreach ($line in (Get-Content -Path $FilePath -Encoding Unicode -ErrorAction Stop)) {
            $line = $line.Trim()
            if (-not $line -or $line -match '^Windows Registry') { continue }
            if ($line.StartsWith('[') -and $line.EndsWith(']')) {
                $key = $line.Substring(1, $line.Length - 2)
            } elseif ($line -match '^"([^"]+)"\s*=\s*(.+)$' -and $key) {
                $raw = $Matches[2].Trim()
                $val = if ($raw.StartsWith('hex:')) { '(binary data)' } else { $raw.Trim('"') }
                $entries += [PSCustomObject]@{ Key = $key; Name = $Matches[1].ToLower(); Value = $val }
            }
        }
    } catch { }
    return $entries
}

function Read-LiveRegistry {
    param([string]$HklmPath)
    $entries = @()
    $psPath = "HKLM:\$HklmPath"
    if (Test-Path $psPath) {
        $props = Get-ItemProperty -Path $psPath -ErrorAction SilentlyContinue
        if ($props) {
            $skip = @('PSPath','PSParentPath','PSChildName','PSProvider','PSDrive')
            foreach ($name in $props.PSObject.Properties.Name) {
                if ($name -in $skip) { continue }
                $val = $props.$name
                $valStr = if ($val -is [int] -or $val -is [long]) {
                    'dword:{0:x8}' -f $val
                } elseif ($val -is [byte[]]) {
                    '(binary data)'
                } else {
                    "$val"
                }
                $entries += [PSCustomObject]@{
                    Key   = "HKEY_LOCAL_MACHINE\$HklmPath"
                    Name  = $name.ToLower()
                    Value = $valStr
                }
            }
        }
    }
    return $entries
}

$allEntries = @()

if ($liveMode) {
    # Read all relevant registry paths directly
    $livePaths = @(
        'SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate'
        'SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate'
        'SOFTWARE\Microsoft\office\ClickToRun\Configuration'
        'SOFTWARE\Microsoft\office\ClickToRun\Updates'
    )
    foreach ($lp in $livePaths) {
        $allEntries += Read-LiveRegistry $lp
    }
    if ($allEntries.Count -eq 0) {
        Write-Host "  WARNING: No registry values found. Office may not be fully configured." -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "  Scanning: $regFolder ..." -ForegroundColor DarkGray
    foreach ($f in (Get-ChildItem $regFolder -Filter '*.txt' -ErrorAction SilentlyContinue | Where-Object { $_.Length -lt 500KB } | Sort-Object Name)) {
        $allEntries += Parse-RegFile $f.FullName
    }
    if ($allEntries.Count -eq 0) {
        Write-Host "  ERROR: No registry data parsed. Ensure files are UTF-16LE encoded .txt reg exports." -ForegroundColor Red
        exit 1
    }
}

function Get-RegValue {
    param([string]$KeyPath, [string]$Name)
    $k = $KeyPath -replace '^HKLM\\','HKEY_LOCAL_MACHINE\' -replace '^HKCU\\','HKEY_CURRENT_USER\'
    ($allEntries | Where-Object { $_.Key -ieq $k -and $_.Name -ieq $Name.ToLower() } | Select-Object -First 1).Value
}

function Get-RegEntriesForPath {
    param([string]$KeyPath)
    $k = $KeyPath -replace '^HKLM\\','HKEY_LOCAL_MACHINE\' -replace '^HKCU\\','HKEY_CURRENT_USER\'
    $allEntries | Where-Object { $_.Key -ieq $k } | Sort-Object Name
}

# ─────────────────────────────────────────────────────────────────────────────
# CHANNEL PRIORITY EVALUATION
# ─────────────────────────────────────────────────────────────────────────────
$cloudPath   = 'HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate'
$gpoPath     = 'HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate'
$configPath  = 'HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration'
$updatesPath = 'HKLM\SOFTWARE\Microsoft\office\ClickToRun\Updates'

$priorities = @(
    [PSCustomObject]@{ P=1; Mgmt='Cloud Update';   RegVal='UpdatePath';         RegPath=$cloudPath }
    [PSCustomObject]@{ P=2; Mgmt='Cloud Update';   RegVal='UpdateBranch';       RegPath=$cloudPath }
    [PSCustomObject]@{ P=3; Mgmt='Policy/GPO';     RegVal='UpdatePath';         RegPath=$gpoPath }
    [PSCustomObject]@{ P=4; Mgmt='Policy/GPO';     RegVal='UpdateBranch';       RegPath=$gpoPath }
    [PSCustomObject]@{ P=5; Mgmt='ODT';            RegVal='UpdateUrl';          RegPath=$configPath }
    [PSCustomObject]@{ P=6; Mgmt='Unmanaged';      RegVal='UnmanagedUpdateUrl'; RegPath=$configPath }
    [PSCustomObject]@{ P=7; Mgmt='Unmanaged';      RegVal='CDNBaseUrl';         RegPath=$configPath }
)

$activeRow = $null
foreach ($row in $priorities) {
    $val = Get-RegValue $row.RegPath $row.RegVal
    $row | Add-Member NoteProperty DetectedValue $val          -Force
    $row | Add-Member NoteProperty ChannelName   (Resolve-ChannelName $val) -Force
    if (-not $activeRow -and $val) { $activeRow = $row }
}

# ─────────────────────────────────────────────────────────────────────────────
# READ KEY SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
$enableAutoRaw      = Get-RegValue $gpoPath   'enableautomaticupdates'
$ignoreGpoRaw       = Get-RegValue $cloudPath 'ignoregpo'
$mgmtComRaw         = Get-RegValue $gpoPath   'officemgmtcom'
$hideRaw            = Get-RegValue $gpoPath   'hideenabledisableupdates'
$updateChannel      = Get-RegValue $configPath 'UpdateChannel'
$updateChannelChg   = Get-RegValue $configPath 'UpdateChannelChanged'
$audienceData       = Get-RegValue $configPath 'AudienceData'
$audienceId         = Get-RegValue $configPath 'AudienceId'
$currentVersion     = Get-RegValue $configPath 'VersionToReport'
$cdnBaseUrl         = Get-RegValue $configPath 'CDNBaseUrl'
$updatesEnabled     = Get-RegValue $configPath 'UpdatesEnabled'
$platform           = Get-RegValue $configPath 'Platform'
$installPath        = Get-RegValue $configPath 'InstallationPath'
$skuBlocked         = Get-RegValue $updatesPath 'UpdatesSkuToSkuBlocked'

function Get-Dword { param([string]$v)
    if ($v -match 'dword:([0-9a-f]+)') { [Convert]::ToInt32($Matches[1], 16) } else { $null }
}
$autoUpdatesDword = Get-Dword $enableAutoRaw
$ignoreGpoDword   = Get-Dword $ignoreGpoRaw
$mgmtComDword     = Get-Dword $mgmtComRaw

# ─────────────────────────────────────────────────────────────────────────────
# CONSOLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
$CW = 96

function Write-Banner {
    param([string]$Title, [string]$Sub = '')
    $top = "  " + [string][char]0x2554 + ([string][char]0x2550) * ($CW - 4) + [string][char]0x2557
    $bot = "  " + [string][char]0x255A + ([string][char]0x2550) * ($CW - 4) + [string][char]0x255D
    $mid = [string][char]0x2551
    function Pad-Center { param([string]$s)
        $inner = $CW - 4
        $l = [int][Math]::Floor(($inner - $s.Length) / 2)
        $r = $inner - $s.Length - $l
        "  $mid" + (" " * $l) + $s + (" " * $r) + $mid
    }
    Write-Host ""
    Write-Host $top -ForegroundColor DarkCyan
    Write-Host (Pad-Center $Title) -ForegroundColor Cyan
    if ($Sub) { Write-Host (Pad-Center $Sub) -ForegroundColor DarkGray }
    Write-Host $bot -ForegroundColor DarkCyan
}

function Write-Section { param([string]$Title)
    Write-Host ""
    Write-Host "  $Title" -ForegroundColor Yellow
    Write-Host ("  " + ([string][char]0x2500) * ($CW - 4)) -ForegroundColor DarkGray
}

function Write-KV { param([string]$Key, [string]$Val, [string]$Note = '', [string]$Color = 'White')
    Write-Host ("    " + $Key.PadRight(32)) -NoNewline -ForegroundColor DarkGray
    Write-Host $Val -NoNewline -ForegroundColor $Color
    if ($Note) { Write-Host "  <- $Note" -ForegroundColor DarkYellow } else { Write-Host "" }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER & MACHINE SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
$modeLabel = if ($liveMode) { "Live Registry: $folderLabel" } else { "Folder: $folderLabel" }
Write-Banner "Microsoft 365 Apps  -  Update Channel Analyzer" $modeLabel

Write-Section "MACHINE SUMMARY"
$activeCh = if ($activeRow) { $activeRow.ChannelName } else { 'Unknown' }
$targetCh = Resolve-ChannelName $updateChannel
$cdnName  = Resolve-ChannelName $cdnBaseUrl

Write-KV "Source:"                 $(if ($liveMode) { "Live registry ($folderLabel)" } else { $folderLabel })
Write-KV "Office version:"        (Nvl $currentVersion)
Write-KV "Platform (bitness):"    (Nvl $platform)
Write-KV "Install path:"          (Nvl $installPath)

$_note  = if ($activeCh -match 'Semi-Annual') { 'Not Monthly Enterprise!' } else { '' }
$_color = if ($activeCh -match 'Semi-Annual') { 'Red' } elseif ($activeCh -match 'Monthly') { 'Green' } else { 'Yellow' }
Write-KV "Active channel:"        (Nvl $activeCh)     $_note  $_color

$_color = if ($targetCh -match 'Monthly') { 'Green' } else { 'Gray' }
Write-KV "UpdateChannel (target):" (Nvl $targetCh)    ''      $_color

$_note  = if ($updateChannelChg -ieq 'False') { 'Switch not applied!' } else { '' }
$_color = if ($updateChannelChg -ieq 'False') { 'Red' } elseif ($updateChannelChg -ieq 'True') { 'Green' } else { 'Gray' }
Write-KV "UpdateChannelChanged:"  (Nvl $updateChannelChg) $_note $_color

$_color = if ($cdnName -match 'Semi-Annual') { 'Yellow' } else { 'White' }
Write-KV "CDNBaseUrl (installed):" (Nvl $cdnName)     ''      $_color

$_color = if ($audienceData -match 'MEC') { 'Cyan' } else { 'Gray' }
Write-KV "Audience data:"         (Nvl $audienceData) ''      $_color

$_color = if ($updatesEnabled -ieq 'True') { 'Green' } else { 'Yellow' }
Write-KV "Updates enabled:"       (Nvl $updatesEnabled) ''    $_color

# ─────────────────────────────────────────────────────────────────────────────
# CHANNEL PRIORITY TABLE
# ─────────────────────────────────────────────────────────────────────────────
Write-Section "CHANNEL PRIORITY TABLE  (Priority 1 wins - first configured value determines active channel)"
Write-Host "  Ref: https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels" -ForegroundColor DarkGray
Write-Host ""

# Column widths
$cP = 4; $cM = 16; $cR = 22; $cV = 40; $cS = 10; $cA = 9

function Table-Row { param([string]$p,[string]$m,[string]$r,[string]$v,[string]$s,[string]$a)
    "  |" + $p.PadRight($cP) + "|" + $m.PadRight($cM) + "|" + $r.PadRight($cR) + "|" + $v.PadRight($cV) + "|" + $s.PadRight($cS) + "|" + $a.PadRight($cA) + "|"
}
$div = "  +" + ("-"*$cP) + "+" + ("-"*$cM) + "+" + ("-"*$cR) + "+" + ("-"*$cV) + "+" + ("-"*$cS) + "+" + ("-"*$cA) + "+"

Write-Host $div -ForegroundColor DarkGray
Write-Host (Table-Row " Pri" " Management" " Registry Value" " Channel / Detected Value" " Status" " Active?") -ForegroundColor White
Write-Host $div -ForegroundColor DarkGray

$mgmtColor = @{ 'Cloud Update'='Magenta'; 'Policy/GPO'='DarkYellow'; 'ODT'='Cyan'; 'Unmanaged'='DarkCyan' }

foreach ($row in $priorities) {
    $isActive = ($activeRow -and $activeRow.P -eq $row.P)
    $hasVal   = ($row.DetectedValue -and $row.DetectedValue -ne '')
    $display  = if ($hasVal -and $row.ChannelName) { $row.ChannelName } elseif ($hasVal) { $row.DetectedValue } else { '-' }
    if ($display.Length -gt $cV - 1) { $display = $display.Substring(0, $cV - 4) + '...' }

    $pStr = (" " + $row.P).PadRight($cP)
    $mStr = (" " + $row.Mgmt).PadRight($cM)
    $rStr = (" " + $row.RegVal).PadRight($cR)
    $vStr = (" " + $display).PadRight($cV)
    $sStr = if ($hasVal) { " [SET]    " } else { " Not Set  " }
    $aStr = if ($isActive) { " * ACTIVE" } else { "         " }

    $pCol = if ($isActive) { 'Green' } else { 'DarkGray' }
    $mCol = if ($isActive) { 'Green' } else { $mgmtColor[$row.Mgmt] }
    $rCol = if ($isActive) { 'Green' } else { 'Gray' }
    $vCol = if ($isActive) { 'Green' } elseif ($hasVal) { 'Cyan' } else { 'DarkGray' }
    $sCol = if ($hasVal) { 'Green' } else { 'DarkGray' }
    $aCol = if ($isActive) { 'Green' } else { 'DarkGray' }

    Write-Host "  |" -NoNewline -ForegroundColor DarkGray
    Write-Host $pStr -NoNewline -ForegroundColor $pCol
    Write-Host "|" -NoNewline -ForegroundColor DarkGray
    Write-Host $mStr -NoNewline -ForegroundColor $mCol
    Write-Host "|" -NoNewline -ForegroundColor DarkGray
    Write-Host $rStr -NoNewline -ForegroundColor $rCol
    Write-Host "|" -NoNewline -ForegroundColor DarkGray
    Write-Host $vStr -NoNewline -ForegroundColor $vCol
    Write-Host "|" -NoNewline -ForegroundColor DarkGray
    Write-Host $sStr -NoNewline -ForegroundColor $sCol
    Write-Host "|" -NoNewline -ForegroundColor DarkGray
    Write-Host $aStr -NoNewline -ForegroundColor $aCol
    Write-Host "|" -ForegroundColor DarkGray
}
Write-Host $div -ForegroundColor DarkGray

if ($activeRow) {
    $activeChDisplay = if ($activeRow.ChannelName) { $activeRow.ChannelName } else { $activeRow.DetectedValue }
    Write-Host "  Active channel source: Priority $($activeRow.P) [$($activeRow.Mgmt)] -> $activeChDisplay" -ForegroundColor $(if ($activeRow.ChannelName -match 'Monthly Enterprise') { 'Green' } else { 'Yellow' })
} else {
    Write-Host "  WARNING: No active channel could be determined from the registry data." -ForegroundColor Red
}

# ─────────────────────────────────────────────────────────────────────────────
# ADDITIONAL SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
Write-Section "ADDITIONAL CONFIGURATION SETTINGS"

$autoStr   = if ($autoUpdatesDword -eq 0) { "DISABLED (dword:00000000)" } elseif ($autoUpdatesDword -eq 1) { "Enabled (dword:00000001)" } else { Nvl $enableAutoRaw 'Not Set' }
$autoCol   = if ($autoUpdatesDword -eq 0) { 'Red' } elseif ($autoUpdatesDword -eq 1) { 'Green' } else { 'DarkGray' }
$autoNote  = if ($autoUpdatesDword -eq 0) { "GPO is blocking all automatic updates!" } else { '' }
Write-KV "Auto updates (GPO):"    $autoStr    $autoNote  $autoCol

$igStr     = if ($ignoreGpoDword -eq 1) { "YES - Cloud Update overrides GPO (dword:1)" } elseif ($ignoreGpoDword -eq 0) { "NO - Cloud Update respects GPO (dword:0)" } else { 'Not Set' }
$igCol     = if ($ignoreGpoDword -eq 1) { 'Green' } elseif ($ignoreGpoDword -eq 0) { 'Yellow' } else { 'DarkGray' }
$igNote    = if ($ignoreGpoDword -eq 0 -and $autoUpdatesDword -eq 0) { "Cloud Update blocked by auto-update GPO" } else { '' }
Write-KV "Cloud ignoregpo:"       $igStr      $igNote    $igCol

$mcStr     = if ($mgmtComDword -eq 1) { "Enabled (dword:1) - SCCM/management triggers updates" } elseif ($mgmtComDword -eq 0) { "Disabled" } else { 'Not Set' }
$mcCol     = if ($mgmtComDword -eq 1) { 'Cyan' } else { 'DarkGray' }
Write-KV "OfficeMgmtCOM:"         $mcStr      ''         $mcCol

$hideStr   = if ($hideRaw -match 'dword:00000001') { "Yes - update toggle hidden from users" } elseif ($hideRaw) { "No" } else { 'Not Set' }
Write-KV "Hide update UI:"        $hideStr

if ($skuBlocked -and $skuBlocked -ne '0') {
    Write-KV "SKU-to-SKU blocked:"   $skuBlocked '' 'Red'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINDINGS & ROOT CAUSE ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
$findings   = [System.Collections.Generic.List[PSCustomObject]]::new()
$actionPlan = [System.Collections.Generic.List[string]]::new()
$step = 0

function Add-Finding { param([string]$Sev, [string]$Msg)
    $findings.Add([PSCustomObject]@{ Sev = $Sev; Msg = $Msg })
}
function Add-Step { param([string]$Header, [string[]]$Lines)
    $script:step++
    $actionPlan.Add("[$script:step] $Header")
    foreach ($l in $Lines) { $actionPlan.Add("    $l") }
}

# F1 – active channel is not MEC
if ($activeRow -and $activeRow.ChannelName -notmatch 'Monthly Enterprise') {
    Add-Finding 'ERROR' "Active channel is '$($activeRow.ChannelName)' (Priority $($activeRow.P), $($activeRow.Mgmt)) - target is Monthly Enterprise Channel."
}

# F2 – CDNBaseUrl is SAC
if ($cdnBaseUrl -match '7ffbc6bf') {
    Add-Finding 'ERROR' "CDNBaseUrl = Semi-Annual Enterprise Channel (SAC). Office was originally installed using a SAC config file (e.g., via SCCM with Channel=SemiAnnual), which stamped SAC into the registry as the base CDN."
    Add-Step "CHANNEL SOURCE - Re-deploy Office with MEC config or configure a management policy" @(
        "Option A (GPO/Intune): Set UpdateBranch = MonthlyEnterprise at Priority 4:",
        "  HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate",
        "Option B (Cloud Update): Set UpdateBranch = MonthlyEnterprise at Priority 2:",
        "  HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate",
        "Option C (ODT/SCCM): Re-deploy with <Updates Channel='MonthlyEnterprise'/> in config XML."
    )
}

# F3 – no channel policy at priorities 1-4
if (-not ($priorities | Where-Object { $_.P -le 4 -and $_.DetectedValue })) {
    Add-Finding 'ERROR' "No channel management policy is configured at priorities 1-4. CDNBaseUrl (Priority 7) wins by default - no management authority is directing the channel to MEC."
}

# F4 – auto updates disabled by GPO
if ($autoUpdatesDword -eq 0) {
    Add-Finding 'ERROR' "GPO has DISABLED automatic updates (enableautomaticupdates=dword:0 at $gpoPath). Even if a channel policy is set, the Office update engine will not run."
    Add-Step "AUTO UPDATES GPO - Re-enable automatic updates" @(
        "Option A: Change the GPO setting enableautomaticupdates to dword:1 (or remove it).",
        "  Path: $gpoPath",
        "Option B: If Cloud Update manages this device, set ignoregpo=dword:1 in:",
        "  $cloudPath",
        "  This allows Cloud Update to override the GPO-disabled auto-update block."
    )
}

# F5 – ignoregpo=0 with auto updates disabled
if ($ignoreGpoDword -eq 0 -and $autoUpdatesDword -eq 0) {
    Add-Finding 'WARN' "Cloud Update respects GPO (ignoregpo=dword:0) and CANNOT override the disabled auto-updates policy. Cloud Update is aware of MEC target but is blocked from completing the switch."
    Add-Step "CLOUD UPDATE ignoregpo - Allow Cloud Update to override GPO" @(
        "Set ignoregpo=dword:00000001 in the Cloud Update policy key:",
        "  $cloudPath",
        "This allows Cloud Update to bypass the GPO auto-update block."
    )
}

# F6 – UpdateChannelChanged = False
if ($updateChannelChg -ieq 'False') {
    Add-Finding 'WARN' "UpdateChannelChanged=False - a channel switch has been requested/staged but has never been successfully applied. The Office client has not yet consumed the new channel setting."
}

# F7 – OfficeMgmtCOM = 1
if ($mgmtComDword -eq 1) {
    Add-Finding 'INFO' "OfficeMgmtCOM=1 - Office is configured to receive update triggers via SCCM/ConfigMgr COM interface. Office will NOT update autonomously; it waits for the management system trigger."
    Add-Step "SCCM/CONFIGMGR - Verify management system is deploying MEC updates" @(
        "a) Confirm the SCCM Software Update Point publishes MEC updates for this device's collection.",
        "b) Verify a Software Update deployment or Office channel-change baseline targets this machine.",
        "c) Check SCCM client logs on the device: CAS.log, UpdatesDeployment.log, UpdatesHandler.log.",
        "d) Run a Software Update Deployment Evaluation Cycle from SCCM client actions."
    )
}

# F8 – Cloud Update is targeting MEC
if ($audienceData -match 'MEC' -or $audienceId -match '55336b82') {
    Add-Finding 'INFO' "Cloud Update IS targeting Monthly Enterprise Channel (AudienceData=$audienceData, AudienceId=$audienceId) but the channel switch is blocked by one or more issues above."
}

# Verify step
Add-Step "VERIFY - Confirm the channel change after fixes are applied" @(
    "a) Re-run this tool and verify Priority 1-4 shows [SET] and * ACTIVE with MEC.",
    "b) Check registry: CDNBaseUrl = http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6",
    "c) Check registry: UpdateChannelChanged = True",
    "d) Verify Office version updates to a Monthly Enterprise Channel build.",
    "e) Ensure the 'Office Automatic Updates 2.0' scheduled task is enabled and runs."
)

Write-Section "FINDINGS  ([!] = Error  [W] = Warning  [i] = Info)"
foreach ($f in $findings) {
    $prefix = switch ($f.Sev) { 'ERROR'{'  [!] '} 'WARN'{'  [W] '} 'INFO'{'  [i] '} default{'  [ ] '} }
    $fCol   = switch ($f.Sev) { 'ERROR'{'Red'} 'WARN'{'Yellow'} 'INFO'{'Cyan'} default{'White'} }
    # Word-wrap to console width
    $words = $f.Msg -split ' '; $line = ''; $cont = '       '
    Write-Host $prefix -NoNewline -ForegroundColor $fCol
    foreach ($w in $words) {
        if (($cont + $line + $w).Length -gt $CW) {
            Write-Host $line.TrimEnd() -ForegroundColor $fCol
            Write-Host $cont -NoNewline
            $line = "$w "
        } else { $line += "$w " }
    }
    if ($line.Trim()) { Write-Host $line.TrimEnd() -ForegroundColor $fCol }
}

Write-Section "ACTION PLAN"
$inStep = $false
foreach ($l in $actionPlan) {
    if ($l -match '^\[\d+\]') {
        if ($inStep) { Write-Host "" }
        Write-Host "  $l" -ForegroundColor White
        $inStep = $true
    } else {
        Write-Host "  $l" -ForegroundColor Gray
    }
}
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
# HTML REPORT GENERATION
# ─────────────────────────────────────────────────────────────────────────────
function Encode-Html { param([string]$T)
    $T -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'",'&#39;'
}

$tierBg    = @{1='#FFF0F0';2='#FFF0F0';3='#FFF5EB';4='#FFF5EB';5='#F0FFFE';6='#F0F8FF';7='#F0F8FF'}
$tierBadge = @{1='#FF6B6B';2='#FF6B6B';3='#FF8C42';4='#FF8C42';5='#4ECDC4';6='#45B7D1';7='#45B7D1'}

$tableRows = ''
foreach ($row in $priorities) {
    $p        = $row.P
    $badge    = $tierBadge[$p]
    $bg       = $tierBg[$p]
    $isActive = ($activeRow -and $activeRow.P -eq $p)
    $hasVal   = ($row.DetectedValue -and $row.DetectedValue -ne '')
    $display  = if ($hasVal) { Encode-Html $row.DetectedValue } else { "<span style='color:#bbb'>&#8212;</span>" }
    $statusIcon  = if ($hasVal) { "&#9989;" } else { "&#10060;" }
    $statusTitle = if ($hasVal) { "Configured" } else { "Not Set" }
    $statusColor = if ($hasVal) { "#2e7d32" } else { "#999" }
    $rowStyle    = if ($isActive) { "background:#E8F5E9;border-left:4px solid #2e7d32" } else { "background:$bg" }
    $activeCell  = if ($isActive) { "<td style='text-align:center;font-weight:bold;color:#2e7d32'>&#9733; ACTIVE</td>" } else { "<td></td>" }
    $rid = "row-$p"

    $tableRows += "<tr style='$rowStyle;cursor:pointer' onclick='toggleDetail(""$rid"")'>"
    $tableRows += "<td style='text-align:center;font-weight:bold;color:$badge;font-size:16px'>$p</td>"
    $tableRows += "<td><span class='mgmt-badge' style='background:$badge'>$(Encode-Html $row.Mgmt)</span> <span class='arrow' id='arrow-$rid'>&#9654;</span></td>"
    $tableRows += "<td><code>$($row.RegVal)</code></td>"
    $tableRows += "<td class='reg-path'><code>$(Encode-Html $row.RegPath)</code></td>"
    $tableRows += "<td class='value-cell'>$display</td>"
    $tableRows += "<td style='text-align:center;color:$statusColor'><span title='$statusTitle'>$statusIcon</span></td>"
    $tableRows += "$activeCell</tr>`n"

    # Expandable detail panel
    $pathEntries = Get-RegEntriesForPath $row.RegPath
    $detailHtml  = ''
    if ($pathEntries.Count -gt 0) {
        $detailHtml += "<div class='detail-header'>&#128194; $(Encode-Html $row.RegPath) <span class='setting-count'>($($pathEntries.Count) settings)</span></div><div class='settings-grid'>"
        foreach ($e in $pathEntries) {
            $hl = if ($e.Name -ieq $row.RegVal.ToLower()) { " style='background:#E8F5E9;border-left:3px solid #2e7d32;padding-left:9px'" } else { '' }
            $detailHtml += "<div class='setting-row'$hl><span class='setting-name'>$(Encode-Html $e.Name)</span><span class='setting-value'>$(Encode-Html $e.Value)</span></div>"
        }
        $detailHtml += "</div>"
    } else { $detailHtml = "<div class='detail-empty'>No registry values found at this path.</div>" }

    $tableRows += "<tr id='$rid' class='detail-row' style='display:none'><td colspan='7' style='padding:0;background:$bg'><div class='detail-panel'>$detailHtml</div></td></tr>`n"
}

# Build findings HTML
$findingsHtml = ''
foreach ($f in $findings) {
    $icon  = switch ($f.Sev) { 'ERROR'{'&#10060;'} 'WARN'{'&#9888;'} 'INFO'{'&#8505;'} default{'&#8226;'} }
    $color = switch ($f.Sev) { 'ERROR'{'#c62828'} 'WARN'{'#e65100'} 'INFO'{'#1565c0'} default{'#333'} }
    $bg2   = switch ($f.Sev) { 'ERROR'{'#fff3f3'} 'WARN'{'#fff8f0'} 'INFO'{'#f0f4ff'} default{'#fff'} }
    $findingsHtml += "<div style='background:$bg2;border-left:4px solid $color;padding:10px 14px;margin-bottom:8px;border-radius:4px;font-size:13px;'>"
    $findingsHtml += "<span style='color:$color;font-size:16px;margin-right:8px;'>$icon</span><strong style='color:$color'>$($f.Sev)</strong> &nbsp; $(Encode-Html $f.Msg)</div>"
}

# Build action plan HTML
$actionHtml = ''
$inBlock = $false
foreach ($l in $actionPlan) {
    if ($l -match '^\[\d+\]') {
        if ($inBlock) { $actionHtml += "</ul>" }
        $actionHtml += "<p style='font-weight:bold;color:#1a3a5c;margin:14px 0 4px 0;font-size:13px;'>$(Encode-Html $l)</p><ul style='margin:0;padding-left:22px;font-size:12px;color:#444;'>"
        $inBlock = $true
    } else {
        $actionHtml += "<li style='margin-bottom:3px;'>$(Encode-Html $l.Trim())</li>"
    }
}
if ($inBlock) { $actionHtml += "</ul>" }

$generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm"
$outputHtml = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Office Update Channel Analysis - $folderLabel</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 40px; background: #f5f6fa; }
h1 { color: #1a3a5c; margin-bottom:4px; }
h2 { color: #1a3a5c; margin-top:36px; }
.subtitle { color: #666; font-size: 14px; margin-bottom: 20px; }
.subtitle a { color: #2b579a; }
table { border-collapse: collapse; width: 100%; max-width: 1150px; font-size: 13px; background: white; box-shadow: 0 2px 8px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }
th { background: #2b579a; color: white; padding: 12px 14px; text-align: left; }
td { padding: 10px 14px; border-bottom: 1px solid #eee; }
tr:hover { filter: brightness(0.97); }
.mgmt-badge { display:inline-block; padding:3px 10px; border-radius:12px; color:white; font-size:11px; font-weight:bold; }
.arrow { font-size:10px; color:#888; transition:transform 0.2s; display:inline-block; }
.arrow.open { transform:rotate(90deg); }
.reg-path code { font-size:11px; color:#555; word-break:break-all; }
.value-cell { font-family:Consolas,monospace; font-size:12px; max-width:300px; word-break:break-all; color:#0066cc; }
code { background:#f0f0f0; padding:2px 5px; border-radius:3px; font-size:12px; }
.info-box { background:white; padding:15px 20px; border-radius:8px; margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,0.1); border-left:4px solid #2b579a; font-size:13px; }
.legend { display:flex; gap:20px; margin:15px 0; font-size:13px; }
.detail-panel { padding:12px 24px 16px 24px; }
.detail-header { font-weight:bold; color:#2b579a; font-size:13px; margin-bottom:8px; }
.setting-count { font-weight:normal; color:#888; font-size:12px; }
.settings-grid { display:grid; grid-template-columns:1fr; gap:2px; }
.setting-row { display:flex; justify-content:space-between; padding:5px 12px; font-size:12px; font-family:Consolas,monospace; border-radius:4px; }
.setting-row:nth-child(odd) { background:rgba(0,0,0,0.03); }
.setting-name { color:#333; font-weight:bold; margin-right:20px; }
.setting-value { color:#0066cc; text-align:right; word-break:break-all; max-width:500px; }
.detail-empty { color:#999; font-style:italic; font-size:13px; padding:5px 0; }
.summary-grid { display:grid; grid-template-columns:200px 1fr; gap:4px 20px; font-size:13px; max-width:700px; }
.summary-key { color:#555; font-weight:bold; }
.summary-val { color:#1a3a5c; }
</style>
</head>
<body>
<h1>&#128736; Office Update Channel Analysis</h1>
<p class="subtitle">
  $modeLabel &nbsp;|&nbsp;
  Generated: $generatedAt &nbsp;|&nbsp;
  Ref: <a href="https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels" target="_blank">Microsoft Learn: Change update channels</a>
</p>

<div class="info-box">
<div class="summary-grid">
  <span class="summary-key">Office version:</span>      <span class="summary-val">$(Nvl $currentVersion)</span>
  <span class="summary-key">Platform:</span>            <span class="summary-val">$(Nvl $platform)</span>
  <span class="summary-key">Active channel:</span>      <span class="summary-val" style="color:$(if($activeCh -match 'Semi-Annual'){'#c62828'}else{'#2e7d32'})"><strong>$activeCh</strong></span>
  <span class="summary-key">Target channel:</span>      <span class="summary-val">$(Nvl $targetCh)</span>
  <span class="summary-key">Channel changed:</span>     <span class="summary-val" style="color:$(if($updateChannelChg -ieq 'False'){'#c62828'}else{'#2e7d32'})">$(Nvl $updateChannelChg)</span>
  <span class="summary-key">Audience data:</span>       <span class="summary-val">$(Nvl $audienceData)</span>
  <span class="summary-key">Auto updates (GPO):</span>  <span class="summary-val" style="color:$(if($autoUpdatesDword -eq 0){'#c62828'}else{'#2e7d32'})">$autoStr</span>
  <span class="summary-key">Cloud ignoregpo:</span>     <span class="summary-val">$igStr</span>
  <span class="summary-key">OfficeMgmtCOM:</span>       <span class="summary-val">$mcStr</span>
</div>
</div>

<div class="legend">
  <span>&#9989; = Configured</span>
  <span>&#10060; = Not Set</span>
  <span>&#9733; = Active (winning priority)</span>
</div>

<h2>Channel Priority Table</h2>
<div class="info-box" style="margin-bottom:10px;">
  <strong>&#9432; How it works:</strong> The first configured value (1st &#8594; 7th) determines the active update channel.
  The winning row is highlighted in green. <strong>&#128073; Click any row</strong> to expand all registry settings at that path.
</div>

<table>
<tr>
  <th>Priority</th><th>Management Type</th><th>Registry Value</th>
  <th>Registry Path</th><th>Detected Value</th><th>Status</th><th>Active?</th>
</tr>
$tableRows
</table>

<h2>Findings &amp; Root Cause Analysis</h2>
$findingsHtml

<h2>Action Plan</h2>
<div style="background:white;padding:16px 20px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.1);max-width:900px;">
$actionHtml
</div>

<h2 style="margin-top:40px;">&#128209; Reference: Update Channel Priority Processing</h2>
<p style="font-size:13px;color:#666;">Source: <a href="https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels" target="_blank">Microsoft Learn — Change update channels for Microsoft 365 Apps</a></p>
<table style="max-width:900px;">
  <tr><th>Priority</th><th>Management Type</th><th>Registry Value</th><th>Registry Path</th></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#FF6B6B">1st</td><td><span class="mgmt-badge" style="background:#FF6B6B">Cloud Update</span></td><td><code>UpdatePath</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#FF6B6B">2nd</td><td><span class="mgmt-badge" style="background:#FF6B6B">Cloud Update</span></td><td><code>UpdateBranch</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Policies\Microsoft\cloud\office\16.0\Common\officeupdate</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#FF8C42">3rd</td><td><span class="mgmt-badge" style="background:#FF8C42">Policy Setting</span></td><td><code>UpdatePath</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#FF8C42">4th</td><td><span class="mgmt-badge" style="background:#FF8C42">Policy Setting</span></td><td><code>UpdateBranch</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Policies\Microsoft\office\16.0\Common\officeupdate</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#4ECDC4">5th</td><td><span class="mgmt-badge" style="background:#4ECDC4">ODT</span></td><td><code>UpdateUrl</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#45B7D1">6th</td><td><span class="mgmt-badge" style="background:#45B7D1">Unmanaged</span></td><td><code>UnmanagedUpdateURL *</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration</code></td></tr>
  <tr><td style="text-align:center;font-weight:bold;color:#45B7D1">7th</td><td><span class="mgmt-badge" style="background:#45B7D1">Unmanaged</span></td><td><code>CDNBaseUrl</code></td><td class="reg-path"><code>HKLM\SOFTWARE\Microsoft\office\ClickToRun\Configuration</code></td></tr>
</table>
<p style="font-size:12px;color:#666;margin-top:8px;">* <em>UnmanagedUpdateURL is only set on unmanaged devices.</em></p>

<div class="info-box" style="margin-top:20px;border-left-color:#e67e22;max-width:900px;">
  <strong>&#9888; Important — Channel changes announced (July 2025):</strong>
  <ul style="margin:8px 0 0 0;padding-left:20px;font-size:13px;">
    <li><strong>Semi-Annual Enterprise Channel (Preview)</strong> is being deprecated. Migrate immediately.</li>
    <li><strong>Semi-Annual Enterprise Channel</strong> is shifting to unattended devices only. Microsoft recommends moving interactive devices to Monthly Enterprise Channel or Current Channel.</li>
    <li>SAEC feature releases now supported for <strong>6 months</strong> (down from 14).</li>
    <li>MEC now includes <strong>2 months</strong> rollback support (up from 1).</li>
  </ul>
</div>

<script>
function toggleDetail(id) {
  var r = document.getElementById(id), a = document.getElementById('arrow-' + id);
  if (r.style.display === 'none') { r.style.display = 'table-row'; if(a) a.classList.add('open'); }
  else { r.style.display = 'none'; if(a) a.classList.remove('open'); }
}
</script>
</body>
</html>
"@

$htmlOut = Join-Path $outputFolder 'OfficePolicies.html'
$outputHtml | Out-File -FilePath $htmlOut -Encoding utf8

Write-Host ""
Write-Host ("  " + ([string][char]0x2500) * ($CW - 4)) -ForegroundColor DarkGray
Write-Host "  HTML report saved: $htmlOut" -ForegroundColor Green
Write-Host ""
try { Start-Process $htmlOut } catch { }
