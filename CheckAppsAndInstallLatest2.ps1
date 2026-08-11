#requires -Version 5.1
<#
.SYNOPSIS
    AVD application version / update report and updater.

.DESCRIPTION
    - Detects installed versions independently from Evergreen/winget.
    - Always returns a report row for every configured application.
    - Uses winget as the primary latest-version source.
    - Uses Evergreen only as an optional fallback.
    - Handles Power BI installer-technology mismatch by downloading the current installer, uninstalling the incompatible old Power BI Desktop package, reinstalling, and verifying the version.
    - Handles Visual Studio updates through the Visual Studio Installer.
    - Can optionally upgrade installed applications with -Upgrade.
    - Can optionally include Windows Updates with -IncludeWindowsUpdate.

.NOTES
    Run elevated for reliable machine-wide detection and updates.
#>

[CmdletBinding()]
param(
    [switch]$Upgrade,
    [switch]$IncludeWindowsUpdate,
    [switch]$WindowsUpdateOnly,
    [switch]$WhatIf,
    [switch]$NoHtml,
    [switch]$HtmlOnlyWhenGreen,
    [switch]$NoCsv,
    [switch]$NoReports,
    [switch]$DisableMsStore
)

if ($NoReports) {
    $NoCsv  = $true
    $NoHtml = $true
}

$ReportPath = Join-Path $env:PUBLIC "AVD-AppUpdateReport.html"
$CsvPath    = Join-Path $env:PUBLIC "AVD-AppUpdateReport.csv"
$LogPath    = Join-Path $env:PUBLIC "AVD-AppUpdateActions.log"

$DownloadDirCandidates = @(
    "C:\Ultima\Patching",
    (Join-Path $env:ProgramData "AVD-AppDownloads")
)

$DownloadDir = $null
foreach ($candidate in $DownloadDirCandidates) {
    try {
        if (-not (Test-Path $candidate)) {
            New-Item -ItemType Directory -Path $candidate -Force | Out-Null
        }
        $DownloadDir = $candidate
        break
    }
    catch {}
}

if (-not $DownloadDir) {
    throw "Could not create a download directory."
}

Write-Host "Download directory: $DownloadDir" -ForegroundColor DarkGray

$WingetTechnologyMismatch = -1978335090
$WingetNoApplicationsFound = -1978335212

# ------------------------------------------------------------
# Applications to track
# ------------------------------------------------------------

$AppsToCheck = @(
    @{
        Name            = "Windows Updates"
        LocalMatch      = "__WINDOWS_UPDATE__"
        LatestProvider  = "None"
        WingetId        = $null
        EvergreenName   = $null
        UpdateMethod    = "WindowsUpdate"
    },
    @{
        Name            = "Microsoft Office 365"
        LocalMatch      = "__OFFICE_C2R__"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.Office"
        EvergreenName   = $null
        UpdateMethod    = "OfficeC2R"
    },
    @{
        Name            = "Google Chrome"
        LocalMatch      = "Google Chrome"
        LatestProvider  = "Winget"
        WingetId        = "Google.Chrome"
        EvergreenName   = "GoogleChrome"
        UpdateMethod    = "Winget"
    },
    @{
        Name            = "Microsoft Edge"
        LocalMatch      = "Microsoft Edge"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.Edge"
        EvergreenName   = "MicrosoftEdge"
        UpdateMethod    = "EdgeUpdate"
    },
    @{
        Name            = "Visual Studio Code"
        LocalMatch      = "Microsoft Visual Studio Code"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.VisualStudioCode"
        EvergreenName   = "MicrosoftVisualStudioCode"
        UpdateMethod    = "Winget"
    },
  #  @{
  #      Name            = "Power BI Desktop"
  #      LocalMatch      = "Microsoft Power BI Desktop"
  #      LatestProvider  = "Winget"
  #      WingetId        = "Microsoft.PowerBI"
  #      EvergreenName   = $null
  #      UpdateMethod    = "PowerBI"
  #  },
  @{
    Name            = "Mozilla Firefox"
    LocalMatch      = "Mozilla Firefox"
    LatestProvider  = "Winget"
    WingetId        = "Mozilla.Firefox"
    EvergreenName   = $null
    UpdateMethod    = "Winget"
},
    @{
        Name            = "Visual Studio"
        LocalMatch      = "Microsoft Visual Studio"
        LatestProvider  = "VisualStudioInstaller"
        WingetId        = $null
        EvergreenName   = $null
        UpdateMethod    = "VisualStudioInstaller"
    },
    @{
        Name            = "Adobe Acrobat Reader"
        LocalMatch      = "Adobe Acrobat Reader"
        LatestProvider  = "Winget"
        WingetId        = "Adobe.Acrobat.Reader.64-bit"
        EvergreenName   = $null
        UpdateMethod    = "Winget"
    },
    @{
        Name            = "Adobe Acrobat (full)"
        LocalMatch      = "Adobe Acrobat"
        LatestProvider  = "None"
        WingetId        = $null
        EvergreenName   = $null
        UpdateMethod    = "None"
    },
    @{
        Name            = "Microsoft Teams"
        LocalMatch      = "Microsoft Teams"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.Teams"
        EvergreenName   = $null
        UpdateMethod    = "Winget"
    },
    @{
        Name            = "OneDrive"
        LocalMatch      = "Microsoft OneDrive"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.OneDrive"
        EvergreenName   = $null
        UpdateMethod    = "Winget"
    },
    @{
        Name            = "GitHub Desktop"
        LocalMatch      = "GitHub Desktop"
        LatestProvider  = "Winget"
        WingetId        = "GitHub.GitHubDesktop"
        EvergreenName   = $null
        UpdateMethod    = "Winget"
    },
    @{
        Name            = "Azure Data Studio"
        LocalMatch      = "Azure Data Studio"
        LatestProvider  = "Winget"
        WingetId        = "Microsoft.AzureDataStudio"
        EvergreenName   = $null
        UpdateMethod    = "Winget"
    }
)

# ------------------------------------------------------------
# Utility helpers
# ------------------------------------------------------------

function Write-ActionLog {
    param([string]$Message)
    try {
        Add-Content -Path $LogPath -Value ("{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
    }
    catch {}
}

function Test-WingetReady {
    try {
        $cmd = Get-Command winget.exe -ErrorAction Stop
        $null = & $cmd.Source --version 2>$null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        return $false
    }
}

function Test-EvergreenReady {
    try {
        $module = Get-Module -ListAvailable -Name Evergreen | Select-Object -First 1
        if (-not $module) { return $false }
        Import-Module Evergreen -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Disable-MsStoreSource {
    if (-not (Test-WingetReady)) { return }
    try {
        $sourceOutput = (& winget source list 2>$null | Out-String)
        if ($sourceOutput -match '(?im)^\s*msstore\s+') {
            & winget source disable msstore 2>$null | Out-Null
        }
    }
    catch {}
}

function Convert-ToVersionObject {
    param([string]$VersionText)
    if ([string]::IsNullOrWhiteSpace($VersionText)) { return $null }

    try {
        $clean = ($VersionText -replace '[^\d\.]', '').Trim('.')
        if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
        return [version]$clean
    }
    catch {
        return $null
    }
}

function Compare-VersionSmart {
    param(
        [string]$Installed,
        [string]$Latest
    )

    if (-not $Installed -or -not $Latest) { return $null }

    $v1 = Convert-ToVersionObject $Installed
    $v2 = Convert-ToVersionObject $Latest

    if ($v1 -and $v2) {
        return [Math]::Sign($v1.CompareTo($v2))
    }

    if ($Installed -eq $Latest) { return 0 }
    return $null
}

# ------------------------------------------------------------
# Installed app detection
# ------------------------------------------------------------

function Get-UninstallEntries {
    $roots = @(
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $entries = foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }

        foreach ($subKey in Get-ChildItem -Path $root -ErrorAction SilentlyContinue) {
            try {
                $p = Get-ItemProperty -Path $subKey.PSPath -ErrorAction Stop
                if ($p.DisplayName) {
                    [PSCustomObject]@{
                        DisplayName     = [string]$p.DisplayName
                        DisplayVersion  = [string]$p.DisplayVersion
                        Publisher       = [string]$p.Publisher
                        InstallLocation = [string]$p.InstallLocation
                        UninstallString = [string]$p.UninstallString
                        RegistryPath    = [string]$subKey.PSPath
                    }
                }
            }
            catch {}
        }
    }

    return @($entries)
}

function Get-GitHubDesktopVersion {
    $roots = @(
        (Join-Path $env:LOCALAPPDATA "GitHubDesktop"),
        (Join-Path $env:ProgramFiles "GitHub Desktop")
    )

    foreach ($root in $roots) {
        if (-not $root -or -not (Test-Path $root)) { continue }

        try {
            $exe = Get-ChildItem -Path $root -Filter "GitHubDesktop.exe" -Recurse -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending |
                Select-Object -First 1

            if ($exe) {
                $version = (Get-Item $exe.FullName).VersionInfo.ProductVersion
                if ($version) { return $version }
            }
        }
        catch {}
    }

    return $null
}

function Get-VSCodeVersion {
    $paths = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Microsoft VS Code\Code.exe"),
        (Join-Path $env:ProgramFiles "Microsoft VS Code\Code.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft VS Code\Code.exe")
    )

    foreach ($path in $paths) {
        if ($path -and (Test-Path $path)) {
            try {
                $version = (Get-Item $path).VersionInfo.ProductVersion
                if ($version) { return $version }
            }
            catch {}
        }
    }

    return $null
}

function Get-VSWherePath {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
        "$env:ProgramFiles\Microsoft Visual Studio\Installer\vswhere.exe"
    )

    foreach ($p in $candidates) {
        if ($p -and (Test-Path $p)) { return $p }
    }

    return $null
}

function Get-VisualStudioInstances {
    $vswhere = Get-VSWherePath
    if (-not $vswhere) { return @() }

    try {
        $json = & $vswhere -all -products * -format json -utf8 2>$null
        if (-not $json) { return @() }

        $instances = $json | ConvertFrom-Json
        return @($instances)
    }
    catch {
        return @()
    }
}

function Get-VisualStudioVersion {
    $instances = @(Get-VisualStudioInstances)

    if ($instances.Count -gt 0) {
        $best = $instances |
            Where-Object { $_.installationVersion } |
            Sort-Object { Convert-ToVersionObject $_.installationVersion } -Descending |
            Select-Object -First 1

        if ($best) {
            return [string]$best.installationVersion
        }
    }

    # Registry fallback
    $entries = @(Get-UninstallEntries)

    $candidates = foreach ($entry in $entries) {
        if ($entry.DisplayName -notmatch 'Visual Studio') { continue }

        $v = Convert-ToVersionObject $entry.DisplayVersion
        if (-not $v) { continue }

        [PSCustomObject]@{
            DisplayVersion = $entry.DisplayVersion
            VersionObject  = $v
        }
    }

    if (-not $candidates) { return $null }

    return (
        $candidates |
        Sort-Object VersionObject -Descending |
        Select-Object -First 1
    ).DisplayVersion
}

function Get-InstalledAppVersion {
    param(
        [Parameter(Mandatory)]
        [string]$DisplayNameMatch,

        [string[]]$ExcludePatterns
    )

    if ($DisplayNameMatch -eq "__OFFICE_C2R__") {
        $key = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $key) {
            try {
                $v = (Get-ItemProperty $key -ErrorAction Stop).VersionToReport
                if ($v) { return [string]$v }
            }
            catch {}
        }
        return $null
    }

    if ($DisplayNameMatch -eq "Microsoft Teams") {
        $packages = @()

        foreach ($packageName in @("MSTeams", "MicrosoftTeams")) {
            try { $packages += @(Get-AppxPackage -AllUsers -Name $packageName -ErrorAction SilentlyContinue) } catch {}
            try { $packages += @(Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue) } catch {}
        }

        $pkg = $packages |
            Where-Object { $_ -and $_.Version } |
            Sort-Object Version -Descending |
            Select-Object -First 1

        if ($pkg) { return $pkg.Version.ToString() }
    }

    if ($DisplayNameMatch -eq "Microsoft Visual Studio Code") {
        $v = Get-VSCodeVersion
        if ($v) { return $v }
    }

    if ($DisplayNameMatch -eq "GitHub Desktop") {
        $v = Get-GitHubDesktopVersion
        if ($v) { return $v }
    }

    if ($DisplayNameMatch -eq "Microsoft Visual Studio") {
        $v = Get-VisualStudioVersion
        if ($v) { return $v }
    }

    $entries = @(Get-UninstallEntries)

    $genericExcludes = @(
        'WebView','WebView2','Updater','AutoUpdate','Maintenance',
        'Service','Helper','Crashpad','Stub','Machine-wide',
        'User Installer','System Installer','Setup'
    )

    $candidates = foreach ($entry in $entries) {
        if ($entry.DisplayName -notlike "*$DisplayNameMatch*") { continue }

        $skip = $false

        foreach ($pattern in $genericExcludes) {
            if ($entry.DisplayName -match [regex]::Escape($pattern)) {
                $skip = $true
                break
            }
        }

        if ($skip) { continue }

        if ($ExcludePatterns) {
            foreach ($pattern in $ExcludePatterns) {
                if ($entry.DisplayName -match [regex]::Escape($pattern)) {
                    $skip = $true
                    break
                }
            }
        }

        if ($skip) { continue }

        $score = 1
        if ($entry.DisplayName -ieq $DisplayNameMatch) { $score = 3 }
        elseif ($entry.DisplayName -ilike "$DisplayNameMatch*") { $score = 2 }

        [PSCustomObject]@{
            DisplayName    = $entry.DisplayName
            DisplayVersion = $entry.DisplayVersion
            Score          = $score
            VersionObject  = Convert-ToVersionObject $entry.DisplayVersion
        }
    }

    if (-not $candidates) { return $null }

    $best = $candidates |
        Sort-Object `
            @{ Expression = { if ($_.VersionObject) { $_.VersionObject } else { [version]"0.0" } }; Descending = $true },
            @{ Expression = { $_.Score }; Descending = $true } |
        Select-Object -First 1

    if ($best -and $best.DisplayVersion) {
        return [string]$best.DisplayVersion
    }

    return $null
}

# ------------------------------------------------------------
# Latest-version detection
# ------------------------------------------------------------

function Get-LatestWingetVersion {
    param(
        [Parameter(Mandatory)]
        [string]$WingetId
    )

    if (-not (Test-WingetReady)) { return $null }

    try {
        $output = @(
            & winget show `
                --id $WingetId `
                --exact `
                --source winget `
                --accept-source-agreements 2>$null
        )

        if ($LASTEXITCODE -ne 0 -or -not $output) { return $null }

        foreach ($line in $output) {
            if ($line -match '^\s*Version\s*:\s*(.+?)\s*$') {
                return $matches[1].Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-LatestEvergreenVersion {
    param(
        [Parameter(Mandatory)]
        [string]$EvergreenName
    )

    if (-not (Test-EvergreenReady)) { return $null }

    try {
        $data = @(Get-EvergreenApp -Name $EvergreenName -ErrorAction SilentlyContinue)

        if (-not $data -or $data.Count -eq 0) { return $null }

        if ($data[0].PSObject.Properties.Name -contains "Channel") {
            $stable = @($data | Where-Object { -not $_.Channel -or $_.Channel -match 'Stable' })
            if ($stable.Count -gt 0) { $data = $stable }
        }

        $rows = foreach ($item in $data) {
            if (-not $item.Version) { continue }

            [PSCustomObject]@{
                Version       = [string]$item.Version
                VersionObject = Convert-ToVersionObject ([string]$item.Version)
            }
        }

        if (-not $rows) { return $null }

        return (
            $rows |
            Sort-Object `
                @{ Expression = { if ($_.VersionObject) { $_.VersionObject } else { [version]"0.0" } }; Descending = $true } |
            Select-Object -First 1
        ).Version
    }
    catch {
        return $null
    }
}

function Get-VisualStudioUpdateStatus {
    # Visual Studio exposes multiple version formats:
    #   - installationVersion: internal/build version, e.g. 17.14.37516.0
    #   - catalog.productDisplayVersion: release/display version, e.g. 17.14.37
    #
    # These must NOT be compared numerically because they are not equivalent
    # version schemes. Instead, use the Visual Studio Installer / vswhere
    # state only to report the installed version and treat the installation
    # as current unless a real installer-driven update operation proves
    # otherwise.

    $instances = @(Get-VisualStudioInstances)

    if ($instances.Count -eq 0) {
        return [PSCustomObject]@{
            Latest = $null
            Status = "Unknown (VS Installer not found)"
            Detail = "No Visual Studio instance returned by vswhere"
        }
    }

    $best = $instances |
        Where-Object { $_.installationVersion } |
        Sort-Object { Convert-ToVersionObject $_.installationVersion } -Descending |
        Select-Object -First 1

    if (-not $best) {
        return [PSCustomObject]@{
            Latest = $null
            Status = "Unknown (VS version unavailable)"
            Detail = "Visual Studio instance found but no installationVersion was returned"
        }
    }

    # Prefer the display/release version only for presentation.
    # Do not compare it to installationVersion.
    $displayVersion = $null

    try {
        if ($best.catalog -and $best.catalog.productDisplayVersion) {
            $displayVersion = [string]$best.catalog.productDisplayVersion
        }
    }
    catch {}

    if (-not $displayVersion) {
        $displayVersion = [string]$best.installationVersion
    }

    return [PSCustomObject]@{
        Latest = $displayVersion
        Status = "Up-to-date"
        Detail = "Visual Studio Installer-managed installation"
    }
}

function Get-LatestVersionForApp {
    param(
        [Parameter(Mandatory)]
        [hashtable]$App,

        [bool]$WingetAvailable,
        [bool]$EvergreenAvailable
    )

    if ($App.LatestProvider -eq "VisualStudioInstaller") {
        $vs = Get-VisualStudioUpdateStatus
        return [PSCustomObject]@{
            Version = $vs.Latest
            Source  = "Visual Studio Installer"
            ForcedStatus = $vs.Status
        }
    }

    $latest = $null
    $source = "-"

    if (
        $App.LatestProvider -eq "Winget" -and
        $WingetAvailable -and
        $App.WingetId
    ) {
        $latest = Get-LatestWingetVersion -WingetId $App.WingetId
        if ($latest) { $source = "winget" }
    }

    if (
        -not $latest -and
        $EvergreenAvailable -and
        $App.EvergreenName
    ) {
        $latest = Get-LatestEvergreenVersion -EvergreenName $App.EvergreenName
        if ($latest) { $source = "Evergreen" }
    }

    return [PSCustomObject]@{
        Version      = $latest
        Source       = $source
        ForcedStatus = $null
    }
}

# ------------------------------------------------------------
# Windows Update
# ------------------------------------------------------------

function Get-WindowsUpdateStatus {
    $pending = $null

    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result   = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
        $pending  = $result.Updates.Count
    }
    catch {
        $pending = $null
    }

    try {
        $os = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction Stop

        $displayVersion = $os.DisplayVersion
        if (-not $displayVersion) { $displayVersion = $os.ReleaseId }

        $installedText = "{0} (Build {1}.{2})" -f $displayVersion, $os.CurrentBuild, $os.UBR
    }
    catch {
        $installedText = "Unknown Windows version"
    }

    if ($null -eq $pending) {
        return [PSCustomObject]@{
            Installed = $installedText
            Latest    = "-"
            Status    = "Unknown (WU query failed)"
            UpgradeTo = "-"
        }
    }

    if ($pending -gt 0) {
        return [PSCustomObject]@{
            Installed = $installedText
            Latest    = "Updates available ($pending)"
            Status    = "Update available"
            UpgradeTo = "Apply Windows Updates"
        }
    }

    return [PSCustomObject]@{
        Installed = $installedText
        Latest    = "-"
        Status    = "Up-to-date"
        UpgradeTo = "-"
    }
}

# ------------------------------------------------------------
# Report results
# ------------------------------------------------------------

function Get-AppResults {
    param(
        [Parameter(Mandatory)]
        [array]$Apps
    )

    $wingetAvailable    = Test-WingetReady
    $evergreenAvailable = Test-EvergreenReady

    if ($DisableMsStore -and $wingetAvailable) {
        Disable-MsStoreSource
    }

    if (-not $wingetAvailable) {
        Write-Warning "winget is unavailable. Installed apps will still be reported."
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($app in $Apps) {

        if ($app.LocalMatch -eq "__WINDOWS_UPDATE__") {
            $wu = Get-WindowsUpdateStatus

            $results.Add([PSCustomObject]@{
                Application  = $app.Name
                Installed    = $wu.Installed
                Latest       = $wu.Latest
                Status       = $wu.Status
                UpgradeTo    = $wu.UpgradeTo
                LatestSource = "Windows Update"
                WingetId     = "-"
            })
            continue
        }

        $excludePatterns = $null
        if ($app.Name -eq "Adobe Acrobat (full)") {
            $excludePatterns = @("Reader")
        }

        $installed = $null

        try {
            $installed = Get-InstalledAppVersion `
                -DisplayNameMatch $app.LocalMatch `
                -ExcludePatterns $excludePatterns
        }
        catch {
            Write-Warning ("Installed-version detection failed for {0}: {1}" -f $app.Name, $_.Exception.Message)
        }

        if (-not $installed) {
            $results.Add([PSCustomObject]@{
                Application  = $app.Name
                Installed    = "-"
                Latest       = "-"
                Status       = "Application not installed"
                UpgradeTo    = "-"
                LatestSource = "-"
                WingetId     = $(if ($app.WingetId) { $app.WingetId } else { "-" })
            })
            continue
        }

        $latestInfo = [PSCustomObject]@{
            Version      = $null
            Source       = "-"
            ForcedStatus = $null
        }

        try {
            $latestInfo = Get-LatestVersionForApp `
                -App $app `
                -WingetAvailable $wingetAvailable `
                -EvergreenAvailable $evergreenAvailable
        }
        catch {}

        $latest = $latestInfo.Version
        $status = $null
        $upgradeTo = "-"

        if ($latestInfo.ForcedStatus) {
            $status = $latestInfo.ForcedStatus
            if ($status -eq "Update available" -and $latest) {
                $upgradeTo = $latest
            }
        }
        elseif ($app.LatestProvider -eq "None") {
            $status = "Installed (not checked)"
        }
        elseif (-not $latest) {
            $status = "Unknown (no latest info)"
        }
        else {
            $comparison = Compare-VersionSmart -Installed $installed -Latest $latest

            if ($null -eq $comparison) {
                $status = "Unknown (version comparison)"
            }
            elseif ($comparison -lt 0) {
                $status = "Update available"
                $upgradeTo = $latest
            }
            elseif ($comparison -eq 0) {
                $status = "Up-to-date"
            }
            else {
                $status = "Ahead of catalog"
            }
        }

        $results.Add([PSCustomObject]@{
            Application  = $app.Name
            Installed    = $installed
            Latest       = $(if ($latest) { $latest } else { "-" })
            Status       = $status
            UpgradeTo    = $upgradeTo
            LatestSource = $latestInfo.Source
            WingetId     = $(if ($app.WingetId) { $app.WingetId } else { "-" })
        })
    }

    return @($results | Sort-Object Application)
}

# ------------------------------------------------------------
# Update helpers
# ------------------------------------------------------------

function Stop-TeamsIfRunning {
    foreach ($name in @("ms-teams", "ms-teams-updater", "Teams", "MSTeams")) {
        Get-Process -Name $name -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch {}
            }
    }
}

function Stop-EdgeIfRunning {
    foreach ($processName in @("msedge", "msedgewebview2")) {
        Get-Process -Name $processName -ErrorAction SilentlyContinue |
            ForEach-Object {
                try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch {}
            }
    }
}

function Get-EdgeUpdateExecutable {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe",
        "$env:ProgramFiles\Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) { return $candidate }
    }

    return $null
}

function Get-EdgeExecutableVersion {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            try {
                $version = (Get-Item $candidate -ErrorAction Stop).VersionInfo.ProductVersion
                if ($version) { return [string]$version }
            }
            catch {}
        }
    }

    return $null
}

function Update-MicrosoftEdge {
    param([switch]$WhatIfMode)

    $edgeUpdater = Get-EdgeUpdateExecutable

    if (-not $edgeUpdater) {
        return [PSCustomObject]@{
            Method   = "EdgeUpdate"
            ExitCode = -1
            Result   = "Microsoft Edge Update executable not found"
        }
    }

    # Microsoft Edge Stable application GUID. This uses Edge's own updater
    # rather than winget so installer-technology mismatch errors are avoided.
    $arguments = '/silent /install appguid={56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}&appname=Microsoft%20Edge&needsadmin=True'

    if ($WhatIfMode) {
        Write-Host ("WhatIf: `"{0}`" {1}" -f $edgeUpdater, $arguments) -ForegroundColor Gray

        return [PSCustomObject]@{
            Method   = "EdgeUpdate"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    try {
        Stop-EdgeIfRunning
        Start-Sleep -Seconds 1

        $before = Get-EdgeExecutableVersion

        Write-Host (
            "Microsoft Edge: invoking Microsoft Edge Update (installed version {0})..." -f
            $(if ($before) { $before } else { "unknown" })
        ) -ForegroundColor Cyan

        $process = Start-Process `
            -FilePath $edgeUpdater `
            -ArgumentList $arguments `
            -Wait `
            -PassThru `
            -NoNewWindow

        # Edge Update can hand work to its service and return before msedge.exe
        # changes, so give it a short window to finish.
        for ($attempt = 1; $attempt -le 12; $attempt++) {
            Start-Sleep -Seconds 5
            $after = Get-EdgeExecutableVersion
            if ($after -and $before -and $after -ne $before) { break }
        }

        return [PSCustomObject]@{
            Method   = "EdgeUpdate"
            ExitCode = $process.ExitCode
            Result   = $(if ($process.ExitCode -eq 0) { "Success" } else { "ExitCode $($process.ExitCode)" })
        }
    }
    catch {
        return [PSCustomObject]@{
            Method   = "EdgeUpdate"
            ExitCode = -1
            Result   = $_.Exception.Message
        }
    }
}

function Update-OfficeC2R {
    param([switch]$WhatIfMode)

    $client = Join-Path $env:ProgramFiles "Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"

    if (-not (Test-Path $client)) {
        return [PSCustomObject]@{
            Method   = "OfficeC2R"
            ExitCode = -1
            Result   = "C2R client not found"
        }
    }

    $arguments = "/update user displaylevel=false forceappshutdown=true"

    if ($WhatIfMode) {
        return [PSCustomObject]@{
            Method   = "OfficeC2R"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    try {
        $process = Start-Process `
            -FilePath $client `
            -ArgumentList $arguments `
            -NoNewWindow `
            -PassThru `
            -Wait

        return [PSCustomObject]@{
            Method   = "OfficeC2R"
            ExitCode = $process.ExitCode
            Result   = $(if ($process.ExitCode -eq 0) { "Success" } else { "ExitCode $($process.ExitCode)" })
        }
    }
    catch {
        return [PSCustomObject]@{
            Method   = "OfficeC2R"
            ExitCode = -1
            Result   = $_.Exception.Message
        }
    }
}

function Invoke-WingetInstallFallback {
    param(
        [Parameter(Mandatory)]
        [string]$WingetId,

        [Parameter(Mandatory)]
        [string]$ApplicationName,

        [switch]$WhatIfMode
    )

    $arguments = @(
        "install",
        "--id", $WingetId,
        "--exact",
        "--source", "winget",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--silent",
        "--force"
    )

    if ($WhatIfMode) {
        Write-Host ("WhatIf: winget {0}" -f ($arguments -join " ")) -ForegroundColor Gray

        return [PSCustomObject]@{
            Method   = "winget-install"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    try {
        Write-Host ("{0}: using winget install --force fallback..." -f $ApplicationName) -ForegroundColor DarkYellow

        $process = Start-Process `
            -FilePath "winget.exe" `
            -ArgumentList $arguments `
            -NoNewWindow `
            -PassThru `
            -Wait

        return [PSCustomObject]@{
            Method   = "winget-install"
            ExitCode = $process.ExitCode
            Result   = $(if ($process.ExitCode -eq 0) { "Success" } else { "ExitCode $($process.ExitCode)" })
        }
    }
    catch {
        return [PSCustomObject]@{
            Method   = "winget-install"
            ExitCode = -1
            Result   = $_.Exception.Message
        }
    }
}

function Invoke-WingetUpgrade {
    param(
        [Parameter(Mandatory)]
        [string]$WingetId,

        [Parameter(Mandatory)]
        [string]$ApplicationName,

        [switch]$WhatIfMode
    )

    if (-not (Test-WingetReady)) {
        return [PSCustomObject]@{
            Method   = "winget"
            ExitCode = -1
            Result   = "winget unavailable"
        }
    }

    if ($ApplicationName -eq "Microsoft Teams" -and -not $WhatIfMode) {
        Stop-TeamsIfRunning
    }

    $arguments = @(
        "upgrade",
        "--id", $WingetId,
        "--exact",
        "--source", "winget",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--silent"
    )

    if ($WhatIfMode) {
        Write-Host ("WhatIf: winget {0}" -f ($arguments -join " ")) -ForegroundColor Gray

        return [PSCustomObject]@{
            Method   = "winget"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    try {
        $process = Start-Process `
            -FilePath "winget.exe" `
            -ArgumentList $arguments `
            -NoNewWindow `
            -PassThru `
            -Wait

        if ($process.ExitCode -eq 0) {
            return [PSCustomObject]@{
                Method   = "winget"
                ExitCode = 0
                Result   = "Success"
            }
        }

        # If WinGet can see the package in the catalog but cannot correlate the
        # existing installation, try a forced install over the top. This is useful
        # for apps such as Firefox and VS Code that may have been installed by a
        # different deployment method.
        if ($process.ExitCode -eq $WingetNoApplicationsFound) {
            Write-Host (
                "{0}: WinGet could not correlate the existing installation; trying install --force fallback..." -f
                $ApplicationName
            ) -ForegroundColor DarkYellow

            return Invoke-WingetInstallFallback `
                -WingetId $WingetId `
                -ApplicationName $ApplicationName
        }

        return [PSCustomObject]@{
            Method   = "winget"
            ExitCode = $process.ExitCode
            Result   = "ExitCode $($process.ExitCode)"
        }
    }
    catch {
        return [PSCustomObject]@{
            Method   = "winget"
            ExitCode = -1
            Result   = $_.Exception.Message
        }
    }
}

function Get-WingetInstallerUrl {
    param(
        [Parameter(Mandatory)]
        [string]$WingetId
    )

    try {
        $output = @(
            & winget show `
                --id $WingetId `
                --exact `
                --source winget `
                --accept-source-agreements 2>$null
        )

        foreach ($line in $output) {
            if ($line -match '^\s*Installer Url\s*:\s*(.+?)\s*$') {
                return $matches[1].Trim()
            }
        }
    }
    catch {}

    return $null
}

function Get-PowerBIUninstallEntries {
    # Only target the normal Power BI Desktop product.
    # Explicitly exclude Power BI Desktop for Report Server.
    $roots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $matches = foreach ($root in $roots) {
        if (-not (Test-Path $root)) { continue }

        foreach ($subKey in Get-ChildItem -Path $root -ErrorAction SilentlyContinue) {
            try {
                $p = Get-ItemProperty -Path $subKey.PSPath -ErrorAction Stop

                if (-not $p.DisplayName) { continue }

                if (
                    $p.DisplayName -match '^Microsoft Power BI Desktop' -and
                    $p.DisplayName -notmatch 'Report Server'
                ) {
                    [PSCustomObject]@{
                        DisplayName          = [string]$p.DisplayName
                        DisplayVersion       = [string]$p.DisplayVersion
                        UninstallString      = [string]$p.UninstallString
                        QuietUninstallString = [string]$p.QuietUninstallString
                        RegistryPath         = [string]$subKey.PSPath
                    }
                }
            }
            catch {}
        }
    }

    return @($matches)
}

function Remove-PowerBIExisting {
    param([switch]$WhatIfMode)

    $entries = @(Get-PowerBIUninstallEntries)

    if ($entries.Count -eq 0) {
        return [PSCustomObject]@{
            ExitCode = 0
            Result   = "Already absent"
        }
    }

    if ($WhatIfMode) {
        foreach ($entry in $entries) {
            Write-Host ("WhatIf: would uninstall {0} {1}" -f $entry.DisplayName, $entry.DisplayVersion) -ForegroundColor Gray
        }

        return [PSCustomObject]@{
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    foreach ($entry in $entries) {
        Write-Host ("Power BI: uninstalling existing {0} {1}..." -f $entry.DisplayName, $entry.DisplayVersion) -ForegroundColor DarkYellow

        $commandLine = $entry.QuietUninstallString

        if ([string]::IsNullOrWhiteSpace($commandLine)) {
            $commandLine = $entry.UninstallString
        }

        if ([string]::IsNullOrWhiteSpace($commandLine)) {
            return [PSCustomObject]@{
                ExitCode = -1
                Result   = "No uninstall command found"
            }
        }

        try {
            # MSI uninstall strings commonly use /I {GUID}; change to /X
            # and force a quiet, non-restarting uninstall.
            if ($commandLine -match '(?i)msiexec(\.exe)?') {
                $guidMatch = [regex]::Match($commandLine, '\{[0-9A-Fa-f\-]{36}\}')

                if (-not $guidMatch.Success) {
                    return [PSCustomObject]@{
                        ExitCode = -1
                        Result   = "Could not determine Power BI MSI product code"
                    }
                }

                $arguments = "/x $($guidMatch.Value) /qn /norestart"

                $process = Start-Process `
                    -FilePath "msiexec.exe" `
                    -ArgumentList $arguments `
                    -Wait `
                    -PassThru `
                    -NoNewWindow
            }
            else {
                # Parse an EXE uninstall string. Power BI's EXE installer
                # supports silent/uninstall command-line switches.
                $exePath = $null
                $existingArgs = ""

                if ($commandLine -match '^\s*"([^"]+)"\s*(.*)$') {
                    $exePath = $matches[1]
                    $existingArgs = $matches[2]
                }
                elseif ($commandLine -match '^\s*(\S+)\s*(.*)$') {
                    $exePath = $matches[1]
                    $existingArgs = $matches[2]
                }

                if (-not $exePath -or -not (Test-Path $exePath)) {
                    return [PSCustomObject]@{
                        ExitCode = -1
                        Result   = "Power BI uninstall executable not found"
                    }
                }

                $arguments = "$existingArgs -uninstall -quiet -norestart".Trim()

                $process = Start-Process `
                    -FilePath $exePath `
                    -ArgumentList $arguments `
                    -Wait `
                    -PassThru `
                    -NoNewWindow
            }

            # MSI 3010 = success, reboot required. We deliberately do not reboot.
            if ($process.ExitCode -notin @(0, 3010)) {
                return [PSCustomObject]@{
                    ExitCode = $process.ExitCode
                    Result   = "Uninstall failed"
                }
            }
        }
        catch {
            return [PSCustomObject]@{
                ExitCode = -1
                Result   = "Uninstall error: $($_.Exception.Message)"
            }
        }
    }

    Start-Sleep -Seconds 3

    $remaining = @(Get-PowerBIUninstallEntries)

    if ($remaining.Count -gt 0) {
        return [PSCustomObject]@{
            ExitCode = -1
            Result   = "Old Power BI still detected after uninstall"
        }
    }

    return [PSCustomObject]@{
        ExitCode = 0
        Result   = "Success"
    }
}

function Get-PowerBIDirectInstaller {
    # Resolve and download the current Power BI Desktop installer BEFORE
    # uninstalling the old version, so a download problem never leaves the
    # image without Power BI.
    $installerUrl = Get-WingetInstallerUrl -WingetId "Microsoft.PowerBI"

    if (-not $installerUrl) {
        return [PSCustomObject]@{
            Success = $false
            Path    = $null
            Result  = "Could not obtain current Power BI installer URL"
        }
    }

    $fileName = [IO.Path]::GetFileName((($installerUrl -split '\?')[0]))

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "PBIDesktopSetup_x64.exe"
    }

    $localPath = Join-Path $DownloadDir $fileName

    try {
        Write-Host "Power BI: downloading current installer before removing the old version..." -ForegroundColor Cyan

        Invoke-WebRequest `
            -UseBasicParsing `
            -Uri $installerUrl `
            -OutFile $localPath `
            -ErrorAction Stop

        if (-not (Test-Path $localPath)) {
            throw "Installer download completed but the file was not found."
        }

        return [PSCustomObject]@{
            Success = $true
            Path    = $localPath
            Result  = "Success"
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Path    = $null
            Result  = "Download failed: $($_.Exception.Message)"
        }
    }
}

function Install-PowerBIDirect {
    param(
        [Parameter(Mandatory)]
        [string]$InstallerPath,

        [switch]$WhatIfMode
    )

    if ($WhatIfMode) {
        Write-Host ("WhatIf: would install Power BI from {0}" -f $InstallerPath) -ForegroundColor Gray

        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    if (-not (Test-Path $InstallerPath)) {
        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = -1
            Result   = "Downloaded installer is missing"
        }
    }

    $extension = [IO.Path]::GetExtension($InstallerPath).ToLowerInvariant()

    try {
        if ($extension -eq ".msi") {
            $arguments = "/i `"$InstallerPath`" /qn /norestart ACCEPT_EULA=1"

            $process = Start-Process `
                -FilePath "msiexec.exe" `
                -ArgumentList $arguments `
                -Wait `
                -PassThru `
                -NoNewWindow
        }
        else {
            $arguments = "-quiet -norestart ACCEPT_EULA=1"

            $process = Start-Process `
                -FilePath $InstallerPath `
                -ArgumentList $arguments `
                -Wait `
                -PassThru `
                -NoNewWindow
        }

        if ($process.ExitCode -in @(0, 3010)) {
            return [PSCustomObject]@{
                Method   = "PowerBI-reinstall"
                ExitCode = $process.ExitCode
                Result   = "Success"
            }
        }

        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = $process.ExitCode
            Result   = "Install failed"
        }
    }
    catch {
        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = -1
            Result   = "Install error: $($_.Exception.Message)"
        }
    }
}

function Reinstall-PowerBI {
    param([switch]$WhatIfMode)

    if ($WhatIfMode) {
        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    # 1. Download first. Do not uninstall anything unless the current
    #    installer has been successfully obtained.
    $download = Get-PowerBIDirectInstaller

    if (-not $download.Success) {
        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = -1
            Result   = $download.Result
        }
    }

    # 2. Remove the old, incompatible Power BI Desktop installation.
    $uninstall = Remove-PowerBIExisting

    if ($uninstall.ExitCode -ne 0) {
        return [PSCustomObject]@{
            Method   = "PowerBI-reinstall"
            ExitCode = $uninstall.ExitCode
            Result   = $uninstall.Result
        }
    }

    # 3. Install the current release.
    return Install-PowerBIDirect -InstallerPath $download.Path
}

function Update-PowerBI {
    param([switch]$WhatIfMode)

    $firstAttempt = Invoke-WingetUpgrade `
        -WingetId "Microsoft.PowerBI" `
        -ApplicationName "Power BI Desktop" `
        -WhatIfMode:$WhatIfMode

    if ($WhatIfMode) {
        return $firstAttempt
    }

    if ($firstAttempt.ExitCode -eq 0) {
        return $firstAttempt
    }

    if ($firstAttempt.ExitCode -eq $WingetTechnologyMismatch) {
        Write-Host "Power BI: winget reported installer technology mismatch." -ForegroundColor DarkYellow
        Write-Host "Power BI: downloading latest installer, uninstalling the old package, then reinstalling..." -ForegroundColor DarkYellow

        return Reinstall-PowerBI
    }

    return $firstAttempt
}

function Update-VisualStudio {
    param([switch]$WhatIfMode)

    $setup = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\setup.exe"

    if (-not (Test-Path $setup)) {
        return [PSCustomObject]@{
            Method   = "VisualStudioInstaller"
            ExitCode = -1
            Result   = "Visual Studio Installer not found"
        }
    }

    if ($WhatIfMode) {
        return [PSCustomObject]@{
            Method   = "VisualStudioInstaller"
            ExitCode = 0
            Result   = "Simulated"
        }
    }

    $instances = @(Get-VisualStudioInstances)

    if ($instances.Count -eq 0) {
        return [PSCustomObject]@{
            Method   = "VisualStudioInstaller"
            ExitCode = -1
            Result   = "No Visual Studio instance found"
        }
    }

    $results = New-Object System.Collections.Generic.List[object]

    foreach ($instance in $instances) {
        if (-not $instance.installationPath) { continue }

        $arguments = @(
            "update",
            "--passive",
            "--norestart",
            "--installPath",
            "`"$($instance.installationPath)`""
        )

        try {
            $process = Start-Process `
                -FilePath $setup `
                -ArgumentList $arguments `
                -Wait `
                -PassThru `
                -NoNewWindow

            $results.Add([PSCustomObject]@{
                Path     = $instance.installationPath
                ExitCode = $process.ExitCode
            })
        }
        catch {
            $results.Add([PSCustomObject]@{
                Path     = $instance.installationPath
                ExitCode = -1
            })
        }
    }

    if ($results.Count -eq 0) {
        return [PSCustomObject]@{
            Method   = "VisualStudioInstaller"
            ExitCode = -1
            Result   = "No Visual Studio instance updated"
        }
    }

    $failed = @($results | Where-Object { $_.ExitCode -ne 0 })

    if ($failed.Count -eq 0) {
        return [PSCustomObject]@{
            Method   = "VisualStudioInstaller"
            ExitCode = 0
            Result   = "Success"
        }
    }

    return [PSCustomObject]@{
        Method   = "VisualStudioInstaller"
        ExitCode = $failed[0].ExitCode
        Result   = "One or more VS updates failed"
    }
}

function Invoke-AppUpgrades {
    param(
        [Parameter(Mandatory)]
        [array]$Results,

        [Parameter(Mandatory)]
        [array]$AppConfig,

        [switch]$IncludeWindowsUpdateMode,
        [switch]$WhatIfMode,
        [switch]$WindowsOnly
    )

    Write-ActionLog "==== Upgrade run started ===="

    if ($IncludeWindowsUpdateMode) {
        $windowsRow = $Results |
            Where-Object { $_.Application -eq "Windows Updates" } |
            Select-Object -First 1

        if ($windowsRow -and $windowsRow.Status -eq "Update available") {
            try {
                if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
                    Import-Module PSWindowsUpdate -ErrorAction Stop | Out-Null

                    if (-not $WhatIfMode) {
                        Get-WindowsUpdate -AcceptAll -Install -AutoReboot:$false |
                            Tee-Object -FilePath $LogPath -Append |
                            Out-Null
                    }
                }
                else {
                    Write-Warning "PSWindowsUpdate is not installed; skipping Windows Updates."
                }
            }
            catch {
                Write-Warning ("Windows Update failed: {0}" -f $_.Exception.Message)
            }
        }

        if ($WindowsOnly) { return }
    }

    $queue = @(
        $Results |
        Where-Object {
            $_.Application -ne "Windows Updates" -and
            $_.Status -eq "Update available"
        }
    )

    if ($queue.Count -eq 0) {
        Write-Host "No application upgrades required." -ForegroundColor Green
        return
    }

    Write-Host ("Upgrading {0} application(s)..." -f $queue.Count) -ForegroundColor Cyan

    $actions = New-Object System.Collections.Generic.List[object]
    $i = 0

    foreach ($row in $queue) {
        $i++

        $config = $AppConfig |
            Where-Object { $_.Name -eq $row.Application } |
            Select-Object -First 1

        if (-not $config) { continue }

        $pct = [int](($i - 1) * 100 / [Math]::Max(1, $queue.Count))

        Write-Progress `
            -Activity "Upgrading applications" `
            -Status ("[{0}/{1}] {2}" -f $i, $queue.Count, $config.Name) `
            -PercentComplete $pct

        $fromVersion = $row.Installed
        $updateResult = $null

        switch ($config.UpdateMethod) {
            "OfficeC2R" {
                $updateResult = Update-OfficeC2R -WhatIfMode:$WhatIfMode
            }

            "PowerBI" {
                $updateResult = Update-PowerBI -WhatIfMode:$WhatIfMode
            }

            "EdgeUpdate" {
                $updateResult = Update-MicrosoftEdge -WhatIfMode:$WhatIfMode
            }

            "VisualStudioInstaller" {
                $updateResult = Update-VisualStudio -WhatIfMode:$WhatIfMode
            }

            "Winget" {
                $updateResult = Invoke-WingetUpgrade `
                    -WingetId $config.WingetId `
                    -ApplicationName $config.Name `
                    -WhatIfMode:$WhatIfMode
            }

            default {
                $updateResult = [PSCustomObject]@{
                    Method   = "-"
                    ExitCode = -1
                    Result   = "No automatic update method configured"
                }
            }
        }

        $toVersion = $fromVersion

        if (-not $WhatIfMode) {
            if ($config.Name -eq "Microsoft Edge") {
                # Edge Update can complete asynchronously via its service.
                # Give it up to one minute to expose the new msedge.exe version.
                for ($verifyAttempt = 1; $verifyAttempt -le 12; $verifyAttempt++) {
                    Start-Sleep -Seconds 5
                    $detected = Get-EdgeExecutableVersion

                    if ($detected) { $toVersion = $detected }

                    if ($row.UpgradeTo -and $row.UpgradeTo -ne "-") {
                        $edgeCmp = Compare-VersionSmart -Installed $toVersion -Latest $row.UpgradeTo
                        if ($null -ne $edgeCmp -and $edgeCmp -ge 0) { break }
                    }
                }
            }
            else {
                Start-Sleep -Seconds 3

                $excludePatterns = $null
                if ($config.Name -eq "Adobe Acrobat (full)") {
                    $excludePatterns = @("Reader")
                }

                try {
                    $detected = Get-InstalledAppVersion `
                        -DisplayNameMatch $config.LocalMatch `
                        -ExcludePatterns $excludePatterns

                    if ($detected) { $toVersion = $detected }
                }
                catch {}
            }
        }

        $verifiedResult = $updateResult.Result

        if (-not $WhatIfMode -and $config.Name -eq "Visual Studio") {
            # Visual Studio uses different internal and display version schemes.
            # A successful Visual Studio Installer exit code is enough here;
            # do not compare installationVersion with productDisplayVersion.
            if ($updateResult.ExitCode -eq 0) {
                $verifiedResult = "Success"
            }
        }
        elseif (-not $WhatIfMode -and $row.UpgradeTo -and $row.UpgradeTo -ne "-") {
            $cmp = Compare-VersionSmart -Installed $toVersion -Latest $row.UpgradeTo

            if ($cmp -ge 0) {
                $verifiedResult = "Success"
            }
            elseif ($updateResult.ExitCode -eq 0) {
                $verifiedResult = "Installer completed but version not updated"
            }
        }

        $action = [PSCustomObject]@{
            Application = $config.Name
            Method      = $updateResult.Method
            From        = $fromVersion
            To          = $toVersion
            Result      = $verifiedResult
            ExitCode    = $updateResult.ExitCode
        }

        $actions.Add($action)

        if ($verifiedResult -eq "Success") {
            Write-Host ("{0}: {1} -> {2} via {3}" -f $config.Name, $fromVersion, $toVersion, $updateResult.Method) -ForegroundColor Green
        }
        elseif ($verifiedResult -eq "Simulated") {
            Write-Host ("{0}: would update {1} -> {2}" -f $config.Name, $fromVersion, $row.UpgradeTo) -ForegroundColor Yellow
        }
        else {
            Write-Host ("{0}: {1} (exit {2})" -f $config.Name, $verifiedResult, $updateResult.ExitCode) -ForegroundColor DarkYellow
        }

        Write-ActionLog (
            "{0}: {1} via {2} ({3} -> {4}) exit {5}" -f
            $action.Application,
            $action.Result,
            $action.Method,
            $action.From,
            $action.To,
            $action.ExitCode
        )
    }

    Write-Progress -Activity "Upgrading applications" -Completed

    Write-Host ""
    Write-Host "Upgrade summary:" -ForegroundColor White
    $actions | Format-Table Application, Method, From, To, Result, ExitCode -AutoSize
    Write-Host "Action log: $LogPath" -ForegroundColor DarkGray
}

# ------------------------------------------------------------
# HTML
# ------------------------------------------------------------

function Convert-HtmlSafe {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Write-ReportHtml {
    param(
        [Parameter(Mandatory)]
        [array]$Results
    )

    $css = @"
<style>
body { font-family: "Segoe UI", Arial, sans-serif; margin: 24px; color: #222; }
h1 { font-size: 22px; margin-bottom: 4px; }
.subtitle { color: #666; margin-bottom: 18px; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 9px 10px; text-align: left; }
th { background: #f3f4f6; }
tr.status-uptodate { background: #e8f5e9; }
tr.status-update { background: #ffebee; }
tr.status-missing { background: #eceff1; }
tr.status-installed { background: #e3f2fd; }
tr.status-unknown { background: #fff8e1; }
.badge { display: inline-block; padding: 3px 8px; border-radius: 999px; font-size: 12px; font-weight: 600; }
.badge-ok { background: #2e7d32; color: white; }
.badge-update { background: #c62828; color: white; }
.badge-missing { background: #607d8b; color: white; }
.badge-installed { background: #1565c0; color: white; }
.badge-unknown { background: #f9a825; color: #111; }
.small { color: #666; font-size: 12px; }
</style>
"@

    $rows = foreach ($row in $Results) {
        $class = "status-unknown"
        $badgeClass = "badge-unknown"

        switch -Regex ($row.Status) {
            '^Up-to-date$' {
                $class = "status-uptodate"
                $badgeClass = "badge-ok"
                break
            }

            '^Update available$' {
                $class = "status-update"
                $badgeClass = "badge-update"
                break
            }

            '^Application not installed$' {
                $class = "status-missing"
                $badgeClass = "badge-missing"
                break
            }

            '^Installed \(not checked\)$' {
                $class = "status-installed"
                $badgeClass = "badge-installed"
                break
            }
        }

        @"
<tr class="$class">
    <td>$(Convert-HtmlSafe $row.Application)</td>
    <td>$(Convert-HtmlSafe $row.Installed)</td>
    <td>$(Convert-HtmlSafe $row.Latest)</td>
    <td><span class="badge $badgeClass">$(Convert-HtmlSafe $row.Status)</span></td>
    <td>$(Convert-HtmlSafe $row.UpgradeTo)</td>
    <td>$(Convert-HtmlSafe $row.LatestSource)</td>
    <td class="small">$(Convert-HtmlSafe $row.WingetId)</td>
</tr>
"@
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>AVD App Update Report</title>
$css
</head>
<body>
<h1>AVD App Update Report</h1>
<div class="subtitle">Computer: $env:COMPUTERNAME</div>

<table>
<thead>
<tr>
    <th>Application</th>
    <th>Installed Version</th>
    <th>Latest Version</th>
    <th>Status</th>
    <th>Upgrade To</th>
    <th>Latest Source</th>
    <th>Winget ID</th>
</tr>
</thead>
<tbody>
$($rows -join "`r`n")
</tbody>
</table>

<p class="small">Generated: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")</p>
</body>
</html>
"@

    $html | Set-Content -Path $ReportPath -Encoding UTF8
    Write-Host "HTML report saved to: $ReportPath"

    try {
        if ($env:USERNAME -ne "SYSTEM") {
            Start-Process -FilePath $ReportPath -ErrorAction Stop
        }
    }
    catch {}
}

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

Write-Host ""
Write-Host "AVD Application Update Check" -ForegroundColor Cyan
Write-Host ("Computer: {0}" -f $env:COMPUTERNAME) -ForegroundColor DarkGray
Write-Host ""

$results = @(Get-AppResults -Apps $AppsToCheck)

if ($results.Count -ne $AppsToCheck.Count) {
    Write-Warning (
        "Expected {0} rows but received {1}." -f
        $AppsToCheck.Count,
        $results.Count
    )
}

$format = "{0,-28} {1,-22} {2,-22} {3,-28} {4,-20}"

Write-Host ($format -f "Application", "Installed", "Latest", "Status", "Upgrade To")
Write-Host ("-" * 125)

foreach ($row in $results) {
    $foreground = "Gray"

    if ($row.Status -eq "Up-to-date") { $foreground = "Green" }
    elseif ($row.Status -eq "Update available") { $foreground = "Red" }
    elseif ($row.Status -eq "Application not installed") { $foreground = "DarkGray" }
    elseif ($row.Status -like "Unknown*") { $foreground = "Yellow" }
    elseif ($row.Status -eq "Installed (not checked)") { $foreground = "Cyan" }
    elseif ($row.Status -eq "Ahead of catalog") { $foreground = "Cyan" }

    Write-Host (
        $format -f
        $row.Application,
        $row.Installed,
        $row.Latest,
        $row.Status,
        $row.UpgradeTo
    ) -ForegroundColor $foreground
}

if (-not $NoCsv) {
    $results |
        Export-Csv `
            -NoTypeInformation `
            -Encoding UTF8 `
            -Path $CsvPath

    Write-Host ""
    Write-Host "CSV saved to: $CsvPath"
}
else {
    Write-Host ""
    Write-Host "Skipping CSV output." -ForegroundColor DarkGray
}

$pendingHtml = $false

if (-not $NoHtml) {
    $issuesNow = @(
        $results |
        Where-Object { $_.Status -ne "Application not installed" } |
        Where-Object {
            $_.Status -eq "Update available" -or
            $_.Status -like "Unknown*"
        }
    )

    if ($HtmlOnlyWhenGreen) {
        if ($issuesNow.Count -eq 0 -and -not $Upgrade) {
            Write-ReportHtml -Results $results
        }
        elseif ($Upgrade) {
            $pendingHtml = $true
            Write-Host "HTML will be generated after upgrades if everything is green." -ForegroundColor DarkYellow
        }
        else {
            Write-Host ("HTML not generated because {0} item(s) need attention." -f $issuesNow.Count) -ForegroundColor DarkYellow
        }
    }
    else {
        Write-ReportHtml -Results $results
    }
}
else {
    Write-Host "Skipping HTML report." -ForegroundColor DarkGray
}

if ($Upgrade) {
    Invoke-AppUpgrades `
        -Results $results `
        -AppConfig $AppsToCheck `
        -IncludeWindowsUpdateMode:$IncludeWindowsUpdate `
        -WhatIfMode:$WhatIf `
        -WindowsOnly:$WindowsUpdateOnly

    if ($pendingHtml -and -not $NoHtml) {
        Write-Host ""
        Write-Host "Rechecking application versions after upgrade..." -ForegroundColor Cyan

        $resultsPost = @(Get-AppResults -Apps $AppsToCheck)

        $issuesPost = @(
            $resultsPost |
            Where-Object { $_.Status -ne "Application not installed" } |
            Where-Object {
                $_.Status -eq "Update available" -or
                $_.Status -like "Unknown*"
            }
        )

        if ($issuesPost.Count -eq 0) {
            Write-ReportHtml -Results $resultsPost
        }
        else {
            Write-Host (
                "Still not fully green; HTML not generated. {0} item(s) still need attention." -f
                $issuesPost.Count
            ) -ForegroundColor DarkYellow
        }
    }
}
