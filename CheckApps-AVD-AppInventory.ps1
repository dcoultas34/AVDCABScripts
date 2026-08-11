param(
  [string]$CsvPath = (Join-Path $PSScriptRoot "Widgets_installed_programs_updated.csv"),
  [string]$ReportRoot = "C:\Ultima\Reporting\AVDAppInventory"
)

$ErrorActionPreference = 'SilentlyContinue'

# ---------------------- Paths ----------------------
$today = Get-Date
$period = $today.ToString("MMMM yyyy")
$monthFolder = Join-Path $ReportRoot $period

if (-not (Test-Path -LiteralPath $monthFolder)) {
  New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $CsvPath)) {
  Write-Error "CSV file not found: $CsvPath"
  exit 1
}

$machine = $env:COMPUTERNAME
$todayIso = $today.ToString('yyyy-MM-dd')

# ---------------------- Evergreen map ----------------------
# Only applications with a known Evergreen application name are queried.
$EvergreenMap = @{

  '7-Zip 25.01 (x64)' = '7Zip'

  'Adobe Acrobat (64-bit)' = 'AdobeAcrobat'

  'Audacity 3.7.7' = 'Audacity'

  'FileZilla 3.69.5' = 'FileZilla'

  'Google Chrome' = 'GoogleChrome'

  'Microsoft 365 Apps for enterprise - en-us' = 'Microsoft365Apps'

  'Microsoft 365 Apps for enterprise - fr-fr' = 'Microsoft365Apps'

  'Microsoft Edge' = 'MicrosoftEdge'

  'Microsoft FSLogix Apps' = 'MicrosoftFSLogixApps'

  'Mozilla Firefox ESR (x64 en-US)' = 'MozillaFirefox'

  'Notepad++ (64-bit x64)' = 'NotepadPlusPlus'

  'VLC media player' = 'VideoLanVlc'

  'Zoom Workplace (64-bit)' = 'Zoom'

}


# ---------------------- Winget map ----------------------
# Power BI Desktop uses Winget for the latest available version.
$WingetMap = @{
  'Microsoft Power BI Desktop (x64)' = 'Microsoft.PowerBI'
}

# ---------------------- Version/release-date lookup map ----------------------
# Used only to enrich the report with historical release dates.
# If an exact old version is not present in Winget, the report simply leaves
# the release date blank rather than guessing.
$VersionDateWingetMap = @{
  '7-Zip 25.01 (x64)'                    = '7zip.7zip'
  'Adobe Acrobat (64-bit)'              = 'Adobe.Acrobat.Pro'
  'Audacity 3.7.7'                      = 'Audacity.Audacity'
  'FileZilla 3.69.5'                    = 'TimKosse.FileZilla.Client'
  'Google Chrome'                       = 'Google.Chrome'
  'Microsoft Edge'                      = 'Microsoft.Edge'
  'Notepad++ (64-bit x64)'              = 'Notepad++.Notepad++'
  'VLC media player'                    = 'VideoLAN.VLC'
  'Zoom Workplace (64-bit)'             = 'Zoom.Zoom'
  'Microsoft Power BI Desktop (x64)'    = 'Microsoft.PowerBI'
}

function Test-WingetReady {
  try {
    $null = Get-Command winget.exe -ErrorAction Stop
    winget --version | Out-Null
    return $true
  }
  catch {
    return $false
  }
}

function Convert-WingetDate {
  param($Value)

  if (-not $Value) {
    return $null
  }

  try {
    return [datetime]::Parse("$Value")
  }
  catch {
    return $null
  }
}

function Get-WingetReleaseInfo {
  param(
    [string]$WingetId,
    [string]$Version
  )

  if (-not $WingetId -or -not (Test-WingetReady)) {
    return $null
  }

  try {
    $args = @(
      'show',
      '--id', $WingetId,
      '--exact',
      '--source', 'winget',
      '--accept-source-agreements'
    )

    if ($Version -and $Version -ne '-') {
      $args += @('--version', $Version)
    }

    $out = @(& winget @args 2>$null)

    if (-not $out -or $LASTEXITCODE -ne 0) {
      return $null
    }

    $resolvedVersion = $null
    $releaseDate = $null

    foreach ($line in $out) {
      if (-not $resolvedVersion -and $line -match '^\s*Version\s*:\s*(.+?)\s*$') {
        $resolvedVersion = $matches[1].Trim()
      }

      if (-not $releaseDate -and $line -match '^\s*Release\s+Date\s*:\s*(.+?)\s*$') {
        $releaseDate = Convert-WingetDate $matches[1].Trim()
      }
    }

    return [pscustomobject]@{
      Version     = $resolvedVersion
      ReleaseDate = $releaseDate
    }
  }
  catch {
    return $null
  }
}

function Get-LatestWingetVersion {
  param([string]$WingetId)

  $info = Get-WingetReleaseInfo -WingetId $WingetId

  if ($info) {
    return $info.Version
  }

  return $null
}

function Get-WingetReleaseDateForVersion {
  param(
    [string]$WingetId,
    [string]$Version
  )

  if (-not $WingetId -or -not $Version -or $Version -eq '-') {
    return $null
  }

  $info = Get-WingetReleaseInfo -WingetId $WingetId -Version $Version

  if ($info) {
    return $info.ReleaseDate
  }

  return $null
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

function Compare-VersionSmart {
  param(
    [string]$Installed,
    [string]$Latest
  )

  if (-not $Installed -or -not $Latest -or $Installed -eq '-' -or $Latest -eq '-') {
    return $null
  }

  try {
    $installedVersion = [version](($Installed -replace '[^\d\.]','').Trim('.'))
    $latestVersion    = [version](($Latest    -replace '[^\d\.]','').Trim('.'))

    return [Math]::Sign($installedVersion.CompareTo($latestVersion))
  }
  catch {
    if ($Installed -eq $Latest) {
      return 0
    }

    return $null
  }
}


function Convert-EvergreenDate {
  param($Value)

  if (-not $Value) {
    return $null
  }

  try {
    return ([datetime]$Value)
  }
  catch {
    try {
      return [datetime]::Parse("$Value")
    }
    catch {
      return $null
    }
  }
}

function Get-EvergreenReleaseInfo {
  param(
    [string]$EvergreenName,
    [string]$InstalledVersion
  )

  if (-not $EvergreenName) {
    return $null
  }

  try {
    $data = @(Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop)

    if (-not $data -or $data.Count -eq 0) {
      return $null
    }

    # Prefer x64 records where Evergreen exposes architecture.
    $filtered = @($data)

    if ($filtered[0].PSObject.Properties.Name -contains 'Architecture') {
      $x64 = @($filtered | Where-Object { $_.Architecture -match 'x64|64|amd64' })
      if ($x64.Count -gt 0) {
        $filtered = $x64
      }
    }

    # Prefer stable/enterprise/monthly records. Explicitly avoid obvious beta,
    # preview, developer and nightly channels where possible.
    if ($filtered[0].PSObject.Properties.Name -contains 'Channel') {
      $stable = @(
        $filtered | Where-Object {
          $_.Channel -match 'Stable|Enterprise|Monthly' -and
          $_.Channel -notmatch 'Beta|Preview|Developer|Dev|Nightly|Canary'
        }
      )

      if ($stable.Count -gt 0) {
        $filtered = $stable
      }
    }

    # Also discard obvious prerelease version strings.
    $filtered = @(
      $filtered | Where-Object {
        -not $_.Version -or "$($_.Version)" -notmatch '(?i)(beta|preview|alpha|nightly|canary|(?:^|[.\-_])?[ab]\d+$)'
      }
    )

    if ($filtered.Count -eq 0) {
      return $null
    }

    $rows = foreach ($item in $filtered) {
      if (-not $item.Version) {
        continue
      }

      $versionObj = $null

      try {
        $versionObj = [version](("$($item.Version)" -replace '[^\d\.]','').Trim('.'))
      }
      catch {}

      # Evergreen apps do not all use the same date property name,
      # so probe the common properties without assuming one schema.
      $releaseDate = $null

      foreach ($propertyName in @(
        'ReleaseDate',
        'ReleaseDateTime',
        'Date',
        'PublishedDate',
        'Published',
        'Updated',
        'LastUpdated'
      )) {
        if ($item.PSObject.Properties.Name -contains $propertyName) {
          $candidateDate = Convert-EvergreenDate $item.$propertyName

          if ($candidateDate) {
            $releaseDate = $candidateDate
            break
          }
        }
      }

      [pscustomobject]@{
        VersionText = "$($item.Version)"
        VersionObj  = $versionObj
        ReleaseDate = $releaseDate
      }
    }

    if (-not $rows) {
      return $null
    }

    $latest = $rows |
      Sort-Object @{ Expression = { if ($_.VersionObj) { $_.VersionObj } else { [version]'0.0' } }; Descending = $true } |
      Select-Object -First 1

    $installedReleaseDate = $null

    if ($InstalledVersion -and $InstalledVersion -ne '-') {
      $installedMatch = $rows |
        Where-Object {
          $_.VersionText -eq $InstalledVersion -or
          (
            $_.VersionObj -and
            (Compare-VersionSmart -Installed $InstalledVersion -Latest $_.VersionText) -eq 0
          )
        } |
        Sort-Object ReleaseDate -Descending |
        Select-Object -First 1

      if ($installedMatch) {
        $installedReleaseDate = $installedMatch.ReleaseDate
      }
    }

    return [pscustomobject]@{
      LatestVersion        = $latest.VersionText
      LatestReleaseDate    = $latest.ReleaseDate
      InstalledReleaseDate = $installedReleaseDate
    }
  }
  catch {
    return $null
  }
}


# ---------------------- Visual Studio 2022 ----------------------
# Full Visual Studio is best detected with vswhere rather than the normal
# Add/Remove Programs entries, because the registry also contains many
# Visual Studio runtimes, installers and shared components.

function Get-VSWherePath {
  $candidates = @(
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
    "$env:ProgramFiles\Microsoft Visual Studio\Installer\vswhere.exe"
  )

  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      return $candidate
    }
  }

  return $null
}

function Get-VisualStudio2022Inventory {
  $vswhere = Get-VSWherePath

  if (-not $vswhere) {
    return @()
  }

  try {
    $json = & $vswhere `
      -all `
      -products * `
      -format json `
      -utf8 2>$null

    if (-not $json) {
      return @()
    }

    $instances = @($json | ConvertFrom-Json)
    $results = @()

    foreach ($instance in $instances) {
      $productId = "$($instance.productId)"

      $wingetId = $null
      $displayName = $null

      switch ($productId) {
        'Microsoft.VisualStudio.Product.Enterprise' {
          $wingetId = 'Microsoft.VisualStudio.2022.Enterprise'
          $displayName = 'Visual Studio 2022 Enterprise'
        }
        'Microsoft.VisualStudio.Product.Professional' {
          $wingetId = 'Microsoft.VisualStudio.2022.Professional'
          $displayName = 'Visual Studio 2022 Professional'
        }
        'Microsoft.VisualStudio.Product.Community' {
          $wingetId = 'Microsoft.VisualStudio.2022.Community'
          $displayName = 'Visual Studio 2022 Community'
        }
        'Microsoft.VisualStudio.Product.BuildTools' {
          $wingetId = 'Microsoft.VisualStudio.2022.BuildTools'
          $displayName = 'Visual Studio 2022 Build Tools'
        }
        default {
          continue
        }
      }

      # Visual Studio exposes two different version schemes:
      # installationVersion is an internal build number such as 17.14.37516.0.
      # catalog.productDisplayVersion is the release/display version such as
      # 17.14.37. WinGet uses the display-style version, so prefer that here.
      $installedVersion = $null

      try {
        if ($instance.catalog -and $instance.catalog.productDisplayVersion) {
          $installedVersion = "$($instance.catalog.productDisplayVersion)"
        }
      }
      catch {}

      if (-not $installedVersion) {
        $installedVersion = "$($instance.installationVersion)"
      }

      $latestVersion = $null
      $installedReleaseDate = $null
      $latestReleaseDate = $null

      if ($wingetReady -and $wingetId) {
        $latestVersion = Get-LatestWingetVersion -WingetId $wingetId

        if ($installedVersion) {
          $installedReleaseDate = Get-WingetReleaseDateForVersion `
            -WingetId $wingetId `
            -Version $installedVersion
        }

        if ($latestVersion) {
          $latestReleaseDate = Get-WingetReleaseDateForVersion `
            -WingetId $wingetId `
            -Version $latestVersion
        }
      }

      $results += [pscustomobject]@{
        Application           = $displayName
        InstalledVersion      = if ($installedVersion) { $installedVersion } else { '-' }
        LatestAvailable       = if ($latestVersion) { $latestVersion } else { '-' }
        LatestSource          = if ($latestVersion) { 'Winget' } else { '-' }
        HasEvergreen          = $false
        InstalledReleaseDate  = $installedReleaseDate
        LatestReleaseDate     = $latestReleaseDate
        InstalledReleaseLabel = $null
      }
    }

    return @($results)
  }
  catch {
    return @()
  }
}


# ---------------------- Vendor historical release metadata ----------------------
# Conservative vendor-documented fallbacks used only when Evergreen/Winget
# do not expose an historical release date for the exact version.
$FSLogixReleaseMap = @{
  '3.26.102.18413' = [datetime]'2026-01-13'
  '3.26.126.19110' = [datetime]'2026-02-10'
}

$ZoomReleaseMap = @{
  '6.7.26346' = [datetime]'2025-12-30'
}

function Get-VendorReleaseDateFallback {
  param(
    [string]$Application,
    [string]$Version
  )

  if (-not $Version -or $Version -eq '-') {
    return $null
  }

  if ($Application -eq 'Microsoft FSLogix Apps' -and $FSLogixReleaseMap.ContainsKey($Version)) {
    return $FSLogixReleaseMap[$Version]
  }

  if ($Application -eq 'Zoom Workplace (64-bit)' -and $ZoomReleaseMap.ContainsKey($Version)) {
    return $ZoomReleaseMap[$Version]
  }

  return $null
}

# ---------------------- Power BI historical release archive ----------------------
# Microsoft publishes a monthly Power BI Desktop archive. Some very old builds
# are no longer present in current Winget manifests, so use the archive as a
# fallback for a month/year release label rather than showing "date unavailable".
#
# These mappings are intentionally conservative: only versions explicitly
# present in Microsoft's archived monthly release history are listed here.
$PowerBIArchiveMap = @{
  '2.64.5285.582'  = 'November 2018'
  '2.65.5313.621'  = 'December 2018'
  '2.66.5376.1681' = 'February 2019'
  '2.67.5404.581'  = 'March 2019'
  '2.68.5432.361'  = 'April 2019'
  '2.69.5467.1251' = 'May 2019'
  '2.70.5494.561'  = 'June 2019'
  '2.71.5523.641'  = 'July 2019'
  '2.72.5556.801'  = 'August 2019'
  '2.73.5586.561'  = 'September 2019'
  '2.74.5619.621'  = 'October 2019'
  '2.75.5649.341'  = 'November 2019'
}

function Get-PowerBIArchiveReleaseLabel {
  param([string]$Version)

  if (-not $Version -or $Version -eq '-') {
    return $null
  }

  if ($PowerBIArchiveMap.ContainsKey($Version)) {
    return $PowerBIArchiveMap[$Version]
  }

  return $null
}

function Get-AllInstalledPrograms {
  $items = @()

  $targets = @(
    @{ Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
    @{ Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
    @{ Hive='CurrentUser';  View='Default';    Path='Software\Microsoft\Windows\CurrentVersion\Uninstall' }
  )

  foreach ($t in $targets) {
    try {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::$($t.Hive),
        [Microsoft.Win32.RegistryView]::$($t.View)
      )

      $key = $base.OpenSubKey($t.Path)
      if (-not $key) { continue }

      foreach ($subName in $key.GetSubKeyNames()) {
        try {
          $sub = $key.OpenSubKey($subName)
          $displayName = $sub.GetValue('DisplayName')
          if (-not $displayName) { continue }

          $items += [pscustomobject]@{
            DisplayName    = "$displayName"
            DisplayVersion = "$($sub.GetValue('DisplayVersion'))"
            Publisher      = "$($sub.GetValue('Publisher'))"
          }
        }
        catch {}
      }
    }
    catch {}
  }

  return $items
}

function Get-InstalledVersionForCsvApp {
  param(
    [string]$CsvName,
    [string]$CsvPublisher,
    [array]$InstalledPrograms
  )

  # Special handling for Office Click-to-Run language entries.
  if ($CsvName -match 'Microsoft 365 Apps for enterprise|Aplicaciones de Microsoft 365') {
    try {
      $office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
      if ($office.VersionToReport) {
        return "$($office.VersionToReport)"
      }
    }
    catch {}
  }

  # First try exact display name.
  $match = $InstalledPrograms |
    Where-Object { $_.DisplayName -ieq $CsvName } |
    Select-Object -First 1

  if ($match -and $match.DisplayVersion) {
    return $match.DisplayVersion
  }

  # Remove trailing version text / architecture text from the CSV name
  # to make matching resilient to version changes.
  $baseName = $CsvName

  $baseName = $baseName -replace '\s+\d+(\.\d+)+.*$',''
  $baseName = $baseName -replace '\s+\(64-bit\)$',''
  $baseName = $baseName -replace '\s+\(x64\)$',''
  $baseName = $baseName -replace '\s+\(64-bit x64\)$',''
  $baseName = $baseName.Trim()

  $candidates = @(
    $InstalledPrograms | Where-Object {
      $_.DisplayName -ilike "$baseName*" -or
      $_.DisplayName -ilike "*$baseName*"
    }
  )

  # Publisher helps avoid false matches where possible.
  if ($CsvPublisher -and $CsvPublisher -ne 'nan') {
    $publisherCandidates = @(
      $candidates | Where-Object {
        $_.Publisher -and (
          $_.Publisher -ilike "*$CsvPublisher*" -or
          $CsvPublisher -ilike "*$($_.Publisher)*"
        )
      }
    )

    if ($publisherCandidates.Count -gt 0) {
      $candidates = $publisherCandidates
    }
  }

  $best = $candidates |
    Where-Object { $_.DisplayVersion } |
    Select-Object -First 1

  if ($best) {
    return $best.DisplayVersion
  }

  # Appx package fallback for modern apps.
  try {
    $appx = Get-AppxPackage -AllUsers |
      Where-Object {
        $_.Name -ilike "*$baseName*" -or
        $_.PackageFullName -ilike "*$baseName*"
      } |
      Sort-Object Version -Descending |
      Select-Object -First 1

    if ($appx) {
      return "$($appx.Version)"
    }
  }
  catch {}

  return '-'
}

# ---------------------- Run inventory ----------------------
$csvApps = Import-Csv -LiteralPath $CsvPath
$installedPrograms = @(Get-AllInstalledPrograms)
$evergreenReady = Test-EvergreenReady
$wingetReady   = Test-WingetReady

if (-not $evergreenReady) {
  Write-Warning "Evergreen module is not installed or could not be loaded. Evergreen-backed latest versions will show as '-'."
}

if (-not $wingetReady) {
  Write-Warning "Winget is not installed or could not be used. Winget-backed latest versions, including Power BI, will show as '-'."
}

$reportRows = foreach ($app in $csvApps) {
  $name = "$($app.Name)".Trim()
  $publisher = "$($app.Publisher)".Trim()

  $installedVersion = Get-InstalledVersionForCsvApp `
    -CsvName $name `
    -CsvPublisher $publisher `
    -InstalledPrograms $installedPrograms

  $latestVersion        = $null
  $latestSource         = '-'
  $installedReleaseDate = $null
  $latestReleaseDate    = $null
  $hasEvergreen         = $false

  if ($WingetMap.ContainsKey($name)) {
    if ($wingetReady) {
      $latestVersion = Get-LatestWingetVersion -WingetId $WingetMap[$name]

      if ($latestVersion) {
        $latestSource = 'Winget'
      }
    }
  }
  else {
    $evergreenName = $null

    if ($EvergreenMap.ContainsKey($name)) {
      $evergreenName = $EvergreenMap[$name]
      $hasEvergreen = $true
    }

    if ($evergreenReady -and $evergreenName) {
      $releaseInfo = Get-EvergreenReleaseInfo `
        -EvergreenName $evergreenName `
        -InstalledVersion $installedVersion

      if ($releaseInfo) {
        $latestVersion        = $releaseInfo.LatestVersion
        $installedReleaseDate = $releaseInfo.InstalledReleaseDate
        $latestReleaseDate    = $releaseInfo.LatestReleaseDate

        if ($latestVersion) {
          $latestSource = 'Evergreen'
        }
      }
    }
  }

  # Evergreen generally exposes only current release metadata. For an older
  # installed version, try the exact historical Winget manifest as a second
  # source for its release date.
  if (
    -not $installedReleaseDate -and
    $installedVersion -and
    $installedVersion -ne '-' -and
    $VersionDateWingetMap.ContainsKey($name) -and
    $wingetReady
  ) {
    $installedReleaseDate = Get-WingetReleaseDateForVersion `
      -WingetId $VersionDateWingetMap[$name] `
      -Version $installedVersion
  }

  # If the latest version came from Winget, also try to enrich it with its
  # manifest Release Date.
  if (
    -not $latestReleaseDate -and
    $latestVersion -and
    $latestVersion -ne '-' -and
    $VersionDateWingetMap.ContainsKey($name) -and
    $wingetReady
  ) {
    $latestReleaseDate = Get-WingetReleaseDateForVersion `
      -WingetId $VersionDateWingetMap[$name] `
      -Version $latestVersion
  }

  # Vendor-specific fallback when neither Evergreen nor exact Winget history
  # returned a release date.
  if (-not $installedReleaseDate -and $installedVersion -and $installedVersion -ne '-') {
    $installedReleaseDate = Get-VendorReleaseDateFallback `
      -Application $name `
      -Version $installedVersion
  }

  if (-not $latestReleaseDate -and $latestVersion -and $latestVersion -ne '-') {
    $latestReleaseDate = Get-VendorReleaseDateFallback `
      -Application $name `
      -Version $latestVersion
  }

  # Firefox ESR: never present a beta/prerelease build as the current stable
  # release. If ESR is not installed, simply leave latest as '-'.
  if (
    $name -match '^Mozilla Firefox ESR' -and
    (
      -not $installedVersion -or
      $installedVersion -eq '-' -or
      ($latestVersion -and "$latestVersion" -match '(?i)(beta|preview|alpha|nightly|canary|[._-]?b\d+$)')
    )
  ) {
    $latestVersion     = $null
    $latestReleaseDate = $null
    $latestSource      = '-'
  }

  [pscustomobject]@{
    Application           = $name
    InstalledVersion      = if ($installedVersion) { $installedVersion } else { '-' }
    LatestAvailable       = if ($latestVersion) { $latestVersion } else { '-' }
    LatestSource          = $latestSource
    HasEvergreen          = $hasEvergreen
    InstalledReleaseDate  = $installedReleaseDate
    LatestReleaseDate     = $latestReleaseDate
    InstalledReleaseLabel = $(if ($name -eq 'Microsoft Power BI Desktop (x64)') { Get-PowerBIArchiveReleaseLabel -Version $installedVersion } else { $null })
  }
}

# Add full Visual Studio 2022 installations as dedicated rows.
# This intentionally does NOT treat "Visual Studio Installer", "Tools for
# Applications", runtimes, SDKs, etc. as the Visual Studio IDE.
$visualStudioRows = @(Get-VisualStudio2022Inventory)

foreach ($vsRow in $visualStudioRows) {
  $alreadyPresent = @(
    $reportRows | Where-Object { $_.Application -ieq $vsRow.Application }
  ).Count -gt 0

  if (-not $alreadyPresent) {
    $reportRows += $vsRow
  }
}


# ---------------------- Output ----------------------
$baseName = "$machine AVD Application Inventory $todayIso"
$outCsv   = Join-Path $monthFolder ($baseName + '.csv')
$outHtml  = Join-Path $monthFolder ($baseName + '.html')

$reportRows |
  Sort-Object @{Expression='HasEvergreen';Descending=$true}, Application |
  Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCsv

$css = @"
<style>
body {
  font-family: Segoe UI, Arial, sans-serif;
  margin: 24px;
  color: #222;
}
h1 {
  margin-bottom: 4px;
}
.subtitle {
  color: #666;
  margin-bottom: 18px;
}
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
  table-layout: auto;
}
th:first-child,
td:first-child {
  width: 1%;
  white-space: nowrap;
}
th {
  background: #f2f4f7;
  border: 1px solid #d9dde3;
  padding: 9px;
  text-align: left;
}
td {
  border: 1px solid #d9dde3;
  padding: 9px;
}
tr:nth-child(even) {
  background: #fafafa;
}
.missing {
  color: #b42318;
  font-weight: 600;
}
.outdated {
  color: #b42318;
  font-weight: 700;
}
.current {
  color: #067647;
  font-weight: 600;
}
.evergreen {
  color: #067647;
  font-weight: 700;
}
.footer {
  margin-top: 16px;
  color: #777;
  font-size: 12px;
}
</style>
"@

$htmlRows = foreach ($row in (
  $reportRows |
  Sort-Object @{Expression='HasEvergreen';Descending=$true}, Application
)) {

  $installedClass = ''
  $latestClass = ''

  $installedDisplay = $row.InstalledVersion
  $latestDisplay    = $row.LatestAvailable

  if ($row.InstalledVersion -eq '-') {
    $installedClass = 'missing'
  }

  if ($row.LatestAvailable -ne '-') {
    $latestClass = 'evergreen'

    # Do not compare "-" against a version and do not append
    # "(release date unavailable)" to applications that are not installed.
    if ($row.InstalledVersion -ne '-') {
      $comparison = Compare-VersionSmart `
        -Installed $row.InstalledVersion `
        -Latest $row.LatestAvailable

      if ($comparison -lt 0) {
        $installedClass = 'outdated'

        if ($row.InstalledReleaseDate) {
          $installedDisplay = "{0} (released {1})" -f `
            $row.InstalledVersion,
            ([datetime]$row.InstalledReleaseDate).ToString('dd/MM/yyyy')
        }
        elseif ($row.InstalledReleaseLabel) {
          $installedDisplay = "{0} ({1} release)" -f `
            $row.InstalledVersion,
            $row.InstalledReleaseLabel
        }
        else {
          $installedDisplay = "{0} (release date unavailable)" -f $row.InstalledVersion
        }

        if ($row.LatestReleaseDate) {
          $latestDisplay = "{0} (available as of {1})" -f `
            $row.LatestAvailable,
            ([datetime]$row.LatestReleaseDate).ToString('dd/MM/yyyy')
        }
      }
      elseif ($comparison -ge 0) {
        $installedClass = 'current'
      }
    }
  }

  "<tr><td>$($row.Application)</td><td class='$installedClass'>$installedDisplay</td><td class='$latestClass'>$latestDisplay</td><td>$($row.LatestSource)</td></tr>"
}

$html = @"
<html>
<head>
<meta charset='utf-8'>
<title>$baseName</title>
$css
</head>
<body>

<h1>AVD Application Inventory</h1>
<div class='subtitle'>
Machine: <strong>$machine</strong> &nbsp; | &nbsp;
Generated: $(Get-Date -Format 'dd/MM/yyyy HH:mm')
</div>

<table>
<thead>
<tr>
<th>Application</th>
<th>Installed Version</th>
<th>Latest Available Version</th>
<th>Source</th>
</tr>
</thead>
<tbody>
$(($htmlRows -join "`r`n"))
</tbody>
</table>

<div class='footer'>
Applications with an Evergreen source are listed first. Installed versions older than the latest available version are shown in red. For outdated applications, the script performs a best-effort historical version lookup using Evergreen and exact Winget manifests. For older Power BI Desktop builds that are no longer available in Winget history, the script falls back to Microsoft's archived monthly Power BI release history and shows the month/year release label. Full Visual Studio 2022 installations are detected with vswhere and checked against the matching WinGet package for Enterprise, Professional, Community or Build Tools. FSLogix and Zoom also use conservative vendor-documented historical release mappings when Evergreen/Winget history is missing. Firefox ESR prerelease/beta values are suppressed, and applications that are not installed are shown simply as '-' without a release-date warning. If no trustworthy historical date or archive label is available for an installed outdated app, the report says release date unavailable rather than guessing.
</div>

</body>
</html>
"@

$html | Set-Content -LiteralPath $outHtml -Encoding UTF8

Write-Host ""
Write-Host "AVD application inventory completed." -ForegroundColor Green
Write-Host "CSV : $outCsv" -ForegroundColor Cyan
Write-Host "HTML: $outHtml" -ForegroundColor Cyan
