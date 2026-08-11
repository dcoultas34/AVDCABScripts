param(
  [string]$ReportRoot = "C:\Ultima\Reporting\AVDPatchReports",

  # Create only the monthly BEFORE baseline and then exit.
  [switch]$BeforeOnly,

  # Create only the AFTER comparison/report.
  # Requires an existing BEFORE baseline for the current month.
  [switch]$AfterOnly
)

if ($BeforeOnly -and $AfterOnly) {
  throw "Use either -BeforeOnly or -AfterOnly, not both."
}


# ---------------------- AVD master images ----------------------
$MasterMap = @(
  @{ Master='CUK-AVDGI-P-001'; Environment='Production' }
  @{ Master='CUK-AVDGI-P-002'; Environment='Production' }
) | ForEach-Object { [pscustomobject]$_ }

# ---------------------- Applications to check ----------------------
$AppsToCheck = @(
  @{ Name="Adobe Acrobat";           IsAdobeFull=$true;  LocalMatch=$null; EvergreenName=$null; WingetId=$null; ExpectEvergreen=$false }
  @{ Name="Microsoft Office 365"; LocalMatch="__OFFICE_C2R__"; EvergreenName="Microsoft365Apps"; PreferredChannel="MonthlyEnterprise"; WingetId="Microsoft.Office"; ExpectEvergreen=$true }
  #@{ Name="Adobe Acrobat Reader";    LocalMatch="Adobe Acrobat Reader"; EvergreenName="AdobeAcrobatReaderDC"; WingetId="Adobe.Acrobat.Reader.64-bit"; ExpectEvergreen=$false }
  @{ Name="Google Chrome";           LocalMatch="Google Chrome";        EvergreenName="GoogleChrome";         WingetId="Google.Chrome";               ExpectEvergreen=$false }
  @{ Name="Mozilla Firefox";         LocalMatch="Mozilla Firefox";      EvergreenName=$null;                   WingetId="Mozilla.Firefox";            ExpectEvergreen=$false }

  @{ Name="Microsoft Edge";          LocalMatch="Microsoft Edge";       EvergreenName="MicrosoftEdge";        WingetId="Microsoft.Edge";              ExpectEvergreen=$false }
  @{ Name="Visual Studio Code";      LocalMatch="Microsoft Visual Studio Code"; EvergreenName="MicrosoftVisualStudioCode"; WingetId="Microsoft.VisualStudioCode"; ExpectEvergreen=$false }
  @{ Name="Power BI Desktop";        LocalMatch="Microsoft Power BI Desktop";   EvergreenName="MicrosoftPowerBIDesktop";    WingetId="Microsoft.PowerBI";         ExpectEvergreen=$false }
  @{ Name="Microsoft Teams";         LocalMatch="Microsoft Teams";      EvergreenName=$null; WingetId="Microsoft.Teams";  ExpectEvergreen=$false }
  @{ Name="OneDrive";                LocalMatch="Microsoft OneDrive";   EvergreenName=$null; WingetId="Microsoft.OneDrive"; ExpectEvergreen=$false }
)

# ---------------------- Utilities ----------------------
function Test-WingetReady  { try { $null = Get-Command winget -ErrorAction Stop; winget --version | Out-Null; $true } catch { $false } }
function Test-EvergreenReady { try { $m = Get-Module -ListAvailable -Name Evergreen; if ($m){Import-Module Evergreen -ErrorAction SilentlyContinue|Out-Null;$true}else{$false} } catch { $false } }

function Get-InstalledVersion {
  param([string]$DisplayNameMatch)

  if ($DisplayNameMatch -eq "Microsoft Teams") {
    $pkgs=@(); foreach($n in 'MSTeams','MicrosoftTeams'){ try{$pkgs+=Get-AppxPackage -AllUsers -Name $n -ErrorAction SilentlyContinue}catch{}; try{$pkgs+=Get-AppxPackage -Name $n -ErrorAction SilentlyContinue}catch{} }
    $pkg=$pkgs|Where-Object{$_}|Sort-Object Version -Descending|Select-Object -First 1
    if($pkg){ return $pkg.Version.ToString() }
  }

  if ($DisplayNameMatch -eq "__OFFICE_C2R__") {
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
      $sub=$base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration')
      $v=$sub.GetValue('VersionToReport'); if($v){return $v}
    }catch{}; return $null
  }

  if ($DisplayNameMatch -eq "Microsoft Visual Studio Code") {
    foreach($p in @("$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe","$env:ProgramFiles\Microsoft VS Code\Code.exe","${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe")){
      if(Test-Path $p){ try{ return (Get-Item $p).VersionInfo.ProductVersion }catch{} }
    }
  }

  $entries=@()
  $targets=@(
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='CurrentUser' ; View='Default'   ; Path='Software\Microsoft\Windows\CurrentVersion\Uninstall'}
  )
  foreach($t in $targets){
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$($t.Hive),[Microsoft.Win32.RegistryView]::$($t.View))
      $key=$base.OpenSubKey($t.Path); if(-not $key){continue}
      foreach($n in $key.GetSubKeyNames()){
        try{ $s=$key.OpenSubKey($n); $dn=$s.GetValue('DisplayName'); if(-not $dn){continue}; $dv=$s.GetValue('DisplayVersion'); $entries+=[pscustomobject]@{DisplayName=$dn;DisplayVersion=$dv} }catch{}
      }
    }catch{}
  }
  $nameExcludes='WebView','Runtime','WebView2','Updater','Update','AutoUpdate','Maintenance','Service','Helper','Crashpad','Stub','Machine-wide','User Installer','System Installer','Setup'
  $cands=@()
  foreach($e in $entries){
    if($e.DisplayName -notlike "*$DisplayNameMatch*"){continue}
    if($nameExcludes | Where-Object { $e.DisplayName -match [regex]::Escape($_) }){continue}
    $dv=$e.DisplayVersion
    $vObj=$null; if($dv){ try{$vObj=[version]($dv -replace '[^\d\.]','')}catch{} }
    $score= if($e.DisplayName -ieq $DisplayNameMatch){3}elseif($e.DisplayName -ilike "$DisplayNameMatch*"){2}else{1}
    $cands += [pscustomobject]@{DisplayName=$e.DisplayName;DisplayVersion=$dv;VersionObj=$vObj;Score=$score}
  }
  if($cands.Count -eq 0){return $null}
  ($cands | Sort-Object @{e='Score';Descending=$true}, @{e='VersionObj';Descending=$true} | Select-Object -First 1).DisplayVersion
}

# -------- Adobe Acrobat (Full/DC/Pro) robust detection --------
function Get-AdobeFullInstalledVersion {
  $paths = @(
    'HKLM:\SOFTWARE\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe Acrobat',
    'HKLM:\SOFTWARE\Adobe\Acrobat',
    'HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat'
  )
  foreach ($root in $paths) {
    if (Test-Path $root) {
      Get-ChildItem $root -EA SilentlyContinue | ForEach-Object {
        foreach ($leaf in 'Installer','CurrentVersion','\DC\Installer','\DC\CurrentVersion') {
          $k = ($_.PsPath + $leaf)
          if (Test-Path $k) {
            try {
              $p = Get-ItemProperty -Path $k -EA Stop
              foreach($name in 'ACPVersion','ProductVersion','Version','PV') {
                $v = $p.$name
                if ($v -and ($v -match '^\d{1,2}\.\d{3}\.\d{5}$')) { return $v }
              }
            } catch {}
          }
        }
      }
    }
  }
  $targets = @(
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'},
    @{Hive='CurrentUser' ; View='Default'   ; Path='Software\Microsoft\Windows\CurrentVersion\Uninstall'}
  )
  foreach($t in $targets){
    try{
      $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$($t.Hive),[Microsoft.Win32.RegistryView]::$($t.View))
      $key=$base.OpenSubKey($t.Path); if(-not $key){continue}
      foreach($n in $key.GetSubKeyNames()){
        try{
          $s=$key.OpenSubKey($n)
          $dn=$s.GetValue('DisplayName'); if(-not $dn){continue}
          if($dn -match '^Adobe\s+Acrobat(?!.*Reader)'){
            foreach ($prop in 'DisplayVersion','BundleVersion') {
              $dv=$s.GetValue($prop); if($dv){ return $dv }
            }
          }
        }catch{}
      }
    }catch{}
  }
  $exe = Get-ChildItem "$env:ProgramFiles\Adobe","${env:ProgramFiles(x86)}\Adobe" -Recurse -Filter Acrobat.exe -EA SilentlyContinue | Select-Object -ExpandProperty FullName
  foreach($p in $exe){
    try{
      $fv = (Get-Item $p).VersionInfo.FileVersion
      if ($fv -and ($fv -match '^\d{1,2}\.\d{3}\.\d{5}$')) { return $fv }
    }catch{}
  }
  return $null
}
function Get-LatestFromAdobeAcrobatFull {
  # Method 1: winget. This is the preferred method because the main script
  # already uses winget for application version lookups.
  try {
    if (Test-WingetReady) {
      $out = winget show `
        --id Adobe.Acrobat.Pro `
        --exact `
        --source winget `
        --accept-source-agreements 2>$null

      if ($out) {
        $verLine = ($out -split "`r?`n") |
          Where-Object { $_ -match '^\s*Version\s*:' } |
          Select-Object -First 1

        if ($verLine) {
          $v = ($verLine -split ':\s*',2)[1].Trim()

          if ($v -match '^\d{2}\.\d{3}\.\d{5}$') {
            return $v
          }
        }
      }
    }
  }
  catch {
    Write-Verbose "winget Acrobat lookup unavailable: $($_.Exception.Message)"
  }

  # Method 2: Adobe lightweight current-version endpoint.
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $v = (Invoke-RestMethod `
      -Uri 'https://armmf.adobe.com/arm-manifests/win/AcrobatDC/acrobat/acrobat/current_version.txt' `
      -UseBasicParsing `
      -TimeoutSec 15 `
      -ErrorAction Stop).Trim()

    if ($v -match '^\d{2}\.\d{3}\.\d{5}$') {
      return $v
    }
  }
  catch {
    Write-Verbose "Adobe current-version endpoint unavailable: $($_.Exception.Message)"
  }

  # Method 3: Adobe Continuous Track release notes.
  try {
    $uri = 'https://helpx.adobe.com/acrobat/release-note/release-notes-acrobat-reader.html'

    $response = Invoke-WebRequest `
      -Uri $uri `
      -UseBasicParsing `
      -TimeoutSec 20 `
      -ErrorAction Stop

    $matches = [regex]::Matches(
      $response.Content,
      '\b\d{2}\.\d{3}\.\d{5}\b'
    )

    if ($matches.Count -gt 0) {
      $versions = foreach ($m in $matches) {
        try {
          [pscustomobject]@{
            Text    = $m.Value
            Version = [version]$m.Value
          }
        }
        catch {}
      }

      $latest = $versions |
        Sort-Object Version -Descending |
        Select-Object -First 1

      if ($latest) {
        return $latest.Text
      }
    }
  }
  catch {
    Write-Verbose "Adobe release-note lookup unavailable: $($_.Exception.Message)"
  }

  return $null
}

function Get-LatestFromEvergreen {
  param([string]$EvergreenName,[string]$PreferredChannel)
  try {
    $data = Get-EvergreenApp -Name $EvergreenName -ErrorAction Stop
    if ($PreferredChannel -and ($data | Get-Member -Name Channel -EA SilentlyContinue)) {
      $pref = $data | Where-Object { $_.Channel -match $PreferredChannel }
      if ($pref) { $data = $pref }
    }
    if ($data -and ($data | Get-Member -Name Channel -EA SilentlyContinue)) {
      $stable = $data | Where-Object { $_.Channel -match 'Stable' }
      if ($stable) { $data = $stable }
    }
    $top = $data | Where-Object { $_.Version } |
           Sort-Object { try{ [version]($_.Version -replace '[^\d\.]','') }catch{ $_.Version } } -Descending |
           Select-Object -First 1
    if ($top) { [pscustomobject]@{ Version = $top.Version } }
  } catch { $null }
}
function Get-LatestFromWinget {
  param([string]$WingetId)
  try{
    $out = winget show --id $WingetId --exact --source winget --accept-source-agreements 2>$null
    if($out){
      $verLine = ($out -split "`r?`n") | Where-Object { $_ -match "^\s*Version\s*:" } | Select-Object -First 1
      $ver = if($verLine){ ($verLine -split ":\s*",2)[1].Trim() } else { $null }
      [pscustomobject]@{ Version=$ver }
    }
  }catch{}
}
function Compare-VersionSmart { param([string]$Installed,[string]$Latest)
  if(-not $Installed -or -not $Latest){ return $null }
  try{ $v1=[version]($Installed -replace '[^\d\.]',''); $v2=[version]($Latest -replace '[^\d\.]',''); [Math]::Sign($v1.CompareTo($v2)) }
  catch{ if($Installed -eq $Latest){0}else{-1} }
}
function Get-RecentOSUpdates {
  try {
    Get-HotFix | Sort-Object InstalledOn -Descending | ForEach-Object {
      [pscustomobject]@{
        Computer    = $env:COMPUTERNAME
        KB          = $_.HotFixID
        Description = $_.Description
        InstalledOn = $(if ($_.InstalledOn) { [datetime]$_.InstalledOn } else { $null })
      }
    }
  } catch { @() }
}

# ---------------------- Local report paths & dates ----------------------
$today     = Get-Date
$todayIso  = $today.ToString('yyyy-MM-dd')
$todayUK   = $today.ToString('dd/MM/yy')
$Period    = $today.ToString("MMMM yyyy")

if (-not (Test-Path -LiteralPath $ReportRoot)) {
  New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
}

$monthFolder = Join-Path $ReportRoot $Period

if (-not (Test-Path -LiteralPath $monthFolder)) {
  New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null
  Write-Host "Created monthly report folder: $monthFolder" -ForegroundColor Green
}

# ---------------------- Channel versions (Evergreen + winget) ----------------------
function Get-ChannelVersions {
  param([string]$EvergreenName,[string]$WingetId,[string]$PreferredChannel)
  $ever = $null; $wing = $null
  if ($EvergreenName -and (Test-EvergreenReady)) { $ever = Get-LatestFromEvergreen -EvergreenName $EvergreenName -PreferredChannel $PreferredChannel }
  if ($WingetId -and (Test-WingetReady))         { $wing = Get-LatestFromWinget    -WingetId       $WingetId }
  [pscustomobject]@{
    EvergreenStable = if($ever  -and $ever.PSObject.Properties['Version']){ $ever.Version } else { $null }
    WingetLatest    = if($wing  -and $wing.PSObject.Properties['Version']){ $wing.Version } else { $null }
  }
}


# ---------------------- Other applications installed/updated this month ----------------------
function Get-ApplicationsChangedThisMonth {
  param(
    [datetime]$MonthStart,
    [datetime]$MonthEnd
  )

  $results = @()

  $targets = @(
    @{ Hive='LocalMachine'; View='Registry64'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
    @{ Hive='LocalMachine'; View='Registry32'; Path='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' }
  )

  foreach ($t in $targets) {
    try {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::$($t.Hive),
        [Microsoft.Win32.RegistryView]::$($t.View)
      )

      $key = $base.OpenSubKey($t.Path)
      if (-not $key) { continue }

      foreach ($name in $key.GetSubKeyNames()) {
        try {
          $sub = $key.OpenSubKey($name)

          $displayName = $sub.GetValue('DisplayName')
          if (-not $displayName) { continue }

          $displayVersion = $sub.GetValue('DisplayVersion')
          $publisher      = $sub.GetValue('Publisher')
          $installDateRaw = $sub.GetValue('InstallDate')

          if (-not $installDateRaw) { continue }

          $installDate = $null

          foreach ($fmt in @('yyyyMMdd','yyyy-MM-dd','dd/MM/yyyy','MM/dd/yyyy')) {
            try {
              $installDate = [datetime]::ParseExact(
                "$installDateRaw",
                $fmt,
                [System.Globalization.CultureInfo]::InvariantCulture
              )
              break
            } catch {}
          }

          if (-not $installDate) {
            try { $installDate = [datetime]::Parse("$installDateRaw") } catch {}
          }

          if ($installDate -and $installDate -ge $MonthStart -and $installDate -lt $MonthEnd) {
            $results += [pscustomobject]@{
              Application = "$displayName"
              Version     = if ($displayVersion) { "$displayVersion" } else { "-" }
              Publisher   = if ($publisher) { "$publisher" } else { "-" }
              ChangeDate  = $installDate
            }
          }
        }
        catch {}
      }
    }
    catch {}
  }

  $results |
    Sort-Object Application, Version, ChangeDate -Unique
}

# ---------------------- Per-AVD-master run ----------------------
$machine  = $env:COMPUTERNAME
$wingetOk = Test-WingetReady
$everOk   = Test-EvergreenReady

$expectedMasters = $MasterMap | Select-Object -ExpandProperty Master
if ($machine -notin $expectedMasters) {
  Write-Warning "This script is intended for AVD master images: $($expectedMasters -join ', '). Current machine: $machine"
}

Write-Host "Reporting for $machine" -ForegroundColor Cyan
Write-Host "Local report folder: $monthFolder" -ForegroundColor Cyan

if (-not $everOk) {
  Write-Warning "Evergreen module not available. Install with: Install-Module Evergreen -Scope AllUsers"
}
if (-not $wingetOk) {
  Write-Warning "winget not available. Some latest-version information may be missing."
}

# ---------------------- OS information ----------------------
try {
  $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue

  $caption = $os.Caption
  if ($reg.DisplayVersion) {
    $caption = "$caption $($reg.DisplayVersion)"
  }

  $buildFull = if ($reg.UBR -ne $null) {
    "$($os.BuildNumber).$($reg.UBR)"
  }
  else {
    "$($os.BuildNumber)"
  }

  $osInfo = [pscustomobject]@{
    Computer  = $machine
    OSName    = $caption
    OSVersion = $os.Version
    OSBuild   = $buildFull
  }
}
catch {
  $osInfo = [pscustomobject]@{
    Computer  = $machine
    OSName    = '-'
    OSVersion = '-'
    OSBuild   = '-'
  }
}

# ---------------------- Check tracked applications ----------------------
$rows = @()

foreach ($app in $AppsToCheck) {

  if ($app.IsAdobeFull) {
    $installed = Get-AdobeFullInstalledVersion

    if (-not $installed) {
      $rows += [pscustomobject]@{
        Computer   = $machine
        Application= $app.Name
        Installed  = '-'
        Latest     = '-'
        Status     = 'Application not installed'
        Css        = ''
        LatestCss  = ''
      }
      continue
    }

    $adobeLatest = Get-LatestFromAdobeAcrobatFull

    if ($adobeLatest) {
      $cmp = Compare-VersionSmart -Installed $installed -Latest $adobeLatest

      if ($cmp -eq 0) {
        $status = "Latest version installed as of $todayUK"
        $css = 'ok'
      }
      elseif ($cmp -lt 0) {
        $status = 'Update available'
        $css = 'bad'
      }
      else {
        $status = "Ahead of catalog as of $todayUK"
        $css = 'ok'
      }

      $rows += [pscustomobject]@{
        Computer   = $machine
        Application= $app.Name
        Installed  = $installed
        Latest     = $adobeLatest
        Status     = $status
        Css        = $css
        LatestCss  = $css
      }
    }
    else {
      $rows += [pscustomobject]@{
        Computer   = $machine
        Application= $app.Name
        Installed  = $installed
        Latest     = '-'
        Status     = 'Unknown (winget and Adobe lookups unavailable)'
        Css        = ''
        LatestCss  = ''
      }
    }

    continue
  }

  $installed = if ($app.LocalMatch) {
    Get-InstalledVersion -DisplayNameMatch $app.LocalMatch
  }
  else {
    $null
  }

  if (-not $installed) {
    $rows += [pscustomobject]@{
      Computer   = $machine
      Application= $app.Name
      Installed  = '-'
      Latest     = '-'
      Status     = 'Application not installed'
      Css        = ''
      LatestCss  = ''
    }
    continue
  }

  $chan = Get-ChannelVersions `
    -EvergreenName $app.EvergreenName `
    -WingetId $app.WingetId `
    -PreferredChannel $app.PreferredChannel

  $latestToShow = $null
  $status = ''
  $css = ''
  $latestCss = ''

  if ($chan.EvergreenStable) {
    $ev = $chan.EvergreenStable
    $latestToShow = "$ev (Latest Evergreen stable version)"

    if ($chan.WingetLatest) {
      $cmpWingVsEver = Compare-VersionSmart -Installed $chan.WingetLatest -Latest $ev
      if ($cmpWingVsEver -gt 0) {
        $latestToShow = "$($chan.WingetLatest) (waiting on Evergreen release)"
        $latestCss = 'info'
      }
    }

    $cmp = Compare-VersionSmart -Installed $installed -Latest $ev

    if ($cmp -eq 0) {
      $status = "Latest Evergreen stable version installed as of $todayUK"
      $css = 'ok'
    }
    elseif ($cmp -lt 0) {
      $status = 'Update available (vs Evergreen Stable)'
      $css = 'bad'
      if (-not $latestCss) {
        $latestCss = 'bad'
      }
    }
    else {
      $status = "Ahead of Evergreen Stable as of $todayUK"
      $css = 'ok'
    }
  }
  else {
    if ($chan.WingetLatest) {
      $cmp = Compare-VersionSmart -Installed $installed -Latest $chan.WingetLatest

      if ($cmp -eq 0) {
        $latestToShow = $chan.WingetLatest
        $status = "Latest version installed as of $todayUK"
        $css = 'ok'
      }
      elseif ($cmp -gt 0) {
        $latestToShow = $chan.WingetLatest
        $status = "Ahead of catalog as of $todayUK"
        $css = 'ok'
      }
      else {
        if ($app.ExpectEvergreen) {
          $latestToShow = "$($chan.WingetLatest) (Evergreen not available)"
          $status = 'Evergreen not available'
          $css = 'bad'
          $latestCss = 'bad'
        }
        else {
          $latestToShow = $chan.WingetLatest
          $status = 'Update available (catalog)'
          $css = 'bad'
          $latestCss = 'bad'
        }
      }
    }
    else {
      $latestToShow = '-'
      $status = 'Unknown (no latest info)'
      $css = ''
    }
  }

  if (-not $latestToShow) {
    $latestToShow = '-'
  }

  $rows += [pscustomobject]@{
    Computer   = $machine
    Application= $app.Name
    Installed  = $installed
    Latest     = $latestToShow
    Status     = $status
    Css        = $css
    LatestCss  = $latestCss
  }
}

# ---------------------- BEFORE / AFTER monthly snapshot ----------------------
$baselinePath = Join-Path $monthFolder "$machine AVD Patching BEFORE.csv"

# BEFORE-only mode:
# Create/refresh the BEFORE snapshot from the current detected application versions,
# then stop before creating AFTER/HTML reports.
if ($BeforeOnly) {
  $rows |
    Sort-Object Application |
    Select-Object Computer,Application,Installed |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $baselinePath -Force

  Write-Host ""
  Write-Host "Created BEFORE snapshot only:" -ForegroundColor Green
  Write-Host $baselinePath -ForegroundColor Cyan
  return
}

# Normal mode:
# First run of the month becomes the BEFORE snapshot and is not overwritten.
if (-not (Test-Path -LiteralPath $baselinePath)) {
  if ($AfterOnly) {
    throw "AFTER-only mode requires an existing BEFORE snapshot for this month: $baselinePath"
  }

  $rows |
    Sort-Object Application |
    Select-Object Computer,Application,Installed |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $baselinePath

  Write-Host "Created BEFORE snapshot: $baselinePath" -ForegroundColor Yellow
}

try {
  $beforeRows = @(Import-Csv -LiteralPath $baselinePath -ErrorAction Stop)
}
catch {
  throw "Unable to read BEFORE snapshot: $baselinePath. $($_.Exception.Message)"
}

$comparisonRows = foreach ($current in ($rows | Sort-Object Application)) {

  $before = $beforeRows |
    Where-Object { $_.Application -eq $current.Application } |
    Select-Object -First 1

  $installedBefore = if ($before -and $before.Installed) {
    $before.Installed
  }
  else {
    '-'
  }

  $installedAfter = if ($current.Installed) {
    $current.Installed
  }
  else {
    '-'
  }

  if ($installedBefore -eq '-' -and $installedAfter -ne '-') {
    $change = 'Installed this month'
  }
  elseif ($installedBefore -ne '-' -and $installedAfter -eq '-') {
    $change = 'Removed this month'
  }
  elseif ($installedBefore -ne $installedAfter) {
    $change = 'Updated'
  }
  else {
    $change = 'No change'
  }

  [pscustomobject]@{
    Computer        = $machine
    Application     = $current.Application
    InstalledBefore = $installedBefore
    InstalledAfter  = $installedAfter
    Latest          = $current.Latest
    Status          = $current.Status
    Change          = $change
    Css             = $current.Css
    LatestCss       = $current.LatestCss
  }
}

# ---------------------- Report files ----------------------
$baseName = "$machine AVD Patching Report $todayIso"

# The first CSV created in the month is the BEFORE snapshot.
# Every subsequent/current comparison export is clearly named AFTER.
$afterCsvPath = Join-Path $monthFolder "$machine AVD Patching AFTER.csv"
$htmlPath     = Join-Path $monthFolder ($baseName + '.html')

$afterCsvCreated = $false

try {
  $comparisonRows |
    Select-Object Computer,Application,InstalledBefore,InstalledAfter,Latest,Status,Change |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $afterCsvPath -ErrorAction Stop

  if (Test-Path -LiteralPath $afterCsvPath) {
    $afterCsvCreated = $true
    Write-Host "Created AFTER snapshot: $afterCsvPath" -ForegroundColor Green
  }
}
catch {
  Write-Error "Failed to create AFTER snapshot: $($_.Exception.Message)"
}

# ---------------------- Other monthly information ----------------------
$osUpdates = @(Get-RecentOSUpdates | Select-Object -First 10)

$monthStart = Get-Date -Year $today.Year -Month $today.Month -Day 1 -Hour 0 -Minute 0 -Second 0
$monthEnd   = $monthStart.AddMonths(1)

$otherAppsChanged = @(
  Get-ApplicationsChangedThisMonth -MonthStart $monthStart -MonthEnd $monthEnd
)

# ---------------------- HTML ----------------------
if ($afterCsvCreated) {

$css = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; }
h1 { margin-bottom: 6px; }
h2 { margin-top: 26px; }
.summary { margin-bottom: 20px; }
.tag { display:inline-block; background:#e5e7eb; padding:3px 10px; border-radius:12px; margin-right:6px; font-size:12px; }
table { border-collapse: collapse; width:100%; margin-top:10px; }
th, td { border:1px solid #ddd; padding:8px; text-align:left; }
th { background:#f3f3f3; }
.ok { color:#2e7d32; font-weight:600; }
.bad { color:#c62828; font-weight:600; }
.info { color:#1d4ed8; font-weight:600; }
.changed { color:#2e7d32; font-weight:600; }
.small { color:#6b7280; font-size:12px; margin-top:18px; }
</style>
"@

$appRows = foreach ($a in ($comparisonRows | Where-Object {
  $_.InstalledBefore -ne '-' -or $_.InstalledAfter -ne '-'
} | Sort-Object Application)) {

  $rowCls = $a.Css
  $latestCls = if ($a.LatestCss) {
    $a.LatestCss
  }
  else {
    $rowCls
  }

  $changeCls = if ($a.Change -eq 'No change') {
    ''
  }
  elseif ($a.Change -eq 'Removed this month') {
    'bad'
  }
  else {
    'changed'
  }

  "<tr><td class='$rowCls'>$($a.Application)</td><td>$($a.InstalledBefore)</td><td class='$rowCls'>$($a.InstalledAfter)</td><td class='$latestCls'>$($a.Latest)</td><td class='$rowCls'>$($a.Status)</td><td class='$changeCls'>$($a.Change)</td></tr>"
}

if (-not $appRows) {
  $appRows = @("<tr><td colspan='6'><em>No matching applications were found.</em></td></tr>")
}

$osRows = foreach ($o in $osUpdates) {
  $installedOn = if ($o.InstalledOn) {
    $o.InstalledOn.ToString('dd/MM/yyyy')
  }
  else {
    '-'
  }

  "<tr><td>$($o.KB)</td><td>$($o.Description)</td><td>$installedOn</td></tr>"
}

if (-not $osRows) {
  $osRows = @("<tr><td colspan='3'><em>No OS update history was returned.</em></td></tr>")
}

$changedAppRows = foreach ($a in $otherAppsChanged) {
  $changedDate = if ($a.ChangeDate) {
    $a.ChangeDate.ToString('dd/MM/yyyy')
  }
  else {
    '-'
  }

  "<tr><td>$($a.Application)</td><td>$($a.Version)</td><td>$($a.Publisher)</td><td>$changedDate</td></tr>"
}

if (-not $changedAppRows) {
  $changedAppRows = @("<tr><td colspan='4'><em>No additional applications with a registry InstallDate in $Period were found.</em></td></tr>")
}

$html = @"
<html>
<head>
<meta charset='utf-8'>
<title>$baseName</title>
$css
</head>
<body>

<h1>AVD Patching Report</h1>

<div class='summary'>
  <span class='tag'>Master: $machine</span>
  <span class='tag'>OS: $($osInfo.OSName)</span>
  <span class='tag'>Build: $($osInfo.OSBuild)</span>
</div>

<h2>Applications - Before and After</h2>
<p class='small'>The first run for this master image in $Period creates the Installed Before baseline. Later runs compare the current version with that baseline.</p>

<table>
<thead>
<tr>
<th>Application</th>
<th>Installed Before</th>
<th>Installed After</th>
<th>Latest</th>
<th>Status</th>
<th>Change</th>
</tr>
</thead>
<tbody>
$(($appRows -join "`r`n"))
</tbody>
</table>

<h2>Recent OS Updates</h2>
<table>
<thead>
<tr>
<th>KB</th>
<th>Description</th>
<th>Installed On</th>
</tr>
</thead>
<tbody>
$(($osRows -join "`r`n"))
</tbody>
</table>

<h2>Other Applications Installed or Updated in $Period</h2>
<p class='small'>Best-effort detection based on the Windows uninstall registry InstallDate. Some installers do not maintain this value, so this section may not capture every application change.</p>

<table>
<thead>
<tr>
<th>Application</th>
<th>Version</th>
<th>Publisher</th>
<th>Install/Update Date</th>
</tr>
</thead>
<tbody>
$(($changedAppRows -join "`r`n"))
</tbody>
</table>

<p class='small'>Generated: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</p>

</body>
</html>
"@

$html | Set-Content -LiteralPath $htmlPath -Encoding UTF8
}
else {
  Write-Warning "HTML report was not created because the AFTER CSV was not successfully created."
}


Write-Host ''
Write-Host 'AVD patch report completed.' -ForegroundColor Green
Write-Host "BEFORE snapshot : $baselinePath" -ForegroundColor DarkCyan

if ($AfterOnly) {
  Write-Host "Mode            : AFTER only" -ForegroundColor DarkGray
}

if ($afterCsvCreated) {
  Write-Host "AFTER snapshot  : $afterCsvPath" -ForegroundColor Cyan
  Write-Host "HTML            : $htmlPath" -ForegroundColor Cyan
}
else {
  Write-Host "AFTER snapshot  : NOT CREATED" -ForegroundColor Red
  Write-Host "HTML            : NOT CREATED" -ForegroundColor Red
}
