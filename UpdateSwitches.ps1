#Useful Modules
Install-Module -name PSWindowsUpdate -Force
Import-Module PSWindowsUpdate
Update-Evergreen

#APP INVENTORY - SHOWCASES CURRENT APPLICATION INVENTORY FOR GIVEN MONTH
.\CheckApps-AVD-AppInventory.ps1


#PATCH REPORTING BEFORE AND AFTER

#STEP 1 BEFORE PACTHING STATUS
..\AVDPatchReports\CheckApps-AVD-Monthly-Patching.ps1 -BeforeOnly

#STEP 4 AFTER PATCHING STATUS AND REPORT CREATION
..\AVDPatchReports\CheckApps-AVD-Monthly-Patching.ps1 -AfterOnly


#APPLICATION AND OS UPDATES

#STEP 2 #Report Only - Shows the status of defined installed applications 
.\CheckAppsAndInstallLatest2.ps1 -NoReports


#UPDATE SWITCHES

# Optional – TEST RUN Apps and Windows updates
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -IncludeWindowsUpdate -WhatIf

# STEP 3 – Full run: Windows OS patches + app upgrades
.\CheckAppsAndInstallLatest2.ps1 -Upgrade -IncludeWindowsUpdate -NoReports






