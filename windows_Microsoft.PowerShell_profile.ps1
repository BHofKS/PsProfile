New-Alias 'cips' Connect-IPPSSession
New-Alias 'caad' connect-azuread
New-Alias 'cmso' connect-msolservice
New-Alias 'ct' connect-microsoftteams
New-Alias 'caz' connect-azaccount
New-Alias 'cl' clear-host
New-Alias 'dmg' disconnect-mggraph

$PSDefaultParameterValues['Get-Help:full'] = $true

$ProfileVersion = '2026050601'  # yyyymmdd##

$MasterUrl = 'https://raw.githubusercontent.com/BHofKS/PsProfile/main/windows_Microsoft.PowerShell_profile.ps1'

# === Path resolution =====================================================
$usersPath = 'C:\Users\bh1.users\'
$adsPath = 'C:\Users\bh1.ads\'
$eidPath = 'C:\Users\bh1\'
if (Test-Path -Path $usersPath -PathType Container) {
    $homePath = $usersPath; $adminPath = $eidPath
}
else {
    $homePath = $eidPath;   $adminPath = $adsPath
}

# === Profile sync (version-based upstream + hash-based reconciliation) ===
$primaryProfile = Join-Path $adminPath 'Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1'
$allProfiles = @(
    $primaryProfile,
    (Join-Path $adminPath 'Documents\PowerShell\Microsoft.PowerShell_profile.ps1'),
    (Join-Path $homePath 'Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1'),
    (Join-Path $homePath 'Documents\PowerShell\Microsoft.PowerShell_profile.ps1')
)

$whoami = [Security.Principal.WindowsIdentity]::GetCurrent().Name

# Step 1: ADS sessions fetch upstream and update primary if a newer version exists
if ($whoami -like 'ADS\*') {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $upstream = (Invoke-WebRequest -Uri $MasterUrl -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop).Content
        if ($upstream -match '(?i)<!DOCTYPE|<html|Sign in to your account') {
            Write-Warning "Profile sync: upstream returned an HTML page (likely auth redirect) -- skipping update."
        }
        elseif ($upstream -match '(?m)^\s*\$ProfileVersion\s*=\s*[''"](\d{10})[''"]') {
            $upstreamVersion = $Matches[1]
            if ($upstreamVersion -gt $ProfileVersion) {
                # Back up the about-to-be-replaced primary before overwriting
                $backupDir = Join-Path $homePath 'Documents\PowerShellProfileBackups'
                if (-not (Test-Path $backupDir)) {
                    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
                }
                if (Test-Path $primaryProfile) {
                    $backupPath = Join-Path $backupDir "profile-$ProfileVersion.ps1"
                    Copy-Item -Path $primaryProfile -Destination $backupPath -Force
                }

                # Write upstream to primary
                $parent = Split-Path $primaryProfile -Parent
                if (-not (Test-Path $parent)) {
                    New-Item -ItemType Directory -Path $parent -Force | Out-Null 
                }
                Set-Content -Path $primaryProfile -Value $upstream -NoNewline -Encoding UTF8
                Write-Output "Profile sync: primary updated $ProfileVersion -> $upstreamVersion"

                # Prune backups, keep newest 5 by version
                $backups = Get-ChildItem -Path $backupDir -Filter 'profile-*.ps1' -File -ErrorAction SilentlyContinue |
                    Sort-Object Name -Descending
                if ($backups -and $backups.Count -gt 5) {
                    $backups | Select-Object -Skip 5 | Remove-Item -Force
                }
            }
        }
        else {
            Write-Warning "Profile sync: upstream had no recognizable version -- skipping update."
        }
    }
    catch {
        # Network down, URL unreachable, etc. -- silently use local.
    }
}

# Step 2: Reconcile remaining copies against primary (ADS) or against own Users-side WindowsPowerShell (Users)
$referenceFile = if ($whoami -like 'ADS\*') {
    $primaryProfile
}
else {
    Join-Path $homePath 'Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1'
}

if (Test-Path $referenceFile) {
    $writable = if ($whoami -like 'ADS\*') {
        $allProfiles | Where-Object { $_ -ne $referenceFile }
    }
    else {
        $allProfiles | Where-Object { $_ -like "$homePath*" -and $_ -ne $referenceFile }
    }
    $refHash = (Get-FileHash $referenceFile -Algorithm MD5).Hash
    foreach ($p in $writable) {
        $needsUpdate = $true
        if (Test-Path $p) {
            $item = Get-Item $p -Force
            if ($item.LinkType) {
                Remove-Item $p -Force
            }
            elseif ((Get-FileHash $p -Algorithm MD5).Hash -eq $refHash) {
                $needsUpdate = $false
            }
        }
        if ($needsUpdate) {
            $parent = Split-Path $p -Parent
            if (-not (Test-Path $parent)) {
                New-Item -ItemType Directory -Path $parent -Force | Out-Null 
            }
            try {
                Copy-Item -Path $referenceFile -Destination $p -Force -ErrorAction Stop
                Write-Output "Profile sync: updated $p"
            }
            catch {
                Write-Warning "Profile sync: could not update $p -- $_"
            }
        }
    }
    Remove-Variable refHash, writable, p, needsUpdate, item, parent -ErrorAction SilentlyContinue
}
Remove-Variable allProfiles, primaryProfile, referenceFile, whoami, upstream, upstreamVersion `
    -ErrorAction SilentlyContinue
# =========================================================================

Set-Location $homePath

function aduc {
    # Start Active Directory Users and Computers
    Start-Process "C:\Windows\system32\dsa.msc"
}

function ce {
    #connect to ExchangeOnline with cmdlet help
    Connect-ExchangeOnline -LoadCmdletHelp
}

function cleanwsus {
    # run cleanup operation on hco-wsus
    $server = Get-WsusServer -PortNumber 80 -Name "hco-wsus-p-app1.ads.ksu.edu"
    Invoke-WsusServerCleanup -CleanupObsoleteComputers -CleanupObsoleteUpdates -CleanupUnneededContentFiles -DeclineExpiredUpdates -DeclineSupersededUpdates -UpdateServer $server
}

function cmg {
    # Connect to Microsoft Graph without the banner
    Connect-MgGraph -NoWelcome
}

function compmgmt {
    # Open Computer Management UI as admin
    Start-Process "C:\Windows\system32\compmgmt.msc" -Verb runas
}

function dce {
    # Disconnect from Exchange Online without answering a prompt
    Disconnect-ExchangeOnline -Confirm:$false
}

function dl {
    #shortcut to the Downloads folder
    Set-Location -Path $homePath/Downloads
}

function dns {
    # Open DNS management
    Start-Process C:\Windows\system32\mmc.exe C:\Windows\system32\dnsmgmt.msc
}

function docs {
    #Shortcut to Documents folder
    Set-Location -Path "$homePath/Documents"
}

function dt {
    # Shortcut to \Users account desktop
    Set-Location -Path "$homePath/Desktop"
}

function du {
    # equivalent of du -sh *
    Get-ChildItem -Path "." -Directory | ForEach-Object { Get-ChildItem -Path $_.FullName -Recurse -File | Measure-Object -Property Length -Sum | Select-Object @{n = "Folder"; e = { $_.Name } }, @{n = "Size (GB)"; e = { "{0:N2}" -f ($_.Sum / 1GB) } } }
}

function edge {
    # Open Microsoft Edge
    Start-Process "msedge.exe"
}

function events {
    # open Event Viewer
    Start-Process "C:\Windows\system32\eventvwr.msc" -Verb runas
}

function explore {
    # open File Explorer as Administrator
    Start-Process -Verb runas explorer.exe
}

function forest {
    # Open MMC that has the Active Directory components
    Start-Process $homePath\Documents\AdminKit\ForestManagement.msc
}

function gh {
    #Shortcut to Source code files
    Set-Location -Path $homePath/Source
}

function hco {
    #shortcut to the Onedrive folder
    Set-Location -Path "$homePath/OneDriveKSU/HCO/"
}

function hm {
    #shortcut to the home folder
    Set-Location -Path "$homePath/"
}

function iis {
    # Opens IIS Manager
    Start-Process "C:\Windows\system32\inetsrv\InetMgr.exe" -Verb runas

}

function la {
    # equivalent of ls -al
    ls -Force -Attributes
}

function local {
    # Open LocalAdminTools MMC from bh1 Desktop
    $toolPath = "Documents\AdminKit\LocalAdminTools.msc"
    Start-Process $homePath$toolPath -Verb runas
}

function lsf {
    # equivalent of ls -l
    Get-ChildItem -File | Format-Table -AutoSize
}

function mmc {
    # Open MMC
    Start-Process C:\Windows\system32\mmc.exe -Verb runas
}

function od {
    #shortcut to the Onedrive folder
    Set-Location -Path "$homePath/OneDriveKSU/"
}

function pmp {
    # Open Patch My PC settings app
    Start-Process "C:\Program Files\Patch My PC\Patch My PC Publishing Service\PatchMyPC-Settings.exe" -Verb runas
}

function pow {
    #shortcut to the Powershell folder
    Set-Location -Path "$homePath/Source/ksuAdminTools/Powershell"
}

function proj {
    #shortcut to the Onedrive folder
    Set-Location -Path "$homePath/OneDriveKSU/HCO/Projects"
}

function rd {
    #Start Remote Desktop
    Start-Process -Path "C:\Windows\system32\mstsc.exe"
}

function restart {
    # Restart for updates
    Start-Process "C:\Windows\system32\cmd.exe /c "shutdown /g /d p:00:00 /c "Restart for updates" /t 0"" -Verb runas
}

function runas {
    # start powershell as administrator
    Start-Process -Verb runas Powershell.exe
}

function secpol {
    # Open Local Security Policy manager
    Start-Process "C:\Windows\system32\secpol.msc" -Verb runas
}

function sm {
    # Start Server Manager
    Start-Process "C:\Windows\system32\ServerManager.exe" -Verb runas
}

function sscm {
    # Start Sql Server Configuration Manager
    Start-Process "C:\Windows\SysWOW64\SQLServerManager16.msc" -Verb runas
}

function ssms {
    # Start SQL Server Management Studio
    Start-Process "ssms.exe" -Verb runas
}

function tasks {
    # Open Task Scheduler
    Start-Process "C:\Windows\system32\taskschd.msc" -Verb runas
}

function update {
    # Update all modules and help with timestamps pre and post
    Get-Date; Update-Module; Update-Help; Get-Date
}

function wsus {
    # Open WSUS management
    Start-Process "C:\Program Files\Update Services\AdministrationSnapin\wsus.msc" -Verb runas
}
#
# === Local server-specific profile =======================================
# Server-specific functions/aliases live in $env:LocalProfile.
# Sourced last so it can override anything defined above.
$env:LocalProfile = 'C:\ProgramData\PowerShell\LocalProfile.ps1'
if (Test-Path $env:LocalProfile) {
    . $env:LocalProfile 
}
# =========================================================================
