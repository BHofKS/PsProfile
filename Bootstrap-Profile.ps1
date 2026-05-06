# bootstrap-profile.ps1 -- Run ONCE per server (or to force-update), elevated, as ADS\bh1.
#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$MasterUrl = 'https://raw.githubusercontent.com/BHofKS/Work_dotfiles/main/windows_Microsoft.PowerShell_profile.ps1',
    [string]$MasterPath  # Optional: skip download, use this local file instead
)

$ErrorActionPreference = 'Stop'

# --- Sanity checks ---
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
$pr = New-Object Security.Principal.WindowsPrincipal($id)
if (-not $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "Run this from an ELEVATED PowerShell session."
}
if ($id.Name -notlike 'ADS\*') {
    throw "Run this as your ADS account (currently: $($id.Name))."
}

# --- Path resolution ---
$usersPath = 'C:\Users\bh1.users\'
$adsPath = 'C:\Users\bh1.ads\'
$eidPath = 'C:\Users\bh1\'
if (Test-Path $usersPath -PathType Container) {
    $homePath = $usersPath; $adminPath = $eidPath
}
else {
    $homePath = $eidPath;   $adminPath = $adsPath
}

# Primary now lives on the ADS side, since 90% of sessions start there
$primaryProfile = Join-Path $adminPath 'Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1'
$secondary = @(
    (Join-Path $adminPath 'Documents\PowerShell\Microsoft.PowerShell_profile.ps1'),
    (Join-Path $homePath 'Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1'),
    (Join-Path $homePath 'Documents\PowerShell\Microsoft.PowerShell_profile.ps1')
)
$localProfilePath = 'C:\ProgramData\PowerShell\LocalProfile.ps1'

# --- 1. ACLs on Users-side Documents and Desktop (skip if OneDrive-redirected) ---
Write-Host "[1/5] Setting ACLs on Users-side Documents and Desktop ..." -ForegroundColor Cyan
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule("ads\bh1", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
$aclTargets = @(
    (Join-Path $homePath 'Documents'),
    (Join-Path $homePath 'Desktop')
)
foreach ($target in $aclTargets) {
    if (-not (Test-Path $target)) {
        Write-Host "      Skip (missing): $target" -ForegroundColor DarkGray
        continue
    }
    $item = Get-Item $target -Force
    $isReparse = ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
    $resolved = $item.FullName
    if ($isReparse -or $resolved -match '(?i)OneDrive') {
        Write-Warning "Skip (OneDrive-redirected): $target -> $resolved"
        continue
    }
    try {
        $a = Get-Acl $target
        $a.SetAccessRule($rule)
        Set-Acl -Path $target -AclObject $a
        Write-Host "      ACL set on $target" -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "Failed to set ACL on $target : $_"
    }
}

# --- 2. Ensure profile directories exist ---
Write-Host "[2/5] Creating profile directories ..." -ForegroundColor Cyan
foreach ($p in @($primaryProfile) + $secondary) {
    $parent = Split-Path $p -Parent
    if (-not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null 
    }
}

# --- 3. Back up the existing primary if present ---

Write-Host "[3/6] Backing up existing primary (if any) ..." -ForegroundColor Cyan
if (Test-Path $primaryProfile) {
    $backupDir = Join-Path $homePath 'Documents\PowerShellProfileBackups'
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

    $existingContent = Get-Content $primaryProfile -Raw
    $backupName = if ($existingContent -match '(?m)^\s*\$ProfileVersion\s*=\s*[''"](\d{10})[''"]') {
        "profile-$($Matches[1]).ps1"
    }
    else {
        "profile-pre-bootstrap-$(Get-Date -Format 'yyyyMMdd-HHmmss').ps1"
    }

    Copy-Item -Path $primaryProfile -Destination (Join-Path $backupDir $backupName) -Force
    Write-Host "      Saved as $backupName" -ForegroundColor Green

    # Prune to newest 5 by name
    Get-ChildItem -Path $backupDir -Filter 'profile-*.ps1' -File |
        Sort-Object Name -Descending |
        Select-Object -Skip 5 |
        Remove-Item -Force
}
else {
    Write-Host "      No existing primary -- nothing to back up." -ForegroundColor Green

}

# --- 4. Obtain the primary profile (URL -> MasterPath -> existing -> notepad) ---
Write-Host "[3/5] Obtaining primary profile content ..." -ForegroundColor Cyan
$obtained = $false

if ($MasterPath) {
    if (Test-Path $MasterPath) {
        Copy-Item -Path $MasterPath -Destination $primaryProfile -Force
        Write-Host "      Copied from local $MasterPath" -ForegroundColor Green
        $obtained = $true
    }
    else {
        Write-Warning "Specified -MasterPath not found: $MasterPath"
    }
}

if (-not $obtained) {
    $tmp = [IO.Path]::GetTempFileName()
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $MasterUrl -OutFile $tmp -UseBasicParsing -ErrorAction Stop
        $head = Get-Content $tmp -TotalCount 5 -ErrorAction Stop
        $headJoined = ($head -join "`n")
        if ($headJoined -match '(?i)<!DOCTYPE|<html|Sign in to your account') {
            Write-Warning "Download returned an HTML page (likely a login redirect). Discarding."
        }
        elseif ($headJoined -notmatch '(?m)^\s*\$ProfileVersion\s*=') {
            Write-Warning "Download did not contain `$ProfileVersion -- not a valid profile. Discarding."
        }
        else {
            Move-Item -Path $tmp -Destination $primaryProfile -Force
            Write-Host "      Downloaded from $MasterUrl" -ForegroundColor Green
            $obtained = $true
        }
    }
    catch {
        Write-Warning "Download failed: $($_.Exception.Message)"
    }
    finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

if (-not $obtained) {
    if (Test-Path $primaryProfile) {
        Write-Host "      Using existing primary (URL unreachable)." -ForegroundColor Yellow
        $obtained = $true
    }
    else {
        Write-Host "      No URL, no local copy. Opening notepad ..." -ForegroundColor Yellow
        Write-Host "      Paste your profile, save, close notepad, then press Enter." -ForegroundColor Yellow
        notepad $primaryProfile
        Read-Host "Press Enter once saved"
        if (-not (Test-Path $primaryProfile)) {
            throw "Primary profile still missing -- aborting." 
        }
        $obtained = $true
    }
}

# Sanity check: the content we just placed should contain a recognizable version
$primaryContent = Get-Content $primaryProfile -Raw
if ($primaryContent -notmatch '(?m)^\s*\$ProfileVersion\s*=\s*[''"](\d{10})[''"]') {
    Write-Warning "Primary has no recognizable `$ProfileVersion line. Self-healing won't be able to detect future upstream updates."
}
else {
    Write-Host "      Primary version: $($Matches[1])" -ForegroundColor Green
}

# --- 5. Distribute primary to the three secondary locations ---
Write-Host "[4/5] Distributing to secondary locations ..." -ForegroundColor Cyan
$primaryHash = (Get-FileHash $primaryProfile -Algorithm MD5).Hash
foreach ($p in $secondary) {
    if (Test-Path $p) {
        $item = Get-Item $p -Force
        if ($item.LinkType) {
            Remove-Item $p -Force
        }
        elseif ((Get-FileHash $p -Algorithm MD5).Hash -eq $primaryHash) {
            Write-Host "      OK    : $p" -ForegroundColor Green
            continue
        }
    }
    Copy-Item -Path $primaryProfile -Destination $p -Force
    Write-Host "      COPIED: $p" -ForegroundColor Green
}

# --- 6. LocalProfile + machine env var ---
Write-Host "[5/5] Setting up LocalProfile and env var ..." -ForegroundColor Cyan
$lpDir = Split-Path $localProfilePath -Parent
if (-not (Test-Path $lpDir)) {
    New-Item -ItemType Directory -Path $lpDir -Force | Out-Null 
}
if (-not (Test-Path $localProfilePath)) {
    @"
# Server-specific PowerShell profile for $env:COMPUTERNAME
# Sourced by the main profile after defaults have loaded.
# Override or add server-specific functions, aliases, and variables here.
"@ | Set-Content -Path $localProfilePath -Encoding UTF8
    Write-Host "      Created $localProfilePath" -ForegroundColor Green
}
else {
    Write-Host "      LocalProfile already exists -- leaving alone." -ForegroundColor Green
}

$lpAcl = Get-Acl $localProfilePath
foreach ($acct in @('ads\bh1','users\bh1')) {
    try {
        $lpRule = New-Object System.Security.AccessControl.FileSystemAccessRule($acct,'Modify','Allow')
        $lpAcl.SetAccessRule($lpRule)
    }
    catch {
        Write-Warning "ACL grant failed for $acct on LocalProfile: $_" 
    }
}
Set-Acl -Path $localProfilePath -AclObject $lpAcl

[Environment]::SetEnvironmentVariable('LocalProfile', $localProfilePath, 'Machine')
Write-Host "      `$env:LocalProfile = $localProfilePath" -ForegroundColor Green

Write-Host "`nDone. Open a new PowerShell session as ADS to verify." -ForegroundColor Green
