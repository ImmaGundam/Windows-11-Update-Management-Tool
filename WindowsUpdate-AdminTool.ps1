<# 
================================================================================
 Windows 10 & 11 Update Admin Tool
 Disables or restores Windows Update while keeping Defender updates.

 Author: https://github.com/ImmaGundam

 WHAT THIS DOES:
 - Disable all Windows OS update mechanisms
 - Clear downloaded update caches and temp files
 - Keep Microsoft Defender signature updates working
 - Restore to default

 REQUIREMENTS:
 - Windows 10 / 11
 - Run as Administrator
================================================================================
#>

# -----------------------------
# Admin Check
# -----------------------------
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "ERROR: Run this script as Administrator." -ForegroundColor Red
    Pause
    exit
}

# -----------------------------
# Helper Functions
# -----------------------------

function Disable-Service {
    param ([string]$ServiceName)

    # Disabling service
    Write-Host "Disabling service: $ServiceName"
    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    Set-Service -Name $ServiceName -StartupType Disabled -ErrorAction SilentlyContinue
}

function Enable-Service {
    param ([string]$ServiceName)

    # Enabling service
    Write-Host "Enabling service: $ServiceName"
    Set-Service -Name $ServiceName -StartupType Manual -ErrorAction SilentlyContinue
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
}

# -----------------------------
# Disable OS Update Functions
# -----------------------------

function Disable-WindowsUpdateServices {
    Write-Host "`nDisabling Windows Update services..." -ForegroundColor Cyan

    $services = @(
        "wuauserv",
        "bits",
        "dosvc",
        "UsoSvc",
        "WaaSMedicSvc"
    )

    foreach ($svc in $services) {
        Disable-Service $svc
    }
}

function Disable-WindowsUpdatePolicy {
    Write-Host "`nDisabling Windows Update via registry policy..." -ForegroundColor Cyan

    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    New-Item -Path $path -Force | Out-Null

    Set-ItemProperty -Path $path -Name NoAutoUpdate -Type DWord -Value 1
}

function Disable-UpdateScheduledTasks {
    Write-Host "`nDisabling Update Orchestrator scheduled tasks..." -ForegroundColor Cyan

    Get-ScheduledTask | Where-Object {
        $_.TaskPath -like "\Microsoft\Windows\UpdateOrchestrator*"
    } | ForEach-Object {
        Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
    }
}

function Set-MeteredEthernet {
    Write-Host "`nSetting Ethernet as metered..." -ForegroundColor Cyan

    try {
        Set-ItemProperty `
            -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\DefaultMediaCost" `
            -Name Ethernet `
            -Value 2 `
            -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to set metered connection (permissions)." -ForegroundColor Yellow
    }
}

# -----------------------------
# Cache / Temp Cleanup
# -----------------------------

function Clear-WindowsUpdateCache {
    Write-Host "`nClearing Windows Update cache and temp files..." -ForegroundColor Cyan

    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Stop-Service bits -Force -ErrorAction SilentlyContinue

    $paths = @(
        "C:\Windows\SoftwareDistribution\Download",
        "C:\Windows\SoftwareDistribution\DataStore",
        "C:\Windows\Temp",
        "$env:TEMP",
        "C:\ProgramData\Microsoft\Windows\DeliveryOptimization\Cache"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            Write-Host "Clearing: $path"
            Remove-Item "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Cache and temp cleanup complete." -ForegroundColor Green
}

# -----------------------------
# Defender-Only Updates
# -----------------------------

function Enable-DefenderOnlyUpdates {
    Write-Host "`nEnabling Defender updates only..." -ForegroundColor Cyan

    # Defender services must remain enabled
    $defenderServices = @("WinDefend", "WdNisSvc", "Sense")

    foreach ($svc in $defenderServices) {
        Write-Host "Enabling Defender service: $svc"
        Set-Service -Name $svc -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service -Name $svc -ErrorAction SilentlyContinue
    }

    # Force Defender to update directly from Microsoft
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates"
    New-Item -Path $path -Force | Out-Null

    Set-ItemProperty `
        -Path $path `
        -Name FallbackOrder `
        -Type String `
        -Value "MMPC"

    Write-Host "Defender updates will continue independently." -ForegroundColor Green
}

# -----------------------------
# Restore / Undo Functions
# -----------------------------

function Restore-WindowsUpdate {
    Write-Host "`nRestoring Windows Update defaults..." -ForegroundColor Cyan

    $services = @("wuauserv","bits","dosvc","UsoSvc","WaaSMedicSvc")
    foreach ($svc in $services) {
        Enable-Service $svc
    }

    Remove-ItemProperty `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" `
        -Name NoAutoUpdate `
        -ErrorAction SilentlyContinue

    Get-ScheduledTask | Where-Object {
        $_.TaskPath -like "\Microsoft\Windows\UpdateOrchestrator*"
    } | ForEach-Object {
        Enable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
    }

    try {
        Set-ItemProperty `
            -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\DefaultMediaCost" `
            -Name Ethernet `
            -Value 1 `
            -ErrorAction SilentlyContinue
    } catch {}

    Remove-ItemProperty `
        -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates" `
        -Name FallbackOrder `
        -ErrorAction SilentlyContinue

    Write-Host "System update behavior restored." -ForegroundColor Green
}

# -----------------------------
# Status
# -----------------------------

function Show-ServiceStatus {
    Write-Host "`nUpdate-related service status:" -ForegroundColor Cyan
    Get-Service wuauserv, bits, dosvc, UsoSvc, WinDefend |
        Select Name, Status, StartType |
        Format-Table -AutoSize
}

# -----------------------------
# Menu Loop
# -----------------------------

do {
    Clear-Host
    Write-Host "===================================================="
    Write-Host "  Windows 10 & 11 Update Admin Tool"
    Write-Host "----------------------------------------------------"
    Write-Host "  https://github.com/ImmaGundam"
    Write-Host "===================================================="
    Write-Host ""
    Write-Host "1. Disable Windows Update services"
    Write-Host "2. Disable Windows Update registry policy"
    Write-Host "3. Disable Update scheduled tasks"
    Write-Host "4. Set Ethernet as metered (Optional)"
    Write-Host ""
    Write-Host "5. Clear update cache & temp files"
    Write-Host "6. Enable Defender updates only"
    Write-Host ""
    Write-Host "7. Run all steps (recommended)"
    Write-Host "8. Restore all settings"
    Write-Host "9. Show service status"
    Write-Host "0. Exit"
    Write-Host ""

    $choice = Read-Host "Select an option"

    switch ($choice) {
        "1" { Disable-WindowsUpdateServices; Pause }
        "2" { Disable-WindowsUpdatePolicy; Pause }
        "3" { Disable-UpdateScheduledTasks; Pause }
        "4" { Set-MeteredEthernet; Pause }
        "5" { Clear-WindowsUpdateCache; Pause }
        "6" { Enable-DefenderOnlyUpdates; Pause }
        "7" {
            Disable-WindowsUpdateServices
            Disable-WindowsUpdatePolicy
            Disable-UpdateScheduledTasks
            Set-MeteredEthernet
            Clear-WindowsUpdateCache
            Enable-DefenderOnlyUpdates
            Show-ServiceStatus
            Pause
        }
        "8" { Restore-WindowsUpdate; Show-ServiceStatus; Pause }
        "9" { Show-ServiceStatus; Pause }
        "0" { break }
        default { Write-Host "Invalid selection." -ForegroundColor Red; Pause }
    }
}
while ($true)

Write-Host "Exiting Windows 10 & 11 Update Tool." -ForegroundColor Cyan
