<#
.SYNOPSIS
    Windows Auto Install - Bootstrap Script
.DESCRIPTION
    BUOC 1: Chay tren PowerShell thong thuong
    - Cai WinGet va Windows Terminal
    - Tai script chinh tu GitHub
    - Mo Windows Terminal moi
.USAGE
    irm YOUR_GITHUB_LINK/install.ps1 | iex
#>

Set-ExecutionPolicy Bypass -Scope Process -Force
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

Clear-Host
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  WINDOWS AUTO INSTALL - V3.0" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Cyan

# ===================================================================
# BUOC 1: KIEM TRA INTERNET
# ===================================================================
Write-Host "[1/3] Kiem tra Internet..." -ForegroundColor Yellow
$canConnect = $false
foreach ($testHost in @('8.8.8.8', '1.1.1.1')) {
    if (Test-Connection -ComputerName $testHost -Count 2 -Quiet) {
        $canConnect = $true
        break
    }
}

if (-not $canConnect) {
    Write-Host "      [LOI] Khong co ket noi Internet!`n" -ForegroundColor Red
    Read-Host "Bam Enter de thoat"
    Exit 1
}
Write-Host "      [OK] Ket noi tot`n" -ForegroundColor Green

# ===================================================================
# BUOC 2: TIM VA CAI WINGET
# ===================================================================
Write-Host "[2/3] Kiem tra WinGet..." -ForegroundColor Yellow

$winget = $null
$paths = @(
    "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
    "C:\Program Files\WindowsApps\*\winget.exe"
)

foreach ($p in $paths) {
    $f = Get-Item $p -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { $winget = $f.FullName; break }
}

if (-not $winget) {
    $cmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($cmd) { $winget = $cmd.Source }
}

if (-not $winget) {
    Write-Host "      Dang cai dat WinGet..." -ForegroundColor Cyan
    try {
        # VCLibs
        Write-Host "      - Tai VCLibs..." -ForegroundColor Gray
        $url = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        $tmp = "$env:TEMP\vc.appx"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing | Out-Null
        Add-AppxPackage -Path $tmp | Out-Null
        Remove-Item $tmp -Force
        
        # UI.Xaml
        Write-Host "      - Tai UI.Xaml..." -ForegroundColor Gray
        $url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        $tmp = "$env:TEMP\ui.appx"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing | Out-Null
        Add-AppxPackage -Path $tmp | Out-Null
        Remove-Item $tmp -Force
        
        # WinGet
        Write-Host "      - Tai WinGet..." -ForegroundColor Gray
        $url = "https://aka.ms/getwinget"
        $tmp = "$env:TEMP\wg.msixbundle"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing | Out-Null
        Add-AppxPackage -Path $tmp | Out-Null
        Remove-Item $tmp -Force
        
        Start-Sleep 5
        $winget = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
        Write-Host "      [OK] WinGet cai xong" -ForegroundColor Green
    } catch {
        Write-Host "      [LOI] Khong cai duoc WinGet`n" -ForegroundColor Red
        Read-Host "Bam Enter de thoat"
        Exit 1
    }
} else {
    Write-Host "      [OK] WinGet da co san" -ForegroundColor Green
}

# Sua loi certificate
Write-Host "      Sua loi certificate..." -ForegroundColor Gray
& $winget source reset --force 2>&1 | Out-Null
Start-Sleep 2
& $winget source update 2>&1 | Out-Null

# ===================================================================
# BUOC 3: CAI WINDOWS TERMINAL
# ===================================================================
Write-Host "`n[3/3] Cai Windows Terminal..." -ForegroundColor Yellow

$wtInstalled = Get-AppxPackage -Name "Microsoft.WindowsTerminal" -ErrorAction SilentlyContinue

if (-not $wtInstalled) {
    Write-Host "      Dang tai va cai..." -ForegroundColor Cyan
    & $winget install Microsoft.WindowsTerminal -e --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
    Start-Sleep 3
    Write-Host "      [OK] Da cai xong" -ForegroundColor Green
} else {
    Write-Host "      [OK] Da co san" -ForegroundColor Green
}

# ===================================================================
# TAI SCRIPT CHINH TU GITHUB
# ===================================================================
Write-Host "`nDang tai script chinh..." -ForegroundColor Yellow

# THAY DOI LINK NAY THANH LINK GITHUB CUA BAN
$mainScriptUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/refs/heads/main/main-script.ps1"
$mainScript = "$env:TEMP\windows_install_main.ps1"

try {
    Invoke-WebRequest -Uri $mainScriptUrl -OutFile $mainScript -UseBasicParsing
    Write-Host "      [OK] Tai thanh cong`n" -ForegroundColor Green
} catch {
    Write-Host "      [LOI] Khong tai duoc script chinh!" -ForegroundColor Red
    Write-Host "      Kiem tra link GitHub cua ban`n" -ForegroundColor Yellow
    Read-Host "Bam Enter de thoat"
    Exit 1
}

# ===================================================================
# MO WINDOWS TERMINAL
# ===================================================================
Write-Host "Mo Windows Terminal..." -ForegroundColor Yellow
Write-Host "Cua so nay se DONG, theo doi cua so MOI`n" -ForegroundColor Cyan

$wt = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"

if (-not (Test-Path $wt)) {
    Write-Host "[CANH BAO] Khong tim thay Terminal!" -ForegroundColor Red
    Write-Host "Chay trong PowerShell thong thuong...`n" -ForegroundColor Yellow
    Start-Sleep 3
    & powershell.exe -NoExit -ExecutionPolicy Bypass -File $mainScript
} else {
    Write-Host "[OK] Khoi dong Terminal moi..." -ForegroundColor Green
    Write-Host ">>> CUA SO NAY DONG SAU 3 GIAY <<<`n" -ForegroundColor Yellow
    
    Start-Sleep 3
    
    # Mo Terminal moi voi Admin
    Start-Process $wt -ArgumentList "PowerShell -NoExit -ExecutionPolicy Bypass -File `"$mainScript`"" -Verb RunAs
    
    # Dong cua so cu
    Start-Sleep 2
    Exit
}
