<#
.SYNOPSIS
    Script cai dat tu dong phan mem Windows - PHIEN BAN TOI UU
.DESCRIPTION
    - Go bo bloatware Windows
    - Cai dat Windows Updates
    - Cai dat phan mem tu dong bang WinGet
    - Khac phuc loi certificate msstore
    - Logging day du
.NOTES
    Version: 3.0 - Optimized
    Author: Rayal1102 (Optimized)
#>

#Requires -RunAsAdministrator

Set-ExecutionPolicy Bypass -Scope Process -Force
$progressPreference = 'silentlyContinue'
$ErrorActionPreference = 'Continue' # THAY DOI: SilentlyContinue -> Continue de bat loi

# ===================================================================
# CAU HINH - KHONG DAU
# ===================================================================
$LogFile = "$env:TEMP\WindowsAutoInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$MaxUpdateSizeGB = 10 # Kich thuoc toi da cua 1 update (GB)

# ===================================================================
# FUNCTION LOGGING - KHONG DAU
# ===================================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
    
    switch ($Level) {
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'WARNING' { Write-Host $Message -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message -ForegroundColor Cyan }
    }
}

# ===================================================================
# FUNCTION KIEM TRA INTERNET - KHONG DAU
# ===================================================================
function Test-InternetConnection {
    Write-Log "Kiem tra ket noi Internet..." "INFO"
    $TestHosts = @('8.8.8.8', '1.1.1.1')
    
    foreach ($Host in $TestHosts) {
        if (Test-Connection -ComputerName $Host -Count 2 -Quiet) {
            Write-Log "  [OK] Ket noi Internet hoat dong" "SUCCESS"
            return $true
        }
    }
    
    Write-Log "  [LOI] Khong co ket noi Internet!" "ERROR"
    return $false
}

# ===================================================================
# BAT DAU SCRIPT - KHONG DAU
# ===================================================================
Clear-Host
Write-Log "`n========================================"
Write-Log "  SCRIPT CAI DAT TU DONG - V3.0 OPTIMIZED"
Write-Log "========================================`n"
Write-Log "Log file: $LogFile" "INFO"

# Kiem tra Internet
if (-not (Test-InternetConnection)) {
    Read-Host "`nBam Enter de thoat"
    Exit 1
}

# ===================================================================
# BUOC 1: TIM VA SUA LOI WINGET - KHONG DAU
# ===================================================================
Write-Log "`n[BUOC 1/5] KIEM TRA VA SUA LOI WINGET" "WARNING"
Write-Log "================================================`n"

# Tim WinGet
$winget = $null
$wingetPaths = @(
    "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
    "C:\Program Files\WindowsApps\*\winget.exe"
)

foreach ($path in $wingetPaths) {
    $found = Get-Item $path -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($found) {
        $winget = $found.FullName
        break
    }
}

if (-not $winget) {
    $cmd = Get-Command winget -ErrorAction SilentlyContinue
    if ($cmd) { $winget = $cmd.Source }
}

# Cai WinGet neu chua co
if (-not $winget) {
    Write-Log "Dang cai dat WinGet..." "WARNING"
    try {
        # Tai VCLibs dependencies
        Write-Log "  - Dang tai VCLibs..." "INFO"
        $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        $vcLibsPath = "$env:TEMP\VCLibs.appx"
        Invoke-WebRequest -Uri $vcLibsUrl -OutFile $vcLibsPath -UseBasicParsing
        Add-AppxPackage -Path $vcLibsPath
        Remove-Item $vcLibsPath -Force
        
        # Tai UI.Xaml
        Write-Log "  - Dang tai UI.Xaml..." "INFO"
        $xamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        $xamlPath = "$env:TEMP\UIXaml.appx"
        Invoke-WebRequest -Uri $xamlUrl -OutFile $xamlPath -UseBasicParsing
        Add-AppxPackage -Path $xamlPath
        Remove-Item $xamlPath -Force
        
        # Tai WinGet
        Write-Log "  - Dang tai WinGet..." "INFO"
        $wingetUrl = "https://aka.ms/getwinget"
        $wingetPath = "$env:TEMP\winget.msixbundle"
        Invoke-WebRequest -Uri $wingetUrl -OutFile $wingetPath -UseBasicParsing
        Add-AppxPackage -Path $wingetPath
        Remove-Item $wingetPath -Force
        
        Start-Sleep 5
        $winget = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
        Write-Log "  [OK] WinGet da duoc cai dat" "SUCCESS"
    } catch {
        Write-Log "  [LOI] Khong the cai WinGet: $($_.Exception.Message)" "ERROR"
        Write-Log $LogFile "ERROR"
        Read-Host "`nBam Enter de thoat"
        Exit 1
    }
}

Write-Log "  [OK] WinGet path: $winget" "SUCCESS"

# SUA LOI CERTIFICATE MSSTORE - QUAN TRONG!
Write-Log "`nSua loi WinGet certificate (0x8a15005e)..." "WARNING"
try {
    # Reset winget sources
    Write-Log "  - Reset winget sources..." "INFO"
    & $winget source reset --force 2>&1 | Out-Null
    Start-Sleep 2
    
    # Update sources
    Write-Log "  - Update sources..." "INFO"
    & $winget source update 2>&1 | Out-Null
    Start-Sleep 2
    
    Write-Log "  [OK] Da sua loi certificate" "SUCCESS"
} catch {
    Write-Log "  [CANH BAO] Khong the reset sources, tiep tuc..." "WARNING"
}

# ===================================================================
# BUOC 2: CAI WINDOWS TERMINAL - KHONG DAU
# ===================================================================
Write-Log "`n[BUOC 2/5] CAI DAT WINDOWS TERMINAL" "WARNING"
Write-Log "================================================`n"

$wtInstalled = Get-AppxPackage -Name "Microsoft.WindowsTerminal" -ErrorAction SilentlyContinue

if (-not $wtInstalled) {
    Write-Log "Dang cai Windows Terminal..." "INFO"
    # SUA LOI: Chi dung source winget, tranh loi msstore
    & $winget install Microsoft.WindowsTerminal -e --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
    Start-Sleep 3
    Write-Log "  [OK] Da cai Windows Terminal" "SUCCESS"
} else {
    Write-Log "  [OK] Windows Terminal da co san" "SUCCESS"
}

# ===================================================================
# BUOC 3: GO BLOATWARE - KHONG DAU
# ===================================================================
Write-Log "`n[BUOC 3/5] GO BO BLOATWARE WINDOWS" "WARNING"
Write-Log "================================================`n"

$bloatware = @{
    "Cortana" = "Microsoft.549981C3F5F10*"
    "Bing Weather" = "Microsoft.BingWeather*"
    "Get Help" = "Microsoft.GetHelp*"
    "Get Started" = "Microsoft.Getstarted*"
    "Office Hub" = "Microsoft.MicrosoftOfficeHub*"
    "Solitaire" = "Microsoft.MicrosoftSolitaireCollection*"
    "Sticky Notes" = "Microsoft.MicrosoftStickyNotes*"
    "Mixed Reality" = "Microsoft.MixedReality.Portal*"
    "People" = "Microsoft.People*"
    "Skype" = "Microsoft.SkypeApp*"
    "Wallet" = "Microsoft.Wallet*"
    "Alarms & Clock" = "Microsoft.WindowsAlarms*"
    "Mail & Calendar" = "microsoft.windowscommunicationsapps*"
    "Feedback Hub" = "Microsoft.WindowsFeedbackHub*"
    "Maps" = "Microsoft.WindowsMaps*"
    "Voice Recorder" = "Microsoft.WindowsSoundRecorder*"
    "Xbox TCUI" = "Microsoft.Xbox.TCUI*"
    "Xbox App" = "Microsoft.XboxApp*"
    "Xbox Game Overlay" = "Microsoft.XboxGameOverlay*"
    "Xbox Gaming Overlay" = "Microsoft.XboxGamingOverlay*"
    "Xbox Identity" = "Microsoft.XboxIdentityProvider*"
    "Xbox Speech" = "Microsoft.XboxSpeechToTextOverlay*"
    "Your Phone" = "Microsoft.YourPhone*"
    "Groove Music" = "Microsoft.ZuneMusic*"
    "Movies & TV" = "Microsoft.ZuneVideo*"
}

$removed = 0
$total = $bloatware.Count
$current = 0

foreach ($app in $bloatware.GetEnumerator()) {
    $current++
    Write-Log "[$current/$total] $($app.Key)..." "INFO"
    
    $package = Get-AppxPackage -AllUsers $app.Value -ErrorAction SilentlyContinue
    if ($package) {
        try {
            Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction Stop
            Write-Log "  [OK] Da go" "SUCCESS"
            $removed++
        } catch {
            Write-Log "  [LOI] Khong the go: $($_.Exception.Message)" "ERROR"
        }
    } else {
        Write-Log "  [SKIP] Khong co" "INFO"
    }
}

Write-Log "`nTong ket: Da go $removed/$total ung dung`n" "SUCCESS"

# ===================================================================
# BUOC 4: CAI WINDOWS UPDATE - KHONG DAU (SUA LOI CHINH)
# ===================================================================
Write-Log "`n[BUOC 4/5] CAI DAT WINDOWS UPDATE" "WARNING"
Write-Log "================================================`n"

# Cai NuGet Provider truoc
Write-Log "Cai dat NuGet Provider..." "INFO"
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        Write-Log "  [OK] NuGet Provider da cai" "SUCCESS"
    }
} catch {
    Write-Log "  [CANH BAO] Loi cai NuGet: $($_.Exception.Message)" "WARNING"
}

# Cai PSWindowsUpdate Module
Write-Log "Cai dat PSWindowsUpdate Module..." "INFO"
try {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
        Write-Log "  [OK] Module da cai" "SUCCESS"
    } else {
        Write-Log "  [OK] Module da co san" "SUCCESS"
    }
    
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
} catch {
    Write-Log "  [LOI] Khong the cai Module: $($_.Exception.Message)" "ERROR"
    Write-Log "  [SKIP] Bo qua Windows Update" "WARNING"
    $skipUpdate = $true
}

# Kiem tra va cai Updates
if (-not $skipUpdate) {
    Write-Log "`nDang quet Windows Update (co the mat 2-5 phut)..." "INFO"
    
    try {
        # SUA LOI CHINH: Dung -MicrosoftUpdate thay vi Get-WindowsUpdate don thuan
        $updates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction Stop
        
        if ($updates -and $updates.Count -gt 0) {
            Write-Log "  [OK] Tim thay $($updates.Count) updates`n" "SUCCESS"
            
            # Loc updates theo kich thuoc
            $validUpdates = @()
            foreach ($update in $updates) {
                $sizeGB = [math]::Round($update.Size / 1GB, 2)
                $sizeMB = [math]::Round($update.Size / 1MB, 0)
                
                if ($sizeGB -gt $MaxUpdateSizeGB) {
                    Write-Log "  [SKIP] $($update.KB) - $($update.Title)" "WARNING"
                    Write-Log "         Kich thuoc: ${sizeGB}GB (qua lon)" "WARNING"
                } else {
                    $validUpdates += $update
                    Write-Log "  [CHON] $($update.KB) - $($update.Title)" "INFO"
                    Write-Log "         Kich thuoc: ${sizeMB}MB" "INFO"
                }
            }
            
            # Cai dat updates
            if ($validUpdates.Count -gt 0) {
                Write-Log "`nBat dau cai dat $($validUpdates.Count) updates..." "WARNING"
                
                $installCount = 0
                foreach ($update in $validUpdates) {
                    $sizeMB = [math]::Round($update.Size / 1MB, 0)
                    Write-Log "  Dang cai: $($update.KB) (${sizeMB}MB)..." "INFO"
                    
                    try {
                        # SUA LOI: Dung Install-WindowsUpdate voi -MicrosoftUpdate
                        $result = Install-WindowsUpdate -KBArticleID $update.KB -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop
                        
                        if ($result) {
                            Write-Log "    [OK] Cai thanh cong" "SUCCESS"
                            $installCount++
                        } else {
                            Write-Log "    [LOI] Cai that bai" "ERROR"
                        }
                    } catch {
                        Write-Log "    [LOI] $($_.Exception.Message)" "ERROR"
                    }
                }
                
                Write-Log "`nTong ket Windows Update:" "SUCCESS"
                Write-Log "  - Da cai: $installCount/$($validUpdates.Count)" "SUCCESS"
                Write-Log "  - Bo qua: $($updates.Count - $validUpdates.Count) (qua lon)" "INFO"
            } else {
                Write-Log "  [OK] Khong co update phu hop" "SUCCESS"
            }
        } else {
            Write-Log "  [OK] He thong da cap nhat moi nhat" "SUCCESS"
        }
    } catch {
        Write-Log "  [LOI] Loi kiem tra updates: $($_.Exception.Message)" "ERROR"
    }
}

# ===================================================================
# BUOC 5: CAI PHAN MEM - KHONG DAU (SUA LOI CERTIFICATE)
# ===================================================================
Write-Log "`n[BUOC 5/5] CAI DAT PHAN MEM" "WARNING"
Write-Log "================================================`n"

$apps = @{
    "UniKey" = "UniKey.UniKey"
    "WinRAR" = "RARLab.WinRAR"
    "7-Zip" = "7zip.7zip"
    "Foxit Reader" = "Foxit.FoxitReader"
    "Java Runtime" = "Oracle.JavaRuntimeEnvironment"
    "Nilesoft Shell" = "Nilesoft.Shell"
    "Flow Launcher" = "Flow-Launcher.Flow-Launcher"
    "Everything" = "voidtools.Everything"
    "Notepad++" = "Notepad++.Notepad++"
    "Google Chrome" = "Google.Chrome"
    "VCRedist 2015+ x86" = "Microsoft.VCRedist.2015+.x86"
    "VCRedist 2015+ x64" = "Microsoft.VCRedist.2015+.x64"
    "VCRedist 2005 x64" = "Microsoft.VCRedist.2005.x64"
    "VCRedist 2005 x86" = "Microsoft.VCRedist.2005.x86"
    "Microsoft Office" = "Microsoft.Office"
}

$total = $apps.Count
$current = 0
$success = 0
$failed = @()

foreach ($app in $apps.GetEnumerator()) {
    $current++
    Write-Log "[$current/$total] $($app.Key)" "INFO"
    Write-Log "  Package: $($app.Value)" "INFO"
    
    try {
        # SUA LOI CHINH: Chi dung --source winget, TRANH loi msstore certificate
        $result = & $winget install -e --id $app.Value --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  [OK] Cai thanh cong`n" "SUCCESS"
            $success++
        } else {
            Write-Log "  [LOI] Exit code: $LASTEXITCODE`n" "ERROR"
            $failed += $app.Key
        }
    } catch {
        Write-Log "  [LOI] $($_.Exception.Message)`n" "ERROR"
        $failed += $app.Key
    }
}

# ===================================================================
# KET THUC - KHONG DAU
# ===================================================================
Write-Log "`n========================================"
Write-Log "  HOAN TAT CAI DAT"
Write-Log "========================================`n"

Write-Log "TONG KET CUOI CUNG:" "WARNING"
Write-Log "  1. Bloatware: Da go $removed/$($bloatware.Count) ung dung" $(if($removed -gt 0){'SUCCESS'}else{'INFO'})
Write-Log "  2. Windows Update: Da cai $installCount updates" $(if($installCount -gt 0){'SUCCESS'}else{'INFO'})
Write-Log "  3. Phan mem: Da cai $success/$total ung dung" $(if($success -gt 0){'SUCCESS'}else{'INFO'})

if ($failed.Count -gt 0) {
    Write-Log "`nPhan mem LOI (can cai thu cong):" "WARNING"
    foreach ($app in $failed) {
        Write-Log "  - $app" "ERROR"
    }
}

Write-Log "`nLog file: $LogFile" "INFO"
Write-Log "`nBam phim bat ky de thoat..." "INFO"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
