<#
.SYNOPSIS
    Windows Auto Install - Bootstrap Script
.DESCRIPTION
    BUOC 1: Chay tren PowerShell thong thuong
    - Cai WinGet va Windows Terminal
    - Tao script chinh
    - Mo Windows Terminal moi
    
    BUOC 2: Chay trong Terminal moi
    - Go bloatware
    - Cai Windows Updates
    - Cai phan mem
.USAGE
    irm YOUR_GITHUB_LINK | iex
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
# TAO SCRIPT CHINH (CHAY TRONG TERMINAL)
# ===================================================================
Write-Host "`nDang chuan bi script chinh..." -ForegroundColor Yellow

$mainScript = "$env:TEMP\windows_install_main.ps1"

$scriptContent = @'
#Requires -RunAsAdministrator
Set-ExecutionPolicy Bypass -Scope Process -Force
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Continue'

$Host.UI.RawUI.WindowTitle = "Windows Auto Install - Main"
$LogFile = "$env:TEMP\WindowsAutoInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$MaxUpdateSizeGB = 10

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$Timestamp] [$Level] $Message" -ErrorAction SilentlyContinue
    switch ($Level) {
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'WARNING' { Write-Host $Message -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        default   { Write-Host $Message -ForegroundColor Cyan }
    }
}

Clear-Host
Write-Log "`n========================================"
Write-Log "  WINDOWS AUTO INSTALL - TERMINAL"
Write-Log "========================================`n"
Write-Log "Log: $LogFile" "INFO"

# Tim WinGet
$winget = $null
$paths = @("$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe", "C:\Program Files\WindowsApps\*\winget.exe")
foreach ($p in $paths) {
    $f = Get-Item $p -EA SilentlyContinue | Select -First 1
    if ($f) { $winget = $f.FullName; break }
}
if (-not $winget) {
    $cmd = Get-Command winget -EA SilentlyContinue
    if ($cmd) { $winget = $cmd.Source }
}

Write-Log "WinGet: $winget`n" "INFO"

# ===================================================================
# BUOC 1: GO BLOATWARE
# ===================================================================
Write-Log "[1/3] GO BLOATWARE" "WARNING"
Write-Log "==================`n"

$bloat = @{
    "Cortana"="Microsoft.549981C3F5F10*";"Bing Weather"="Microsoft.BingWeather*"
    "Get Help"="Microsoft.GetHelp*";"Get Started"="Microsoft.Getstarted*"
    "Office Hub"="Microsoft.MicrosoftOfficeHub*";"Solitaire"="Microsoft.MicrosoftSolitaireCollection*"
    "Sticky Notes"="Microsoft.MicrosoftStickyNotes*";"Mixed Reality"="Microsoft.MixedReality.Portal*"
    "People"="Microsoft.People*";"Skype"="Microsoft.SkypeApp*"
    "Wallet"="Microsoft.Wallet*";"Alarms"="Microsoft.WindowsAlarms*"
    "Mail"="microsoft.windowscommunicationsapps*";"Feedback"="Microsoft.WindowsFeedbackHub*"
    "Maps"="Microsoft.WindowsMaps*";"Voice Recorder"="Microsoft.WindowsSoundRecorder*"
    "Xbox TCUI"="Microsoft.Xbox.TCUI*";"Xbox App"="Microsoft.XboxApp*"
    "Xbox Overlay"="Microsoft.XboxGameOverlay*";"Xbox Gaming"="Microsoft.XboxGamingOverlay*"
    "Xbox Identity"="Microsoft.XboxIdentityProvider*";"Xbox Speech"="Microsoft.XboxSpeechToTextOverlay*"
    "Your Phone"="Microsoft.YourPhone*";"Groove"="Microsoft.ZuneMusic*"
    "Movies"="Microsoft.ZuneVideo*"
}

$removed = 0
$i = 0
foreach ($app in $bloat.GetEnumerator()) {
    $i++
    $pkg = Get-AppxPackage -AllUsers $app.Value -EA SilentlyContinue
    if ($pkg) {
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -EA Stop
            Write-Log "[$i/$($bloat.Count)] $($app.Key) - [OK]" "SUCCESS"
            $removed++
        } catch {
            Write-Log "[$i/$($bloat.Count)] $($app.Key) - [LOI]" "ERROR"
        }
    } else {
        Write-Log "[$i/$($bloat.Count)] $($app.Key) - [SKIP]" "INFO"
    }
}
Write-Log "`nDa go: $removed/$($bloat.Count)`n" "SUCCESS"

# ===================================================================
# BUOC 2: WINDOWS UPDATE
# ===================================================================
Write-Log "`n[2/3] WINDOWS UPDATE" "WARNING"
Write-Log "==================`n"

$skipUpdate = $false
$needsReboot = $false
$installed = 0

# Cai NuGet
try {
    if (-not (Get-PackageProvider -Name NuGet -EA SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }
} catch {}

# Cai PSWindowsUpdate
try {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -EA SilentlyContinue
        Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -AllowClobber -EA Stop
        Write-Log "  [OK] Module cai xong" "SUCCESS"
    }
    Import-Module PSWindowsUpdate -Force -EA Stop
} catch {
    Write-Log "  [LOI] Khong cai duoc module" "ERROR"
    Write-Log "  [SKIP] Bo qua Update`n" "WARNING"
    $skipUpdate = $true
}

if (-not $skipUpdate) {
    Write-Log "Quet updates..." "INFO"
    try {
        $updates = Get-WindowsUpdate -MicrosoftUpdate -EA Stop
        
        if ($updates -and $updates.Count -gt 0) {
            Write-Log "  Tim thay: $($updates.Count) updates`n" "SUCCESS"
            
            $valid = @()
            $skippedReboot = @()
            
            foreach ($u in $updates) {
                $sizeGB = [math]::Round($u.Size/1GB, 2)
                $sizeMB = [math]::Round($u.Size/1MB, 0)
                
                if ($sizeGB -gt $MaxUpdateSizeGB) {
                    Write-Log "  [SKIP] $($u.KB) (${sizeGB}GB - qua lon)" "WARNING"
                    continue
                }
                
                if ($u.RebootRequired) {
                    $skippedReboot += $u
                    Write-Log "  [SKIP] $($u.KB) (${sizeMB}MB - can restart)" "WARNING"
                } else {
                    $valid += $u
                    Write-Log "  [CHON] $($u.KB) (${sizeMB}MB)" "INFO"
                }
            }
            
            if ($skippedReboot.Count -gt 0) {
                Write-Log "`nBo qua $($skippedReboot.Count) updates can restart" "WARNING"
            }
            
            if ($valid.Count -gt 0) {
                Write-Log "`nCai $($valid.Count) updates..." "WARNING"
                foreach ($u in $valid) {
                    try {
                        $r = Install-WindowsUpdate -KBArticleID $u.KB -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -EA Stop
                        if ($r) {
                            Write-Log "  [OK] $($u.KB)" "SUCCESS"
                            $installed++
                        }
                    } catch {}
                }
                Write-Log "`nCai xong: $installed/$($valid.Count)`n" "SUCCESS"
            } else {
                Write-Log "`nKhong co update nao phu hop`n" "INFO"
            }
            
            if ($skippedReboot.Count -gt 0) {
                $needsReboot = $true
            }
            
        } else {
            Write-Log "  [OK] He thong moi nhat`n" "SUCCESS"
        }
    } catch {
        Write-Log "  [LOI] Kiem tra updates that bai`n" "ERROR"
    }
}

# ===================================================================
# BUOC 3: CAI PHAN MEM
# ===================================================================
Write-Log "`n[3/3] CAI PHAN MEM" "WARNING"
Write-Log "==================`n"

$apps = @{
    "UniKey"="UniKey.UniKey"
    "WinRAR"="RARLab.WinRAR"
    "7-Zip"="7zip.7zip"
    "Foxit Reader"="Foxit.FoxitReader"
    "Java"="Oracle.JavaRuntimeEnvironment"
    "Nilesoft Shell"="Nilesoft.Shell"
    "Flow Launcher"="Flow-Launcher.Flow-Launcher"
    "Everything"="voidtools.Everything"
    "Notepad++"="Notepad++.Notepad++"
    "Chrome"="Google.Chrome"
    "VC++ 2015+ x86"="Microsoft.VCRedist.2015+.x86"
    "VC++ 2015+ x64"="Microsoft.VCRedist.2015+.x64"
    "VC++ 2005 x64"="Microsoft.VCRedist.2005.x64"
    "VC++ 2005 x86"="Microsoft.VCRedist.2005.x86"
}

$success = 0
$i = 0
foreach ($app in $apps.GetEnumerator()) {
    $i++
    Write-Log "[$i/$($apps.Count)] $($app.Key)..." "INFO"
    try {
        & $winget install -e --id $app.Value --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  [OK]" "SUCCESS"
            $success++
        } else {
            Write-Log "  [LOI]" "ERROR"
        }
    } catch {
        Write-Log "  [LOI]" "ERROR"
    }
}
Write-Log "`nCai xong: $success/$($apps.Count)`n" "SUCCESS"

# ===================================================================
# BUOC 4: CAI OFFICE 365 (CHI WORD, EXCEL, POWERPOINT)
# ===================================================================
Write-Log "`n[BONUS] CAI OFFICE 365" "WARNING"
Write-Log "==================`n"

Write-Log "Dang cai Office 365 (Word, Excel, PowerPoint)..." "INFO"
Write-Log "Qua trinh nay mat 10-20 phut, vui long cho..." "WARNING"

$officeInstalled = $false

try {
    # Tai ODT
    Write-Log "  - Tai Office Deployment Tool..." "INFO"
    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20196.exe"
    $odtPath = "$env:TEMP\ODT.exe"
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing | Out-Null
    
    # Giai nen ODT
    Write-Log "  - Giai nen ODT..." "INFO"
    $odtFolder = "$env:TEMP\ODT"
    if (Test-Path $odtFolder) { Remove-Item $odtFolder -Recurse -Force }
    New-Item -ItemType Directory -Path $odtFolder -Force | Out-Null
    Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$odtFolder" -Wait
    
    # Tao config XML (chi cai Word, Excel, PowerPoint)
    Write-Log "  - Tao cau hinh..." "INFO"
    $configPath = "$odtFolder\config.xml"
    $config = @'
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <Language ID="vi-vn" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Lync" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
</Configuration>
'@
    $config | Out-File -FilePath $configPath -Encoding UTF8
    
    # Cai dat Office
    Write-Log "  - Bat dau cai dat (10-20 phut)..." "WARNING"
    Set-Location $odtFolder
    $process = Start-Process -FilePath ".\setup.exe" -ArgumentList "/configure `"$configPath`"" -PassThru -Wait
    
    if ($process.ExitCode -eq 0) {
        Write-Log "  [OK] Office cai thanh cong" "SUCCESS"
        $officeInstalled = $true
    } else {
        Write-Log "  [LOI] Cai Office that bai (Code: $($process.ExitCode))" "ERROR"
    }
    
    # Don dep
    Set-Location $env:TEMP
    Remove-Item $odtFolder -Recurse -Force -EA SilentlyContinue
    Remove-Item $odtPath -Force -EA SilentlyContinue
    
} catch {
    Write-Log "  [LOI] Loi cai Office: $($_.Exception.Message)" "ERROR"
}

if ($officeInstalled) {
    Write-Log "`nLuu y: Office can dang nhap Microsoft Account de kich hoat" "WARNING"
    Write-Log "Mo Word hoac Excel de dang nhap`n" "INFO"
}

# ===================================================================
# KET THUC
# ===================================================================
Write-Log "`n========================================"
Write-Log "  HOAN TAT"
Write-Log "========================================`n"
Write-Log "TONG KET:" "WARNING"
Write-Log "  - Bloatware: $removed/$($bloat.Count)" $(if($removed -gt 0){'SUCCESS'}else{'INFO'})
Write-Log "  - Updates: $installed" $(if($installed -gt 0){'SUCCESS'}else{'INFO'})
Write-Log "  - Phan mem: $success/$($apps.Count)" $(if($success -gt 0){'SUCCESS'}else{'INFO'})
Write-Log "  - Office: $(if($officeInstalled){'Da cai'}else{'Khong cai'})" $(if($officeInstalled){'SUCCESS'}else{'INFO'})
Write-Log "`nLog: $LogFile" "INFO"

if ($needsReboot) {
    Write-Log "`n========================================" "WARNING"
    Write-Log "  HE THONG CAN RESTART" "WARNING"
    Write-Log "========================================`n" "WARNING"
    Write-Log "Ly do: Co updates can restart de hoan tat" "INFO"
    Write-Log "`nBan co muon restart ngay? (Y/N)" "WARNING"
    
    $response = Read-Host
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Log "`nRestart trong 10 giay (Ctrl+C de huy)..." "WARNING"
        Start-Sleep 3
        shutdown /r /t 10 /c "Restart de hoan tat Windows Updates"
    } else {
        Write-Log "`nKhong restart. Hay restart sau!`n" "WARNING"
    }
}

Write-Log "`nBam phim bat ky de dong..." "INFO"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
'@

$scriptContent | Out-File -FilePath $mainScript -Encoding UTF8 -Force
Write-Host "      [OK] Script da san sang`n" -ForegroundColor Green

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
