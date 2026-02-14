<#
.SYNOPSIS
    Windows Auto Install - One-liner Installer
.DESCRIPTION
    Script nay duoc thiet ke de chay qua irm (Invoke-RestMethod)
    Cach dung: irm https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/install.ps1 | iex
.NOTES
    Version: 3.0 - GitHub Ready
#>

# Bat dau - Khong can #Requires vi chay qua irm
$Host.UI.RawUI.WindowTitle = "Windows Auto Install - v3.0"

# Kiem tra Admin
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "`n[LOI] Script can quyen Administrator!" -ForegroundColor Red
    Write-Host "Cach su dung dung:" -ForegroundColor Yellow
    Write-Host "  1. Mo PowerShell as Administrator" -ForegroundColor Cyan
    Write-Host "  2. Chay lai lenh: irm YOUR_GITHUB_LINK | iex`n" -ForegroundColor Cyan
    Start-Sleep 5
    Exit 1
}

Set-ExecutionPolicy Bypass -Scope Process -Force
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Continue'

# Cau hinh
$LogFile = "$env:TEMP\WindowsAutoInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$MaxUpdateSizeGB = 10

# Function Write-Log
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

# Function Test Internet
function Test-Internet {
    foreach ($testHost in @('8.8.8.8', '1.1.1.1')) {
        if (Test-Connection -ComputerName $testHost -Count 2 -Quiet) { return $true }
    }
    return $false
}

Clear-Host
Write-Log "`n========================================"
Write-Log "  WINDOWS AUTO INSTALL - V3.0"
Write-Log "========================================`n"
Write-Log "Log: $LogFile" "INFO"

# Kiem tra Internet
if (-not (Test-Internet)) {
    Write-Log "`n[LOI] Khong co ket noi Internet!`n" "ERROR"
    Start-Sleep 5
    Exit 1
}

# ===================================================================
# BUOC 1: TIM VA SUA WINGET
# ===================================================================
Write-Log "`n[1/5] KIEM TRA WINGET" "WARNING"
Write-Log "==================`n"

$winget = $null
$paths = @(
    "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe",
    "C:\Program Files\WindowsApps\*\winget.exe"
)
foreach ($p in $paths) {
    $f = Get-Item $p -EA SilentlyContinue | Select -First 1
    if ($f) { $winget = $f.FullName; break }
}
if (-not $winget) {
    $cmd = Get-Command winget -EA SilentlyContinue
    if ($cmd) { $winget = $cmd.Source }
}

if (-not $winget) {
    Write-Log "Cai dat WinGet..." "WARNING"
    try {
        # VCLibs
        $url = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        $tmp = "$env:TEMP\vc.appx"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
        Add-AppxPackage -Path $tmp; Remove-Item $tmp -Force
        
        # UI.Xaml
        $url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        $tmp = "$env:TEMP\ui.appx"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
        Add-AppxPackage -Path $tmp; Remove-Item $tmp -Force
        
        # WinGet
        $url = "https://aka.ms/getwinget"
        $tmp = "$env:TEMP\wg.msixbundle"
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
        Add-AppxPackage -Path $tmp; Remove-Item $tmp -Force
        
        Start-Sleep 5
        $winget = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
        Write-Log "  [OK] WinGet cai xong" "SUCCESS"
    } catch {
        Write-Log "  [LOI] Khong cai duoc WinGet`n" "ERROR"
        Start-Sleep 5
        Exit 1
    }
}

Write-Log "  [OK] Path: $winget" "SUCCESS"

# Sua loi certificate
Write-Log "`nSua loi certificate..." "WARNING"
try {
    & $winget source reset --force 2>&1 | Out-Null
    Start-Sleep 2
    & $winget source update 2>&1 | Out-Null
    Write-Log "  [OK] Da sua" "SUCCESS"
} catch {}

# ===================================================================
# BUOC 2: GO BLOATWARE
# ===================================================================
Write-Log "`n[2/5] GO BLOATWARE" "WARNING"
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
# BUOC 3: WINDOWS UPDATE
# ===================================================================
Write-Log "`n[3/5] WINDOWS UPDATE" "WARNING"
Write-Log "==================`n"

$skipUpdate = $false

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
            foreach ($u in $updates) {
                $sizeGB = [math]::Round($u.Size/1GB, 2)
                $sizeMB = [math]::Round($u.Size/1MB, 0)
                if ($sizeGB -gt $MaxUpdateSizeGB) {
                    Write-Log "  [SKIP] $($u.KB) (${sizeGB}GB)" "WARNING"
                } else {
                    $valid += $u
                    Write-Log "  [OK] $($u.KB) (${sizeMB}MB)" "INFO"
                }
            }
            
            if ($valid.Count -gt 0) {
                Write-Log "`nCai $($valid.Count) updates..." "WARNING"
                $installed = 0
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
            }
        } else {
            Write-Log "  [OK] He thong moi nhat`n" "SUCCESS"
        }
    } catch {
        Write-Log "  [LOI] Kiem tra updates that bai`n" "ERROR"
    }
}

# ===================================================================
# BUOC 4: CAI PHAN MEM
# ===================================================================
Write-Log "`n[4/5] CAI PHAN MEM" "WARNING"
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
    "Office"="Microsoft.Office"
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
            Write-Log "  [LOI] Code: $LASTEXITCODE" "ERROR"
        }
    } catch {
        Write-Log "  [LOI]" "ERROR"
    }
}
Write-Log "`nCai xong: $success/$($apps.Count)`n" "SUCCESS"

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
Write-Log "`nLog: $LogFile" "INFO"
Write-Log "`nBam phim bat ky de thoat..." "INFO"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
