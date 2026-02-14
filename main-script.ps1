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
Write-Log "[1/4] GO BLOATWARE" "WARNING"
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
Write-Log "`n[2/4] WINDOWS UPDATE" "WARNING"
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
                # Xu ly truong hop Size = null hoac 0
                $size = if ($u.Size) { $u.Size } else { 0 }
                $sizeGB = [math]::Round($size/1GB, 2)
                $sizeMB = [math]::Round($size/1MB, 0)
                
                # Hien thi kich thuoc
                $sizeText = if ($sizeMB -gt 0) { "${sizeMB}MB" } else { "Unknown" }
                
                # Lay ten update (uu tien Title, neu khong co thi dung KB)
                $updateName = if ($u.Title) { $u.Title } else { $u.KB }
                
                # Bo qua updates qua lon
                if ($sizeGB -gt $MaxUpdateSizeGB) {
                    Write-Log "  [SKIP] $updateName" "WARNING"
                    Write-Log "         Kich thuoc: ${sizeGB}GB (qua lon)" "WARNING"
                    continue
                }
                
                # Kiem tra co can reboot khong
                if ($u.RebootRequired) {
                    $skippedReboot += $u
                    Write-Log "  [SKIP] $updateName" "WARNING"
                    Write-Log "         Kich thuoc: $sizeText (can restart)" "WARNING"
                } else {
                    $valid += $u
                    Write-Log "  [CHON] $updateName" "INFO"
                    Write-Log "         KB: $($u.KB) | Kich thuoc: $sizeText" "INFO"
                }
            }
            
            if ($skippedReboot.Count -gt 0) {
                Write-Log "`nBo qua $($skippedReboot.Count) updates can restart" "WARNING"
            }
            
            if ($valid.Count -gt 0) {
                Write-Log "`nBat dau cai $($valid.Count) updates..." "WARNING"
                $currentUpdate = 0
                foreach ($u in $valid) {
                    $currentUpdate++
                    $size = if ($u.Size) { $u.Size } else { 0 }
                    $sizeMB = [math]::Round($size/1MB, 0)
                    $sizeText = if ($sizeMB -gt 0) { "${sizeMB}MB" } else { "Unknown" }
                    $updateName = if ($u.Title) { $u.Title } else { $u.KB }
                    
                    Write-Log "`n[$currentUpdate/$($valid.Count)] $updateName" "INFO"
                    Write-Log "         KB: $($u.KB) | Kich thuoc: $sizeText" "INFO"
                    Write-Log "         Dang cai dat..." "INFO"
                    try {
                        $r = Install-WindowsUpdate -KBArticleID $u.KB -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -EA Stop
                        if ($r) {
                            Write-Log "         [OK] Cai thanh cong" "SUCCESS"
                            $installed++
                        } else {
                            Write-Log "         [LOI] Cai that bai" "ERROR"
                        }
                    } catch {
                        Write-Log "         [LOI] $($_.Exception.Message)" "ERROR"
                    }
                }
                Write-Log "`nTong ket: Cai thanh cong $installed/$($valid.Count) updates`n" "SUCCESS"
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
Write-Log "`n[3/4] CAI PHAN MEM" "WARNING"
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
Write-Log "`n[4/4] CAI OFFICE 365" "WARNING"
Write-Log "==================`n"

Write-Log "Dang cai Office 365 (Word, Excel, PowerPoint)..." "INFO"
Write-Log "Qua trinh nay mat 10-20 phut, vui long cho..." "WARNING"

$officeInstalled = $false

try {
    # Tai ODT
    Write-Log "  - Tai Office Deployment Tool..." "INFO"
    
    # Link moi nhat tu Microsoft (cap nhat 2025)
    # Neu link nay het han, tim link moi tai: https://www.microsoft.com/en-us/download/details.aspx?id=49117
    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe"
    $odtPath = "$env:TEMP\ODT.exe"
    
    try {
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing | Out-Null
    } catch {
        # Thu link du phong
        Write-Log "  - Link chinh loi, thu link du phong..." "WARNING"
        $odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
        Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing | Out-Null
    }
    
    # Giai nen ODT
    Write-Log "  - Giai nen ODT..." "INFO"
    $odtFolder = "$env:TEMP\ODT"
    if (Test-Path $odtFolder) { Remove-Item $odtFolder -Recurse -Force }
    New-Item -ItemType Directory -Path $odtFolder -Force | Out-Null
    Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$odtFolder" -Wait
    
    # Tao config XML
    Write-Log "  - Tao cau hinh..." "INFO"
    $configPath = "$odtFolder\config.xml"
    
    # Tao noi dung XML
    $xmlContent = @"
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
"@
    
    $xmlContent | Out-File -FilePath $configPath -Encoding UTF8
    
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
    Write-Log "  [INFO] Bo qua Office, tiep tuc cac buoc khac..." "WARNING"
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
