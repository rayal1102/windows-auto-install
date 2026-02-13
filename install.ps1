Set-ExecutionPolicy Bypass -Scope Process -Force
$progressPreference = 'silentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

Clear-Host
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  SCRIPT CAI DAT TU DONG PHAN MEM" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Cyan

# ===================================================================
# BUOC 1: TIM WINGET
# ===================================================================
Write-Host "[1/2] Kiem tra WinGet..." -ForegroundColor Yellow

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

if (-not $winget) {
    Write-Host "      Dang cai dat WinGet..." -ForegroundColor Cyan
    try {
        $url = "https://aka.ms/getwinget"
        $temp = "$env:TEMP\winget.msixbundle"
        Invoke-WebRequest -Uri $url -OutFile $temp -UseBasicParsing | Out-Null
        Add-AppxPackage -Path $temp | Out-Null
        Remove-Item $temp -Force
        Start-Sleep 3
        $winget = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    } catch {
        Write-Host "      [LOI] Khong the cai WinGet!" -ForegroundColor Red
        Read-Host "`nBam Enter de thoat"
        Exit
    }
}

Write-Host "      [OK] Da san sang" -ForegroundColor Green

# ===================================================================
# BUOC 2: CAI WINDOWS TERMINAL
# ===================================================================
Write-Host "`n[2/2] Cai dat Windows Terminal..." -ForegroundColor Yellow
& $winget install Microsoft.WindowsTerminal -e --silent --accept-source-agreements --accept-package-agreements *>$null
Write-Host "      [OK] Hoan thanh" -ForegroundColor Green
Write-Host "      Doi 3 giay de Terminal khoi dong..." -ForegroundColor Cyan
Start-Sleep 3

# ===================================================================
# TAO SCRIPT CHINH
# ===================================================================
Write-Host "`nDang chuan bi script..." -ForegroundColor Yellow

$mainScript = "$env:TEMP\main_install.ps1"

@'
Set-ExecutionPolicy Bypass -Scope Process -Force
$ErrorActionPreference = 'SilentlyContinue'
$progressPreference = 'silentlyContinue'

Clear-Host
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host "  WINDOWS TERMINAL - CAI DAT TU DONG" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Cyan

# ===================================================================
# PHAN 1: GO BO BLOATWARE
# ===================================================================
Write-Host "BUOC 1: GO BO UNG DUNG WINDOWS KHONG CAN THIET" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Yellow

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
    Write-Host "[$current/$total] $($app.Key)..." -NoNewline -ForegroundColor Cyan
    
    $package = Get-AppxPackage -AllUsers $app.Value -ErrorAction SilentlyContinue
    if ($package) {
        Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction SilentlyContinue
        Write-Host " [DA GO]" -ForegroundColor Green
        $removed++
    } else {
        Write-Host " [KHONG CO]" -ForegroundColor Gray
    }
}

Write-Host "`nTong ket: Da go $removed/$total ung dung`n" -ForegroundColor Green
Start-Sleep 2

# ===================================================================
# PHAN 2: TIM WINGET
# ===================================================================
Write-Host "`nBUOC 2: KIEM TRA WINGET" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Yellow

$winget = Get-Command winget -ErrorAction SilentlyContinue
if (-not $winget) {
    $winget = Get-Item "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe" -ErrorAction SilentlyContinue
}
if (-not $winget) {
    $winget = Get-Item "C:\Program Files\WindowsApps\*\winget.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
}
$wingetPath = if ($winget.Source) { $winget.Source } else { $winget.FullName }

Write-Host "[OK] Tim thay WinGet: $wingetPath`n" -ForegroundColor Green
Start-Sleep 1

# ===================================================================
# PHAN 3: CAI MODULE WINDOWS UPDATE
# ===================================================================
Write-Host "`nBUOC 3: CAI DAT MODULE WINDOWS UPDATE" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Yellow

if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "Dang tai module PSWindowsUpdate..." -ForegroundColor Cyan
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted | Out-Null
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false | Out-Null
    Write-Host "[OK] Da cai dat module`n" -ForegroundColor Green
} else {
    Write-Host "[OK] Module da co san`n" -ForegroundColor Green
}

Import-Module PSWindowsUpdate -ErrorAction SilentlyContinue
Start-Sleep 1

# ===================================================================
# PHAN 4: KIEM TRA VA CAI WINDOWS UPDATE
# ===================================================================
Write-Host "`nBUOC 4: KIEM TRA WINDOWS UPDATE" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Yellow

Write-Host "Dang quet cap nhat (co the mat 1-2 phut)...`n" -ForegroundColor Cyan

$updates = Get-WindowsUpdate -ErrorAction SilentlyContinue

if ($updates) {
    $totalUpdates = $updates.Count
    Write-Host "Tim thay $totalUpdates ban cap nhat`n" -ForegroundColor Cyan
    
    $validUpdates = @()
    $maxSizeGB = 10
    
    Write-Host "Dang phan tich cac ban cap nhat...`n" -ForegroundColor Cyan
    foreach ($update in $updates) {
        $sizeGB = [math]::Round($update.Size / 1GB, 2)
        $sizeMB = [math]::Round($update.Size / 1MB, 0)
        
        if ($sizeGB -gt $maxSizeGB) {
            Write-Host "  [BO QUA] $($update.Title)" -ForegroundColor Red
            Write-Host "           Kich thuoc: ${sizeGB}GB (qua lon)`n" -ForegroundColor Gray
        } else {
            $validUpdates += $update
            Write-Host "  [CHON] $($update.Title)" -ForegroundColor Green
            Write-Host "         Kich thuoc: ${sizeMB}MB`n" -ForegroundColor Gray
        }
    }
    
    if ($validUpdates.Count -gt 0) {
        Write-Host "`n--- BAT DAU TAI VA CAI DAT ---`n" -ForegroundColor Yellow
        
        $installCount = 0
        $totalValid = $validUpdates.Count
        $currentUpdate = 0
        
        foreach ($update in $validUpdates) {
            $currentUpdate++
            $sizeMB = [math]::Round($update.Size / 1MB, 0)
            
            Write-Host "[$currentUpdate/$totalValid] $($update.Title)" -ForegroundColor Cyan
            Write-Host "             Kich thuoc: ${sizeMB}MB" -ForegroundColor Gray
            Write-Host "             Dang tai va cai dat..." -NoNewline -ForegroundColor Yellow
            
            try {
                Install-WindowsUpdate -KBArticleID $update.KB -AcceptAll -IgnoreReboot -Confirm:$false *>$null
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Host " [OK]`n" -ForegroundColor Green
                    $installCount++
                } else {
                    Write-Host " [LOI]`n" -ForegroundColor Red
                }
            } catch {
                Write-Host " [LOI]`n" -ForegroundColor Red
            }
        }
        
        Write-Host "`nTong ket Windows Update:" -ForegroundColor Yellow
        Write-Host "  - Da cai thanh cong: $installCount/$totalValid" -ForegroundColor Green
        Write-Host "  - Bo qua (qua lon): $($totalUpdates - $totalValid)`n" -ForegroundColor Gray
    } else {
        Write-Host "`n[OK] Khong co ban cap nhat nao phu hop`n" -ForegroundColor Green
    }
} else {
    Write-Host "[OK] He thong da cap nhat moi nhat`n" -ForegroundColor Green
}

Start-Sleep 2

# ===================================================================
# PHAN 5: CAI DAT PHAN MEM
# ===================================================================
Write-Host "`nBUOC 5: CAI DAT PHAN MEM" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Yellow

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

foreach ($app in $apps.GetEnumerator()) {
    $current++
    Write-Host "[$current/$total] $($app.Key)" -ForegroundColor Cyan
    Write-Host "          Package: $($app.Value)" -ForegroundColor Gray
    Write-Host "          Dang tai va cai dat..." -NoNewline -ForegroundColor Yellow
    
    & $wingetPath install -e --id $app.Value --silent --accept-source-agreements --accept-package-agreements *>$null
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host " [OK]`n" -ForegroundColor Green
        $success++
    } else {
        Write-Host " [LOI]`n" -ForegroundColor Red
    }
}

# ===================================================================
# KET THUC
# ===================================================================
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host "  HOAN TAT CAI DAT" -ForegroundColor Yellow
Write-Host "================================================`n" -ForegroundColor Cyan

Write-Host "TONG KET CUOI CUNG:" -ForegroundColor Yellow
Write-Host "  1. Bloatware: Da go $removed/$($bloatware.Count) ung dung" -ForegroundColor $(if($removed -gt 0){'Green'}else{'Gray'})
Write-Host "  2. Windows Update: Da cai $installCount ban cap nhat" -ForegroundColor $(if($installCount -gt 0){'Green'}else{'Gray'})
Write-Host "  3. Phan mem: Da cai $success/$total ung dung" -ForegroundColor $(if($success -gt 0){'Green'}else{'Gray'})
Write-Host ""

Write-Host "Bam phim bat ky de dong cua so..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
'@ | Out-File -FilePath $mainScript -Encoding UTF8 -Force

Write-Host "[OK] Da tao script`n" -ForegroundColor Green

# ===================================================================
# MO WINDOWS TERMINAL - SUA LOI
# ===================================================================
Write-Host "Dang mo Windows Terminal..." -ForegroundColor Yellow
Write-Host "Cua so hien tai se DONG, vui long theo doi cua so MOI`n" -ForegroundColor Cyan

$wt = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"

# Kiem tra Windows Terminal
if (-not (Test-Path $wt)) {
    Write-Host "[LOI] Khong tim thay Windows Terminal!" -ForegroundColor Red
    Write-Host "Cai dat lai Terminal..." -ForegroundColor Yellow
    & $winget install Microsoft.WindowsTerminal -e --accept-source-agreements --accept-package-agreements
    Start-Sleep 5
}

# Mo Terminal va DONG cua so cu
if (Test-Path $wt) {
    Write-Host "[OK] Dang mo Terminal moi..." -ForegroundColor Green
    Write-Host ">>> CUA SO NAY SE TU DONG DONG SAU 3 GIAY <<<`n" -ForegroundColor Yellow
    
    Start-Sleep 3
    
    # Mo Terminal voi quyen Admin
    Start-Process $wt -ArgumentList "new-tab PowerShell -NoExit -Command `"& '$mainScript'`"" -Verb RunAs
    
    # Doi 2 giay roi dong cua so cu
    Start-Sleep 2
    Exit
} else {
    Write-Host "[CANH BAO] Van khong tim thay Terminal!" -ForegroundColor Red
    Write-Host "Chay truc tiep trong cua so nay...`n" -ForegroundColor Yellow
    Start-Sleep 3
    Clear-Host
    & powershell.exe -NoExit -File $mainScript
}
