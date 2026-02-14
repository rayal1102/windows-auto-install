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
            $skippedLarge = @()
            
            foreach ($u in $updates) {
                # Xu ly truong hop Size = null hoac 0
                $size = if ($u.Size) { $u.Size } else { 0 }
                $sizeGB = [math]::Round($size/1GB, 2)
                $sizeMB = [math]::Round($size/1MB, 0)
                
                # Hien thi kich thuoc
                $sizeText = if ($sizeMB -gt 0) { "${sizeMB}MB" } else { "Unknown" }
                
                # Lay ten update
                $updateName = if ($u.Title) { $u.Title } else { $u.KB }
                
                # Bo qua updates qua lon
                if ($size -gt 0 -and $sizeGB -gt $MaxUpdateSizeGB) {
                    $skippedLarge += $u
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
            if ($skippedLarge.Count -gt 0) {
                Write-Log "Bo qua $($skippedLarge.Count) updates qua lon" "WARNING"
            }
            
            if ($valid.Count -gt 0) {
                Write-Log "`nBat dau cai $($valid.Count) updates..." "WARNING"
                Write-Log "Dang tai va cai dat tat ca cung luc (co the mat 5-15 phut)...`n" "INFO"
                
                try {
                    # CAI TAT CA UPDATES CUNG LUC - GIONG CHAY THU CONG
                    # Khong dung -KBArticleID vi mot so updates khong ho tro
                    $result = Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -EA Stop
                    
                    if ($result) {
                        # Dem so luong cai thanh cong
                        $installed = ($result | Where-Object { $_.Result -eq 'Installed' -or $_.Result -eq 'Downloaded' }).Count
                        
                        Write-Log "`nChi tiet ket qua:" "INFO"
                        foreach ($r in $result) {
                            $status = if ($r.Result -eq 'Installed') { "SUCCESS" } else { "WARNING" }
                            $title = if ($r.Title) { $r.Title } else { $r.KB }
                            Write-Log "  [$($r.Result)] $title" $status
                        }
                        
                        Write-Log "`nTong ket: Cai thanh cong $installed/$($valid.Count) updates" "SUCCESS"
                    } else {
                        Write-Log "`nKhong co ket qua tra ve" "WARNING"
                    }
                } catch {
                    Write-Log "`n[LOI] Loi khi cai updates: $($_.Exception.Message)" "ERROR"
                    Write-Log "Thu cai tung update rieng le..." "WARNING"
                    
                    # Neu cai tat ca that bai, thu tung cai mot
                    $installed = 0
                    $currentUpdate = 0
                    foreach ($u in $valid) {
                        $currentUpdate++
                        $updateName = if ($u.Title) { $u.Title } else { $u.KB }
                        
                        Write-Log "`n[$currentUpdate/$($valid.Count)] $updateName" "INFO"
                        try {
                            $r = Install-WindowsUpdate -KBArticleID $u.KB -MicrosoftUpdate -AcceptAll -IgnoreReboot -Confirm:$false -EA Stop
                            if ($r) {
                                Write-Log "  [OK] Cai thanh cong" "SUCCESS"
                                $installed++
                            }
                        } catch {
                            Write-Log "  [LOI] $($_.Exception.Message)" "ERROR"
                        }
                    }
                    Write-Log "`nTong ket: Cai thanh cong $installed/$($valid.Count) updates" "SUCCESS"
                }
                
                Write-Host ""
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

# Tai danh sach apps tu file JSON
$appsJsonUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/refs/heads/main/apps.json"
$appsJsonPath = "$env:TEMP\apps.json"

Write-Log "Tai danh sach phan mem..." "INFO"
try {
    Invoke-WebRequest -Uri $appsJsonUrl -OutFile $appsJsonPath -UseBasicParsing | Out-Null
    $appsConfig = Get-Content $appsJsonPath -Raw | ConvertFrom-Json
    Write-Log "  [OK] Da tai danh sach`n" "SUCCESS"
} catch {
    Write-Log "  [LOI] Khong tai duoc danh sach, dung danh sach mac dinh" "WARNING"
    $appsConfig = @{
        essential_apps = @{
            apps = @(
                @{name="UniKey"; id="UniKey.UniKey"; description="Go tieng Viet"},
                @{name="WinRAR"; id="RARLab.WinRAR"; description="Giai nen file"},
                @{name="7-Zip"; id="7zip.7zip"; description="Giai nen mien phi"},
                @{name="Chrome"; id="Google.Chrome"; description="Trinh duyet"}
            )
        }
    } | ConvertTo-Json -Depth 5 | ConvertFrom-Json
}

# Tai module GUI
$guiModuleUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/refs/heads/main/AppSelector.psm1"
$guiModulePath = "$env:TEMP\AppSelector.psm1"

try {
    Invoke-WebRequest -Uri $guiModuleUrl -OutFile $guiModulePath -UseBasicParsing | Out-Null
    Import-Module $guiModulePath -Force
    $useGUI = $true
} catch {
    Write-Log "Khong tai duoc GUI module, dung che do text" "WARNING"
    $useGUI = $false
}

# Hien thi va chon apps
$essentialApps = $appsConfig.essential_apps.apps
$selectedEssential = @()

if ($useGUI) {
    Write-Log "Mo giao dien chon phan mem..." "INFO"
    Write-Host ""
    Write-Host ">>> DANG MO CUA SO CHON PHAN MEM <<<" -ForegroundColor Yellow
    Write-Host "    Vui long chon apps trong cua so moi mo ra`n" -ForegroundColor Cyan
    
    try {
        $selectedEssential = Show-AppSelector -Apps $essentialApps -Title "CHON PHAN MEM CO BAN" -Description "Tick chon cac phan mem ban muon cai dat"
        
        if ($selectedEssential.Count -gt 0) {
            Write-Log "Da chon $($selectedEssential.Count)/$($essentialApps.Count) phan mem`n" "INFO"
        } else {
            Write-Log "Khong chon phan mem nao`n" "WARNING"
        }
    } catch {
        Write-Log "Loi GUI: $($_.Exception.Message)" "ERROR"
        Write-Log "Chuyen sang che do text...`n" "WARNING"
        $useGUI = $false
    }
}

# Fallback: Text mode neu GUI khong hoat dong
if (-not $useGUI -or $selectedEssential.Count -eq 0) {
    Write-Log "PHAN MEM CO BAN (Khuyến nghị):" "WARNING"
    Write-Host ""
    
    for ($i = 0; $i -lt $essentialApps.Count; $i++) {
        $app = $essentialApps[$i]
        Write-Host "  [$($i+1)] $($app.name)" -ForegroundColor Cyan
        if ($app.description) {
            Write-Host "      $($app.description)" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
    Write-Host "HUONG DAN:" -ForegroundColor Yellow
    Write-Host "  - Nhap so thu tu cach nhau boi dau phay (vd: 1,2,5,10)" -ForegroundColor Cyan
    Write-Host "  - Nhap 'ALL' de chon tat ca" -ForegroundColor Cyan
    Write-Host "  - Bam Enter de chon tat ca`n" -ForegroundColor Cyan
    
    $essentialChoice = Read-Host "Chon phan mem co ban"
    
    if (-not $essentialChoice -or $essentialChoice.Trim().ToLower() -eq "all") {
        $selectedEssential = $essentialApps
        Write-Log "Chon tat ca $($essentialApps.Count) phan mem co ban`n" "INFO"
    } else {
        $selectedEssential = @()
        $indices = $essentialChoice -split "," | ForEach-Object { $_.Trim() }
        foreach ($idx in $indices) {
            if ($idx -match '^\d+$') {
                $arrayIdx = [int]$idx - 1
                if ($arrayIdx -ge 0 -and $arrayIdx -lt $essentialApps.Count) {
                    $selectedEssential += $essentialApps[$arrayIdx]
                }
            }
        }
        Write-Log "Chon $($selectedEssential.Count)/$($essentialApps.Count) phan mem co ban`n" "INFO"
    }
}

# Cai dat cac apps da chon
if ($selectedEssential.Count -gt 0) {
    Write-Log "Bat dau cai dat phan mem co ban..." "WARNING"
    Write-Host ""
    
    $success = 0
    $updated = 0
    $failed = 0
    
    for ($i = 0; $i -lt $selectedEssential.Count; $i++) {
        $app = $selectedEssential[$i]
        Write-Log "[$($i+1)/$($selectedEssential.Count)] $($app.name)..." "INFO"
        
        $checkResult = & $winget list --id $app.id --exact 2>&1
        $isInstalled = $checkResult -match $app.id
        
        try {
            if ($isInstalled) {
                Write-Log "  → Da cai, dang cap nhat..." "INFO"
                & $winget upgrade --id $app.id --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "  [OK] Cap nhat thanh cong" "SUCCESS"
                    $updated++
                    $success++
                } elseif ($LASTEXITCODE -eq -1978335189) {
                    Write-Log "  [OK] Da la phien ban moi nhat" "SUCCESS"
                    $success++
                } else {
                    Write-Log "  [LOI] Loi cap nhat" "ERROR"
                    $failed++
                }
            } else {
                Write-Log "  → Chua cai, dang cai dat..." "INFO"
                & $winget install -e --id $app.id --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
                
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "  [OK] Cai dat thanh cong" "SUCCESS"
                    $success++
                } else {
                    Write-Log "  [LOI] Loi cai dat" "ERROR"
                    $failed++
                }
            }
        } catch {
            Write-Log "  [LOI] $($_.Exception.Message)" "ERROR"
            $failed++
        }
    }
    
    Write-Host ""
    Write-Log "Tong ket phan mem co ban:" "INFO"
    Write-Log "  - Thanh cong: $success/$($selectedEssential.Count)" "SUCCESS"
    if ($updated -gt 0) {
        Write-Log "  - Da cap nhat: $updated" "INFO"
    }
    if ($failed -gt 0) {
        Write-Log "  - That bai: $failed" "WARNING"
    }
    Write-Host ""
} else {
    Write-Log "Khong chon phan mem nao`n" "WARNING"
}

# ===================================================================
# BONUS: CHON THEM PHAN MEM TUY CHON
# ===================================================================
Write-Log "`n[BONUS] CHON THEM PHAN MEM" "WARNING"
Write-Log "==================`n"

Write-Log "Ban co muon cai them phan mem khac khong?" "INFO"
$wantMore = Read-Host "Chon them phan mem? (Y/N)"

if ($wantMore -eq 'Y' -or $wantMore -eq 'y') {
    Write-Host ""
    
    # Lay danh sach optional apps
    $allOptionalApps = @()
    $optionalCategories = $appsConfig.optional_apps.categories
    
    foreach ($category in $optionalCategories.PSObject.Properties) {
        $categoryName = $category.Name
        $apps = $category.Value
        
        foreach ($app in $apps) {
            $allOptionalApps += @{
                name = $app.name
                id = $app.id
                description = "[$categoryName] $($app.description)"
            }
        }
    }
    
    $selectedOptional = @()
    
    if ($useGUI) {
        Write-Log "Mo giao dien chon phan mem tuy chon..." "INFO"
        Write-Host ""
        Write-Host ">>> DANG MO CUA SO CHON PHAN MEM TUY CHON <<<" -ForegroundColor Yellow
        Write-Host "    Vui long chon apps trong cua so moi mo ra`n" -ForegroundColor Cyan
        
        try {
            $selectedOptional = Show-AppSelector -Apps $allOptionalApps -Title "CHON PHAN MEM TUY CHON" -Description "Tick chon cac phan mem ban muon cai them"
            
            if ($selectedOptional.Count -gt 0) {
                Write-Log "Da chon $($selectedOptional.Count) phan mem tuy chon`n" "INFO"
            } else {
                Write-Log "Khong chon phan mem nao`n" "INFO"
            }
        } catch {
            Write-Log "Loi GUI: $($_.Exception.Message)" "ERROR"
            $selectedOptional = @()
        }
    } else {
        # Text mode fallback
        Write-Log "Hien thi danh sach phan mem tuy chon..." "INFO"
        Write-Host ""
        
        Write-Host "DANH SACH PHAN MEM TUY CHON" -ForegroundColor Yellow
        Write-Host "===========================`n" -ForegroundColor Yellow
        
        $allOptionalAppsText = @{}
        foreach ($category in $optionalCategories.PSObject.Properties) {
            $categoryName = $category.Name
            $apps = $category.Value
            
            Write-Host "$categoryName" -ForegroundColor Cyan
            for ($i = 0; $i -lt $apps.Count; $i++) {
                $app = $apps[$i]
                Write-Host "  [$($i+1)] $($app.name)" -ForegroundColor Gray
                if ($app.description) {
                    Write-Host "      $($app.description)" -ForegroundColor DarkGray
                }
                $allOptionalAppsText[$app.name] = $app.id
            }
            Write-Host ""
        }
        
        Write-Host "HUONG DAN:" -ForegroundColor Yellow
        Write-Host "  - Nhap ten phan mem (vd: Firefox, Discord)" -ForegroundColor Cyan
        Write-Host "  - Nhap nhieu phan mem cach nhau boi dau phay" -ForegroundColor Cyan
        Write-Host "  - Nhap 'ALL' de chon tat ca" -ForegroundColor Cyan
        Write-Host "  - Bam Enter de bo qua`n" -ForegroundColor Cyan
        
        $selection = Read-Host "Chon phan mem"
        
        if ($selection -and $selection.Trim() -ne "") {
            if ($selection.Trim().ToLower() -eq "all") {
                $selectedOptional = $allOptionalApps
            } else {
                $selections = $selection -split ","
                foreach ($sel in $selections) {
                    $appName = $sel.Trim()
                    if ($allOptionalAppsText.ContainsKey($appName)) {
                        $selectedOptional += @{
                            name = $appName
                            id = $allOptionalAppsText[$appName]
                        }
                    } else {
                        Write-Log "  [SKIP] Khong tim thay: $appName" "WARNING"
                    }
                }
            }
        }
    }
    
    if ($selectedOptional.Count -gt 0) {
        Write-Host ""
        Write-Log "Dang cai $($selectedOptional.Count) phan mem..." "WARNING"
        Write-Host ""
        
        $bonusSuccess = 0
        $bonusFailed = 0
        $bonusUpdated = 0
        
        for ($bonusIdx = 0; $bonusIdx -lt $selectedOptional.Count; $bonusIdx++) {
            $app = $selectedOptional[$bonusIdx]
            $appName = $app.name
            $appId = $app.id
            
            Write-Log "[$($bonusIdx+1)/$($selectedOptional.Count)] $appName..." "INFO"
            
            $checkResult = & $winget list --id $appId --exact 2>&1
            $isInstalled = $checkResult -match $appId
            
            try {
                if ($isInstalled) {
                    Write-Log "  → Da cai, dang cap nhat..." "INFO"
                    & $winget upgrade --id $appId --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "  [OK] Cap nhat thanh cong" "SUCCESS"
                        $bonusSuccess++
                        $bonusUpdated++
                    } elseif ($LASTEXITCODE -eq -1978335189) {
                        Write-Log "  [OK] Da la phien ban moi nhat" "SUCCESS"
                        $bonusSuccess++
                    } else {
                        Write-Log "  [LOI] Loi cap nhat" "ERROR"
                        $bonusFailed++
                    }
                } else {
                    Write-Log "  → Chua cai, dang cai dat..." "INFO"
                    & $winget install -e --id $appId --source winget --silent --accept-source-agreements --accept-package-agreements 2>&1 | Out-Null
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "  [OK] Cai dat thanh cong" "SUCCESS"
                        $bonusSuccess++
                    } else {
                        Write-Log "  [LOI] Loi cai dat" "ERROR"
                        $bonusFailed++
                    }
                }
            } catch {
                Write-Log "  [LOI] $($_.Exception.Message)" "ERROR"
                $bonusFailed++
            }
        }
        
        Write-Host ""
        Write-Log "Tong ket phan mem tuy chon:" "INFO"
        Write-Log "  - Thanh cong: $bonusSuccess/$($selectedOptional.Count)" "SUCCESS"
        if ($bonusUpdated -gt 0) {
            Write-Log "  - Da cap nhat: $bonusUpdated" "INFO"
        }
        if ($bonusFailed -gt 0) {
            Write-Log "  - That bai: $bonusFailed" "WARNING"
        }
        Write-Host ""
    } else {
        Write-Log "Khong chon phan mem nao`n" "INFO"
    }
} else {
    Write-Log "Bo qua cai them phan mem`n" "INFO"
}

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
