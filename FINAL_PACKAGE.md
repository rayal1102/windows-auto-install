# ===================================================================
# WINDOWS AUTO INSTALL - PACKAGE HOAN CHINH CUOI CUNG
# ===================================================================

## DANH SACH TAT CA FILES

### FILES CHINH (4 files - BAT BUOC)
1. **install.ps1** - Bootstrap script (chay dau tien)
2. **main-script.ps1** - Script chinh (chay trong Terminal)  
3. **apps.json** - Danh sach apps
4. **AppSelector.psm1** - Module GUI chon apps

### FILES BO SUNG (2 files - TUY CHON)
5. **test-office.ps1** - Test cai Office rieng
6. **README.md** - Huong dan su dung

---

## TINH NANG MOI TRONG PHIEN BAN NAY

### 1. XOA CACHE WINDOWS UPDATE (MOI!)
- Tu dong xoa thu muc `C:\Windows\SoftwareDistribution`
- Don dep cache truoc khi update
- Tang toc do va giam loi

### 2. TUY CHON BO QUA WINDOWS UPDATE (MOI!)
```
Ban co muon cap nhat Windows Update khong?
Cap nhat Windows? (Y/N)

→ Chon N de bo qua, nhay thang den cai apps
→ Chon Y de tien hanh update
```

### 3. GIAO DIEN GUI CHON APPS (MOI!)
- Tick chon apps bang chuot
- Tim kiem apps
- Nut "Chon tat ca"
- Dem tu dong so app da chon

### 4. DANH SACH APPS RIENG BIET (MOI!)
- File apps.json rieng
- De chinh sua, them/bot apps
- 14 apps co ban + 35+ apps tuy chon

### 5. CAI DAT THONG MINH
- Kiem tra app da cai → Update
- App chua cai → Install moi
- Hien thi tien trinh chi tiet

---

## HUONG DAN CAI DAT STEP-BY-STEP

### BUOC 1: TAO REPOSITORY TREN GITHUB

1. Vao https://github.com/new
2. Repository name: `windows-auto-install`
3. **QUAN TRONG:** Chon **Public** (khong phai Private!)
4. Click "Create repository"

### BUOC 2: UPLOAD 4 FILE CHINH

**Cach 1: Qua giao dien web (DE NHAT)**

1. Trong repository vua tao, click **"Add file"** > **"Upload files"**

2. Keo tha 4 file:
   - install.ps1
   - main-script.ps1
   - apps.json  
   - AppSelector.psm1

3. Commit message: "Initial commit"

4. Click **"Commit changes"**

**Cach 2: Qua Git command line**

```bash
git clone https://github.com/YOUR_USERNAME/windows-auto-install
cd windows-auto-install
# Copy 4 file vao thu muc nay
git add .
git commit -m "Initial commit"
git push
```

### BUOC 3: SUA LINK TRONG CAC FILE

**⚠️ QUAN TRONG: Phai sua 3 link trong cac file**

#### 3.1. Sua file install.ps1

Mo file `install.ps1` tren GitHub > Click **Edit**

Tim dong:
```powershell
$mainScriptUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/windows-auto-install/main/main-script.ps1"
```

Thay `YOUR_USERNAME` bang username GitHub cua ban:
```powershell
$mainScriptUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/main-script.ps1"
```

Click **"Commit changes"**

#### 3.2. Sua file main-script.ps1

Mo file `main-script.ps1` > Click **Edit**

Tim 2 dong sau va thay `YOUR_USERNAME`:

```powershell
# DONG 1:
$appsJsonUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/windows-auto-install/main/apps.json"

# DONG 2:
$guiModuleUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/windows-auto-install/main/AppSelector.psm1"
```

Thay thanh:
```powershell
$appsJsonUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/apps.json"
$guiModuleUrl = "https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/AppSelector.psm1"
```

Click **"Commit changes"**

### BUOC 4: LAY LINK SU DUNG

1. Vao file `install.ps1` tren GitHub
2. Click nut **"Raw"** (goc tren ben phai)
3. Copy URL tren trinh duyet

Vi du:
```
https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/install.ps1
```

### BUOC 5: TAO LINK NGAN (TUY CHON)

De de nho hon, tao link ngan:

1. Vao https://is.gd
2. Paste link RAW vua copy
3. Tao link ngan, vi du: `is.gd/PCsetup`

### BUOC 6: SU DUNG

**Cach 1: Link day du**
```powershell
irm https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/install.ps1 | iex
```

**Cach 2: Link ngan**
```powershell
irm is.gd/PCsetup | iex
```

---

## LUONG HOAT DONG CHI TIET

```
PowerShell (cu)
===============
install.ps1
│
├─> Kiem tra Internet
├─> Cai WinGet (neu chua co)
├─> Sua loi certificate
├─> Cai Windows Terminal
├─> Tai main-script.ps1 tu GitHub
│
└─> Mo Terminal moi
    Dong PowerShell cu

            ↓

Windows Terminal (moi)
======================
main-script.ps1
│
├─> [1/4] GO BLOATWARE
│   └─> Xoa 25 apps (Cortana, Xbox...)
│
├─> [2/4] WINDOWS UPDATE (MOI!)
│   ├─> Hoi: "Cap nhat Windows? (Y/N)"
│   ├─> Neu Y:
│   │   ├─> Xoa cache SoftwareDistribution
│   │   ├─> Quet updates
│   │   ├─> Hien thi danh sach
│   │   └─> Cai dat (bo qua updates can restart)
│   └─> Neu N: Bo qua, nhay sang buoc 3
│
├─> [3/4] CAI PHAN MEM
│   ├─> Tai apps.json tu GitHub
│   ├─> Tai AppSelector.psm1 tu GitHub
│   ├─> Mo CUA SO GUI #1: Essential Apps
│   │   ├─> Tick chon 14 apps co ban
│   │   └─> Click OK
│   └─> Cai dat apps da chon
│
├─> [BONUS] CHON THEM PHAN MEM
│   ├─> Hoi: "Chon them? (Y/N)"
│   ├─> Neu Y:
│   │   ├─> Mo CUA SO GUI #2: Optional Apps
│   │   ├─> Tick chon 35+ apps tuy chon
│   │   └─> Click OK
│   └─> Cai dat apps tuy chon
│
├─> [4/4] CAI OFFICE 365
│   ├─> Tai Office Deployment Tool
│   ├─> Cai Word, Excel, PowerPoint
│   └─> Bo qua Outlook, Teams, OneNote
│
└─> KET THUC
    ├─> Hien thi tong ket
    ├─> Hoi co restart khong
    └─> Done!
```

---

## OUTPUT MAU

```
========================================
  WINDOWS AUTO INSTALL - TERMINAL
========================================

[1/4] GO BLOATWARE
==================

[1/25] Cortana - [OK]
[2/25] Xbox - [OK]
...
[25/25] Movies - [SKIP]

Da go: 15/25

[2/4] WINDOWS UPDATE
==================

Ban co muon cap nhat Windows Update khong?
Cap nhat Windows? (Y/N) y

Don dep thu muc Windows Update...
  - Dung Windows Update service...
  - Xoa cache Windows Update...
  [OK] Da don dep cache

Quet updates...
  Tim thay: 8 updates

  [CHON] 2025-01 Cumulative Update for .NET Framework
         KB: KB5066128 | Kich thuoc: 179MB
  [SKIP] 2026-02 Security Update
         Kich thuoc: 89GB (qua lon)
  ...

Bat dau cai 5 updates...
...

[3/4] CAI PHAN MEM
==================

Tai danh sach phan mem...
  [OK] Da tai danh sach

Mo giao dien chon phan mem...

>>> DANG MO CUA SO CHON PHAN MEM <<<
    Vui long chon apps trong cua so moi mo ra

[GUI mo ra... tick chon... OK]

Da chon 9/14 phan mem

[1/9] UniKey...
  → Chua cai, dang cai dat...
  [OK] Cai dat thanh cong
...

[BONUS] CHON THEM PHAN MEM
==================

Ban co muon cai them phan mem khac khong?
Chon them phan mem? (Y/N) y

>>> DANG MO CUA SO TUY CHON <<<

[GUI mo ra... tick chon... OK]

Da chon 5 phan mem tuy chon
...

[4/4] CAI OFFICE 365
==================

Dang cai Office 365 (Word, Excel, PowerPoint)...
  - Tai Office Deployment Tool...
  - Cai dat (10-20 phut)...
  [OK] Office cai thanh cong

========================================
  HOAN TAT
========================================

TONG KET:
  - Bloatware: 15/25
  - Updates: 5
  - Phan mem: 14/14
  - Office: Da cai

Log: C:\Users\...\Temp\WindowsAutoInstall_20260214_103045.log

Ban co muon restart ngay? (Y/N) n
Khong restart. Hay restart sau!

Bam phim bat ky de dong...
```

---

## CAU TRUC THU MUC

```
windows-auto-install/
│
├── install.ps1              # Bootstrap (6KB)
├── main-script.ps1          # Script chinh (28KB)
├── apps.json                # Danh sach apps (7KB)
├── AppSelector.psm1         # Module GUI (6KB)
├── test-office.ps1          # Test Office (10KB)
└── README.md                # Huong dan (5KB)

TONG: ~62KB
```

---

## CHINH SUA DANH SACH APPS

### Them app moi vao apps.json

1. Mo file `apps.json` tren GitHub
2. Click **Edit**
3. Them app:

```json
{
  "name": "Zalo",
  "id": "Zalo.Zalo",
  "description": "Nhan tin Viet Nam"
}
```

4. Commit changes

### Tim WinGet ID

```powershell
winget search "Ten App"

# Vi du:
> winget search "Zalo"
Name  Id         Version
-----------------------
Zalo  Zalo.Zalo  24.1.2
      ^^^^^^^^^^
      Copy ID nay
```

---

## KHAC PHUC SU CO

### Loi 1: "scripts is disabled"

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
irm YOUR_LINK | iex
```

### Loi 2: "404 Not Found"

- Repository phai la **PUBLIC**
- Link phai co `raw.githubusercontent.com`
- Da sua `YOUR_USERNAME` chua?

### Loi 3: GUI khong mo

- Script tu dong chuyen sang TEXT MODE
- Neu muon force TEXT:
  - Xoa dong `$useGUI = $true` trong script

### Loi 4: Office loi 404

- Link ODT het han
- Xem file `FIX_ODT_404.txt`

### Loi 5: Windows Update khong chay

- Chon **Y** khi duoc hoi
- Neu van loi → Chon **N** de bo qua

---

## TEST TUNG THANH PHAN

### Test Bootstrap:
```powershell
irm YOUR_LINK/install.ps1 -OutFile test.ps1
.\test.ps1
```

### Test Script chinh:
```powershell
irm YOUR_LINK/main-script.ps1 -OutFile main.ps1
.\main.ps1
```

### Test Office rieng:
```powershell
irm YOUR_LINK/test-office.ps1 -OutFile test-office.ps1
.\test-office.ps1
```

### Test GUI:
```powershell
irm YOUR_LINK/AppSelector.psm1 -OutFile AppSelector.psm1
Import-Module .\AppSelector.psm1
$apps = @(@{name="Test";id="Test.Test";description="Test app"})
Show-AppSelector -Apps $apps -Title "TEST"
```

---

## YEU CAU HE THONG

- Windows 10 (build 1809+) hoac Windows 11
- PowerShell 5.1+
- Internet connection
- Quyen Administrator
- ~5GB dung luong trong (cho updates + apps)

---

## TINH NANG MO RONG

### Them bloatware moi

Sua file `main-script.ps1`, tim `$bloat`:
```powershell
$bloat = @{
    "TenApp"="Microsoft.AppName*"
}
```

### Thay doi kich thuoc update toi da

Sua `$MaxUpdateSizeGB` trong `main-script.ps1`:
```powershell
$MaxUpdateSizeGB = 10  # 10GB
```

### Disable GUI, chi dung TEXT

Them dong nay dau file `main-script.ps1`:
```powershell
$useGUI = $false
```

---

## HO TRO

- GitHub Issues: https://github.com/YOUR_USERNAME/windows-auto-install/issues
- Log file: `%TEMP%\WindowsAutoInstall_*.log`
- WinGet docs: https://learn.microsoft.com/en-us/windows/package-manager/winget/

---

## CHANGELOG

### v3.0 (2026-02-14) - PHIEN BAN NAY
- [NEW] Xoa cache SoftwareDistribution truoc khi update
- [NEW] Tuy chon bo qua Windows Update
- [NEW] GUI tick chon apps
- [NEW] File apps.json rieng
- [NEW] Kiem tra app da cai → update thay vi cai lai
- [FIX] Windows Update khong cai duoc driver
- [FIX] WinGet certificate error
- [FIX] Link ODT het han

### v2.0 (2026-01-20)
- 2-file structure
- Fix encoding errors
- Add Office installation

### v1.0 (2026-01-15)
- Initial release

---

## LICENSE

MIT License - Free to use, modify, and distribute

---

## TOM TAT NHANH

1. **Upload 4 files** len GitHub (Public)
2. **Sua 3 links** trong install.ps1 va main-script.ps1
3. **Copy link RAW** cua install.ps1
4. **Chay:** `irm YOUR_LINK | iex`
5. **Done!**

---

Vi du day du username `rayal1102`:

```powershell
irm https://raw.githubusercontent.com/rayal1102/windows-auto-install/main/install.ps1 | iex
```

Hoac link ngan:
```powershell
irm is.gd/PCsetup | iex
```

CHUC MAY MAN! 🎉
