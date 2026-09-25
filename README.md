<p align="center">
  <img src="app_icon.png" width="160" alt="SLAM Auto-Filler Phoenix Logo" />
</p>

# 🚆 SLAM Jobcard Auto-Filler Bot (Indian Railways)

An automated desktop application designed for Indian Railways Electric Locomotive Sheds (ELS / RPME / PPIO) to automatically parse shed inspection Excel files, stage jobcards in the **SLAM** (Software for Locomotive Asset Management) portal, and dispatch them to respective sections.

---

## 🌟 Key Features

- **Standalone Windows Executable (`SLAM_Auto_Filler.exe`)**:
  - 100% self-contained binary (~74 MB) that runs on **any Windows PC** (Windows 10 / 11).
  - No Python, Node.js, or administrative installations required.

- **Universal Multi-Browser Support (Zero Download Setup)**:
  - Automatically detects and uses installed browsers in sequence:
    1. **Bundled Chromium** (if Playwright browser cache is available).
    2. **Microsoft Edge** (`channel="msedge"`) — **pre-installed on every Windows 10 & Windows 11 PC**.
    3. **Google Chrome** (`channel="chrome"`).
  - Ensures immediate, out-of-the-box execution on any railway office computer.

- **Pre-Configured RPME Shed Presets**:
  - **Jobcard Type**: `Schedule Attention` (value `2`)
  - **Equipment Required**: `Jobcard Without Eq` (value `2`)
  - **Source (Obs Type)**: `Testing Remarks` (value `2`)
  - **Kind**: `Minor` (value `2`)

- **Per-Entry Source (Obs Type) Customization**:
  - Individual entries can have custom Source / Observation Types (`Testing Remarks`, `Driver Remarks`, `Officer Remarks`, `Special Instructions`, `Investigation Remarks`, `Seasonal Remarks`, etc.).
  - **Preview Table Column**: Displays the assigned Obs Type for every booking row.
  - **`⚡ Apply to All Entries`**: One-click button in settings to bulk-apply the selected Obs Type across all loaded rows.

- **In-GUI Message & Booking Editor**:
  - Edit any defect description message or section directly after loading.
  - Double-click any row, press `Enter`, or click **"✏️ Edit Selected Message"** to open the multiline editor.
  - Quick-toggle section buttons (`M1`, `M2`, `E3`, `E3A`, `E4`, `E5A`, `E5B`, `E8`, `BRS`, `LAB`, `PPIO`, `QAI`).
  - Add custom bookings with **"➕ Add Booking"** or delete rows with **"🗑️ Delete Selected"**.

- **Automated Authentication & Daily OTP Caching**:
  - Direct GUI management for **Username** and **Password** with show/hide password toggle (`👁️`).
  - Auto-solves math CAPTCHA (`a + b = ?`).
  - **1 OTP Per Day**: Caches today's OTP for the full day (`user_settings.json`). Subsequent runs throughout the day authenticate automatically without prompting again.

- **Mandatory Pre-Save Verification Modal**:
  - Automatically pauses once all observations are staged into the SLAM `#tbldata1` table.
  - Displays a full review table showing Job No, Section, S.No, Description, Source, and Kind.
  - Requires explicit confirmation:
    - 🟢 **`YES, PROCEED & SAVE TO SECTIONS`**: Commits save to SLAM and dispatches to sections.
    - 🔴 **`STOP & DO NOT SAVE`**: Aborts without saving, leaving observations staged in the portal for manual review.

- **Browser Session Retention & Resilience**:
  - Browser stays open after auto-fill completes or stops for visual inspection.
  - Reuses the existing open browser across consecutive Excel runs.
  - Resilient single-threaded worker eliminates Playwright greenlet and asyncio event loop conflicts.
  - Interrupted navigation protection: automatically waits for server postbacks and redirects (e.g. `SendJobcards.aspx?Type=1`) to settle before navigating.

- **Automatic GitHub Updates & Version Checking**:
  - Automatically queries GitHub Releases for updates on application startup.
  - Interactive **Update Dialog** displays what's new (release notes/changelog) with one-click download.
  - In-app progress indicator and automatic launch of the latest version.
  - Manual **"🔄 Check Updates"** button available in the top banner.

- **Intelligent Excel Parser & Telegram Desktop Integration**:
  - Auto-scans Desktop, `Downloads\Telegram Desktop`, and `Downloads` folders for `.xlsx` booking files.
  - Extracts Locomotive numbers (e.g. `30389`, `37229`, `37827`), Schedules (`IA`, `IB`, `GC`, `IC`, `ET`), and Dates.
  - Normalizes single sections and multi-section strings (e.g. `E3A & E5A`, `M1 / M2`, `E8 / E4`, `E4/E5B`, `BRS`).
  - Resilient against variations in headers (`SL. NO`, `S.NO`, `SR. NO`) and title banner collision.

---

## 🚀 How to Run

### Method 1: Standalone Executable (Recommended for Any PC)
Simply double-click:
```
SLAM_Auto_Filler.exe
```
Works on any Windows computer without installing Python or dependencies.

### Method 2: Batch Launcher
Double-click:
```
run_bot.bat
```
Automatically launches `SLAM_Auto_Filler.exe` if present, or runs via Python.

### Method 3: Run via Python (Development Mode)
```powershell
python main.py
```
Or launch the GUI directly:
```powershell
python gui.py
```

### Method 4: Command-Line (CLI) Mode
```powershell
python main.py --file "C:\Users\HP\Desktop\37229 - ET - 20-09-26 - LBR Bookings.xlsx"
```
Optional CLI parameters:
- `--loco 37229`: Override locomotive number.
- `--type "Schedule Attention"`: Override Jobcard Type.
- `--source "Testing Remarks"`: Override Source / Obs Type.
- `--kind "Minor"`: Override Kind (`Minor` or `Major`).
- `--headless`: Run browser hidden in background.
- `--otp 123456`: Pre-supply OTP.

---

## 🔨 How to Rebuild the Executable

To compile a new standalone `.exe` after making code changes:
1. Double-click **`build_exe.bat`**, OR run:
```powershell
python build_exe.py
```
2. The compiled binary will be placed at:
- `SLAM_Auto_Filler.exe`
- `dist\SLAM_Auto_Filler.exe`

---

## 📁 Project Structure

| File / Folder | Description |
|---|---|
| **`SLAM_Auto_Filler.exe`** | Standalone Windows executable ready for distribution to any PC. |
| **`gui.py`** | Tkinter desktop GUI with preview table, editor dialog, and verification modal. |
| **`slam_bot.py`** | Playwright browser automation engine with multi-browser fallback and safe navigation. |
| **`excel_parser.py`** | Excel parser extracting Loco, Schedule, Date, and normalizing section codes. |
| **`config.py`** | Portal endpoints, DOM section IDs, default values, and settings persistence. |
| **`main.py`** | Unified application entry point supporting both GUI and CLI modes. |
| **`build_exe.py`** | Automated PyInstaller packaging script with Playwright bundling. |
| **`build_exe.bat`** | 1-click batch script to rebuild the standalone `.exe`. |
| **`run_bot.bat`** | 1-click Windows batch launcher. |
| **`.gitignore`** | Git ignore rules protecting credentials, binaries, and build artifacts. |
| **`requirements.txt`** | Python dependencies (`playwright`, `openpyxl`). |

---

## ⚙️ Configuration & Section Codes

### Valid SLAM Section Codes
| Code | DOM ID | Description |
|---|---|---|
| `BRS` | 58 | Brake System |
| `E3` | 46 | Auxiliary Machines / Converters |
| `E3A` | 48 | Traction Motor & Power Circuit |
| `E4` | 49 | Electronics & Microprocessor |
| `E5A` | 50 | Control Circuit / Cab Equipment |
| `E5B` | 51 | Battery & Auxiliary Circuit |
| `E8` | 53 | High Voltage / Roof Equipment |
| `LAB` | 55 | Oil / Chemical Laboratory |
| `M1` | 56 | Mechanical - Bogie & Suspension |
| `M2` | 57 | Mechanical - Pneumatic & Undergear |
| `PPIO` | 60 | Progress, Planning & Inspection Office |
| `QAI` | 45 | Quality Assurance & Inspection |

### Source (Observation Types)
- **`Testing Remarks`** (Default - Value `2`)
- **`Driver Remarks`** (Value `1`)
- **`Officer Remarks`** (Value `3`)
- **`Special Instructions`** (Value `4`)
- **`Investigation Remarks`** (Value `5`)
- **`Seasonal Remarks`** (Value `6`)
- **`Before Test Remarks`** (Value `7`)
- **`Super Check`** (Value `11`)
- **`Out Station Remarks`** (Value `13`)
- **`DDS Remarks`** (Value `14`)