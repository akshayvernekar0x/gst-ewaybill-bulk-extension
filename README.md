# 🧾 GST E-Way Bill Bulk Extension Tool

> Automate bulk extension of GST E-Way Bills on [ewaybillgst.gov.in](https://ewaybillgst.gov.in) — built with Python + Selenium.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-Browser%20Automation-43B02A?logo=selenium&logoColor=white)
![Version](https://img.shields.io/badge/Version-6.0-orange)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgray?logo=windows)

---

## 📌 What It Does

Extending E-Way Bills one by one on the GST portal is painfully slow.
This tool automates the entire process in bulk — upload your Excel file,
log in once, and the script handles the rest automatically.

- ✅ Reads EWB numbers, truck numbers, place, pincode, and distance from your Excel file
- ✅ Full GUI window — no black console, no technical knowledge needed
- ✅ Logs into the portal manually (no credentials stored anywhere)
- ✅ Extends each bill automatically with human-like typing to bypass bot detection
- ✅ Catches portal alerts (expired, 8-hour rule, not found) and logs them with the exact message
- ✅ Auto-retry on failed bills — 1 automatic retry before marking as failed
- ✅ Saves a progress report every 5 bills — stop anytime without losing data
- ✅ Chrome auto-hides after login — only the GUI is visible while it works
- ✅ License key system — control access remotely via Google Sheets

---

## 🆕 What's New in v6.0

| # | Change | Effect |
|---|---|---|
| v6-1 | Anti-bot fully restored | Won't get detected or blocked by portal |
| v6-2 | `webdriver-manager` import fixed | No more buried imports, EXE builds cleaner |
| v6-3 | `io.StringIO` fix | Removed redundant inner import |
| v6-4 | Login sleep 5s → 1s | Saves time on every run |
| v6-5 | Smart alert wait after Go | No more 2s dead wait per bill |
| v6-6 | Smart pincode state wait | Waits for state to load, not a fixed sleep |
| v6-7 | Submit sleep 3s → 1s | Saves ~2s per bill (~5 min for 144 bills) |
| v6-8 | Report saves every 5 bills | Not writing Excel 144 times |
| v6-9 | Auto retry on failure | Fewer false failures from portal lag |
| v6-10 | `js_fill` combined JS call | Faster field filling |
| v6-11 | Bill list logged as DEBUG | Clean console on startup |
| v6-12 | Full GUI window | No black console — professional interface |

---

## 🖥️ GUI Preview

The tool runs with a full graphical interface:

- **Excel file picker** — browse and select your bills file
- **Live stat counters** — Total / Processed / Success / Failed
- **Progress bar** — with bill count and percentage
- **Current bill strip** — shows EWB, truck, place being processed
- **Live log window** — colour coded (green = success, yellow = retry, red = error)
- **Start / Stop / Open Report** buttons

---

## 🛡️ Anti-Bot Protection

The script uses 4 layers of protection so the portal cannot detect automation:

| Layer | What It Does |
|---|---|
| Chrome flags | Hides `AutomationControlled` banner and automation extension |
| User-Agent rotation | Picks a random real Chrome UA each run |
| CDP script | Hides `navigator.webdriver`, fakes plugins and language |
| Human behaviour | Char-by-char typing, random delays, random scrolls, typo simulation |

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| `Python 3.8+` | Core language |
| `Selenium` | Browser automation (Chrome) |
| `webdriver-manager` | Auto-downloads correct ChromeDriver — no manual setup |
| `Pandas` | Excel/CSV reading and report generation |
| `Tkinter` | Full GUI window |
| `Requests` | License verification via Google Sheets CSV |
| `openpyxl` | Excel report output |

---

## ⚙️ Setup (Run from Source)

### 1. Install dependencies

```bash
pip install selenium pandas requests openpyxl webdriver-manager
```

> No need to manually download ChromeDriver — `webdriver-manager` handles it automatically.

### 2. Configure your license sheet URL

```python
LICENSE_SHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/pub?output=csv"
```

Your Google Sheet should have these columns:

| LicenseKey | ExpiryDate |
|---|---|
| ABCD-1234 | 31-Dec-2026 |

### 3. Configure fallback defaults (optional)

If your Excel file doesn't have Place / Pincode / Distance columns,
the script uses these defaults:

```python
DEFAULT_PLACE    = "Banglore"
DEFAULT_PINCODE  = "123456"
DEFAULT_DISTANCE = "100"
```

---

## 🚀 How to Run

```bash
python ewaybill_v6_0.py
```

1. Enter your license key when prompted
2. Click **Browse** and select your Excel file
3. Click **Start**
4. A Chrome window opens — log in manually, complete CAPTCHA and OTP
5. Chrome hides automatically after login
6. The script extends all bills in the background — GUI shows live progress
7. Click **Open Report** anytime to see the results Excel file

---

## 📋 Excel File Format

Your input file must have at minimum an EWB column.
All other columns are optional — the script uses defaults if missing.

| Column Type | Accepted Names |
|---|---|
| E-Way Bill No | `EWAY BILL NO`, `EWB_NO`, `EWB NO`, `EWAYBILLNO` |
| Truck / Vehicle | `TRUCK NO`, `VEHICLE NO`, `VEHICLE`, `TRUCK` |
| Current Place | `CURRENT PLACE`, `CUR PLACE`, `PLACE` |
| Pincode | `PINCODE`, `PIN CODE`, `PIN` |
| Distance (km) | `DISTANCE`, `DIST`, `APPROX DISTANCE` |

> Headers can be on any of the first 5 rows — auto-detected.

---

## 📊 Output Report

An Excel report (`extension_report_YYYYMMDD_HHMMSS.xlsx`) is saved every 5 bills
and always at the end:

| EWB No | Truck No | Status | Message | Time |
|---|---|---|---|---|
| 123456789012 | KA01AB1234 | SUCCESS | Extended successfully | 14:32:01 |
| 234567890123 | MH02CD5678 | FAILED | Portal blocked: EWB is expired | 14:32:18 |

---

## 🔐 License System

The license system lets you:

- Issue unique keys to each client
- Set an expiry date per key
- Revoke access instantly by deleting the row from your Google Sheet
- No code changes needed — managed 100% remotely

---

## 📁 Project Structure

```
ewaybill_v6_0.py              ← Main script
logo.ico                      ← App icon (optional)
ewaybill_selenium.log         ← Auto-generated log file
extension_report_*.xlsx       ← Auto-generated report files
debug_*.png                   ← Debug screenshots (first bill only)
```

---

## 🏗️ Build as EXE

To distribute as a standalone Windows EXE (no Python needed on user's PC):

```bash
pyinstaller --onefile --noconsole --icon=logo.ico --add-data "logo.ico;." ^
  --hidden-import=selenium ^
  --hidden-import=webdriver_manager ^
  --hidden-import=webdriver_manager.chrome ^
  --hidden-import=webdriver_manager.core.os_manager ^
  --hidden-import=pkg_resources ^
  --hidden-import=tkinter ^
  --hidden-import=openpyxl ^
  --hidden-import=requests ^
  ewaybill_v6_0.py
```

The built EXE will be in the `dist/` folder.

> ChromeDriver is NOT bundled in the EXE — it auto-downloads at runtime
> and is cached locally. Chrome updates never break the tool.

---

## ⚠️ Disclaimer

This tool is built for legitimate business use to save time on
repetitive GST portal tasks. Use it responsibly and in accordance
with GST regulations. The author is not responsible for any misuse.

---

## 👨‍💻 Built By

**Akshay Vernekar** — Python automation for real-world business problems.

---

*If this tool saved you time, give it a ⭐ on GitHub!*
