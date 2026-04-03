# 🧾 GST E-Way Bill Bulk Extension Tool

> Automate bulk extension of GST E-Way Bills on [ewaybillgst.gov.in](https://ewaybillgst.gov.in) — built with Python + Selenium.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-Browser%20Automation-43B02A?logo=selenium&logoColor=white)
![Version](https://img.shields.io/badge/Version-4.0-orange)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)

---

## 📌 What It Does

Extending E-Way Bills one by one on the GST portal is painfully slow. This tool automates the entire process in bulk — you just upload your Excel file, log in once, and the script takes care of the rest.

- ✅ Reads EWB numbers and vehicle numbers from your Excel/CSV file
- ✅ Logs into the portal — manually (no credential storage)
- ✅ Extends each bill automatically with configured reason, remarks, place, pincode & distance
- ✅ Catches any portal alert (expired, 8-hour rule, not found) and logs it with the exact message
- ✅ Saves a progress report after **every single bill** — stop anytime without losing data
- ✅ License key system — control access remotely via Google Sheets

---

## 🆕 What's New in v4.0

| Fix | Description |
|-----|-------------|
| **FIX 1** Alert Detection | Portal alerts (8hr rule, expired, not found, etc.) are caught, dismissed, and logged as `FAILED` with the exact portal message |
| **FIX 2** Partial Report Saving | Report is saved after every bill — stop the script anytime and your data is safe |
| **FIX 3** Manual Login | Username/password auto-fill removed. You type manually. Added 5-second wait so CSRF token loads fully — eliminates *"Invalid CSRF token"* errors |
| **FIX 4** License Key System | On launch, user must enter a valid license key verified remotely against a Google Sheet — you control access and expiry from anywhere |

---

## 🛠️ Tech Stack

| Tool | Purpose |
|------|---------|
| `Python 3.8+` | Core language |
| `Selenium` | Browser automation (Chrome) |
| `Pandas` | Excel/CSV reading and report generation |
| `Tkinter` | GUI dialogs (file picker, license key input) |
| `Requests` | License verification via Google Sheets CSV |
| `openpyxl` | Excel report output |

---

## ⚙️ Setup

### 1. Install dependencies

```bash
pip install selenium pandas requests openpyxl
```

### 2. Place `chromedriver.exe` in the same folder as the script

Download the matching version from [chromedriver.chromium.org](https://chromedriver.chromium.org/downloads).

### 3. Configure `LICENSE_SHEET_URL`

Publish your Google Sheet as CSV (File → Share → Publish to web → CSV), then paste the URL:

```python
LICENSE_SHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv"
```

Your sheet should have two columns:

| LicenseKey | ExpiryDate |
|------------|------------|
| ABCD-1234  | 30-Jun-2026 |

### 4. Configure extension settings

Edit the `CONFIG` block in the script:

```python
CONFIG = {
    "reason":          "Others",
    "remarks":         "Others",
    "current_place":   "Bangalore",
    "current_pincode": "562109",
    "approx_distance": "6",
}
```

---

## 🚀 How to Run

```bash
python ewaybill_v4_0.py
```

1. Enter your license key when prompted
2. Select your Excel file (must contain `EWAY BILL NO` and optionally `TRUCK NO` columns)
3. A Chrome window opens — **log in manually**, complete CAPTCHA and OTP
4. The script detects login and starts extending bills automatically
5. A report Excel file is saved after every bill

---

## 📋 Excel File Format

Your input file should have at minimum one of these column names:

| Accepted EWB Columns | Accepted Vehicle Columns |
|----------------------|--------------------------|
| `EWAY BILL NO`, `EWB_NO`, `EWB NO`, `EWAYBILLNO`, `EWAY_BILL_NO`, `ewb_no` | `TRUCK NO`, `VEHICLE NO`, `VEHICLE`, `TRUCK`, etc. |

Headers can be on any of the first 5 rows — the script auto-detects them.

---

## 📊 Output Report

An Excel report (`extension_report_YYYYMMDD_HHMMSS.xlsx`) is generated and updated after every bill:

| EWB No | Truck No | Status | Message | Time |
|--------|----------|--------|---------|------|
| 123456789012 | KA01AB1234 | SUCCESS | Extended successfully | 14:32:01 |
| 234567890123 | MH02CD5678 | FAILED | Portal blocked: EWB is expired | 14:32:18 |

---

## 🔐 License System

The license system lets you:
- Issue unique keys to each client
- Set an expiry date per key
- Revoke access instantly by deleting the row from your Google Sheet
- No code changes needed — it's all managed remotely

---

## 📁 Project Structure

```
ewaybill_v4_0.py          # Main script
chromedriver.exe          # Place here (not included)
ewaybill_selenium.log     # Auto-generated log file
extension_report_*.xlsx   # Auto-generated report files
debug_*.png               # Debug screenshots (first bill only)
```

---

## ⚠️ Disclaimer

This tool is built for legitimate business use to save time on repetitive GST portal tasks. Use it responsibly and in accordance with GST regulations. The author is not responsible for any misuse.

---

## 👨‍💻 Built By

**Akshay** — Python automation for real-world business problems.

---

*If this tool saved you time, give it a ⭐ on GitHub!*
