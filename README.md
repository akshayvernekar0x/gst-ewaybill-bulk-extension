# 🧾 GST E-Way Bill Bulk Extension Tool

> Desktop automation tool for bulk extension of GST E-Way Bills on [ewaybillgst.gov.in](https://ewaybillgst.gov.in) — built for internal business use.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![Version](https://img.shields.io/badge/Version-5.0-orange)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Use](https://img.shields.io/badge/Use-Internal%20Business-green)

---

## 📌 What It Does

Extending E-Way Bills one by one on the GST portal is time-consuming and error-prone.  
This tool automates the entire process in bulk — select your Excel file, log in once, and the tool handles the rest.

- ✅ Reads E-Way Bill numbers and vehicle numbers from Excel / CSV file
- ✅ Opens Chrome browser and navigates to the GST portal automatically
- ✅ Login is always done **manually by the user** — no credentials are stored
- ✅ Extends each bill automatically with configured reason, remarks, place, pincode and distance
- ✅ Saves a progress report after **every single bill** — safe to stop anytime
- ✅ Generates a full Excel report and log file for audit tracking

---

## 🆕 What's New in v5.0

| Feature | Description |
|---|---|
| **v5-1** Smart Input Handling | Text fields filled with controlled timing for consistent and accurate data entry |
| **v5-2** Stable Delays | Structured delays between field fills and button clicks for reliable processing |
| **v5-3** Page Stability Check | Scroll verification before filling starts to ensure page is fully loaded |
| **v5-4** Per-Bill Excel Data | Current place, pincode, and distance can be set per row in Excel — falls back to defaults if columns are missing |

---

## 🛠️ Tech Stack

| Tool | Purpose |
|---|---|
| `Python 3.8+` | Core programming language |
| `Chrome Browser Automation` | Controls Chrome browser to fill and submit forms |
| `Pandas` | Excel and CSV data reading and report generation |
| `Tkinter` | Built-in GUI — file picker and license key dialog |
| `Requests` | License key verification |
| `openpyxl` | Excel report output |

---

## ⚙️ Setup

### 1. Install dependencies

```
pip install selenium pandas requests openpyxl webdriver-manager
```

### 2. Place `chromedriver.exe` in the same folder as the script

Download the matching version from [googlechromelabs.github.io/chrome-for-testing](https://googlechromelabs.github.io/chrome-for-testing/)


## 🚀 How to Run

**Option A — Run as .exe (recommended for end users)**
```
Double-click ewaybill_v5_0.exe
```

**Option B — Run as Python script**
```
python ewaybill_v5_0.py
```

1. Enter your license key when prompted
2. Select your Excel file from the file picker
3. A Chrome window opens — **log in manually**, complete CAPTCHA and OTP
4. Tool detects login and starts processing all bills automatically
5. Report is saved after every bill in the same folder

---

## 📋 Excel File Format

Your input Excel file should contain at minimum:

| Accepted EWB Column Names | Accepted Vehicle Column Names |
|---|---|
| `EWAY BILL NO`, `EWB_NO`, `EWB NO`, `EWAYBILLNO`, `EWAY_BILL_NO` | `TRUCK NO`, `VEHICLE NO`, `VEHICLE`, `TRUCK`, `VEH NO` |

**Optional per-bill columns (v5.0):**

| Column | Purpose | Default if missing |
|---|---|---|
| `CURRENT PLACE` | Place of current location | Banglore |
| `PINCODE` | Current pincode | 12345 |
| `DISTANCE` | Approx distance in km | 100 |

Headers can be on any of the first 5 rows — auto-detected.

---

## 📊 Output Report

An Excel report (`extension_report_YYYYMMDD_HHMMSS.xlsx`) is generated and updated after every bill:

| EWB No | Truck No | Status | Message | Time |
|---|---|---|---|---|
| 123456789012 | KA01AB1234 | SUCCESS | Extended successfully | 14:32:01 |
| 234567890123 | MH02CD5678 | FAILED | Portal returned error | 14:32:18 |

---

## 🔐 License Key

The license key is used as a **controlled release mechanism** for internal deployment:

- Ensures users operate on the latest stable version
- Enables controlled updates during initial deployment phase
- Planned to be removed or simplified after full stabilization

---

## 📁 Project Structure

```
ewaybill_v5_0.py           # Main script (v5.0)
hook-selenium.py           # Build hook for packaging
chromedriver.exe           # Place here (not included)
ewaybill.log      # Auto-generated log file
extension_report_*.xlsx    # Auto-generated report files
```

---

## ⚠️ Disclaimer

This tool is built for legitimate internal business use to save time on repetitive GST portal tasks.  
Use responsibly and in accordance with applicable GST regulations.  
The author is not responsible for any misuse.

---

## 👨‍💻 Built By

**Akshay Vernekar** — automation solutions for real-world business problems.  

---

*If this tool helped you, give it a ⭐ on GitHub!*
