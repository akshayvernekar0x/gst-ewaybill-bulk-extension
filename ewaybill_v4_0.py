"""
GST E-Way Bill Bulk Extension - Browser Automation
Works with: ewaybillgst.gov.in
Built by Akshay
Version 4.0

CHANGES FROM v3.1:
  [FIX 1] Alert popup detection  - Any portal alert (8hr rule, expired, not found etc.)
                                    is caught, dismissed, and logged as FAILED with the
                                    exact message the portal showed.
  [FIX 2] Partial report saving  - Report is saved after EVERY bill, not just at the end.
                                    Stop the script anytime and the report is still there.
  [FIX 3] Manual login           - Username/password auto-fill removed. You type manually.
                                    Added 5-second wait so CSRF token loads → no more
                                    "Invalid CSRF token" on first attempt.
  [FIX 4] License key system     - On launch, user must enter a valid license key.
                                    Keys are checked against your Google Sheet remotely.
                                    You control access & expiry from anywhere.
"""

# ─────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────
import io
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

# ─────────────────────────────────────────────
# ★  LICENSE CONFIG  ★
# Replace YOUR_SHEET_ID with the actual ID from your Google Sheet URL.
# The sheet must be published to web as CSV (File → Share → Publish to web → CSV).
# Sheet columns: LicenseKey | ExpiryDate (format: DD-Mon-YYYY e.g. 30-Jun-2026)
# ─────────────────────────────────────────────
LICENSE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv"
)

# ─────────────────────────────────────────────
# LOGGING (Windows safe)
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler("ewaybill_selenium.log", encoding="utf-8"),
    ] + (
        [logging.StreamHandler(
            io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        )] if sys.stdout and hasattr(sys.stdout, "buffer") else []
    ),
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIG  (only non-login settings remain here)
# ─────────────────────────────────────────────
CONFIG = {
    "reason":          "Others",
    "remarks":         "Others",
    "current_place":   "Bangalore",
    "current_pincode": "12345",
    "approx_distance": "6",
}

LOGIN_URL  = "https://ewaybillgst.gov.in/Login.aspx"
EXTEND_URL = "https://ewaybillgst.gov.in/BillGeneration/EwbExtension.aspx"

EWB_COLS   = ["EWAY BILL NO", "EWB_NO", "EWB NO", "EWAYBILLNO", "EWAY_BILL_NO", "ewb_no"]
TRUCK_COLS = ["TRUCK NO", "TRUCK_NO", "VEHICLE NO", "VEHICLE_NO", "VEH NO",
              "VEH_NO", "VEHICLE", "TRUCK", "truck no", "vehicle no"]

# Global report filename — set once at startup so partial saves overwrite same file
REPORT_FILE = ""

# ─────────────────────────────────────────────
# FIX 4 — LICENSE KEY CHECK
# ─────────────────────────────────────────────
def check_license():
    """
    Shows a popup asking for the license key.
    Fetches valid keys + expiry dates from your Google Sheet.
    Returns True if valid, False if invalid/expired/no internet.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    key = simpledialog.askstring(
        "GST Automator — License",
        "Enter your License Key:\n",
        parent=root,
    )
    root.destroy()

    if not key:
        messagebox.showerror("License Error", "No license key entered. Exiting.")
        return False

    key = key.strip().upper()

    print("  Verifying license key...")
    try:
        resp = requests.get(LICENSE_SHEET_URL, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        messagebox.showerror(
            "License Error",
            f"Could not reach license server.\nCheck your internet connection.\n\n{e}",
        )
        return False

    # Parse CSV — expects columns: LicenseKey, ExpiryDate
    from io import StringIO
    df = pd.read_csv(StringIO(resp.text), dtype=str)
    df.columns = [c.strip().upper() for c in df.columns]

    # Find the key
    key_col = next((c for c in df.columns if "LICENSE" in c or "KEY" in c), None)
    exp_col = next((c for c in df.columns if "EXPIRY" in c or "DATE" in c), None)

    if key_col is None:
        messagebox.showerror("License Error", "License sheet format incorrect. Contact Akshay.")
        return False

    df[key_col] = df[key_col].str.strip().str.upper()
    match = df[df[key_col] == key]

    if match.empty:
        messagebox.showerror("License Error", "Invalid license key.\nContact Akshay to get access.")
        return False

    # Check expiry if column exists
    if exp_col:
        expiry_str = str(match.iloc[0][exp_col]).strip()
        try:
            expiry_date = datetime.strptime(expiry_str, "%d-%b-%Y")
            if datetime.now() > expiry_date:
                messagebox.showerror(
                    "License Expired",
                    f"Your license expired on {expiry_str}.\nContact Akshay to renew.",
                )
                return False
            days_left = (expiry_date - datetime.now()).days
            print(f"  ✓ License valid. Expires: {expiry_str}  ({days_left} days left)")
        except Exception:
            pass  # If date format is wrong, skip expiry check silently

    print("  ✓ License verified successfully!")
    return True


# ─────────────────────────────────────────────
# STARTUP BANNER
# ─────────────────────────────────────────────
def print_banner():
    print("")
    print("=" * 60)
    print("   ███████  ██     ██  ██████                            ")
    print("   ██       ██     ██  ██   ██                           ")
    print("   █████    ██  █  ██  ██████                            ")
    print("   ██       ██ ███ ██  ██   ██                           ")
    print("   ███████   ███ ███   ██████                            ")
    print("")
    print("      GST E-Way Bill Bulk Extension Tool")
    print("               Built by Akshay")
    print("               Version  4.0")
    print("=" * 60)
    print("")


# ─────────────────────────────────────────────
# GUI — FILE PICKER
# ─────────────────────────────────────────────
def get_file_path():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select your E-Way Bill Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")],
    )
    root.destroy()
    return file_path


# ─────────────────────────────────────────────
# READ EXCEL
# ─────────────────────────────────────────────
def load_bills(filepath):
    path = Path(filepath)
    if not path.exists():
        log.error(f"File not found: {filepath}")
        sys.exit(1)

    log.info(f"Reading file: {filepath}")
    df = col_ewb = col_truck = None

    for header_row in range(5):
        _df = (
            pd.read_excel(filepath, dtype=str, header=header_row)
            if path.suffix.lower() in (".xlsx", ".xls")
            else pd.read_csv(filepath, dtype=str, header=header_row)
        )
        _df.columns = [str(c).strip() for c in _df.columns]
        _ce = next((c for c in EWB_COLS if c in _df.columns), None)
        _ct = next((c for c in TRUCK_COLS if c in _df.columns), None)
        if _ce:
            col_ewb = _ce
            col_truck = _ct
            df = _df
            log.info(
                f"Found EWB column: '{col_ewb}' | Truck column: '{col_truck}' "
                f"at header row {header_row}"
            )
            break

    if df is None or col_ewb is None:
        log.error("Could not find EWAY BILL NO column. Please name it: EWAY BILL NO")
        sys.exit(1)

    bills = []
    for _, row in df.iterrows():
        ewb = str(row.get(col_ewb, "") or "").strip()
        ewb = "".join(c for c in ewb if c.isdigit())
        if len(ewb) < 10:
            continue
        truck = ""
        if col_truck:
            truck = str(row.get(col_truck, "") or "").strip()
        bills.append({"ewb": ewb, "truck": truck})

    seen = {}
    unique = []
    for b in bills:
        if b["ewb"] not in seen:
            seen[b["ewb"]] = True
            unique.append(b)

    log.info(f"Loaded {len(unique)} E-Way Bills.")
    for b in unique:
        log.info(f"  EWB: {b['ewb']}  |  Truck: {b['truck'] or '(not found)'}")
    return unique


# ─────────────────────────────────────────────
# BROWSER SETUP
# ─────────────────────────────────────────────
def get_chromedriver_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "chromedriver.exe")


def create_driver(headless=False):
    opts = Options()
    opts.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    opts.add_argument("--start-maximized")
    opts.add_argument("--remote-allow-origins=*")

    driver_path = get_chromedriver_path()
    log.info(f"Using chromedriver at: {driver_path}")

    if not os.path.exists(driver_path):
        log.error(f"chromedriver.exe NOT FOUND at: {driver_path}")
        raise FileNotFoundError(f"chromedriver.exe not found at {driver_path}")

    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=opts)
    return driver


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def wait_click(driver, by, sel, timeout=10):
    try:
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, sel)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        return False


def screenshot(driver, name):
    try:
        driver.save_screenshot(f"debug_{name}.png")
    except Exception:
        pass


def dump_fields(driver, label):
    log.info(f"=== FIELD DUMP: {label} ===")
    for tag in ["input", "select"]:
        for el in driver.find_elements(By.TAG_NAME, tag):
            try:
                if not el.is_displayed():
                    continue
                fid   = el.get_attribute("id") or ""
                ftype = el.get_attribute("type") or ""
                fph   = el.get_attribute("placeholder") or ""
                fro   = el.get_attribute("readonly") or ""
                log.info(
                    f"  {tag.upper():6} id={fid!r:40} type={ftype!r:10} "
                    f"ph={fph!r:20} readonly={fro!r}"
                )
            except Exception:
                pass
    log.info("=== END FIELD DUMP ===")


def js_fill(driver, el, value):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)
    driver.execute_script("arguments[0].value = '';", el)
    driver.execute_script("arguments[0].value = arguments[1];", el, value)
    driver.execute_script(
        """
        arguments[0].dispatchEvent(new Event('input',  {bubbles: true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
        arguments[0].dispatchEvent(new Event('blur',   {bubbles: true}));
        """,
        el,
    )
    time.sleep(0.3)


def find_visible_input(driver, keywords):
    for tag in ["input", "select"]:
        for el in driver.find_elements(By.TAG_NAME, tag):
            try:
                if not el.is_displayed():
                    continue
                if el.get_attribute("readonly") or el.get_attribute("disabled"):
                    continue
                fid   = (el.get_attribute("id")          or "").lower()
                fname = (el.get_attribute("name")        or "").lower()
                fph   = (el.get_attribute("placeholder") or "").lower()
                combo = fid + " " + fname + " " + fph
                if any(k.lower() in combo for k in keywords):
                    return el
            except Exception:
                continue
    return None


# ─────────────────────────────────────────────
# FIX 3 — LOGIN (manual, CSRF-safe)
# ─────────────────────────────────────────────
def do_login(driver):
    log.info("Opening login page...")
    driver.get(LOGIN_URL)

    # ── FIX 3: Wait 5 seconds so CSRF token fully loads before user touches anything
    # This is what caused "Invalid CSRF token" on first attempt.
    time.sleep(5)

    # ── Username/password auto-fill REMOVED (FIX 3)
    # You type everything manually — no more auto-fill conflicts.

    print("")
    print("=" * 60)
    print("  DO THIS NOW IN THE CHROME WINDOW:")
    print("  1. Type your Username")
    print("  2. Type your Password")
    print("  3. Type the CAPTCHA")
    print("  4. Click LOGIN")
    print("  5. Enter OTP from your phone")
    print("  6. Submit OTP")
    print("  Script waits 5 minutes for you...")
    print("=" * 60)
    print("")

    deadline = time.time() + 300
    while time.time() < deadline:
        try:
            url = driver.current_url.lower()
            if "mainmenu" in url or ("main" in url and "login" not in url):
                log.info("[OK] Login complete!")
                time.sleep(2)
                return True
        except Exception:
            pass
        left = int(deadline - time.time())
        if left % 30 == 0 and left > 0:
            print(f"  Waiting... {left}s left. Complete login + OTP in the browser!")
        time.sleep(1)

    log.error("Login timeout.")
    return False


# ─────────────────────────────────────────────
# FIX 2 — SAVE REPORT (consistent filename)
# ─────────────────────────────────────────────
def save_report(results):
    """
    Saves report to the same file every time (REPORT_FILE set at startup).
    Called after every single bill so partial runs are never lost.
    """
    global REPORT_FILE
    rows = [
        {
            "EWB No":   r["ewb_no"],
            "Truck No": r["truck"],
            "Status":   "SUCCESS" if r["success"] else "FAILED",
            "Message":  r["message"],
            "Time":     r.get("time", ""),
        }
        for r in results
    ]
    df = pd.DataFrame(rows)
    df.to_excel(REPORT_FILE, index=False)
    log.info(f"Report updated: {REPORT_FILE}  ({len(results)} bills so far)")


# ─────────────────────────────────────────────
# EXTEND ONE BILL
# ─────────────────────────────────────────────
def extend_one(driver, bill, num, total, debug_mode):
    ewb_no   = bill["ewb"]
    truck_no = bill["truck"]
    result   = {
        "ewb_no":  ewb_no,
        "truck":   truck_no,
        "success": False,
        "message": "",
        "time":    datetime.now().strftime("%H:%M:%S"),
    }

    try:
        log.info(f"  [{num}/{total}] EWB: {ewb_no}  Truck: {truck_no}")

        driver.get(EXTEND_URL)

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_ContentPlaceHolder1_txt_no")
                )
            )
        except Exception:
            time.sleep(3)

        # ── Enter EWB number ───────────────────────────────────
        ewb_field = find_visible_input(driver, ["txt_no", "txtewbno", "ewbno", "ewb_no"])
        if not ewb_field:
            inputs = [
                i for i in driver.find_elements(By.XPATH, "//input[@type='text']")
                if i.is_displayed()
            ]
            ewb_field = inputs[0] if inputs else None
        if not ewb_field:
            result["message"] = "Could not find EWB input"
            return result

        js_fill(driver, ewb_field, ewb_no)

        # ── Click Go ───────────────────────────────────────────
        log.info(f"  [{num}/{total}] Clicking Go...")
        go_clicked = False
        for by, sel in [
            (By.ID,    "ctl00_ContentPlaceHolder1_Btn_go"),
            (By.XPATH, "//input[@value='Go' or @value='GO']"),
            (By.XPATH, "//button[text()='Go']"),
        ]:
            if wait_click(driver, by, sel, timeout=8):
                go_clicked = True
                break
        if not go_clicked:
            result["message"] = "Could not click Go"
            return result

        # ── FIX 1: Check for alert popup IMMEDIATELY after Go ─────────────────
        # The portal shows a JS alert for errors like:
        #   "extension is allowed only 8 hours before expiry"
        #   "EWB is already expired"
        #   "EWB not found"
        #   "You are not authorized to extend this EWB"
        #   ...and any other alert the portal may show.
        # We wait up to 3 seconds for any alert to appear.
        # If found → capture exact message → dismiss → mark FAILED → skip to next bill.
        time.sleep(2)
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text.strip()
            alert.accept()  # clicks OK
            result["message"] = f"Portal blocked: {alert_text}"
            log.warning(
                f"  [{num}/{total}] BLOCKED by portal alert: {alert_text}"
            )
            return result  # ← skip to next bill immediately, no time wasted
        except Exception:
            pass  # No alert appeared — continue normally
        # ─────────────────────────────────────────────────────────────────────

        # ── Wait for Yes radio to appear ───────────────────────
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "rbn_extent_0"))
            )
        except Exception:
            time.sleep(2)

        # ── Click YES radio ────────────────────────────────────
        log.info(f"  [{num}/{total}] Clicking Yes...")
        yes_done = False
        for by, sel in [
            (By.ID,    "rbn_extent_0"),
            (By.XPATH, "//input[@name='ctl00$ContentPlaceHolder1$rbn_extent'][@type='radio'][1]"),
            (By.XPATH, "//input[@type='radio'][1]"),
        ]:
            if wait_click(driver, by, sel, timeout=8):
                yes_done = True
                log.info(f"    Yes radio clicked (sel={sel})")
                break
        if not yes_done:
            log.info("    Yes radio not clicked - trying JS fallback")
            try:
                r = driver.find_element(By.ID, "rbn_extent_0")
                driver.execute_script("arguments[0].click();", r)
                yes_done = True
            except Exception:
                pass

        time.sleep(2)

        if debug_mode and num == 1:
            dump_fields(driver, "AFTER_YES_CLICKED")
            screenshot(driver, f"after_yes_{ewb_no}")

        # ── Reason ────────────────────────────────────────────
        try:
            reason_el = None
            for rid in ["ddl_extend", "ddlExtend", "ddl_Extend"]:
                try:
                    reason_el = driver.find_element(By.ID, rid)
                    break
                except Exception:
                    pass
            if not reason_el:
                reason_el = find_visible_input(driver, ["extend", "reason", "rsn"])
            if reason_el and reason_el.tag_name == "select":
                s = Select(reason_el)
                for opt in s.options:
                    if CONFIG["reason"].lower() in opt.text.lower():
                        s.select_by_visible_text(opt.text)
                        log.info(f"    Reason set: {opt.text}")
                        break
        except Exception as e:
            log.info(f"    Reason error: {e}")

        # ── Remarks ───────────────────────────────────────────
        try:
            rem = find_visible_input(driver, ["remark", "remarks", "txtremarks", "txtremark"])
            if rem:
                js_fill(driver, rem, CONFIG["remarks"])
                log.info(f"    Remarks filled (id={rem.get_attribute('id')})")
        except Exception as e:
            log.info(f"    Remarks error: {e}")

        # ── Current Place ─────────────────────────────────────
        try:
            place = None
            try:
                place = driver.find_element(
                    By.XPATH, "//input[@placeholder='Current Place']"
                )
            except Exception:
                pass
            if not place:
                place = find_visible_input(
                    driver, ["curplace", "currplace", "currentplace", "cur_place"]
                )
            if not place:
                try:
                    place = driver.find_element(
                        By.XPATH,
                        "//td[normalize-space(text())='Current Place']/following-sibling::td//input",
                    )
                except Exception:
                    pass
            if place and place.is_displayed():
                js_fill(driver, place, CONFIG["current_place"])
                log.info(f"    Current Place filled (id={place.get_attribute('id')})")
            else:
                log.info("    Current Place: field NOT found")
        except Exception as e:
            log.info(f"    Current Place error: {e}")

        # ── Pincode → state auto-fill ─────────────────────────
        try:
            pin = None
            for pid in ["txtFromEnteredPinCode", "txtFromEnteredPin", "txtEnteredPinCode"]:
                try:
                    pin = driver.find_element(By.ID, pid)
                    break
                except Exception:
                    pass
            if not pin:
                pin = find_visible_input(
                    driver, ["EnteredPin", "enteredpin", "curpin", "currpin"]
                )
            if pin and pin.is_displayed():
                js_fill(driver, pin, CONFIG["current_pincode"])
                log.info(f"    Pincode filled (id={pin.get_attribute('id')})")
                from selenium.webdriver.common.keys import Keys
                pin.send_keys(Keys.RETURN)
                log.info("    Pincode Enter pressed - waiting for state to load...")
                time.sleep(3)
                try:
                    state_dd = driver.find_element(By.ID, "drp_from")
                    sel = Select(state_dd)
                    loaded = sel.first_selected_option.text.strip()
                    log.info(f"    State auto-loaded: {loaded}")
                except Exception:
                    pass
            else:
                log.info("    Current Pincode: field NOT found")
        except Exception as e:
            log.info(f"    Pincode error: {e}")

        # ── Distance ──────────────────────────────────────────
        def fill_distance():
            try:
                d = driver.find_element(By.ID, "txtDistance")
                if d.is_displayed():
                    for attempt in range(3):
                        js_fill(driver, d, CONFIG["approx_distance"])
                        time.sleep(0.5)
                        actual = d.get_attribute("value").strip()
                        if actual == CONFIG["approx_distance"]:
                            log.info(
                                f"    Distance: {actual} km (confirmed on attempt {attempt+1})"
                            )
                            return True
                        log.info(f"    Distance attempt {attempt+1}: got '{actual}' — retrying")
                        time.sleep(0.5)
                    log.info("    Distance still wrong after 3 attempts")
            except Exception as e:
                log.info(f"    Distance error: {e}")
            return False

        fill_distance()

        # ── Vehicle No ────────────────────────────────────────
        try:
            veh = None
            try:
                veh = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtVehicleNo")
            except Exception:
                pass
            if not veh:
                veh = find_visible_input(driver, ["vehicle", "vehicleno", "vehno", "veh_no"])
            if veh and veh.is_displayed():
                js_fill(driver, veh, truck_no)
                log.info(f"    Vehicle No: {truck_no} (id={veh.get_attribute('id')})")
            else:
                log.info("    Vehicle No: field NOT found")
        except Exception as e:
            log.info(f"    Vehicle error: {e}")

        # ── Final distance check ───────────────────────────────
        try:
            d = driver.find_element(By.ID, "txtDistance")
            actual = d.get_attribute("value").strip()
            if actual != CONFIG["approx_distance"]:
                log.info(
                    f"    Distance reset after vehicle — correcting to {CONFIG['approx_distance']}"
                )
                js_fill(driver, d, CONFIG["approx_distance"])
                time.sleep(0.5)
                log.info(f"    Distance final value: {d.get_attribute('value')}")
        except Exception as e:
            log.info(f"    Final distance check: {e}")

        if debug_mode:
            screenshot(driver, f"before_submit_{ewb_no}")

        # ── Submit ────────────────────────────────────────────
        log.info(f"  [{num}/{total}] Clicking Submit...")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        submitted = False
        for by, sel in [
            (By.XPATH, "//input[@value='Submit' or @value='SUBMIT']"),
            (By.XPATH, "//button[contains(text(),'Submit')]"),
            (By.ID,    "btnSubmit"),
            (By.ID,    "btnsubmit"),
            (By.ID,    "Btnsubmit"),
            (By.XPATH, "//input[@type='submit']"),
        ]:
            try:
                el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((by, sel)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                submitted = True
                log.info("    Submit clicked")
                break
            except Exception:
                continue

        if not submitted:
            result["message"] = "Submit button not found"
            screenshot(driver, f"err_nosubmit_{ewb_no}")
            return result

        # ── Wait for portal result ─────────────────────────────
        SUCCESS_WORDS = ["extended", "valid till", "validity", "successfully"]
        FAILURE_WORDS = ["error", "invalid", "cannot extend", "not found"]

        time.sleep(3)
        try:
            WebDriverWait(driver, 30).until(
                lambda d: (
                    any(
                        w in d.find_element(By.TAG_NAME, "body").text.lower()
                        for w in SUCCESS_WORDS + FAILURE_WORDS
                    )
                    or len(
                        d.find_elements(
                            By.XPATH, "//input[@value='Exit' or @value='EXIT']"
                        )
                    ) > 0
                )
            )
        except Exception:
            time.sleep(5)

        page = driver.find_element(By.TAG_NAME, "body").text.lower()
        if any(w in page for w in SUCCESS_WORDS):
            result["success"] = True
            result["message"] = "Extended successfully"
            log.info(f"  [{num}/{total}] SUCCESS - Extended!")
        elif any(w in page for w in FAILURE_WORDS):
            result["message"] = "Portal returned error - check EWB on portal"
            log.warning(f"  [{num}/{total}] Portal error for {ewb_no}")
        else:
            result["success"] = True
            result["message"] = "Submitted (verify on portal)"
            log.info(f"  [{num}/{total}] Submitted")

    except Exception as e:
        result["message"] = f"Error: {str(e)[:150]}"
        screenshot(driver, f"err_{ewb_no}")
        log.warning(f"  [{num}/{total}] ERROR: {result['message']}")

    return result


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    global REPORT_FILE

    # 1. Banner
    print_banner()

    # 2. FIX 4 — License check
    if not check_license():
        time.sleep(3)
        return

    # 3. File picker
    file_path = get_file_path()
    if not file_path:
        print("No file selected. Exiting.")
        time.sleep(2)
        return

    # 4. Load bills
    bills = load_bills(file_path)

    # 5. Set a single report filename for this entire run (FIX 2)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    REPORT_FILE = f"extension_report_{ts}.xlsx"
    log.info(f"Report file for this run: {REPORT_FILE}")

    # 6. Launch browser
    log.info("Launching Chrome...")
    driver = create_driver(headless=False)

    results = []
    try:
        if not do_login(driver):
            log.error("Login failed or timed out.")
            return

        log.info(f"Starting extension of {len(bills)} bills...")
        for i, bill in enumerate(bills, 1):
            result = extend_one(driver, bill, i, len(bills), debug_mode=(i == 1))
            results.append(result)

            # ── FIX 2: Save after EVERY bill ──────────────────
            # Even if you stop the script midway, the report is up to date.
            save_report(results)

    except Exception as e:
        log.error(f"Critical Error: {e}")

    finally:
        # ── FIX 2: Save in finally block too (catches crashes/force-stops)
        if results:
            save_report(results)
        try:
            driver.quit()
        except Exception:
            pass
        log.info("Browser closed.")

    # 7. Final summary
    ok   = sum(1 for r in results if r["success"])
    fail = len(results) - ok

    print("")
    print("=" * 60)
    print(f"   SUCCESS : {ok}")
    print(f"   FAILED  : {fail}")
    print(f"   TOTAL   : {len(results)}")
    print(f"   REPORT  : {REPORT_FILE}")
    print("=" * 60)
    print("           Developed by Akshay")
    print("=" * 60)

    print("\nProcess Complete. This window will close in 10 seconds...")
    time.sleep(10)


if __name__ == "__main__":
    main()
