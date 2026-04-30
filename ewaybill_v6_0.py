"""
GST E-Way Bill Bulk Extension - Browser Automation
Works with: ewaybillgst.gov.in
Built by Akshay
Version 6.0

CHANGES FROM v5.0:
  [v6-1]  Anti-bot restored    - create_driver() now has full anti-detection:
                                  AutomationControlled flag disabled,
                                  useAutomationExtension off, User-Agent rotation,
                                  CDP script to hide navigator.webdriver.
  [v6-2]  Import fix           - webdriver_manager import moved to top with all
                                  other imports (was buried at line 296).
  [v6-3]  StringIO fix         - Removed redundant 'from io import StringIO' inside
                                  check_license(); now uses io.StringIO correctly.
  [v6-4]  Login sleep reduced  - time.sleep(5) at login page start reduced to 1s.
  [v6-5]  Smart alert wait     - time.sleep(2) after Go click replaced with
                                  WebDriverWait(2) for alert — exits as soon as
                                  alert appears or doesn't, no dead wait.
  [v6-6]  Smart pincode wait   - time.sleep(2) after pincode Enter replaced with
                                  WebDriverWait watching state dropdown to load.
  [v6-7]  Submit sleep reduced - time.sleep(3) before result check reduced to 1s.
                                  WebDriverWait(30) below handles the real wait.
  [v6-8]  Report save batched  - save_report() was called after every single bill
                                  (144x for 144 bills). Now saves every 5 bills
                                  and immediately on any failure.
  [v6-9]  Auto retry           - Failed bills get 1 automatic retry after 3s gap.
  [v6-10] js_fill optimised    - Two separate JS execute_script calls to set value
                                  combined into one call.
  [v6-11] Load bills quiet     - Per-bill logging in load_bills() changed from
                                  INFO to DEBUG — no more 144-line console spam
                                  on startup.
  [v6-12] Version string fixed - Banner and docstring now correctly say 6.0.

CHANGES FROM v5.1 → v5.2 (kept):
  [v5-1] Human-like typing    - Text fields typed char-by-char (10-20ms/key).
                                 1-in-10 chance of a typo + backspace correction.
                                 Applied only to long fields (EWB, place, vehicle).
                                 Short fields (pincode, distance) still use js_fill.
  [v5-2] Random delays        - Between field fills: 0.1-0.3s
                                 Before button clicks: 0.1-0.2s
                                 After page load:      0.2-1.2s
  [v5-3] Random scroll        - 50% chance of a single gentle scroll before filling.
  [v5-4] Excel per-bill data  - current_place, current_pincode, approx_distance
                                 read from Excel columns per bill.
                                 Falls back to DEFAULT_* values if columns missing.
"""

# ─────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────
import io
import logging
import os
import random
import sys
import time
import threading
import tkinter.ttk as ttk
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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.common.exceptions import NoAlertPresentException, TimeoutException

# [v6-2] Moved here from line 296 — must be with all other imports
from webdriver_manager.chrome import ChromeDriverManager


# ─────────────────────────────────────────────
# ★  LICENSE CONFIG  ★
# ─────────────────────────────────────────────
LICENSE_SHEET_URL = (
    "LICENSE_SHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/pub?output=csv"
)

# ─────────────────────────────────────────────
# LOGGING (Windows safe)
# ─────────────────────────────────────────────
_handlers = []
try:
    _log_path = os.path.join(os.path.dirname(sys.executable)
                             if getattr(sys, "frozen", False)
                             else os.path.dirname(os.path.abspath(__file__)),
                             "ewaybill_selenium.log")
    _handlers.append(logging.FileHandler(_log_path, encoding="utf-8"))
except Exception:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=_handlers if _handlers else [logging.NullHandler()],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# CONFIG
# DEFAULT_* values used only when Excel column is missing.
# ─────────────────────────────────────────────
CONFIG = {
    "reason":  "Others",
    "remarks": "Others",
}

# ── v5-4: Fallback defaults if Excel columns are absent ──
DEFAULT_PLACE    = "Banglore"
DEFAULT_PINCODE  = "123456"
DEFAULT_DISTANCE = "100"

LOGIN_URL  = "https://ewaybillgst.gov.in/Login.aspx"
EXTEND_URL = "https://ewaybillgst.gov.in/BillGeneration/EwbExtension.aspx"

EWB_COLS   = ["EWAY BILL NO", "EWB_NO", "EWB NO", "EWAYBILLNO", "EWAY_BILL_NO", "ewb_no"]
TRUCK_COLS = ["TRUCK NO", "TRUCK_NO", "VEHICLE NO", "VEHICLE_NO", "VEH NO",
              "VEH_NO", "VEHICLE", "TRUCK", "truck no", "vehicle no"]

# ── v5-4: Excel column names for per-bill fields ──
PLACE_COLS    = ["CURRENT PLACE", "CUR PLACE", "PLACE", "CURRENT_PLACE", "CURPLACE"]
PINCODE_COLS  = ["PINCODE", "PIN CODE", "PIN", "PIN_CODE", "CURRENT PINCODE", "CUR PINCODE"]
DISTANCE_COLS = ["DISTANCE", "DIST", "APPROX DISTANCE", "APPROX_DISTANCE"]

# [v6-1] User-Agent pool for rotation
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
]

# Global report filename
REPORT_FILE = ""


# ─────────────────────────────────────────────
# LICENSE KEY CHECK
# ─────────────────────────────────────────────
def check_license():
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

    # [v6-3] Use io.StringIO directly — removed redundant inner import
    df = pd.read_csv(io.StringIO(resp.text), dtype=str)
    df.columns = [c.strip().upper() for c in df.columns]

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
            pass

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
    print("               Version  6.0")            # [v6-12]
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
# READ EXCEL  (v5-4: also reads place/pincode/distance columns)
# ─────────────────────────────────────────────
def load_bills(filepath):
    path = Path(filepath)
    if not path.exists():
        log.error(f"File not found: {filepath}")
        sys.exit(1)

    log.info(f"Reading file: {filepath}")
    df = col_ewb = col_truck = None
    col_place = col_pin = col_dist = None

    for header_row in range(5):
        _df = (
            pd.read_excel(filepath, dtype=str, header=header_row)
            if path.suffix.lower() in (".xlsx", ".xls")
            else pd.read_csv(filepath, dtype=str, header=header_row)
        )
        _df.columns = [str(c).strip() for c in _df.columns]
        _ce = next((c for c in EWB_COLS if c in _df.columns), None)
        if _ce:
            col_ewb   = _ce
            col_truck = next((c for c in TRUCK_COLS   if c in _df.columns), None)
            col_place = next((c for c in PLACE_COLS   if c in _df.columns), None)
            col_pin   = next((c for c in PINCODE_COLS if c in _df.columns), None)
            col_dist  = next((c for c in DISTANCE_COLS if c in _df.columns), None)
            df = _df
            log.info(
                f"Found EWB column: '{col_ewb}' | Truck: '{col_truck}' | "
                f"Place: '{col_place}' | Pincode: '{col_pin}' | Distance: '{col_dist}' "
                f"at header row {header_row}"
            )
            break

    if df is None or col_ewb is None:
        log.error("Could not find EWAY BILL NO column. Please name it: EWAY BILL NO")
        sys.exit(1)

    # Log fallback warnings
    if not col_place:
        log.info(f"  [v5-4] No PLACE column found → using fallback: '{DEFAULT_PLACE}'")
    if not col_pin:
        log.info(f"  [v5-4] No PINCODE column found → using fallback: '{DEFAULT_PINCODE}'")
    if not col_dist:
        log.info(f"  [v5-4] No DISTANCE column found → using fallback: '{DEFAULT_DISTANCE}'")

    bills = []
    seen  = {}
    for _, row in df.iterrows():
        ewb = str(row.get(col_ewb, "") or "").strip()
        ewb = "".join(c for c in ewb if c.isdigit())
        if len(ewb) < 10:
            continue
        if ewb in seen:
            continue
        seen[ewb] = True

        truck = ""
        if col_truck:
            truck = str(row.get(col_truck, "") or "").strip()

        place = str(row.get(col_place, "") or "").strip() if col_place else ""
        if not place:
            place = DEFAULT_PLACE

        pin = str(row.get(col_pin, "") or "").strip() if col_pin else ""
        if not pin:
            pin = DEFAULT_PINCODE

        dist = str(row.get(col_dist, "") or "").strip() if col_dist else ""
        if not dist:
            dist = DEFAULT_DISTANCE

        bills.append({
            "ewb":   ewb,
            "truck": truck,
            "place": place,
            "pin":   pin,
            "dist":  dist,
        })

    log.info(f"Loaded {len(bills)} unique E-Way Bills.")

    # [v6-11] Changed from INFO to DEBUG — stops 144-line console spam on startup
    for b in bills:
        log.debug(
            f"  EWB: {b['ewb']}  Truck: {b['truck'] or '(none)'}  "
            f"Place: {b['place']}  Pin: {b['pin']}  Dist: {b['dist']}"
        )
    return bills


# ─────────────────────────────────────────────
# BROWSER SETUP
# [v6-1] Full anti-bot detection restored
# [v6-2] ChromeDriverManager import moved to top
# ─────────────────────────────────────────────
def get_chrome_binary():
    """Find Chrome exe path — needed when running as frozen EXE."""
    import winreg
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Google\Chrome\Application"
        )
        path, _ = winreg.QueryValueEx(key, "")
        if os.path.exists(path):
            return path
    except Exception:
        pass
    common = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        rf"C:\Users\{os.getenv('USERNAME')}\AppData\Local\Google\Chrome\Application\chrome.exe",
    ]
    for p in common:
        if os.path.exists(p):
            return p
    return None
def create_driver(headless=False):
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--remote-allow-origins=*")
    chrome_bin = get_chrome_binary()
    if chrome_bin:
        opts.binary_location = chrome_bin
        log.info(f"Chrome binary: {chrome_bin}")
    else:
        log.warning("Chrome binary not found — using default")

    # [v6-1] Hide all automation signals
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    # [v6-1] Rotate User-Agent — looks like a real browser session
    ua = random.choice(USER_AGENTS)
    opts.add_argument(f"user-agent={ua}")
    log.info(f"User-Agent selected: {ua}")

    # Auto-downloads correct ChromeDriver for installed Chrome version
    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=opts)

    # [v6-1] CDP script — hides navigator.webdriver and other bot signals
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins',   { get: () => [1, 2, 3, 4, 5] });
            Object.defineProperty(navigator, 'languages', { get: () => ['en-IN', 'en', 'hi'] });
            window.chrome = { runtime: {} };
            Object.defineProperty(navigator, 'platform',  { get: () => 'Win32' });
        """
    })

    log.info("Chrome started with anti-bot protection.")
    return driver


# ─────────────────────────────────────────────
# HELPERS — timing
# ─────────────────────────────────────────────
def rand_sleep(lo, hi):
    """Sleep for a random duration between lo and hi seconds."""
    time.sleep(random.uniform(lo, hi))


def pre_click():
    """Short random pause before any button click — v5-2."""
    rand_sleep(0.1, 0.2)


def post_load():
    """Random pause after a page/section loads — v5-2."""
    rand_sleep(0.2, 1.2)


def between_fields():
    """Random pause between filling consecutive fields — v5-2."""
    rand_sleep(0.1, 0.3)


# ─────────────────────────────────────────────
# HELPERS — interaction
# ─────────────────────────────────────────────
def wait_click(driver, by, sel, timeout=10):
    try:
        el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, sel)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        pre_click()
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
                fid   = el.get_attribute("id")          or ""
                ftype = el.get_attribute("type")         or ""
                fph   = el.get_attribute("placeholder")  or ""
                fro   = el.get_attribute("readonly")      or ""
                log.info(
                    f"  {tag.upper():6} id={fid!r:40} type={ftype!r:10} "
                    f"ph={fph!r:20} readonly={fro!r}"
                )
            except Exception:
                pass
    log.info("=== END FIELD DUMP ===")


def js_fill(driver, el, value):
    """
    Fast JS fill — used for short/numeric fields (pincode, distance).
    [v6-10] Combined two separate value-set JS calls into one.
    """
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.15)
    # [v6-10] Single JS call sets value and fires all events
    driver.execute_script(
        """
        arguments[0].value = '';
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input',  {bubbles: true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
        arguments[0].dispatchEvent(new Event('blur',   {bubbles: true}));
        """,
        el, value,
    )
    time.sleep(0.2)


def human_type(driver, el, text):
    """
    v5-1: Type text character-by-character with human-like delays.
    - 10-20ms per keystroke
    - 1-in-10 chance of a typo followed by backspace correction
    Applied to long text fields only (EWB number, place, vehicle no).
    """
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.15)
    el.clear()
    time.sleep(0.1)

    for ch in text:
        # 1-in-10 typo chance
        if random.randint(1, 10) == 1:
            typo_char = random.choice("qwertyuiopasdfghjklzxcvbnm1234567890")
            el.send_keys(typo_char)
            time.sleep(random.uniform(0.04, 0.09))
            el.send_keys(Keys.BACK_SPACE)
            time.sleep(random.uniform(0.03, 0.07))

        el.send_keys(ch)
        time.sleep(random.uniform(0.01, 0.02))   # 10–20 ms per key

    driver.execute_script(
        """
        arguments[0].dispatchEvent(new Event('input',  {bubbles: true}));
        arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
        arguments[0].dispatchEvent(new Event('blur',   {bubbles: true}));
        """,
        el,
    )
    time.sleep(0.2)


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


def maybe_scroll(driver):
    """
    v5-3: 50% chance of a single gentle scroll before filling starts.
    Fast and non-disruptive.
    """
    if random.random() < 0.5:
        scroll_y = random.randint(80, 220)
        driver.execute_script(f"window.scrollBy(0, {scroll_y});")
        time.sleep(0.2)
        driver.execute_script(f"window.scrollBy(0, -{scroll_y});")
        time.sleep(0.15)


# ─────────────────────────────────────────────
# LOGIN (manual, CSRF-safe)
# ─────────────────────────────────────────────
def do_login(driver):
    log.info("Opening login page...")
    driver.get(LOGIN_URL)
    time.sleep(1)   # [v6-4] Reduced from 5s — page just needs to start loading

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
                time.sleep(1)   # brief settle after login
                # Hide Chrome window after login is complete
                driver.execute_script("document.body.style.display='none';")
                driver.set_window_position(-10000, 0)   # moves window off screen
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
# SAVE REPORT
# ─────────────────────────────────────────────
def save_report(results):
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
    place    = bill["place"]
    pin      = bill["pin"]
    dist     = bill["dist"]

    result = {
        "ewb_no":  ewb_no,
        "truck":   truck_no,
        "success": False,
        "message": "",
        "time":    datetime.now().strftime("%H:%M:%S"),
    }

    try:
        log.info(f"  [{num}/{total}] EWB: {ewb_no}  Truck: {truck_no}  Place: {place}  Pin: {pin}  Dist: {dist}")

        driver.get(EXTEND_URL)
        post_load()   # v5-2: random wait after page load

        # v5-3: random scroll before starting
        maybe_scroll(driver)

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_ContentPlaceHolder1_txt_no")
                )
            )
        except Exception:
            time.sleep(2)

        # ── Enter EWB number (human_type) ─────────────────────
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

        human_type(driver, ewb_field, ewb_no)   # v5-1
        between_fields()                          # v5-2

        # ── Click Go ──────────────────────────────────────────
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

        # ── [v6-5] Smart alert check after Go ─────────────────
        # Replaces time.sleep(2) — exits as soon as alert appears or doesn't
        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            alert     = driver.switch_to.alert
            alert_text = alert.text.strip()
            alert.accept()
            result["message"] = f"Portal blocked: {alert_text}"
            log.warning(f"  [{num}/{total}] BLOCKED by portal alert: {alert_text}")
            return result
        except (NoAlertPresentException, TimeoutException):
            pass   # No alert — all good, continue

        # ── Wait for Yes radio ────────────────────────────────
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "rbn_extent_0"))
            )
        except Exception:
            time.sleep(2)

        # ── Click YES radio ───────────────────────────────────
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
            try:
                r = driver.find_element(By.ID, "rbn_extent_0")
                driver.execute_script("arguments[0].click();", r)
                yes_done = True
            except Exception:
                pass

        between_fields()   # v5-2

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

        between_fields()   # v5-2

        # ── Remarks ───────────────────────────────────────────
        try:
            rem = find_visible_input(driver, ["remark", "remarks", "txtremarks", "txtremark"])
            if rem:
                js_fill(driver, rem, CONFIG["remarks"])
                log.info(f"    Remarks filled (id={rem.get_attribute('id')})")
        except Exception as e:
            log.info(f"    Remarks error: {e}")

        between_fields()   # v5-2

        # ── Current Place (human_type, per-bill from Excel) ───
        try:
            place_el = None
            try:
                place_el = driver.find_element(
                    By.XPATH, "//input[@placeholder='Current Place']"
                )
            except Exception:
                pass
            if not place_el:
                place_el = find_visible_input(
                    driver, ["curplace", "currplace", "currentplace", "cur_place"]
                )
            if not place_el:
                try:
                    place_el = driver.find_element(
                        By.XPATH,
                        "//td[normalize-space(text())='Current Place']/following-sibling::td//input",
                    )
                except Exception:
                    pass
            if place_el and place_el.is_displayed():
                human_type(driver, place_el, place)   # v5-1 + v5-4
                log.info(f"    Current Place: {place} (id={place_el.get_attribute('id')})")
            else:
                log.info("    Current Place: field NOT found")
        except Exception as e:
            log.info(f"    Current Place error: {e}")

        between_fields()   # v5-2

        # ── Pincode (fast js_fill, per-bill from Excel) ───────
        try:
            pin_el = None
            for pid in ["txtFromEnteredPinCode", "txtFromEnteredPin", "txtEnteredPinCode"]:
                try:
                    pin_el = driver.find_element(By.ID, pid)
                    break
                except Exception:
                    pass
            if not pin_el:
                pin_el = find_visible_input(
                    driver, ["EnteredPin", "enteredpin", "curpin", "currpin"]
                )
            if pin_el and pin_el.is_displayed():
                js_fill(driver, pin_el, pin)
                log.info(f"    Pincode: {pin} (id={pin_el.get_attribute('id')})")
                pin_el.send_keys(Keys.RETURN)
                log.info("    Pincode Enter pressed - waiting for state to load...")

                # [v6-6] Smart wait — watches state dropdown instead of fixed sleep
                try:
                    WebDriverWait(driver, 5).until(
                        lambda d: (
                            Select(d.find_element(By.ID, "drp_from"))
                            .first_selected_option.text.strip() not in ("", "Select State")
                        )
                    )
                    state_dd = driver.find_element(By.ID, "drp_from")
                    loaded   = Select(state_dd).first_selected_option.text.strip()
                    log.info(f"    State auto-loaded: {loaded}")
                except Exception:
                    time.sleep(1)   # fallback if dropdown not found
            else:
                log.info("    Current Pincode: field NOT found")
        except Exception as e:
            log.info(f"    Pincode error: {e}")

        between_fields()   # v5-2

        # ── Distance (fast js_fill, per-bill from Excel) ──────
        def fill_distance():
            try:
                d = driver.find_element(By.ID, "txtDistance")
                if d.is_displayed():
                    for attempt in range(3):
                        js_fill(driver, d, dist)   # v5-4: per-bill value
                        time.sleep(0.4)
                        actual = d.get_attribute("value").strip()
                        if actual == dist:
                            log.info(
                                f"    Distance: {actual} km (confirmed on attempt {attempt+1})"
                            )
                            return True
                        log.info(f"    Distance attempt {attempt+1}: got '{actual}' — retrying")
                        time.sleep(0.3)
                    log.info("    Distance still wrong after 3 attempts")
            except Exception as e:
                log.info(f"    Distance error: {e}")
            return False

        fill_distance()
        between_fields()   # v5-2

        # ── Vehicle No (human_type) ───────────────────────────
        try:
            veh = None
            try:
                veh = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtVehicleNo")
            except Exception:
                pass
            if not veh:
                veh = find_visible_input(driver, ["vehicle", "vehicleno", "vehno", "veh_no"])
            if veh and veh.is_displayed():
                human_type(driver, veh, truck_no)   # v5-1
                log.info(f"    Vehicle No: {truck_no} (id={veh.get_attribute('id')})")
            else:
                log.info("    Vehicle No: field NOT found")
        except Exception as e:
            log.info(f"    Vehicle error: {e}")

        # ── Final distance check ──────────────────────────────
        try:
            d = driver.find_element(By.ID, "txtDistance")
            actual = d.get_attribute("value").strip()
            if actual != dist:
                log.info(f"    Distance reset after vehicle — correcting to {dist}")
                js_fill(driver, d, dist)
                time.sleep(0.4)
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
                pre_click()   # v5-2
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

        # ── Wait for portal result ────────────────────────────
        SUCCESS_WORDS = ["extended", "valid till", "validity", "successfully"]
        FAILURE_WORDS = ["error", "invalid", "cannot extend", "not found"]

        time.sleep(1)   # [v6-7] Reduced from 3s — WebDriverWait below handles the real wait
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
# GUI LOG HANDLER — routes log output to GUI window
# ─────────────────────────────────────────────
class GUILogHandler(logging.Handler):
    def __init__(self, app):
        super().__init__()
        self.app = app

    def emit(self, record):
        msg = self.format(record)
        m   = msg.lower()
        if any(w in m for w in ["success", "extended!", "ok]"]):
            tag = "ok"
        elif record.levelname == "WARNING" or "retry" in m or "retrying" in m:
            tag = "warn"
        elif record.levelname in ("ERROR", "CRITICAL"):
            tag = "err"
        elif any(w in m for w in ["launch", "login", "start", "loaded", "reading", "report", "chrome", "connecting"]):
            tag = "info"
        else:
            tag = "def"
        self.app.append_log(msg, tag)


# ─────────────────────────────────────────────
# GUI APP CLASS
# ─────────────────────────────────────────────
class EWBApp:
    def __init__(self, root):
        self.root     = root
        self.root.title("GST E-Way Bill Extension  v6.0  —  by Akshay")
        self.root.geometry("720x640")
        self.root.resizable(False, False)

        self.file_var  = tk.StringVar()
        self.stop_flag = False
        self.driver    = None

        self._build()
        self._hook_logger()

    # ── Build UI ──────────────────────────────────────────────
    def _build(self):
        BG = "#f3f4f6"
        self.root.configure(bg=BG)

        # Header
        hdr = tk.Frame(self.root, bg="#185FA5", height=54)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="GST E-Way Bill Extension",
                 bg="#185FA5", fg="white",
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=16, pady=14)
        tk.Label(hdr, text="v6.0  —  built by Akshay",
                 bg="#185FA5", fg="#a8c8e8",
                 font=("Segoe UI", 9)).pack(side="right", padx=16)

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=10)

        # File picker
        fp = tk.LabelFrame(body, text=" Excel File ", bg=BG,
                            font=("Segoe UI", 9))
        fp.pack(fill="x", pady=(0, 8))
        tk.Entry(fp, textvariable=self.file_var, width=70,
                 font=("Segoe UI", 9)).pack(side="left", padx=8, pady=6)
        tk.Button(fp, text="Browse", command=self.browse,
                  font=("Segoe UI", 9)).pack(side="left", pady=6, padx=(0, 8))

        # Stat boxes
        sf = tk.Frame(body, bg=BG)
        sf.pack(fill="x", pady=(0, 8))
        self.stat_vars = {}
        for label, key, color in [
            ("Total",     "total",   "#185FA5"),
            ("Processed", "done",    "#444444"),
            ("Success",   "success", "#166534"),
            ("Failed",    "failed",  "#b91c1c"),
        ]:
            box = tk.Frame(sf, bg="white", relief="solid", bd=1)
            box.pack(side="left", expand=True, fill="x", padx=4)
            v = tk.StringVar(value="0")
            self.stat_vars[key] = v
            tk.Label(box, textvariable=v, bg="white", fg=color,
                     font=("Segoe UI", 22, "bold")).pack(pady=(8, 0))
            tk.Label(box, text=label, bg="white", fg="#888",
                     font=("Segoe UI", 9)).pack(pady=(0, 8))

        # Progress bar
        pf = tk.Frame(body, bg=BG)
        pf.pack(fill="x", pady=(0, 4))
        self.progress_var = tk.DoubleVar(value=0)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("b.Horizontal.TProgressbar",
                         troughcolor="#e5e7eb", background="#185FA5", thickness=10)
        ttk.Progressbar(pf, variable=self.progress_var, maximum=100,
                         style="b.Horizontal.TProgressbar").pack(fill="x", pady=(0, 4))
        self.lbl_progress = tk.Label(pf, text="0 of 0 bills  —  0%",
                                      bg=BG, fg="#666", font=("Segoe UI", 9))
        self.lbl_progress.pack(anchor="w")

        self.lbl_current = tk.Label(body, text="Current: —",
                                     bg=BG, fg="#444", font=("Segoe UI", 9, "italic"))
        self.lbl_current.pack(anchor="w", pady=(0, 6))

        # Log window
        lf = tk.LabelFrame(body, text=" Live Log ", bg=BG, font=("Segoe UI", 9))
        lf.pack(fill="both", expand=True, pady=(0, 8))
        self.log_text = tk.Text(lf, height=11, bg="#0d1117", fg="#8b949e",
                                 font=("Consolas", 9), state="disabled",
                                 wrap="none", relief="flat", bd=0)
        sb = tk.Scrollbar(lf, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True, padx=4, pady=4)
        self.log_text.tag_config("ok",   foreground="#3fb950")
        self.log_text.tag_config("warn", foreground="#d29922")
        self.log_text.tag_config("err",  foreground="#f85149")
        self.log_text.tag_config("info", foreground="#58a6ff")
        self.log_text.tag_config("def",  foreground="#8b949e")

        # Buttons
        bf = tk.Frame(body, bg=BG)
        bf.pack(fill="x")
        self.btn_start = tk.Button(
            bf, text="▶  Start", command=self.start,
            bg="#185FA5", fg="white", activebackground="#0c447c",
            activeforeground="white", font=("Segoe UI", 10, "bold"),
            relief="flat", bd=0, padx=20, pady=9)
        self.btn_start.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self.btn_stop = tk.Button(
            bf, text="■  Stop", command=self.stop,
            bg="#fee2e2", fg="#b91c1c", font=("Segoe UI", 10),
            relief="flat", bd=0, padx=16, pady=9, state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 4))

        self.btn_report = tk.Button(
            bf, text="Open Report", command=self.open_report,
            bg="#f3f4f6", fg="#444", font=("Segoe UI", 10),
            relief="groove", bd=1, padx=16, pady=9)
        self.btn_report.pack(side="left")

    # ── Logger hook ───────────────────────────────────────────
    def _hook_logger(self):
        h = GUILogHandler(self)
        h.setFormatter(logging.Formatter(
            "%(asctime)s  %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger(__name__).addHandler(h)

    # ── Actions ───────────────────────────────────────────────
    def browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV", "*.csv")])
        if path:
            self.file_var.set(path)

    def append_log(self, msg, tag="def"):
        def _do():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n", tag)
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _do)

    def _set_stats(self, total=None, done=None, success=None, failed=None, current=""):
        if total   is not None: self.stat_vars["total"].set(str(total))
        if done    is not None: self.stat_vars["done"].set(str(done))
        if success is not None: self.stat_vars["success"].set(str(success))
        if failed  is not None: self.stat_vars["failed"].set(str(failed))
        t   = int(self.stat_vars["total"].get()   or 0)
        d   = int(self.stat_vars["done"].get()    or 0)
        pct = int(d / t * 100) if t > 0 else 0
        self.progress_var.set(pct)
        self.lbl_progress.config(text=f"{d} of {t} bills  —  {pct}%")
        if current:
            self.lbl_current.config(text=f"Current:  {current}")

    def update_stats(self, **kwargs):
        self.root.after(0, lambda: self._set_stats(**kwargs))

    def start(self):
        if not self.file_var.get():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
        if not check_license():
            return
        self.stop_flag = False
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self._run, daemon=True).start()

    def stop(self):
        self.stop_flag = True
        self.append_log("Stop requested — finishing current bill...", "warn")
        self.btn_stop.config(state="disabled")

    def open_report(self):
        global REPORT_FILE
        if REPORT_FILE and os.path.exists(REPORT_FILE):
            os.startfile(REPORT_FILE)
        else:
            messagebox.showinfo("Report", "No report file yet.")

    def _reset_buttons(self):
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")

    # ── Automation thread ─────────────────────────────────────
    def _run(self):
        global REPORT_FILE
        try:
            bills = load_bills(self.file_var.get())
            total = len(bills)
            self.update_stats(total=total, done=0, success=0, failed=0)

            ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
            REPORT_FILE = f"extension_report_{ts}.xlsx"
            log.info(f"Report file: {REPORT_FILE}")

            log.info("Launching Chrome...")
            self.driver = create_driver(headless=False)

            if not do_login(self.driver):
                self.append_log("Login failed or timed out.", "err")
                self.root.after(0, self._reset_buttons)
                return

            results  = []
            ok_count = 0
            fail_count = 0

            for i, bill in enumerate(bills, 1):
                if self.stop_flag:
                    self.append_log("Stopped by user.", "warn")
                    break

                self.update_stats(
                    done=i - 1, success=ok_count, failed=fail_count,
                    current=f"{bill['ewb']}  |  Truck: {bill['truck']}  |  Place: {bill['place']}"
                )

                result = extend_one(self.driver, bill, i, total, debug_mode=(i == 1))

                if not result["success"]:
                    log.info(f"  Retrying {bill['ewb']} in 3s...")
                    time.sleep(3)
                    result = extend_one(self.driver, bill, i, total, debug_mode=False)

                results.append(result)
                if result["success"]:
                    ok_count += 1
                else:
                    fail_count += 1

                self.update_stats(done=i, success=ok_count, failed=fail_count)

                if i % 5 == 0 or not result["success"]:
                    save_report(results)

            save_report(results)
            self.append_log(
                f"Done!  Success: {ok_count}   Failed: {fail_count}   Report: {REPORT_FILE}",
                "ok"
            )
            self.update_stats(done=len(results), success=ok_count,
                              failed=fail_count, current="Complete")

        except Exception as e:
            self.append_log(f"Critical error: {e}", "err")
            log.error(f"Critical error: {e}")
        finally:
            try:
                if self.driver:
                    self.driver.quit()
                    self.driver = None
            except Exception:
                pass
            self.root.after(0, self._reset_buttons)
# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    root = tk.Tk()
    app  = EWBApp(root)
    root.mainloop()

if __name__ == "__main__":
        main()
