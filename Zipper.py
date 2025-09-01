import time
import os
from datetime import datetime
from pathlib import Path
import platform
from typing import List
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import gspread
from oauth2client.service_account import ServiceAccountCredentials


# -------------------------
# CONFIG
# -------------------------
from dotenv import load_dotenv   # <-- NEW

# Load variables from .env
load_dotenv()

# =========================
# CONFIG — now pulled from environment
# =========================
ODOO_URL   = os.getenv("ODOO_URL")
DB         = os.getenv("ODOO_DB")
USERNAME   = os.getenv("ODOO_USERNAME")
PASSWORD   = os.getenv("ODOO_PASSWORD")
EXPECTED_NAME_HINT = "Standard Items Stock"  # only for logging

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1fnOSIWQa_mbfMHdgPatjYEIhG3kQlzPy0djHG8TOszk/edit?gid=1326846174"
SHEET_NAME = "Zipper Raw"

KEEP_BROWSER_ON_ERROR = True


# -------------------------
# HELPERS
# -------------------------
def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)

def default_download_dir() -> Path:
    """Return OS default Downloads folder."""
    home = Path.home()
    sys = platform.system().lower()
    if "windows" in sys:
        return Path(os.path.join(os.environ.get("USERPROFILE", str(home)), "Downloads"))
    else:
        return home / "Downloads"

def ensure_dir(p: Path) -> Path:
    p.mkdir(parents=True, exist_ok=True)
    return p

def file_snapshot(dirs: List[Path]):
    """Return set of existing file paths (excluding .crdownload) across dirs."""
    snap = set()
    for d in dirs:
        if not d.exists():
            continue
        for f in d.iterdir():
            if f.is_file() and not f.name.endswith(".crdownload"):
                snap.add(f.resolve())
    return snap

def wait_for_new_download(dirs: List[Path], before_set, timeout: int = 180) -> Path:
    """
    Detect a new file appearing in 'dirs' compared to before_set.
    Wait until any .crdownload for it is gone; return the final Path.
    """
    log(f"Watching for new download in: {', '.join(map(str, dirs))} (timeout {timeout}s)")
    start = time.time()

    def partial_exists_for(path: Path) -> bool:
        for d in dirs:
            if (d / (path.name + ".crdownload")).exists():
                return True
        return False

    while time.time() - start < timeout:
        after = file_snapshot(dirs)
        new_files = list(after - before_set)
        if new_files:
            candidate = max(new_files, key=lambda p: p.stat().st_mtime)
            if candidate.exists() and not partial_exists_for(candidate):
                return candidate
        time.sleep(1)

    raise TimeoutError("Download did not appear/complete within timeout (snapshot mode).")

def wait_for_download_since(dirs: List[Path], since_ts: float, timeout: int = 180) -> Path:
    """
    Return the newest file whose mtime >= since_ts and not a .crdownload.
    Handles same-filename overwrite cases.
    """
    log(f"Waiting for file modified since {datetime.fromtimestamp(since_ts).strftime('%H:%M:%S')} (timeout {timeout}s)")
    start = time.time()
    while time.time() - start < timeout:
        candidates = []
        for d in dirs:
            if not d.exists():
                continue
            for f in d.iterdir():
                if f.is_file() and not f.name.endswith(".crdownload"):
                    try:
                        mtime = f.stat().st_mtime
                    except Exception:
                        continue
                    if mtime >= since_ts:
                        candidates.append((mtime, f.resolve()))
        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            newest = candidates[0][1]
            # ensure no corresponding .crdownload exists
            if not newest.with_name(newest.name + ".crdownload").exists():
                return newest
        time.sleep(1)
    raise TimeoutError("No freshly modified file detected within timeout (mtime mode).")

def pick_download_dirs(configured_dir: Path) -> List[Path]:
    """
    Candidate download dirs:
    1) Configured Selenium dir
    2) OS default Downloads
    """
    dirs = []
    if configured_dir:
        ensure_dir(configured_dir)
        dirs.append(configured_dir)
    dirs.append(default_download_dir())
    # de-dupe
    seen, uniq = set(), []
    for d in dirs:
        rp = d.resolve()
        if rp not in seen:
            uniq.append(rp)
            seen.add(rp)
    return uniq

def print_recent_files(dirs: List[Path], topn: int = 3):
    for d in dirs:
        try:
            files = [f for f in d.iterdir() if f.is_file() and not f.name.endswith(".crdownload")]
            files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            head = files[:topn]
            if head:
                log(f"Top {topn} recent in {d}:")
                for f in head:
                    log(f"  - {f.name}  mtime={datetime.fromtimestamp(f.stat().st_mtime).strftime('%H:%M:%S')}")
        except Exception as e:
            log(f"Could not list {d}: {e}")


# -------------------------
# SELENIUM OPTIONS
# -------------------------
CONFIGURED_DIR = ensure_dir(Path.cwd() / "download")
CANDIDATE_DIRS = pick_download_dirs(CONFIGURED_DIR)
log("Candidate download directories:")
for d in CANDIDATE_DIRS:
    log(f"  - {d}")

options = webdriver.ChromeOptions()
options.add_argument("--headless=new")  # enable if needed
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_experimental_option("prefs", {
    "download.default_directory": str(CONFIGURED_DIR),
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True,
})

driver = None
success = False

try:
    log("Booting ChromeDriver...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    # Allow downloads via CDP (extra safety)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": str(CONFIGURED_DIR)
        })
        log(f"CDP download path set to: {CONFIGURED_DIR}")
    except Exception as e:
        log(f"CDP setup skipped ({e})")

    # -------------------------
    # LOGIN
    # -------------------------
    log(f"Opening Odoo URL: {ODOO_URL}")
    driver.get(ODOO_URL)

    log("Waiting for login form...")
    wait.until(EC.presence_of_element_located((By.NAME, "login")))
    driver.find_element(By.NAME, "login").clear()
    driver.find_element(By.NAME, "login").send_keys(EMAIL)
    driver.find_element(By.NAME, "password").clear()
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    log("Submitting login...")
    driver.find_element(By.XPATH, "//button").click()

    log("Waiting for workspace to load after login...")
    time.sleep(8)

    # -------------------------
    # SEARCH
    # -------------------------
    log('Searching for "Standard items Stock"...')
    search_box = wait.until(EC.presence_of_element_located((By.TAG_NAME, "input")))
    search_box.clear()
    search_box.send_keys("Standard items Stock")
    search_box.send_keys(Keys.ENTER)

    log("Waiting for list/table to render...")
    time.sleep(60)

    # -------------------------
    # SELECT ALL ROWS
    # -------------------------
    log("Selecting all rows...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//table/thead/tr/th[1]"))).click()
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Select all')]"))).click()
    time.sleep(3)

    # -------------------------
    # EXPORT
    # -------------------------
    log("Opening 'Action' dropdown...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Action')]"))).click()
    time.sleep(1)

    log("Clicking 'Export'...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Export')]"))).click()

    # Take a pre-download snapshot across candidate dirs
    before = file_snapshot(CANDIDATE_DIRS)

    # Export modal
    log("Waiting for Export modal...")
    time.sleep(2)
    select_xpath = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select"
    dropdown_el = wait.until(EC.presence_of_element_located((By.XPATH, select_xpath)))

    sel = Select(dropdown_el)
    try:
        log('Selecting "00-Ranak"...')
        sel.select_by_visible_text("00-Ranak")
        chosen = "00-Ranak"
    except Exception as e:
        log(f'"00-Ranak" not found ({e}). Selecting first option...')
        if not sel.options:
            raise RuntimeError("No options in export dropdown!")
        sel.select_by_index(0)
        chosen = sel.options[0].text.strip()
    log(f'✅ Chosen template: {chosen}')

    time.sleep(1)

    # Record timing around the click for mtime-based detection
    export_click_ready_ts = time.time()
    log("Confirming export...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//footer//button[contains(., 'Export')]"))).click()
    time.sleep(1)
    export_clicked_ts = time.time()

    log("Export clicked. Monitoring for download (mtime-based first)...")

    # -------- Primary: mtime-based (handles same-name overwrite) --------
    try:
        latest_file = wait_for_download_since(CANDIDATE_DIRS, since_ts=export_clicked_ts, timeout=180)
    except TimeoutError as e:
        log(f"mtime-based detection timed out: {e}")
        # -------- Fallback: snapshot-based --------
        log("Falling back to snapshot-based detection...")
        latest_file = wait_for_new_download(CANDIDATE_DIRS, before_set=before, timeout=60)

    log(f"✅ Download complete: {latest_file} (dir: {latest_file.parent})")
    if EXPECTED_NAME_HINT not in latest_file.name:
        log(f"ℹ️ Note: filename doesn't contain hint '{EXPECTED_NAME_HINT}'. Name: {latest_file.name}")

    # (Optional) Show recent files for verification
    print_recent_files(CANDIDATE_DIRS, topn=3)

    # -------------------------
    # LOAD FILE
    # -------------------------
    log("Reading file into DataFrame...")
    if latest_file.suffix.lower() in [".xlsx", ".xls"]:
        df = pd.read_excel(latest_file)
    elif latest_file.suffix.lower() == ".csv":
        df = pd.read_csv(latest_file)
    else:
        raise ValueError(f"Unsupported file type: {latest_file.suffix}")
    log(f"DataFrame shape: {df.shape}")

    log("Keeping first 10 columns (A:J)...")
    df = df.iloc[:, :10]

    # -------------------------
    # GOOGLE SHEETS
    # -------------------------
    log("Authorizing Google Sheets...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    log("Opening spreadsheet and worksheet...")
    spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
    worksheet = spreadsheet.worksheet(SHEET_NAME)

    log("Clearing range A:J...")
    worksheet.batch_clear(["A:J"])

    log("Preparing values and uploading...")
    values = [df.columns.values.tolist()] + df.values.tolist()
    worksheet.update(f"A1:J{len(values)}", values)
    log("✅ Uploaded to Google Sheet.")

    # -------------------------
    # CLEANUP
    # -------------------------
    try:
        os.remove(latest_file)
        log(f"🗑️ Deleted local file: {latest_file}")
    except Exception as e:
        log(f"⚠️ Could not delete file: {e}")

    success = True

except Exception as e:
    log(f"❌ ERROR: {e}")
finally:
    if driver and (success or not KEEP_BROWSER_ON_ERROR):
        log("Closing browser...")
        driver.quit()
    log(f"Done. success={success}, keep_on_error={KEEP_BROWSER_ON_ERROR}")
    if driver and not success and KEEP_BROWSER_ON_ERROR:
        log("Browser left open for inspection. Press Ctrl+C to stop the script when done.")
        while True:
            time.sleep(1)
