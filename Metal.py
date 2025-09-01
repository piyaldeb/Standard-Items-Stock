import sys
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from pathlib import Path
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

logging.basicConfig(level=logging.INFO)

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

# Use project download folder
DOWNLOAD_PATH = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_PATH, exist_ok=True)

FILE_PATTERN = "Standard Items Stock*"  # pattern to match downloaded file
OUTPUT_FILE_NAME = "Metal Raw.xlsx"

# Google Sheet config
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1fnOSIWQa_mbfMHdgPatjYEIhG3kQlzPy0djHG8TOszk/edit#gid=463655666"
SHEET_NAME = "Metal Raw"
CREDENTIALS_FILE = "credentials.json"  # your service account JSON


def wait_for_download(download_dir: str, pattern: str, timeout: int = 180) -> Path:
    """
    Wait for a file matching 'pattern' to appear in 'download_dir' and for any .crdownload to finish.
    Returns the Path to the newest matching file.
    """
    start = time.time()
    latest_file = None
    logging.info(f"⏳ Waiting for download into {download_dir} (pattern: '{pattern}')")

    while time.time() - start < timeout:
        matches = list(Path(download_dir).glob(pattern))
        if matches:
            # newest match by mtime
            candidate = max(matches, key=os.path.getmtime)
            crdownload = candidate.with_name(candidate.name + ".crdownload")
            if not crdownload.exists():
                latest_file = candidate
                break
        time.sleep(1)

    if not latest_file:
        raise TimeoutError(f"Download did not complete in {timeout}s (pattern='{pattern}')")

    logging.info(f"✅ Download complete: {latest_file}")
    return latest_file


def main():
    logging.info("✅ Starting Metal.py...")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # uncomment to run headless
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    prefs = {"download.default_directory": DOWNLOAD_PATH,
             "download.prompt_for_download": False,
             "safebrowsing.enabled": True}
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # -------------------------
        # OPEN ODOO LOGIN
        # -------------------------
        driver.get(ODOO_URL)
        logging.info(f"🌐 Opened {ODOO_URL}")

        wait.until(EC.presence_of_element_located((By.NAME, "login")))
        driver.find_element(By.NAME, "login").send_keys(EMAIL)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(.,'Log in')]").click()
        logging.info("🔑 Submitted login credentials")
        time.sleep(10)  # wait for dashboard

        # -------------------------
        # SEARCH "Standard Items Stock"
        # -------------------------
        search_box = wait.until(EC.presence_of_element_located((By.TAG_NAME, "input")))
        search_box.send_keys("Standard items Stock")
        search_box.send_keys(Keys.ENTER)
        time.sleep(20)

        # click table header (select all rows)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//table/thead/tr/th[1]"))).click()
        time.sleep(10)

        # click "Select all"
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Select all')]"))).click()
        time.sleep(10)

        # click "Action" dropdown
        action_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Action')]")))
        action_btn.click()
        time.sleep(5)

        # click "Export"
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Export')]"))).click()
        time.sleep(2)

        # =========================
        # (1) SELECT "00-Ranak" IN EXPORT MODAL
        # =========================
        # Use the absolute XPath you shared earlier (adjust if Odoo updates the DOM)
        select_xpath = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div[2]/div[3]/div/select"
        dropdown_el = wait.until(EC.presence_of_element_located((By.XPATH, select_xpath)))

        sel = Select(dropdown_el)
        try:
            sel.select_by_visible_text("00-Ranak")
            logging.info('📌 Selected template: "00-Ranak"')
        except Exception as e:
            # Fallback to first option if not present
            if not sel.options:
                raise RuntimeError("No options found in export dropdown!") from e
            sel.select_by_index(0)
            logging.info(f'📌 "00-Ranak" not found; selected first option: "{sel.options[0].text.strip()}"')

        time.sleep(1)

        # =========================
        # (2) CONFIRM EXPORT
        # =========================
        export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//footer//button[contains(., 'Export')]")))
        export_btn.click()
        logging.info("📤 Export confirmed, waiting for file to download...")
        time.sleep(2)

        # =========================
        # (3) WAIT FOR DOWNLOAD
        # =========================
        latest_file = wait_for_download(DOWNLOAD_PATH, FILE_PATTERN, timeout=180)

        # --- read Excel/CSV ---
        if latest_file.suffix.lower() in [".xlsx", ".xls"]:
            df = pd.read_excel(latest_file)
        elif latest_file.suffix.lower() == ".csv":
            df = pd.read_csv(latest_file)
        else:
            raise ValueError(f"Unsupported file type: {latest_file.suffix}")

        # Keep only columns A:J
        df = df.iloc[:, :10]

        # Save locally (as requested)
        out_file = os.path.join(DOWNLOAD_PATH, OUTPUT_FILE_NAME)
        df.to_excel(out_file, index=False)
        logging.info(f"✅ File saved as: {out_file}")

        # -------------------------
        # UPLOAD TO GOOGLE SHEETS (A:J) BATCH
        # -------------------------
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
        worksheet = spreadsheet.worksheet(SHEET_NAME)

        # Prepare data to update: include header
        data_to_update = [df.columns.tolist()] + df.values.tolist()

        # Batch update A:J
        cell_range = f"A1:J{len(data_to_update)}"
        worksheet.update(cell_range, data_to_update)

        logging.info("✅ Data uploaded successfully to Google Sheet (columns A:J)")

        # delete original downloaded file
        try:
            os.remove(latest_file)
            logging.info(f"🗑️ Original file deleted: {latest_file}")
        except Exception as e:
            logging.warning(f"⚠️ Could not delete file: {e}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
