import sys
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from pathlib import Path
import pandas as pd

logging.basicConfig(level=logging.INFO)

# -------------------------
# CONFIG
# -------------------------
ODOO_URL = "https://taps.odoo.com/web#action=menu&cids=3"
EMAIL = "ranak@texzipperbd.com"
PASSWORD = "2326"

# Use project download folder
DOWNLOAD_PATH = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_PATH, exist_ok=True)

FILE_PATTERN = "Standard Items Stock*"  # pattern to match downloaded file
OUTPUT_FILE_NAME = "Metal Raw.xlsx"

def main():
    logging.info("✅ Starting Metal.py...")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # uncomment to run headless
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    prefs = {"download.default_directory": DOWNLOAD_PATH}
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
        time.sleep(5)

        # confirm Export in popup
        wait.until(EC.element_to_be_clickable((By.XPATH, "//footer//button[contains(., 'Export')]"))).click()

        # -------------------------
        # WAIT FOR DOWNLOAD AND LOAD FILE
        # -------------------------
        timeout = 60  # wait max 60 seconds
        start = time.time()
        latest_file = None

        while time.time() - start < timeout:
            files = list(Path(DOWNLOAD_PATH).glob(FILE_PATTERN))
            if files:
                latest_file = max(files, key=os.path.getctime)
                break
            time.sleep(1)

        if not latest_file:
            print(f"❌ File not found in {DOWNLOAD_PATH} after waiting!")
            driver.quit()
            exit()

        print(f"✅ Latest file found: {latest_file}")

        # --- read Excel ---
        if latest_file.suffix.lower() in [".xlsx", ".xls"]:
            df = pd.read_excel(latest_file)
        elif latest_file.suffix.lower() == ".csv":
            df = pd.read_csv(latest_file)
        else:
            raise ValueError(f"Unsupported file type: {latest_file.suffix}")

        out_file = os.path.join(DOWNLOAD_PATH, OUTPUT_FILE_NAME)
        df.to_excel(out_file, index=False)
        logging.info(f"✅ File saved as: {out_file}")

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
