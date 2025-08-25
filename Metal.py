import sys
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from pathlib import Path
import pandas as pd

logging.basicConfig(level=logging.INFO)

def main():
    logging.info("✅ Starting Metal.py...")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # 🔗 URL
        url = "https://taps.odoo.com/web#action=menu&cids=3"
        driver.get(url)
        logging.info(f"🌐 Opened {url}")

        # --- login if needed ---
        # wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='login']")))

        # --- do your export steps here ---
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div/div"))).click()
        logging.info("📥 Clicked dropdown for Export")

        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Export')]"))).click()
        logging.info("📤 Clicked Export")

        # --- file handling ---
        time.sleep(10)  # wait for file to download
        download_dir = Path.home() / "Downloads"

        # specify exact filename of downloaded file (change if needed)
        downloaded_file_name = "Metal Stock (pending.stock.config)"
        downloaded_file_path = download_dir / downloaded_file_name

        if not downloaded_file_path.exists():
            logging.warning(f"⚠️ File not found: {downloaded_file_path}")
            return

        logging.info(f"📂 Found downloaded file: {downloaded_file_path}")

        # read Excel (even if extension is .config)
        df = pd.read_excel(downloaded_file_path)
        out_file = download_dir / "Metal Raw.xlsx"
        df.to_excel(out_file, index=False)
        logging.info(f"✅ File saved as: {out_file}")

        # delete the original file
        try:
            os.remove(downloaded_file_path)
            logging.info(f"🗑️ Original file deleted: {downloaded_file_path}")
        except Exception as e:
            logging.warning(f"⚠️ Could not delete file: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
