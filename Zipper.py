import time
import os
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# -------------------------
# CONFIG
# -------------------------
ODOO_URL = "https://taps.odoo.com/web#action=menu&cids=1&menu_id=957"
EMAIL = "ranak@texzipperbd.com"
PASSWORD = "2326"

# Use project download folder
DOWNLOAD_PATH = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_PATH, exist_ok=True)

FILE_PATTERN = "Standard Items Stock*"  # match partial name in case it changes slightly

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1fnOSIWQa_mbfMHdgPatjYEIhG3kQlzPy0djHG8TOszk/edit?gid=1326846174"
SHEET_NAME = "Zipper Raw"

# -------------------------
# SELENIUM (HEADLESS)
# -------------------------
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
prefs = {"download.default_directory": DOWNLOAD_PATH}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get(ODOO_URL)
wait = WebDriverWait(driver, 30)

# -------------------------
# LOGIN
# -------------------------
time.sleep(5)
driver.find_element(By.NAME, "login").send_keys(EMAIL)
driver.find_element(By.NAME, "password").send_keys(PASSWORD)
driver.find_element(By.XPATH, "//button").click()
time.sleep(10)

# -------------------------
# SEARCH "Standard items Stock"
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
export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Export')]"))).click()
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

# read CSV (adjust if Excel)
df = pd.read_csv(latest_file)
df = df.iloc[:, :10]  # keep A:J only

# -------------------------
# UPLOAD TO GOOGLE SHEETS
# -------------------------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
worksheet = spreadsheet.worksheet(SHEET_NAME)

worksheet.clear()
worksheet.update([df.columns.values.tolist()] + df.values.tolist())

print("✅ Data uploaded successfully to Google Sheet.")

# -------------------------
# DELETE FILE AFTER UPLOAD
# -------------------------
try:
    os.remove(latest_file)
    print(f"🗑️ File deleted: {latest_file}")
except Exception as e:
    print(f"⚠️ Could not delete file: {e}")

driver.quit()
