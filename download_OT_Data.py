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
import re
from pathlib import Path
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from datetime import datetime
import pytz
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# === Setup Logging ===
# This sets up logging to the console (GitHub Actions will capture this)
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# === Setup: Linux-compatible download directory ===
download_dir = os.path.join(os.getcwd(), "download")
os.makedirs(download_dir, exist_ok=True)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")  # Comment this line for debug
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

pattern = "Monthty Manhours Report"

def is_file_downloaded():
    return any(Path(download_dir).glob(f"*{pattern}*.xlsx"))

# === Debugging Loop ===
while True:  # Infinite loop until the file is downloaded
    try:
        log.info("Attempting to start the browser...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        wait = WebDriverWait(driver, 20)

        log.info("Navigating to login page...")
        driver.get("https://taps.odoo.com")
        wait.until(EC.presence_of_element_located((By.NAME, "login"))).send_keys("supply.chain3@texzipperbd.com")
        driver.find_element(By.NAME, "password").send_keys("@Shanto@86")
        driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]").click()
        time.sleep(2)

        time.sleep(2)
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".modal-backdrop")))
        except:
            pass

        switcher_span = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
            "div.o_menu_systray div.o_switch_company_menu > button > span"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", switcher_span)
        switcher_span.click()
        time.sleep(2)

        log.info("Selecting 'Zipper' company...")
        target_div = wait.until(EC.element_to_be_clickable((By.XPATH,
            "//div[contains(@class, 'log_into')][span[contains(text(), 'Zipper')]]"
        )))
        driver.execute_script("arguments[0].scrollIntoView(true);", target_div)
        target_div.click()
        time.sleep(4)

        # Going to attendence module
        driver.get("https://taps.odoo.com/web#action=699&model=hr.attendance&view_type=list&menu_id=484&cids=1")
        wait.until(EC.presence_of_element_located((By.XPATH, "//html")))
        time.sleep(5)
        #Step 2    
        log.info("Clicking on Attendece button option button...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/header/nav/div[1]/a[4]"))).click()
        time.sleep(8)
        #Step 3
        log.info("Clicking on Report TYPE option LIST TO see all the report list...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[1]/div[1]/div[2]/div/select"))).click()
        time.sleep(8)
        
        #Step 4
        log.info("Clicking on Report TYPE OT Analysis...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[1]/div[1]/div[2]/div/select/option[13]"))).click()
        time.sleep(5)
        
        # Step 5
        log.info("Clicking on Mode mean seleting by company...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[3]/div[2]/div/select"))).click()
        time.sleep(5)
        
        # Step 6
        log.info("Clicking on Mode by company...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[3]/div[2]/div/select/option[3]"))).click()
        time.sleep(5)
        
        # Step 7
        log.info("Clicking on company list to get compnay list ...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[4]/div[2]/div/div[1]/div/div/input"))).click()
        time.sleep(10)
        
        
        # 
        
        # Step 8
        log.info("Clicking on company Zipper ...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[2]/div[4]/div[2]/div/div[1]/div/div/ul/li[1]/a"))).click()
        time.sleep(10)
        
        # /html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[2]    
        
        # Input date value 
        # === 1. Get today's date
        today = datetime.today()

        # === 2. Calculate start and end of "office month" (26th → 25th)
        if today.day >= 26:
            # We're after the 26th: start = 26th current month, end = 25th next month
            start_date = datetime(today.year, today.month, 26)
            # handle December -> January
            if today.month == 12:
                end_date = datetime(today.year + 1, 1, 25)
            else:
                end_date = datetime(today.year, today.month + 1, 25)
        else:
            # We're before 26th: start = 26th previous month, end = 25th current month
            if today.month == 1:
                start_date = datetime(today.year - 1, 12, 26)
            else:
                start_date = datetime(today.year, today.month - 1, 26)
            end_date = datetime(today.year, today.month, 25)

        # Format for Selenium input
        start_str = start_date.strftime("%d/%m/%Y")
        end_str = end_date.strftime("%d/%m/%Y")

        # === 3. Selenium: find input fields and send dates
        start_xpath = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[1]/div[2]/div[2]/div/div/input"
        end_xpath   = "/html/body/div[2]/div[2]/div/div/div/div/main/div/div/div/div/div/div[1]/div[3]/div[2]/div/div/input"

        time.sleep(3)

        # Start date input
        start_input = driver.find_element(By.XPATH, start_xpath)
        start_input.send_keys(Keys.CONTROL + 'a')
        start_input.send_keys(Keys.BACKSPACE)
        start_input.send_keys(start_str)

        time.sleep(2)

        # End date input
        end_input = driver.find_element(By.XPATH, end_xpath)
        end_input.send_keys(Keys.CONTROL + 'a')
        end_input.send_keys(Keys.BACKSPACE)
        end_input.send_keys(end_str)

        time.sleep(2)

        # Step 9
        log.info("Clicking on Export Excel to download the file ...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div/div/div/footer/footer/button[2]"))).click()
        time.sleep(40)

        if is_file_downloaded():
            log.info("✅ File download complete!")
            files = list(Path(download_dir).glob(f"*{pattern}*.xlsx"))
            if len(files) > 1:
                files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
                for file in files[1:]:
                    file.unlink()
            driver.quit()
            break  # Exit the loop after file download is complete
        else:
            log.warning("⚠️ File not downloaded. Retrying...")

    except Exception as e:
        log.error(f"❌ Error occurred: {e}\nRetrying in 10 seconds...\n")
        try:
            driver.quit()
        except:
            pass
        time.sleep(10)

# === Step: Upload to Google Sheets ===
try:
    log.info("Checking for downloaded files...")
    files = list(Path(download_dir).glob(f"*{pattern}*.xlsx"))
    if not files:
        raise Exception("No matching file found.")

    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    latest_file = files[0]
    df = pd.read_excel(latest_file,sheet_name=0)
    time.sleep(8)
    log.info("✅ File loaded into DataFrame.")

    # Use credentials stored in gcreds.json (created in GitHub Action)
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # Use google-auth to load credentials
    creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=scope)
    log.info("✅ Successfully loaded credentials.")

    # Use gspread to authorize and access Google Sheets
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1-kBuln5CnKucuHqYG4vvgttJ8DqeJALvr4TjAYuVkXs")
    worksheet = sheet.worksheet("ZIP_OT_NEW_DATA")

    if df.empty:
        print("Skip: DataFrame is empty, not pasting to sheet.")
    else:
        worksheet.batch_clear(["B1:BZ1000"])
        time.sleep(4)
        set_with_dataframe(worksheet, df, row=1, col=2)
        print("Data pasted to Google Sheet (Sheet4).")
        local_tz = pytz.timezone('Asia/Dhaka')
        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update("E1", [[f"{local_time}"]])
        log.info(f"✅ Data pasted & timestamp updated: {local_time}")

    

except Exception as e:
    log.error(f"❌ Error while pasting to Google Sheets: {e}")
