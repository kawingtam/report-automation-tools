from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import os
import time
import pandas as pd
import glob
import shutil

# User inputs
DOWNLOAD_DIR = input("Enter temporary download folder path: ").strip().strip('"')
REPORT_URL = input("Enter report dashboard URL: ").strip().strip('"')
INPUT_EXCEL = input("Enter award/task input Excel file path: ").strip().strip('"')

# Configurations
MAX_WAIT = 30
REPORT_TIMEOUT = 600
WAIT_AFTER_SEARCH = 300  # seconds

# Load award/task pairs from Excel
df = pd.read_excel(INPUT_EXCEL)
df["TASK_NUMBER"] = df["TASK_NUMBER"].fillna("")

# Track failures
failed_reports = []

def initialize_browser():
    print("🛠️ Initializing browser configuration...")
    chrome_options = Options()
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-extensions")
    service = ChromeService(ChromeDriverManager().install())
    browser = webdriver.Chrome(service=service, options=chrome_options)
    browser.maximize_window()
    print("✅ Browser initialized successfully")
    return browser

def wait_for_user_ready():
    """
    Wait for the user to confirm they are logged in and ready to proceed.
    This is useful when the BI dashboard requires a manual login.
    """
    while True:
        user_input = input("\n🔑 Please log in to the dashboard. Type 'yes' when ready to start downloading: ").strip().lower()
        if user_input == "yes":
            print("✅ User confirmed, proceeding with report downloads...")
            break
        else:
            print("ℹ️ Waiting for user to be ready... type 'yes' when done logging in.")

def optimized_locator(browser, field_name, locators):
    print(f"   🔎 Searching for {field_name} field...")
    for by, locator in locators:
        try:
            element = WebDriverWait(browser, MAX_WAIT).until(
                EC.presence_of_element_located((by, locator))
            )
            print(f"   ✅ Found {field_name} using locator: {locator}")
            return element
        except:
            print(f"   ⚠️ Locator failed for {field_name}: {locator}")
    raise Exception(f"❌ Could not locate {field_name}")

def enter_award_number(browser, award_number):
    print(f"   🖊️ Entering Award Number: {award_number}")
    award_locators = [
        (By.XPATH, "//label[contains(.,'Award Number')]/following::input[1]"),
        (By.NAME, "award_number"),
        (By.XPATH, "//input[contains(@id,'award')]")
    ]
    field = optimized_locator(browser, "Award Number", award_locators)
    field.clear()
    time.sleep(0.5)
    field.send_keys(award_number)

def enter_task_number(browser, task_number):
    print(f"   🖊️ Entering Task Number: {task_number}")
    task_locators = [
        (By.XPATH, "//label[contains(.,'Task Number')]/following::input[1]"),
        (By.NAME, "task_number"),
        (By.XPATH, "//input[contains(@id,'task')]")
    ]
    field = optimized_locator(browser, "Task Number", task_locators)
    field.clear()
    time.sleep(0.5)
    field.send_keys(task_number)
    field.send_keys(Keys.TAB)

def click_apply_button(browser):
    print("   ⏩ Clicking Apply button...")
    apply_button = WebDriverWait(browser, MAX_WAIT).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@value='Apply' and contains(@class,'promptApplyButton')]"))
    )
    apply_button.click()

def open_report_in_new_tab(browser, award_number, task_number):
    report_name = f"{award_number}-{task_number}"
    print(f"\n📂 Opening report tab for {report_name}")
    browser.execute_script("window.open('');")
    browser.switch_to.window(browser.window_handles[-1])
    browser.get(REPORT_URL)

    # Wait for page load
    WebDriverWait(browser, MAX_WAIT).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )

    enter_award_number(browser, award_number)
    enter_task_number(browser, task_number)
    click_apply_button(browser)

    print(f"   ⏳ Waiting {WAIT_AFTER_SEARCH} seconds for report {report_name} to load...")
    time.sleep(WAIT_AFTER_SEARCH)
    print(f"   ✅ Search completed for {report_name}")

def click_show_all_if_present(browser, award_number, task_number):
    report_name = f"{award_number}-{task_number}"
    print(f"🔎 Now finding 'Show All' button for report {report_name}...")
    try:
        show_all = WebDriverWait(browser, 3).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//img[contains(@title, 'Display maximum') and contains(@src, 'showallrows')]")
            )
        )
        print(f"   📌 Found 'Show All' button, clicking...")
        show_all.click()
        time.sleep(2)
    except:
        print(f"   ℹ️ No 'Show All' button for {report_name}")

def export_to_pdf(browser, award_number, task_number, dest_path):
    report_name = f"{award_number}-{task_number}"
    print(f"💾 Now starting download for report {report_name}...")
    try:
        # Ensure destination folder exists
        os.makedirs(dest_path, exist_ok=True)

        # Track current files
        existing_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf")))

        # Click second print button
        print(f"   📄 Now clicking dropdown menu button...")
        second_button = WebDriverWait(browser, MAX_WAIT).until(
            EC.element_to_be_clickable((By.XPATH, "(//a[contains(@title, 'Print in different format')])[2]"))
        )
        second_button.click()

        # Find "Printable PDF"
        print(f"   🔍 Looking for 'Printable PDF' option...")
        menu_items = WebDriverWait(browser, MAX_WAIT).until(
            EC.presence_of_all_elements_located((By.XPATH, "//td[contains(@class,'MenuItemTextCell')]"))
        )
        for item in menu_items:
            if "Printable PDF" in item.text:
                print(f"   ✅ Found 'Printable PDF', clicking...")
                item.click()
                break
        else:
            raise Exception("Printable PDF option not found")

        # Wait dynamically for new PDF to appear
        print(f"   ⏳ Waiting for PDF download to complete...")
        timeout = 60
        start_time = time.time()
        downloaded_file = None
        while time.time() - start_time < timeout:
            all_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf")))
            new_files = all_files - existing_files
            if new_files:
                downloaded_file = new_files.pop()
                # Check file size stabilization
                size1 = os.path.getsize(downloaded_file)
                time.sleep(1)
                size2 = os.path.getsize(downloaded_file)
                if size1 == size2:
                    break
            time.sleep(1)

        if not downloaded_file:
            raise Exception("No new PDF file detected after waiting")

        # Move and rename to DEST_PATH folder
        new_filename = os.path.join(dest_path, f"{award_number}_{task_number}.pdf")
        shutil.move(downloaded_file, new_filename)
        print(f"   ✅ Report saved as {new_filename}")

    except Exception as e:
        print(f"   ❌ Failed to export {report_name}: {e}")
        failed_reports.append((award_number, task_number))



if __name__ == "__main__":
    browser = initialize_browser()
    browser.get(REPORT_URL)

    # Wait for user to complete manual login
    wait_for_user_ready()

    # Clean up Excel columns
    df['TASK_NUMBER'] = df['TASK_NUMBER'].fillna('')
    df['DEST_PATH'] = df['DEST_PATH'].astype(str).str.strip()  # remove extra spaces/newlines

    try:
        # Phase 1: open all reports in new tabs
        for _, row in df.iterrows():
            award = str(row['AWARD_NUMBER']).strip()
            task = str(row['TASK_NUMBER']).strip()
            open_report_in_new_tab(browser, award, task)

        # Optional wait after opening all tabs
        print("\n⏳ Waiting a few seconds to ensure all reports have loaded...")
        time.sleep(5)  # adjust if needed

        # Phase 2: process each tab (skip first if main dashboard)
        for idx, handle in enumerate(browser.window_handles[1:], start=1):
            browser.switch_to.window(handle)
            award = str(df.iloc[idx-1]['AWARD_NUMBER']).strip()
            task = str(df.iloc[idx-1]['TASK_NUMBER']).strip()
            dest_path = str(df.iloc[idx-1]['DEST_PATH']).strip()  # ensure no whitespace

            print(f"\n=== Processing report {award}-{task} ===")
            print(f"DEBUG: dest_path -> {dest_path}")  # debug print

            click_show_all_if_present(browser, award, task)
            export_to_pdf(browser, award, task, dest_path)

        # Final failure report
        print("\n=== Download Failures ===")
        if failed_reports:
            for aw, ts in failed_reports:
                print(f"❌ {aw} / {ts}")
        else:
            print("✅ All reports downloaded successfully")

    finally:
        input("\nPress Enter to close browser...")
        browser.quit()

