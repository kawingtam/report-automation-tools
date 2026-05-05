import os
import time
import shutil
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import sys

# User inputs
DOWNLOAD_DIR = input("Enter temporary download folder path: ").strip().strip('"')
REPORT_URL = input("Enter report dashboard URL: ").strip().strip('"')
INPUT_EXCEL = input("Enter award/task input Excel file path: ").strip().strip('"')

# Configurations
MAX_WAIT = 300
REPORT_TIMEOUT = 600

# Load award/task pairs from Excel
df = pd.read_excel(INPUT_EXCEL)
df["TASK_NUMBER"] = df["TASK_NUMBER"].fillna("")
df["DEST_PATH"] = df["DEST_PATH"].fillna("")

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

def optimized_locator(browser, field_name, locators):
    print(f"🔍 Locating element: {field_name}")
    for by, locator in locators:
        try:
            element = WebDriverWait(browser, MAX_WAIT).until(
                EC.presence_of_element_located((by, locator)))
            browser.execute_script("arguments[0].style.border='2px solid green';", element)
            print(f"✅ Located {field_name} using: {by}='{locator}'")
            return element
        except Exception as e:
            print(f"⚠️ Attempt failed with {by}='{locator}': {str(e)}")
    browser.save_screenshot(f"element_not_found_{field_name.replace(' ', '_')}.png")
    raise Exception(f"❌ Could not locate {field_name}")

def enter_award_number(browser, award_number):
    print("\n=== Entering Award Number ===")
    award_locators = [
        (By.XPATH, "//label[contains(.,'Award Number')]/following::input[1]"),
        (By.NAME, "award_number"),
        (By.XPATH, "//input[contains(@id,'award')]")
    ]
    field = optimized_locator(browser, "Award Number", award_locators)
    field.clear()
    time.sleep(0.5)
    field.send_keys(award_number)
    print(f"✏️ Entered Award Number: {award_number}")

def enter_task_number(browser, task_number):
    print("\n=== Entering Task Number ===")
    wait = WebDriverWait(browser, MAX_WAIT)
    task_input = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//label[contains(.,'Task Number')]/following::input[1]")))

    browser.execute_script("arguments[0].scrollIntoView(true);", task_input)
    task_input.click()
    task_input.clear()
    time.sleep(0.5)
    task_input.send_keys(task_number)
    task_input.send_keys(Keys.TAB)
    print(f"✏️ Entered Task Number: {task_number}")

def click_apply_button(browser):
    print("\n=== Clicking Apply Button ===")
    wait = WebDriverWait(browser, MAX_WAIT)
    apply_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//input[@value='Apply' and contains(@class,'promptApplyButton')]")))
    apply_button.click()
    print("🔄 Apply button clicked - generating report...")

def wait_for_report(browser, timeout=REPORT_TIMEOUT, min_wait_seconds=40):
    print("\n=== Waiting for Report Generation ===")
    start_time = time.time()
    report_xpath = "//*[contains(@class,'PivotTable') or contains(@id,'report') or contains(@class,'ReportContent')]"
    report_detected_time = None

    while True:
        elapsed = int(time.time() - start_time)
        if elapsed % 10 == 0:
            print(f"⏳ Still waiting... {elapsed}s elapsed")

        try:
            elements = browser.find_elements(By.XPATH, report_xpath)
            if elements and report_detected_time is None:
                report_detected_time = time.time()
                print("✅ Report detected! Waiting to meet minimum wait time...")

            if report_detected_time:
                remaining = min_wait_seconds - (time.time() - report_detected_time)
                if remaining <= 0:
                    print(f"✅ Minimum {min_wait_seconds}s wait completed after report detection.")
                    return
                else:
                    print(f"⏳ Waiting {remaining:.1f}s more before proceeding to download.")

        except Exception as e:
            print(f"⚠️ Exception during wait: {e}")

        if time.time() - start_time > timeout:
            screenshot_path = f"report_timeout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            browser.save_screenshot(screenshot_path)
            raise TimeoutError(f"❌ Report did not load in time. Screenshot saved: {screenshot_path}")

        time.sleep(2)
def wait_for_two_print_buttons(browser, timeout=120):
    print("\n=== Waiting for at least 2 print buttons to appear ===")
    start_time = time.time()
    while True:
        elapsed = int(time.time() - start_time)
        time_left = timeout - elapsed
        if time_left < 0:
            print("\n⚠️ Timeout waiting for 2 print buttons.")
            return False
        try:
            buttons = browser.find_elements(By.XPATH, "//a[contains(@title, 'Print in different format')]")
            if len(buttons) >= 2 and all(button.is_displayed() and button.is_enabled() for button in buttons[:2]):
                print(f"\n✅ Found {len(buttons)} print buttons.")
                return True
        except Exception:
            pass
        
        print(f"\r⏳ Waiting for 2 print buttons: {time_left:3d}s left...", end="")
        sys.stdout.flush()
        time.sleep(1)


def export_to_pdf(browser):
    print("\n=== Starting PDF Export Process ===")
    try:
        wait = WebDriverWait(browser, MAX_WAIT)
        print_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[@title='Print in different format' and @name='ReportLinkMenu']")))
        browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", print_button)
        print("🖨️ Clicking Print button...")
        browser.execute_script("arguments[0].click();", print_button)

        pdf_option_xpath = "//a[contains(.,'Printable PDF') and @role='menuitem']"
        pdf_option = wait.until(EC.element_to_be_clickable((By.XPATH, pdf_option_xpath)))
        print("📄 Clicking 'Printable PDF' option...")
        browser.execute_script("arguments[0].click();", pdf_option)
        print("✅ PDF export initiated")
        return True

    except Exception as e:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        browser.save_screenshot(f"pdf_export_error_{timestamp}.png")
        raise Exception(f"❌ PDF export failed: {str(e)}")

def verify_download(timeout=120):
    print("\n=== Verifying Download ===")
    start_time = time.time()
    initial_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.pdf')))
    while True:
        elapsed = int(time.time() - start_time)
        time_left = timeout - elapsed
        if time_left < 0:
            raise Exception(f"❌ Download not found or incomplete after {timeout}s")

        current_files = set(glob.glob(os.path.join(DOWNLOAD_DIR, '*.pdf')))
        new_files = current_files - initial_files
        if new_files:
            newest_pdf = max(new_files, key=os.path.getmtime)
            initial_size = os.path.getsize(newest_pdf)
            time.sleep(2)
            current_size = os.path.getsize(newest_pdf)
            if initial_size == current_size and current_size > 1024:
                print(f"✅ Download verified: {newest_pdf} ({current_size / 1024:.2f} KB)")
                return newest_pdf
            else:
                print(f"⏳ File still downloading: {newest_pdf} ({current_size} bytes) — {time_left}s left")
        else:
            print(f"⏳ No PDF found yet. Waiting... {time_left}s left")

        time.sleep(3)

def rename_latest_pdf(award_number, task_number, dest_path):
    print("\n=== Renaming and Moving downloaded PDF ===")
    pdf_files = glob.glob(os.path.join(DOWNLOAD_DIR, '*.pdf'))
    if not pdf_files:
        raise Exception("No PDF files found to rename!")
    latest_file = max(pdf_files, key=os.path.getmtime)

    if task_number:
        new_filename = f"{award_number}_{task_number}.pdf"
    else:
        new_filename = f"{award_number}.pdf"

    os.makedirs(dest_path, exist_ok=True)
    new_filepath = os.path.join(dest_path, new_filename)
    print(f"📁 Moving '{os.path.basename(latest_file)}' to '{new_filepath}'")
    shutil.move(latest_file, new_filepath)
    return new_filepath

def generate_report(browser, award_number, task_number, dest_path, max_retries=2):
    retry_count = 0
    while retry_count <= max_retries:
        try:
            print(f"\n=== Attempt {retry_count + 1} of {max_retries + 1} ===")
            click_apply_button(browser)
            time.sleep(3)
            wait_for_report(browser, timeout=REPORT_TIMEOUT)
            time.sleep(2)

            # WAIT for two print buttons before exporting PDF
            if not wait_for_two_print_buttons(browser):
                raise Exception("Less than 2 print buttons found, cannot proceed.")

            export_to_pdf(browser)
            downloaded_pdf_path = verify_download()
            new_pdf_path = rename_latest_pdf(award_number, task_number, dest_path)
            return new_pdf_path

        except Exception as e:
            retry_count += 1
            if retry_count > max_retries:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                browser.save_screenshot(f"final_attempt_error_{timestamp}.png")
                raise Exception(f"❌ Failed after {max_retries} attempts.\nLast error: {e}")
            print(f"⚠️ Retry #{retry_count} due to: {e}")
            browser.refresh()
            WebDriverWait(browser, MAX_WAIT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            enter_award_number(browser, award_number)
            enter_task_number(browser, task_number)
            time.sleep(3)


def main(award_number, task_number, dest_path, browser):
    try:
        print("\n" + "=" * 50)
        print(f"🚀 Starting report for Award: {award_number}, Task: {task_number}")
        print("=" * 50 + "\n")
        browser.refresh()
        WebDriverWait(browser, MAX_WAIT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        enter_award_number(browser, award_number)
        enter_task_number(browser, task_number)
        generate_report(browser, award_number, task_number, dest_path)
        print(f"\n✅ Finished report for {award_number} / {task_number}")

    except Exception as e:
        print(f"\n❌ Error for {award_number} / {task_number}: {e}")
        screenshot_name = f"error_{award_number}_{task_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        browser.save_screenshot(screenshot_name)
        print(f"🖼️ Screenshot saved as {screenshot_name}")

if __name__ == "__main__":
    browser = initialize_browser()
    try:
        browser.get(REPORT_URL)
        WebDriverWait(browser, MAX_WAIT).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        for idx, row in df.iterrows():
            AWARD_NUMBER = str(row["AWARD_NUMBER"]).strip()
            TASK_NUMBER = str(row["TASK_NUMBER"]).strip()
            DEST_PATH = str(row["DEST_PATH"]).strip()
            main(AWARD_NUMBER, TASK_NUMBER, DEST_PATH, browser)
    finally:
        browser.quit()
        print("🛑 Browser closed")
