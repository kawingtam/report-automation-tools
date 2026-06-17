import os
import re
import sys
import time
import glob
import shutil
from datetime import datetime

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    WebDriverException,
)

from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# Configurations
# ============================================================

DOWNLOAD_DIR = input("Enter temporary download folder path: ").strip().strip('"')

INPUT_FILE = input("Enter award/task input Excel file path: ").strip().strip('"')

REPORT_URL = input("Enter report dashboard URL: ").strip().strip('"')

MAX_WAIT = 300
REPORT_TIMEOUT = 600
DOWNLOAD_TIMEOUT = 180
MIN_WAIT_AFTER_REPORT_DETECTED = 40


# ============================================================
# Helper functions
# ============================================================

def clean_excel_value(value):
    """
    Clean values read from Excel.
    Prevents blank cells from becoming 'nan'.
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() == "nan":
        return ""

    if text.endswith(".0"):
        text = text[:-2]

    return text


def safe_filename_part(text):
    """
    Remove characters that are invalid in Windows filenames.
    """
    text = str(text).strip()
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    return text


def load_award_input():
    """
    Load award/task/destination rows from the input Excel file.

    Required columns:
    - AWARD_NUMBER
    - TASK_NUMBER
    - DEST_PATH
    """
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Input file not found: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, dtype=str)
    df = df.fillna("")

    required_columns = ["AWARD_NUMBER", "TASK_NUMBER", "DEST_PATH"]
    missing_columns = [c for c in required_columns if c not in df.columns]

    if missing_columns:
        raise Exception(
            "Input Excel file is missing required column(s): "
            + ", ".join(missing_columns)
        )

    rows = []

    for idx, row in df.iterrows():
        award_number = clean_excel_value(row["AWARD_NUMBER"])
        task_number = clean_excel_value(row["TASK_NUMBER"])
        dest_path = clean_excel_value(row["DEST_PATH"])

        if not award_number:
            print(f"⚠️ Skipping row {idx + 2}: missing AWARD_NUMBER")
            continue

        if not dest_path:
            print(
                f"⚠️ Row {idx + 2}: DEST_PATH is blank. "
                f"Using DOWNLOAD_DIR instead."
            )
            dest_path = DOWNLOAD_DIR

        rows.append(
            {
                "AWARD_NUMBER": award_number,
                "TASK_NUMBER": task_number,
                "DEST_PATH": dest_path,
            }
        )

    if not rows:
        raise Exception("No valid award rows found in the input Excel file.")

    return rows


def initialize_browser():
    print("🛠️ Initializing browser configuration...")

    chrome_options = Options()
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
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


def hide_page_banner(browser):
    """
    Hide a temporary page banner if it appears.
    This avoids Selenium click failures caused by the banner blocking or shifting elements.
    """
    try:
        hidden_count = browser.execute_script(
            """
            const keywords = [
                "temporary outage",
                "system maintenance",
                "planned maintenance",
                "data freeze",
                "systems impacted"
            ];

            const nodes = Array.from(
                document.querySelectorAll("div, table, tbody, tr, td, span")
            );

            let hidden = 0;

            for (const el of nodes) {
                const text = (el.innerText || "").trim();
                if (!text) continue;

                const matches = keywords.some(k => text.includes(k));
                if (!matches) continue;

                let target = el;

                while (target.parentElement) {
                    const r = target.getBoundingClientRect();
                    const pr = target.parentElement.getBoundingClientRect();

                    const parentLooksLikeBanner =
                        pr.top >= -10 &&
                        pr.top <= 180 &&
                        pr.height <= 200 &&
                        pr.width >= window.innerWidth * 0.5;

                    if (parentLooksLikeBanner) {
                        target = target.parentElement;
                    } else {
                        break;
                    }
                }

                const r = target.getBoundingClientRect();

                if (
                    r.top >= -10 &&
                    r.top <= 180 &&
                    r.height <= 200 &&
                    r.width >= window.innerWidth * 0.5
                ) {
                    target.style.setProperty("display", "none", "important");
                    hidden++;
                }
            }

            return hidden;
            """
        )

        if hidden_count:
            print(f"✅ Hidden temporary banner element(s): {hidden_count}")

    except Exception as e:
        print(f"⚠️ Could not hide temporary banner, continuing anyway: {e}")


def wait_for_page_ready(browser, timeout=MAX_WAIT):
    """
    Wait for page body and document ready state.
    Then hide any temporary page banner.
    """
    WebDriverWait(browser, timeout).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )

    try:
        WebDriverWait(browser, 30).until(
            lambda d: d.execute_script("return document.readyState")
            in ["interactive", "complete"]
        )
    except Exception:
        pass

    time.sleep(1)
    hide_page_banner(browser)


def safe_click(browser, element, description="element"):
    """
    Safer Selenium click:
    - hides temporary page banner
    - scrolls element to center
    - tries normal click
    - falls back to JavaScript click if normal click is blocked
    """
    hide_page_banner(browser)

    try:
        browser.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            element,
        )
        time.sleep(0.5)

        try:
            element.click()
        except ElementClickInterceptedException:
            print(f"⚠️ Normal click intercepted for {description}; using JavaScript click.")
            browser.execute_script("arguments[0].click();", element)

    except StaleElementReferenceException:
        raise

    except WebDriverException as e:
        print(f"⚠️ Normal click failed for {description}; trying JavaScript click.")
        try:
            browser.execute_script("arguments[0].click();", element)
        except Exception:
            raise e


def clear_and_type(browser, element, text, field_name):
    """
    Clear an input field and type text.
    Uses CTRL+A instead of only clear(), which works better on many web form fields.
    """
    safe_click(browser, element, field_name)

    try:
        element.send_keys(Keys.CONTROL, "a")
        time.sleep(0.2)
        element.send_keys(Keys.BACKSPACE)
    except Exception:
        browser.execute_script(
            """
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """,
            element,
        )

    time.sleep(0.3)

    if text:
        element.send_keys(text)

    time.sleep(0.3)
    element.send_keys(Keys.TAB)


def optimized_locator(browser, field_name, locators):
    print(f"🔍 Locating element: {field_name}")

    last_error = None

    for by, locator in locators:
        try:
            hide_page_banner(browser)

            element = WebDriverWait(browser, MAX_WAIT).until(
                EC.presence_of_element_located((by, locator))
            )

            browser.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                element,
            )

            browser.execute_script(
                "arguments[0].style.border='2px solid green';",
                element,
            )

            print(f"✅ Located {field_name} using: {by}='{locator}'")
            return element

        except Exception as e:
            last_error = e
            print(f"⚠️ Attempt failed with {by}='{locator}': {e}")

    screenshot_name = (
        f"element_not_found_{field_name.replace(' ', '_')}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    )
    browser.save_screenshot(screenshot_name)

    raise Exception(
        f"❌ Could not locate {field_name}. "
        f"Screenshot saved: {screenshot_name}. "
        f"Last error: {last_error}"
    )


# ============================================================
# Report field entry
# ============================================================

def enter_award_number(browser, award_number):
    print("\n=== Entering Award Number ===")
    hide_page_banner(browser)

    award_locators = [
        (By.XPATH, "//label[contains(.,'Award Number')]/following::input[1]"),
        (By.NAME, "award_number"),
        (
            By.XPATH,
            "//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'award')]",
        ),
    ]

    field = optimized_locator(browser, "Award Number", award_locators)
    clear_and_type(browser, field, award_number, "Award Number")

    print(f"✏️ Entered Award Number: {award_number}")


def enter_task_number(browser, task_number):
    print("\n=== Entering Task Number ===")
    hide_page_banner(browser)

    task_locators = [
        (By.XPATH, "//label[contains(.,'Task Number')]/following::input[1]"),
        (By.NAME, "task_number"),
        (
            By.XPATH,
            "//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'task')]",
        ),
    ]

    field = optimized_locator(browser, "Task Number", task_locators)
    clear_and_type(browser, field, task_number, "Task Number")

    if task_number:
        print(f"✏️ Entered Task Number: {task_number}")
    else:
        print("✏️ Task Number is blank; field was cleared.")


def click_apply_button(browser):
    print("\n=== Clicking Apply Button ===")
    hide_page_banner(browser)

    wait = WebDriverWait(browser, MAX_WAIT)

    apply_buttons = wait.until(
        lambda d: [
            b
            for b in d.find_elements(
                By.XPATH,
                "//input[(@value='Apply' or @title='Apply') and contains(@class,'promptApplyButton')]",
            )
            if b.is_displayed() and b.is_enabled()
        ]
    )

    apply_button = apply_buttons[-1]
    safe_click(browser, apply_button, "Apply button")

    print("🔄 Apply button clicked - generating report...")


# ============================================================
# Report wait and PDF export
# ============================================================

def wait_for_report(browser, timeout=REPORT_TIMEOUT, min_wait_seconds=MIN_WAIT_AFTER_REPORT_DETECTED):
    print("\n=== Waiting for Report Generation ===")

    start_time = time.time()
    report_detected_time = None

    report_xpath = (
        "//*[contains(@class,'PivotTable') "
        "or contains(@id,'report') "
        "or contains(@class,'ReportContent') "
        "or contains(@class,'ViewContent')]"
    )

    last_printed_elapsed = -1

    while True:
        hide_page_banner(browser)

        elapsed = int(time.time() - start_time)

        if elapsed % 10 == 0 and elapsed != last_printed_elapsed:
            print(f"⏳ Still waiting... {elapsed}s elapsed")
            last_printed_elapsed = elapsed

        try:
            elements = browser.find_elements(By.XPATH, report_xpath)
            visible_elements = [e for e in elements if e.is_displayed()]

            if visible_elements and report_detected_time is None:
                report_detected_time = time.time()
                print("✅ Report detected! Waiting to meet minimum wait time...")

            if report_detected_time:
                remaining = min_wait_seconds - (time.time() - report_detected_time)

                if remaining <= 0:
                    print(f"✅ Minimum {min_wait_seconds}s wait completed after report detection.")
                    return

                print(f"⏳ Waiting {remaining:.1f}s more before proceeding to download.")

        except Exception as e:
            print(f"⚠️ Exception during report wait: {e}")

        if time.time() - start_time > timeout:
            screenshot_path = f"report_timeout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            browser.save_screenshot(screenshot_path)
            raise TimeoutError(
                f"❌ Report did not load in time. Screenshot saved: {screenshot_path}"
            )

        time.sleep(2)


def wait_for_print_buttons(browser, timeout=120, preferred_count=3):
    print("\n=== Waiting for print buttons to appear ===")

    start_time = time.time()

    while True:
        hide_page_banner(browser)

        elapsed = int(time.time() - start_time)
        time_left = timeout - elapsed

        if time_left < 0:
            print("\n⚠️ Timeout waiting for print buttons.")
            return []

        try:
            buttons = browser.find_elements(
                By.XPATH,
                "//a[contains(@title, 'Print in different format') and @name='ReportLinkMenu']",
            )

            visible_buttons = []

            for b in buttons:
                try:
                    if not b.is_displayed() or not b.is_enabled():
                        continue

                    rect = browser.execute_script(
                        """
                        const r = arguments[0].getBoundingClientRect();
                        return {
                            top: r.top,
                            left: r.left,
                            width: r.width,
                            height: r.height
                        };
                        """,
                        b,
                    )

                    if rect["width"] <= 0 or rect["height"] <= 0:
                        continue

                    visible_buttons.append((rect["top"], rect["left"], b))

                except StaleElementReferenceException:
                    continue

            # Sort visually from top-to-bottom, then left-to-right.
            # This makes "first button" mean the first visible one on the page.
            visible_buttons.sort(key=lambda x: (x[0], x[1]))

            sorted_buttons = [item[2] for item in visible_buttons]

            if len(sorted_buttons) >= preferred_count:
                print(f"\n✅ Found {len(sorted_buttons)} visible print buttons.")
                print("✅ Using the FIRST visible print button.")
                return sorted_buttons

            # Fallback: if at least one print button is visible for a while, allow it.
            if len(sorted_buttons) >= 1 and elapsed >= 30:
                print(
                    f"\n✅ Found {len(sorted_buttons)} visible print button(s). "
                    "Proceeding with the first available button."
                )
                return sorted_buttons

        except Exception as e:
            print(f"\n⚠️ Error while checking print buttons: {e}")

        print(f"\r⏳ Waiting for print buttons: {time_left:3d}s left...", end="")
        sys.stdout.flush()
        time.sleep(1)


def export_to_pdf(browser):
    print("\n=== Starting PDF Export Process ===")

    try:
        hide_page_banner(browser)

        # The report page should normally show multiple matching print buttons.
        print_buttons = wait_for_print_buttons(
            browser,
            timeout=120,
            preferred_count=3,
        )

        if not print_buttons:
            raise Exception("No visible print buttons found, cannot proceed.")

        # IMPORTANT:
        # Click the FIRST visible print button, not the last one.
        print_button = print_buttons[0]

        print("🖨️ Clicking FIRST Print button...")
        safe_click(browser, print_button, "First Print button")

        hide_page_banner(browser)

        wait = WebDriverWait(browser, MAX_WAIT)

        pdf_option_xpath = "//a[contains(.,'Printable PDF') and @role='menuitem']"
        pdf_option = wait.until(
            EC.element_to_be_clickable((By.XPATH, pdf_option_xpath))
        )

        print("📄 Clicking 'Printable PDF' option...")
        safe_click(browser, pdf_option, "Printable PDF option")

        print("✅ PDF export initiated")
        return True

    except Exception as e:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_name = f"pdf_export_error_{timestamp}.png"
        browser.save_screenshot(screenshot_name)
        raise Exception(f"❌ PDF export failed: {e}. Screenshot saved: {screenshot_name}")


# ============================================================
# Download verification and file moving
# ============================================================

def get_pdf_files():
    return set(glob.glob(os.path.join(DOWNLOAD_DIR, "*.pdf")))


def verify_download(initial_pdf_files, timeout=DOWNLOAD_TIMEOUT):
    print("\n=== Verifying Download ===")

    start_time = time.time()

    while True:
        elapsed = int(time.time() - start_time)
        time_left = timeout - elapsed

        if time_left < 0:
            raise Exception(f"❌ Download not found or incomplete after {timeout}s")

        current_pdf_files = get_pdf_files()
        new_pdf_files = current_pdf_files - initial_pdf_files

        partial_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))

        if new_pdf_files:
            newest_pdf = max(new_pdf_files, key=os.path.getmtime)

            initial_size = os.path.getsize(newest_pdf)
            time.sleep(2)
            current_size = os.path.getsize(newest_pdf)

            if initial_size == current_size and current_size > 1024:
                print(
                    f"✅ Download verified: {newest_pdf} "
                    f"({current_size / 1024:.2f} KB)"
                )
                return newest_pdf

            print(
                f"⏳ PDF still downloading: {newest_pdf} "
                f"({current_size} bytes) — {time_left}s left"
            )

        elif partial_files:
            print(f"⏳ Chrome download still in progress — {time_left}s left")

        else:
            print(f"⏳ No new PDF found yet. Waiting... {time_left}s left")

        time.sleep(3)


def rename_and_move_pdf(downloaded_pdf_path, award_number, task_number, dest_path):
    print("\n=== Renaming and Moving Downloaded PDF ===")

    if not downloaded_pdf_path or not os.path.exists(downloaded_pdf_path):
        raise Exception(f"Downloaded PDF not found: {downloaded_pdf_path}")

    award_safe = safe_filename_part(award_number)
    task_safe = safe_filename_part(task_number)

    if task_safe:
        new_filename = f"{award_safe}_{task_safe}.pdf"
    else:
        new_filename = f"{award_safe}.pdf"

    os.makedirs(dest_path, exist_ok=True)

    new_filepath = os.path.join(dest_path, new_filename)

    if os.path.exists(new_filepath):
        print(f"⚠️ Existing file found and will be overwritten: {new_filepath}")
        os.remove(new_filepath)

    print(f"📁 Moving '{os.path.basename(downloaded_pdf_path)}' to '{new_filepath}'")
    shutil.move(downloaded_pdf_path, new_filepath)

    print(f"✅ PDF saved as: {new_filepath}")
    return new_filepath


# ============================================================
# Main report generation workflow
# ============================================================

def reset_report_page_and_enter_values(browser, award_number, task_number):
    """
    Reload the report page, hide the banner, and enter award/task values.
    Used before each report and again after retry failures.
    """
    print("\n=== Loading Report Page ===")
    browser.get(REPORT_URL)
    wait_for_page_ready(browser)

    enter_award_number(browser, award_number)
    enter_task_number(browser, task_number)

    time.sleep(2)
    hide_page_banner(browser)


def generate_report(browser, award_number, task_number, dest_path, max_retries=2):
    retry_count = 0

    while retry_count <= max_retries:
        try:
            print(f"\n=== Attempt {retry_count + 1} of {max_retries + 1} ===")

            hide_page_banner(browser)

            click_apply_button(browser)

            time.sleep(3)
            hide_page_banner(browser)

            wait_for_report(browser, timeout=REPORT_TIMEOUT)

            time.sleep(2)
            hide_page_banner(browser)

            initial_pdf_files = get_pdf_files()

            export_to_pdf(browser)

            downloaded_pdf_path = verify_download(
                initial_pdf_files=initial_pdf_files,
                timeout=DOWNLOAD_TIMEOUT,
            )

            new_pdf_path = rename_and_move_pdf(
                downloaded_pdf_path=downloaded_pdf_path,
                award_number=award_number,
                task_number=task_number,
                dest_path=dest_path,
            )

            return new_pdf_path

        except Exception as e:
            retry_count += 1

            if retry_count > max_retries:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_name = (
                    f"final_attempt_error_"
                    f"{safe_filename_part(award_number)}_"
                    f"{safe_filename_part(task_number)}_"
                    f"{timestamp}.png"
                )
                browser.save_screenshot(screenshot_name)

                raise Exception(
                    f"❌ Failed after {max_retries + 1} attempts.\n"
                    f"Last error: {e}\n"
                    f"Screenshot saved: {screenshot_name}"
                )

            print(f"\n⚠️ Retry #{retry_count} due to: {e}")
            print("🔄 Reloading report page and re-entering values before retry...")

            reset_report_page_and_enter_values(browser, award_number, task_number)


def process_one_report(browser, award_number, task_number, dest_path):
    try:
        print("\n" + "=" * 70)
        print(f"🚀 Starting report for Award: {award_number}, Task: {task_number}")
        print("=" * 70)

        reset_report_page_and_enter_values(browser, award_number, task_number)

        output_pdf = generate_report(
            browser=browser,
            award_number=award_number,
            task_number=task_number,
            dest_path=dest_path,
            max_retries=2,
        )

        print(f"\n✅ Finished report for {award_number} / {task_number}")
        print(f"📄 Output PDF: {output_pdf}")

        return True

    except Exception as e:
        print(f"\n❌ Error for {award_number} / {task_number}: {e}")

        screenshot_name = (
            f"error_{safe_filename_part(award_number)}_"
            f"{safe_filename_part(task_number)}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        )

        try:
            browser.save_screenshot(screenshot_name)
            print(f"🖼️ Screenshot saved as {screenshot_name}")
        except Exception:
            pass

        return False


# ============================================================
# Program entry point
# ============================================================

if __name__ == "__main__":
    rows = load_award_input()

    browser = initialize_browser()

    success_count = 0
    fail_count = 0

    try:
        print("\nOpening report page...")
        print("If login page appears, complete login in the browser window.")
        print("The program will wait for the report fields to become available.\n")

        browser.get(REPORT_URL)
        wait_for_page_ready(browser)

        for idx, row in enumerate(rows, start=1):
            award_number = row["AWARD_NUMBER"]
            task_number = row["TASK_NUMBER"]
            dest_path = row["DEST_PATH"]

            print("\n" + "#" * 70)
            print(f"Processing row {idx} of {len(rows)}")
            print("#" * 70)

            ok = process_one_report(
                browser=browser,
                award_number=award_number,
                task_number=task_number,
                dest_path=dest_path,
            )

            if ok:
                success_count += 1
            else:
                fail_count += 1

    finally:
        browser.quit()
        print("\n🛑 Browser closed")

        print("\n" + "=" * 70)
        print("RUN SUMMARY")
        print("=" * 70)
        print(f"✅ Successful reports: {success_count}")
        print(f"❌ Failed reports: {fail_count}")
        print("=" * 70)