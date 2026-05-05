import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os
from datetime import datetime


def extract_codes_from_page(driver):
    time.sleep(2)
    page_source = driver.page_source
    pattern = r"\b(\d{7})-(\d{1,6})-([A-Z]{5})\b"
    matches = re.findall(pattern, page_source)
    return matches


def group_and_save_to_excel(all_extractions, output_file):
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    rows = []

    for extraction in all_extractions:
        codes = extraction["codes"]
        note = extraction["note"]

        grouped = {}

        for project_num, task_num, award_num in codes:
            if award_num not in grouped:
                grouped[award_num] = {
                    "project_num": project_num,
                    "task_nums": set()
                }

            grouped[award_num]["task_nums"].add(task_num)

        for award_num, data in grouped.items():
            project_num = data["project_num"]
            task_nums = ";".join(sorted(data["task_nums"], key=int))
            rows.append([project_num, task_nums, award_num, note])

    new_df = pd.DataFrame(
        rows,
        columns=["Project Number", "Task Numbers", "Award Number", "Note"]
    )

    if os.path.exists(output_file):
        existing_df = pd.read_excel(output_file)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    combined_df.to_excel(output_file, index=False)

    print(f"\nData appended and saved to: {output_file}")


def main():
    output_dir = input("Enter output folder path: ").strip().strip('"')
    target_url = input("Enter website URL to extract from: ").strip().strip('"')

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_dir, f"output_{timestamp}.xlsx")

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    options.add_argument("--log-level=3")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    driver.get(target_url)

    print("Waiting 60 seconds to allow manual login...")
    time.sleep(60)

    all_extractions = []

    while True:
        ready = input("Ready to extract from this page? (yes/no): ").strip().lower()

        if ready not in ["yes", "y"]:
            print("Exiting extraction loop.")
            break

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

        codes = extract_codes_from_page(driver)

        if not codes:
            print("No codes found on this page.")
        else:
            note = input("Enter a note for this extraction, for example 'current': ").strip()
            all_extractions.append({
                "codes": codes,
                "note": note
            })
            print(f"Extracted {len(codes)} codes with note '{note}'.")

        more = input("Run extraction on another page? (yes/no): ").strip().lower()

        if more not in ["yes", "y"]:
            break

    if all_extractions:
        group_and_save_to_excel(all_extractions, output_file)
    else:
        print("No data extracted, nothing saved.")


if __name__ == "__main__":
    main()