import re
import pandas as pd
import fitz  # PyMuPDF
import os


def normalize(text):
    return str(text).strip().upper()


def extract_task_numbers(text):
    return re.findall(r"\d+", str(text))


def parse_allowed_table(excel_path):
    df = pd.read_excel(excel_path)

    allowed_pairs = set()

    for _, row in df.iterrows():
        award = normalize(row["Award"])
        tasks = extract_task_numbers(row["PI Task"])

        for task in tasks:
            allowed_pairs.add((task, award))

    print(f"Loaded {len(allowed_pairs)} allowed task/award combinations.")
    return allowed_pairs


def extract_tasks_and_award(page_text):
    """
    Reads a PDF format that contains project, task, and award information.

    Example format:

    Category: Project Info Task Info Award Info
    Number:   1234567     619       ABCDE

    Returns:
    tasks = ["619"]
    award = "ABCDE"
    """

    # Primary method for your PDF
    match = re.search(
        r"Number:\s+(\d+)\s+(\d+)\s+([A-Za-z0-9_-]+)",
        page_text,
        re.IGNORECASE
    )

    if match:
        task = match.group(2)
        award = normalize(match.group(3))
        return [task], award

    # Backup method if line breaks are weird
    text = page_text.replace("\n", " ")

    match = re.search(
        r"Category:\s*Project Info\s*Task Info\s*Award Info\s*Number:\s+(\d+)\s+(\d+)\s+([A-Za-z0-9_-]+)",
        text,
        re.IGNORECASE
    )

    if match:
        task = match.group(2)
        award = normalize(match.group(3))
        return [task], award

    return [], None


def filter_pdf(pdf_path, allowed_pairs, output_path, result_log):
    src = fitz.open(pdf_path)
    out = fitz.open()

    results = []
    kept = 0
    deleted = 0

    for page_num in range(len(src)):
        page = src[page_num]
        text = page.get_text("text")

        page_tasks, award = extract_tasks_and_award(text)

        matched_pairs = []
        should_keep = False

        if award and page_tasks:
            for task in page_tasks:
                if (task, award) in allowed_pairs:
                    should_keep = True
                    matched_pairs.append(f"{task}-{award}")

        results.append({
            "page": page_num + 1,
            "tasks_found": ",".join(page_tasks),
            "award_found": award if award else "",
            "matched_pairs": ",".join(matched_pairs),
            "action": "KEEP" if should_keep else "DELETE"
        })

        if should_keep:
            out.insert_pdf(src, from_page=page_num, to_page=page_num)
            kept += 1
        else:
            deleted += 1

    if kept == 0:
        print("\nWARNING: No pages matched your table. Nothing saved.")
        print("Check results.csv to see what task/award values were detected.")
    else:
        out.save(output_path)
        print(f"\nSaved filtered PDF: {output_path}")

    out.close()
    src.close()

    pd.DataFrame(results).to_csv(result_log, index=False)

    print(f"Log saved: {result_log}")
    print(f"\nSummary: Kept={kept}, Deleted={deleted}")


def main():
    print("=== PDF FILTER TOOL ===")

    pdf_path = input("Enter FULL path to PDF: ").strip().strip('"')
    table_path = input("Enter FULL path to Excel table: ").strip().strip('"')
    output_name = input("Enter OUTPUT PDF name (e.g. result.pdf): ").strip().strip('"')

    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    output_dir = os.path.dirname(pdf_path)
    output_path = os.path.join(output_dir, output_name)
    result_log = os.path.join(output_dir, "results.csv")

    allowed_pairs = parse_allowed_table(table_path)
    filter_pdf(pdf_path, allowed_pairs, output_path, result_log)

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()