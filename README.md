# 📒 Report Automation Tools

A Python automation toolkit for reducing repetitive reporting work through browser automation, Excel processing, PDF filtering, and summary generation.

This project started as a practical way to make recurring reporting workflows faster, cleaner, and less error-prone. Over time, it became a small collection of reusable scripts for downloading reports, processing workbooks, validating data, and generating structured outputs.

I think of this repository as an automation journal: each script solves a real workflow problem and documents part of my learning journey in Python automation. 🌱

---

## ✨ Project Highlights

* Automated repetitive report-downloading workflows using Selenium
* Processed and validated Excel workbooks using OpenPyXL and Pandas
* Filled monthly commitment, expense, and balance values into existing workbooks
* Supported both single-tab and multi-tab workbook structures
* Generated overdraft and PTA summary reports from Excel source data
* Filtered PDF pages based on award/task mappings
* Converted repetitive manual reporting steps into reusable automation workflows

---

## 🗂️ Project Structure

```text
report-automation-tools/
├── selenium_reports/
│   ├── download_expenditure_balance_reports.py
│   ├── download_expenditure_detail_reports.py
│   └── extract_award_task_codes.py
│
├── excel_processing/
│   ├── single_tab/
│   │   ├── fill_monthly_commitments.py
│   │   ├── fill_monthly_expense_balances.py
│   │   └── generate_overdraft_summary.py
│   │
│   ├── multi_tab/
│   │   ├── fill_current_month_commit_multitab.py
│   │   └── fill_current_month_expense_multitab_v3.py
│   │
│   └── utility_extract_merged_name_id_pairs.py
│
├── pdf_tools/
│   └── filter_pdf_by_award_task.py
│
├── requirements.txt
├── .gitignore
└── README.md
```

---

## 🧰 Tools Included

### 📥 Selenium Report Automation

| Script                                    | What it does                                                                         |
| ----------------------------------------- | ------------------------------------------------------------------------------------ |
| `download_expenditure_balance_reports.py` | Automates batch downloading of balance-style PDF reports from a report dashboard     |
| `download_expenditure_detail_reports.py`  | Automates batch downloading of detail-style PDF reports from a report dashboard      |
| `extract_award_task_codes.py`             | Extracts project, task, and award-style codes from a webpage and saves them to Excel |

---

### 📊 Excel Processing — Single-Tab Workbooks

| Script                                        | What it does                                                                                 |
| --------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `single_tab/fill_monthly_commitments.py`      | Fills monthly commitment values and commitment detail tables into single-tab Excel workbooks |
| `single_tab/fill_monthly_expense_balances.py` | Fills monthly expense and balance values into single-tab workbooks and validates totals      |
| `single_tab/generate_overdraft_summary.py`    | Generates a clean overdraft summary workbook from Excel source data                          |

---

### 📊 Excel Processing — Multi-Tab Workbooks

| Script                                                | What it does                                                                                    |
| ----------------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| `multi_tab/fill_current_month_commit_multitab.py`     | Fills current-month commitment values across multiple workbook tabs                             |
| `multi_tab/fill_current_month_expense_multitab_v3.py` | Fills current-month expense values across multiple workbook tabs and writes a processing report |

---

### 🧩 Excel Utility Scripts

| Script                                    | What it does                                                                          |
| ----------------------------------------- | ------------------------------------------------------------------------------------- |
| `utility_extract_merged_name_id_pairs.py` | Extracts name/ID-style pairs from merged-cell Excel worksheets and writes them to CSV |

---

### 📄 PDF Processing

| Script                        | What it does                                                                          |
| ----------------------------- | ------------------------------------------------------------------------------------- |
| `filter_pdf_by_award_task.py` | Filters PDF pages based on allowed award/task combinations from an Excel mapping file |

---

## 🛠️ Tech Stack

* Python
* Selenium
* Pandas
* OpenPyXL
* PyMuPDF
* WebDriver Manager
* Git / GitHub

---

## 💡 Why I Built This

Reporting work often includes many repeated manual steps, such as downloading files, checking Excel reports, copying values, validating totals, and filtering long PDFs.

I built these tools to improve productivity by reducing repetitive manual steps and making reporting workflows more efficient, consistent, and reusable.

This project helped me practice:

* browser automation
* file handling
* Excel workbook processing
* PDF text extraction and filtering
* data validation
* error handling
* reusable script design
* organizing code for a public GitHub portfolio

---

## 🚀 Getting Started

Clone the repository:

```bash
git clone https://github.com/kawingtam/report-automation-tools.git
cd report-automation-tools
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## 🌷 Example Usage

Run a script directly with Python:

```bash
python excel_processing/single_tab/generate_overdraft_summary.py
```

Or run a multi-tab workbook automation script:

```bash
python excel_processing/multi_tab/fill_current_month_commit_multitab.py
```

Most scripts prompt the user for inputs such as:

* source Excel file path
* target workbook or folder path
* output folder path
* month label
* report/dashboard URL

This keeps the scripts reusable without exposing private local paths, credentials, or internal links.

---

## 🔒 Privacy & Safety

This public version does **not** include:

* private report files
* credentials
* internal dashboard URLs
* personal local file paths
* confidential input data
* real output workbooks or PDFs

The scripts are designed to accept user-provided paths and URLs at runtime.

---

## 📝 Project Reflection

This project represents how I like to solve workflow problems:

1. Notice a repetitive manual process
2. Break the process into clear steps
3. Build a small automation tool
4. Test it with realistic files
5. Add validation and error handling
6. Clean the script so it can be reused safely

It is both a practical automation toolkit and a record of my growth in Python scripting. ✨

---

## 🌟 Status

Organized as a public portfolio project and open to future improvements.
