# 📒 Report Automation Tools

A Python automation toolkit for making reporting workflows faster, cleaner, and less repetitive.

This project started as a practical way to reduce manual reporting work. Over time, it became a small collection of tools for downloading reports, processing Excel workbooks, generating summaries, and filtering PDF files.

I think of this repository as an automation journal: each script solves a real workflow problem and documents part of my learning journey in Python automation. 🌱

---

## ✨ Project Highlights

- Built reusable Python scripts for repetitive reporting tasks
- Automated report downloading with Selenium
- Processed and validated Excel workbooks using OpenPyXL and Pandas
- Filtered PDF pages based on award/task mappings
- Generated structured summary reports from pivot-style Excel files
- Improved productivity by turning repetitive manual reporting steps into reusable automation workflows


---

## 🗂️ Project Structure

~~~text
report-automation-tools/
├── selenium_reports/
│   ├── download_expenditure_balance_reports.py
│   ├── download_expenditure_detail_reports.py
│   └── extract_award_task_codes.py
│
├── excel_processing/
│   ├── fill_monthly_commitments.py
│   ├── fill_monthly_expense_balances.py
│   └── generate_overdraft_summary.py
│
├── pdf_tools/
│   └── filter_pdf_by_award_task.py
│
├── requirements.txt
└── README.md
~~~

---

## 🧰 Tools Included

### 📥 Selenium Report Automation

| Script | What it does |
|---|---|
| `download_expenditure_balance_reports.py` | Automates batch downloading of expenditure balance PDF reports |
| `download_expenditure_detail_reports.py` | Automates batch downloading of expenditure detail PDF reports |
| `extract_award_task_codes.py` | Extracts project, task, and award codes from a webpage and saves them to Excel |

### 📊 Excel Processing

| Script | What it does |
|---|---|
| `fill_monthly_commitments.py` | Fills monthly commitment values and commitment detail tables into Excel workbooks |
| `fill_monthly_expense_balances.py` | Fills monthly expense values into balance workbooks and validates remaining balances |
| `generate_overdraft_summary.py` | Reads a pivot-style Excel report and generates a clean overdraft summary workbook |

### 📄 PDF Processing

| Script | What it does |
|---|---|
| `filter_pdf_by_award_task.py` | Filters PDF pages based on allowed award/task combinations from an Excel file |

---

## 🛠️ Tech Stack

- Python
- Selenium
- Pandas
- OpenPyXL
- PyMuPDF
- WebDriver Manager
- Git / GitHub

---

## 💡 Why I Built This

Reporting work often includes many repeated manual steps, such as downloading files, checking Excel reports, copying values, validating totals, and filtering long PDFs.

I built these tools to improve productivity by reducing repetitive manual steps, making the reporting process more efficient, consistent, and less error-prone.

This project helped me practice:

- workflow automation
- file handling
- Excel data processing
- PDF processing
- Selenium browser automation
- writing reusable scripts
- organizing code for a public GitHub portfolio

---

## 🚀 Getting Started

Clone the repository:

~~~bash
git clone https://github.com/kawingtam/report-automation-tools.git
cd report-automation-tools
~~~

Install dependencies:

~~~bash
pip install -r requirements.txt
~~~

---

## 🌷 Example Usage

Run a script directly with Python:

~~~bash
python excel_processing/generate_overdraft_summary.py
~~~

Most scripts prompt the user for inputs such as:

- source Excel file path
- input folder path
- output folder path
- month label
- report/dashboard URL

This keeps the scripts reusable without exposing private local paths or internal links.

---

## 🔒 Privacy & Safety

This public version does **not** include:

- private report files
- credentials
- internal dashboard URLs
- personal local file paths
- confidential input data

The scripts are designed to accept user-provided paths and URLs at runtime.

---

## 📝 Project Reflection

This project represents how I like to solve problems:

1. Notice a repetitive workflow  
2. Build a small tool to reduce manual work  
3. Test it with real files  
4. Improve the script as new issues appear  
5. Organize the final version so it can be reused  

It is both a practical automation toolkit and a record of my growth in Python scripting. ✨

---

## 🌟 Status

Organized as a portfolio project and open to future improvements.
