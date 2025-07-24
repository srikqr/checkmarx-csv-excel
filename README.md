Checkmarx CSV → Consolidated Excel Report
A lightweight, open-source utility by srikqr – MIT License

✨ What this tool does
Transforms a raw Checkmarx CSV export into a polished, management-ready Excel workbook:

Feature	Details
🛠 Auto-installs dependencies	Installs pandas, openpyxl, xlsxwriter, and numpy on first run.
📑 Cleans & consolidates	Drops findings with blank/“None” severity; groups by Vulnerability Type; merges all instances with clear numbering and a blank line between each.
🎯 Severity aware	Recognises Critical, High, Medium, Low, and Information; colours cells and builds a dedicated sheet for each level.
📊 Beautiful Excel output	Wide wrapped columns, frozen header row, auto-filter, zebra striping, and colour-coded Severity cells.
⚡ Zero-config CLI	One command converts CSV → XLSX in seconds.
📂 Folder structure
text
.
├─ checkmarx_consolidated_final.py   ← single-file script
└─ README.md                         ← you’re here
🚀 Quick-start
bash
# 1 ) Clone or copy the script
git clone https://github.com/srikqr/checkmarx-csv-excel.git
cd checkmarx-csv-excel

# 2 ) Run against a Checkmarx CSV export
python checkmarx_consolidated_final.py CheckmarxReport.csv FinalReport.xlsx

# 3 ) Open FinalReport.xlsx — each severity has its own tab!
Tip: Add the command to your CI pipeline so every scan automatically generates a board-ready XLSX.

🔍 How the workbook looks
Sheet	What’s inside
All Findings	Every vulnerability (one row per Vulnerability Type)
Critical Issues	Only rows whose highest severity is Critical
High Issues	…likewise for High
Medium Issues	…
Low Issues	…
Information Issues	Informational findings (if any)
Column layout
Column	Meaning
Vulnerability Type	Checkmarx “Query” name
Occurrences	Numbered list of instances, comma-separated fields, blank line between each
Severity	Highest severity in the group (Critical ▶ Information)
Total Findings	Count of merged instances
Occurrence example
text
1) Source File=src/auth/login.js, Line=53, Column=12, Function=loginHandler

2) Source File=src/auth/register.js, Line=78, Column=8, Function=registerUser
🛠️ Installation notes
Python 3.7 + recommended.

The script will pip install any missing packages automatically; internet access is therefore required on first run.

⚙️ Configuration
If you need to change which CSV fields appear in the Occurrences list:

Open checkmarx_consolidated_final.py.

Edit the DETAIL_FIELDS list (raw CSV column names).

Update the matching LABEL_MAP entries for friendlier labels.

Save and re-run — no other tweaks required.

📄 License
Released under the MIT License — free for personal and commercial use.
Copyright © 2025 srikqr

You are welcome to clone, fork, modify, and redistribute the script.
Please keep this README and the original license header intact.

🤝 Contributing
Pull requests and issues are very welcome!

Fork the repo

Create a feature branch

Commit your changes with clear messages

Open a PR — we’ll review ASAP

Happy reporting!
