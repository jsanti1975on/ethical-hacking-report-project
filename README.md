# ethical-hacking-report-project
This is a temp report for the student. The idea is to automate daily findings. No screenshots or exploits will live in the repo.

# 🛡️ Automated Pentest Daily Report Builder

A fully local, VBA-driven system for generating structured daily wrap-up reports during a month-long ethical hacking exam. Built with Excel + Word automation. No internet required. No cloud integration.

---

## 📌 Overview

This project automates the creation of daily penetration test reports using:

- ✔️ Microsoft Excel (for tracking findings)
- ✔️ Word (for professional report formatting)
- ✔️ VBA (to tie it all together)
- ✔️ Local screenshot evidence folders

Each report summarizes key findings, includes screenshots, and formats them for submission or presentation — all fully offline.

---

## ⚙️ Features

- Prompted report generation per day
- Screenshot auto-insertion with captions
- Fully local & offline; no data ever leaves the machine
- Easy to reuse for 30-day exams
- Clean, readable Word reports using a pre-built `.docm` template
- Modular VBA design for future expansion (PowerShell, Python, etc.)

---

```bash
📦PentestReporting/
├── DailyReport_Template.docm # Word report template
├── Pentest_DailyLog.xlsx # Master Excel workbook
├── Evidence/
│ ├── 2025-11-17/
│ │ ├── 01_meterpreter.png
│ │ ├── 02_ldap.png
│ │ └── 03_cred_dump.png
│ └── ...
└── README.md # This file
```


---

## 🔢 Excel Structure (`DailyLog` Sheet)

Each row represents a single finding for a specific day.

| Column | Name         | Description                        |
|--------|--------------|------------------------------------|
| A      | Date         | YYYY-MM-DD                         |
| B      | FindingID    | Unique ID per finding              |
| C      | Title        | Short summary of issue             |
| D      | Severity     | Low / Medium / High / Critical     |
| E      | Description  | Detailed explanation               |
| F      | Host         | Target system name/IP              |
| G–I    | Screenshot1–3| Screenshot filenames (max 3)       |
| J      | Notes        | Analyst notes                      |

---

## 🧭 Daily Workflow

1. **Perform Pentest Activities**
   - Use Wireshark, Mimikatz, Meterpreter, etc.
   - Save screenshots locally (max 3 per finding)

2. **Organize Evidence**
   - Save screenshots under:  
     `Evidence/YYYY-MM-DD/`

3. **Log Findings in Excel**
   - Enter each finding in the `DailyLog` worksheet
   - Link screenshot filenames (e.g. `01_ldap.png`)

4. **Run the VBA Macro**
   - Run `GenerateDailyReport` from Excel
   - Select the report date
   - Select the matching evidence folder

5. **Review Generated Word Report**
   - Word opens automatically
   - All findings and screenshots are inserted
   - Save the report manually (e.g. `Daily_Report_YYYY-MM-DD.docx`)

---

## 🚀 How to Use

1. ✅ Open `Pentest_DailyLog.xlsx`
2. 🛠️ Enable macros (Developer tab > Macros)
3. ▶️ Run `GenerateDailyReport`
4. 💬 Follow the prompts:
    - Enter date (e.g. `2025-11-17`)
    - Select folder: `Evidence\2025-11-17\`
5. 💾 Save the Word report

---

## 🔐 Privacy & Compliance

- **All operations are 100% local**
- No cloud storage
- No webhooks or APIs
- No online services used
- Compliant with exam & red team privacy requirements

---

## 🧩 Expansion Ideas

Optional add-ons (not required):

- PowerShell scripts to gather host/AD data
- Python log parsers (e.g. for Wireshark or ARP tables)
- Batch reporting across multiple days
- Charts or timeline visualizations in Excel/Word

---

## 🧠 Credits

Developed as a utility tool for ethical hacking students and red team operators who need reliable and repeatable daily documentation systems.

---

📅 Built for 30-day ethical hacking final exams.  
🛠️ Powered by Excel + Word + VBA.  
🧪 Evidence stays local. Reports stay clean.


## 📁 Folder & File Structure

