# WAT Data Automation Tool – v1.1.1
## 📖 Description
Version 1.1.1 is a continuation and structural upgrade of the original WAT Data Automation Tool.
This release introduces a multi‑class architecture, centralized error logging, and an updated Histogram Plot dashboard, while preserving the core purpose of automating semiconductor wafer acceptance test (WAT) deliverables.
The tool converts raw .wat files into structured Excel workbooks, generates per‑Unit and per‑Wafer summaries, validates specifications, and produces interactive capability plots with Cp/Cpk statistics. Built with Python, Tkinter, OpenPyXL, Matplotlib, NumPy, and Pandas, it streamlines engineering workflows and ensures reproducible, audit‑ready insights.

---

## 📌 Disclaimer
This project is a portfolio demonstration built entirely with synthetic/dummy data.
While the workbook structure and formatting are inspired by typical engineering workflows, all headers, values, and examples have been replaced with generic placeholders.
No proprietary intellectual property, client data, or company‑specific conventions are included.
Its sole purpose is to showcase automation techniques, reproducible workflows, and technical skills in Python, Tkinter, OpenPyXL, Matplotlib, NumPy, and Pandas.

---

## 🚀 Features
## 🔧 New in v1.1.1
- Multi‑class architecture (Parser, Builder, Summary Generator, Histogram Viewer, Logger, GUI Controller)
- Centralized error logging with timestamped files
- Automatic 30‑day log cleanup
- Updated Histogram Plot dashboard with improved layout and visuals

## .wat to Excel Conversion
- Converts raw .wat files into structured Excel workbooks
- Dynamic workbook naming (e.g., Wafer 1~N)
- Serves as the reference for:
- Per Unit Data
- Per Wafer Summary
- Summary Sheet
- Histogram Plot

## Per Unit Data
- Extracts wafer IDs, parameters, and site measurements
- Produces a structured sheet for traceability and analysis
## Per Wafer
- Transposes site values per wafer
- Adds AVERAGE and STDEV formulas
- Maps Spec HI, Spec LO, and Unit
- Professionally formatted with merged headers, borders, and autofit
## Summary Sheet
- Consolidates per‑parameter statistics
- Includes Spec HI/LO, Mean, Std Dev, CpK, CpK Hi, CpK Lo
- Professionally formatted for quick engineering review
## Histogram Plot (Capability Plot)
- Updated v1.0.0 dashboard
- Interactive parameter selection
- ±3σ normal curve overlay
- Cp/Cpk statistics displayed in a side panel
- Clean GUI with scrollable logs and status messages

## 🛠️ Tech Stack
- Python (automation & GUI)
- Tkinter (user interface)
- OpenPyXL (Excel file handling)
- Matplotlib (histogram visualization)
- NumPy (statistics)
- Pandas (data handling)

---

## 📦 Required Packages
The dependencies are listed in [`requirements.txt`](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1/blob/main/requirements.txt)

Install them with:
```bash
pip install -r requirements.txt
```

---
## ⚡ Usage Workflow
1. **Load Input File**  
   Select a raw `.wat` file (e.g., [Dummy data.wat](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.wat)).
2. **Run Automation**  
Generates:
• Per Unit Data sheet
• Per Wafer Summary sheet
3. **Generate Summary Sheet**  
Produces a consolidated capability summary with Spec HI/LO, Mean, Std Dev, CpK, CpK Hi, and CpK Lo.
4. **Explore Histogram Plot** 
Updated v1.1.0 dashboard with ±3σ overlays and Cp/Cpk statistics.
- Review Outputs
- Excel workbook
- Dashboard interface
- Capability plot

---

## 📂 Sample Files
- Input: [Dummy data.wat](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.wat)  
- Output: [Dummy data.xlsx](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.xlsx) 

These files are synthetic examples included for demonstration only.

---

## 📸 Screenshots
- **WAT Data Automation v1.1.0 Dashboard Screenshot**
![WAT Data Automation v1.1.0 Dashboard.png](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1/blob/main/WAT%20Data%20Automation%20v1.1.1%20Dashboard.png)
- **Histogram Dashboard Screenshot** 
![Histogram Viewer Dashboard.png](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1/blob/main/Histogram%20Viewer%20Dashboard.png)
- **Sample Output Capability Plot (Param G)**
![Sample Histogram Plot.png](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1/blob/main/Param_G.png)

---

## 🌟 Impact
- Reduces manual effort in semiconductor WAT deliverables
- Ensures reproducibility with deterministic spec mapping
- Improves accuracy with Cp/Cpk capability statistics
- Enhances usability with a polished GUI and centralized logging

---

## 📦 Download
Release v1.1.0 with histogram viewer and spec mapping is available here:  
➡️ [Download WAT Data Automation Tool v1.1.1](https://github.com/roannelafuente/WAT-Data-Automation-v1.1.1/releases/tag/v1.1.1)

▶️ **Usage**: Run the .exe to launch the dashboard and explore the features.

---

## 👩‍💻 Author
**Rose Anne Lafuente**  
Licensed Electronics Engineer | Product Engineer II | Python Automation  
GitHub: [@roannelafuente](https://github.com/roannelafuente)  
LinkedIn: [Rose Anne Lafuente](www.linkedin.com/in/rose-anne-lafuente)
