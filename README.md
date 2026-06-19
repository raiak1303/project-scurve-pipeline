# 📊 Project S-Curve Pipeline
### Automated Project Progress Dashboard | Excel → S-Curve → Delay Analysis

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)
![Pandas](https://img.shields.io/badge/Pandas-2.0+-green?style=flat-square&logo=pandas)
![Plotly](https://img.shields.io/badge/Plotly-Interactive-orange?style=flat-square&logo=plotly)
![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=flat-square&logo=windows)

---

## 🚀 What This Does

A fully automated pipeline for **project controls professionals** that:

- 📁 **Watches your Excel file** for any changes automatically
- 📈 **Generates S-Curves** (Schedule Progress % + Cost + Earned Value) instantly
- ⚠️ **Detects and explains delays** task by task with severity rating
- 💰 **Calculates EVM metrics** — CPI, SPI, Cost Variance, Schedule Variance
- 🌐 **Produces an interactive HTML dashboard** — no software needed to view
- 🔄 **Real-time updates** — save your Excel → dashboard refreshes automatically

> Built to eliminate hours of manual S-curve updates every week.

---

## 📸 Dashboard Preview

### S-Curve — Schedule Progress
> Planned % vs Actual % over the project timeline

### S-Curve — Cost & Earned Value
> PV (Budget) vs AC (Actual Cost) vs EV (Earned Value)

### Delay Analysis Table
> Color-coded severity per task: 🟢 On Track | 🟡 Minor | 🟠 Moderate | 🔴 Critical

---

## 📁 Files in This Repo

| File | Description |
|------|-------------|
| `pipeline.py` | Main pipeline script — reads Excel, builds S-curves, generates dashboard |
| `Project_Progress_Report.xlsx` | Sample project data with 12 tasks, cost tracking & config sheet |
| `RUN_PIPELINE.bat` | Double-click to run on Windows — no Command Prompt needed |
| `S_Curve_Dashboard.html` | Pre-generated sample dashboard — open in any browser |
| `README.md` | This file |

---

## ⚙️ How It Works

```
Excel File (Tasks + Costs)
        ↓
  pipeline.py watches for changes (watchdog)
        ↓
  Reads data with Pandas
        ↓
  Calculates S-Curves + EVM metrics
        ↓
  Generates interactive HTML dashboard (Plotly)
        ↓
  Opens in your browser automatically
```

---

## 🛠️ Installation

### Requirements
- Python 3.8+
- Windows OS

### Install dependencies
```bash
pip install pandas openpyxl watchdog plotly
```

---

## 🚀 Quick Start

### Option 1 — Double Click (Easiest)
1. Download all files into one folder
2. Double-click `RUN_PIPELINE.bat`
3. Dashboard opens in your browser ✅

### Option 2 — Command Prompt
```bash
cd "C:\Your\Folder\Path"
python pipeline.py
```

---

## 📊 Dashboard KPIs Explained

| KPI | What It Means |
|-----|---------------|
| **Overall Progress %** | Average actual completion across all tasks |
| **PV (Planned Value)** | Budget allocated for work scheduled to date |
| **AC (Actual Cost)** | Money actually spent to date |
| **EV (Earned Value)** | Budget for work actually completed |
| **CPI** | Cost Performance Index — >1 means under budget ✅ |
| **SPI** | Schedule Performance Index — >1 means ahead of schedule ✅ |

---

## 🔧 Customizing Column Names

Your Excel file may use different column names — no problem!

### Easy Way (No coding)
Open `Project_Progress_Report.xlsx` → go to **⚙️ Config** sheet → update column B with your column names → save → run pipeline.

### Manual Way
Open `pipeline.py` and edit the `COL` dictionary:

```python
COL = {
    "task_id":          "Your_Task_ID_Column",
    "task_name":        "Your_Task_Name_Column",
    "planned_start":    "Your_Planned_Start_Column",
    "planned_finish":   "Your_Planned_Finish_Column",
    "actual_start":     "Your_Actual_Start_Column",
    "actual_finish":    "Your_Actual_Finish_Column",
    "planned_progress": "Your_Planned_%_Column",
    "actual_progress":  "Your_Actual_%_Column",
    "planned_cost":     "Your_Planned_Cost_Column",
    "actual_cost":      "Your_Actual_Cost_Column",
}
```

---

## 🏗️ Compatible With

- ✅ Microsoft Excel (.xlsx)
- ✅ Primavera P6 exports (Excel format)
- ✅ Any project tracking spreadsheet

---

## 💡 Use Cases

- Construction project progress reporting
- EPC project controls & scheduling
- Weekly/monthly S-curve generation
- Project delay analysis & reporting
- Earned Value Management (EVM) tracking

---

## 🧰 Tech Stack

| Tool | Purpose |
|------|---------|
| Python | Core pipeline logic |
| Pandas | Data reading & processing |
| Plotly | Interactive S-curve charts |
| OpenPyXL | Excel file reading/writing |
| Watchdog | File system monitoring |
| HTML/CSS/JS | Dashboard frontend |

---

## 👤 Author

**Anand** — Project Controls Professional  
Passionate about automating manual reporting processes in construction & engineering projects.

📧 Connect on [LinkedIn](https://linkedin.com/in/linkedin.com/in/raiak-planningeng)  
🐙 GitHub: [raiak1303](https://github.com/raiak1303)

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

⭐ **If this helped you, give it a star!** ⭐
