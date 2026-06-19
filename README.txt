═══════════════════════════════════════════════════════════════
  PROJECT PROGRESS PIPELINE – Quick Start Guide
═══════════════════════════════════════════════════════════════

FILES IN THIS PACKAGE
─────────────────────
  Project_Progress_Report.xlsx   → Your project data (sample included)
  pipeline.py                    → The pipeline script (watches & generates)
  S_Curve_Dashboard.html         → Output dashboard (auto-generated)
  README.txt                     → This file


STEP 1 – INSTALL PYTHON DEPENDENCIES
─────────────────────────────────────
Open Command Prompt and run:

  pip install pandas openpyxl watchdog plotly


STEP 2 – SET UP YOUR FOLDER
────────────────────────────
1. Create a folder, e.g.:   C:\ProjectData
2. Copy both files there:
     - Project_Progress_Report.xlsx
     - pipeline.py


STEP 3 – CONFIGURE THE PIPELINE
─────────────────────────────────
Open pipeline.py in Notepad (or VS Code) and edit lines 15–17:

  WATCH_FOLDER   = r"C:\ProjectData"            ← your folder path
  EXCEL_FILENAME = "Project_Progress_Report.xlsx"  ← your file name
  OUTPUT_HTML    = "S_Curve_Dashboard.html"     ← output file name


STEP 4 – CUSTOMIZE COLUMN NAMES (if your file is different)
─────────────────────────────────────────────────────────────
Option A – Use the Excel Config sheet:
  Open Project_Progress_Report.xlsx → go to "⚙️ Config" sheet
  Change the values in column B to match YOUR column names.
  The pipeline reads this sheet automatically!

Option B – Edit pipeline.py directly:
  Find the COL dictionary (around line 20) and update the values:

  COL = {
      "task_id":          "YOUR_TASK_ID_COLUMN",
      "task_name":        "YOUR_TASK_NAME_COLUMN",
      "planned_start":    "YOUR_PLANNED_START_COLUMN",
      "planned_finish":   "YOUR_PLANNED_FINISH_COLUMN",
      "actual_start":     "YOUR_ACTUAL_START_COLUMN",
      "actual_finish":    "YOUR_ACTUAL_FINISH_COLUMN",
      "planned_progress": "YOUR_PLANNED_%_COLUMN",
      "actual_progress":  "YOUR_ACTUAL_%_COLUMN",
      "planned_cost":     "YOUR_PLANNED_COST_COLUMN",
      "actual_cost":      "YOUR_ACTUAL_COST_COLUMN",
  }


STEP 5 – RUN THE PIPELINE
──────────────────────────
Double-click: run_pipeline.bat  (or run in Command Prompt):

  cd C:\ProjectData
  python pipeline.py

The dashboard (S_Curve_Dashboard.html) will:
  ✅ Open in your browser automatically
  ✅ Regenerate every time you save the Excel file
  ✅ Show 3 S-curves: Schedule %, Cost PV, Earned Value
  ✅ Show a delay analysis table for every task


WHAT THE DASHBOARD SHOWS
──────────────────────────
  📈 S-Curve 1 – Schedule Progress (%)
     Planned % vs Actual % over time

  💰 S-Curve 2 – Cost & Earned Value
     PV (budget) vs AC (spent) vs EV (earned)
     Gap between EV and AC = cost overrun
     Gap between EV and PV = schedule delay in $ terms

  ⚠️  Delay Analysis Table
     Start delay / Finish delay per task
     Color-coded severity: 🟢 On Track | 🟡 Minor | 🟠 Moderate | 🔴 Critical

  📊 KPIs
     CPI (Cost Performance Index) – >1 means under budget
     SPI (Schedule Performance Index) – >1 means ahead of schedule


USING YOUR REAL P6 / PRIMAVERA DATA
─────────────────────────────────────
1. Export from P6 as Excel (.xlsx)
2. Put the exported file in C:\ProjectData
3. Update column names in ⚙️ Config sheet or COL dictionary
4. Run pipeline.py → done!


SUPPORT
────────
Built for interview demonstration.
Pipeline reads Excel → calculates S-curves → generates HTML dashboard.
No internet required to run the dashboard once generated.

═══════════════════════════════════════════════════════════════
