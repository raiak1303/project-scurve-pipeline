"""
╔══════════════════════════════════════════════════════════════════╗
║  PROJECT PROGRESS PIPELINE  –  S-Curve + Delay Analysis         ║
║  Usage:  python pipeline.py                                      ║
║  Config: Edit CONFIG section below OR use the ⚙️ Config sheet   ║
╚══════════════════════════════════════════════════════════════════╝
"""

# ─────────────────────────────────────────────────────────────────
#  ⚙️  CONFIG  –  Change paths & column names here
# ─────────────────────────────────────────────────────────────────
WATCH_FOLDER   = WATCH_FOLDER = r"C:\Users\raiak\OneDrive\Desktop\Projectdata"        # Folder to watch
EXCEL_FILENAME = "Project_Progress_Report.xlsx"  # File to watch
TASK_SHEET     = "Project_Tasks"
COST_SHEET     = "Cost_Tracker"
OUTPUT_HTML    = "S_Curve_Dashboard.html"   # Output dashboard

# Column name mapping – update these if your file uses different headers
COL = {
    "task_id":          "Task_ID",
    "task_name":        "Task_Name",
    "planned_start":    "Planned_Start",
    "planned_finish":   "Planned_Finish",
    "actual_start":     "Actual_Start",
    "actual_finish":    "Actual_Finish",
    "planned_progress": "Planned_Progress_%",
    "actual_progress":  "Actual_Progress_%",
    "status":           "Status",
    "planned_cost":     "Planned_Cost_USD",
    "actual_cost":      "Actual_Cost_USD",
    "earned_value":     "Earned_Value_USD",
}
# ─────────────────────────────────────────────────────────────────

import os, sys, time, json, warnings
from pathlib import Path
from datetime import date, timedelta, datetime

import pandas as pd
import numpy as np
warnings.filterwarnings("ignore")

# ── Try to load watchdog (optional for auto-watch) ───────────────
try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    HAS_WATCHDOG = True
except ImportError:
    HAS_WATCHDOG = False

EXCEL_PATH = Path(WATCH_FOLDER) / EXCEL_FILENAME
OUTPUT_PATH = Path(WATCH_FOLDER) / OUTPUT_HTML


# ═══════════════════════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════════════════════
def load_data():
    df_tasks = pd.read_excel(EXCEL_PATH, sheet_name=TASK_SHEET)
    df_costs = pd.read_excel(EXCEL_PATH, sheet_name=COST_SHEET)

    # Try to read column mapping from Config sheet
    try:
        df_cfg = pd.read_excel(EXCEL_PATH, sheet_name="⚙️ Config", header=1)
        for _, row in df_cfg.iterrows():
            field = str(row.get("Field", "")).strip()
            col_name = str(row.get("Current Column Name in Excel", "")).strip()
            if field in COL and col_name and col_name != "nan":
                COL[field] = col_name
    except Exception:
        pass  # Use defaults if config sheet missing

    # Rename columns to standard internal names
    inv = {v: k for k, v in COL.items()}
    df_tasks.rename(columns={v: k for v, k in inv.items() if v in df_tasks.columns}, inplace=True)
    df_costs.rename(columns={v: k for v, k in inv.items() if v in df_costs.columns}, inplace=True)

    # Parse dates
    for col in ["planned_start", "planned_finish", "actual_start", "actual_finish"]:
        if col in df_tasks.columns:
            df_tasks[col] = pd.to_datetime(df_tasks[col], errors="coerce")

    # Normalize progress 0-100
    for col in ["planned_progress", "actual_progress"]:
        if col in df_tasks.columns:
            if df_tasks[col].max() <= 1.0:
                df_tasks[col] = df_tasks[col] * 100

    # Derive earned value if missing
    if "earned_value" not in df_costs.columns and "planned_cost" in df_costs.columns:
        prog_map = df_tasks.set_index("task_id")["actual_progress"].to_dict() if "task_id" in df_tasks.columns else {}
        df_costs["earned_value"] = df_costs.apply(
            lambda r: r.get("planned_cost", 0) * prog_map.get(r.get("task_id", ""), 0) / 100,
            axis=1
        )

    return df_tasks, df_costs


# ═══════════════════════════════════════════════════════════════
#  S-CURVE CALCULATION
# ═══════════════════════════════════════════════════════════════
def build_s_curves(df_tasks, df_costs):
    # Date range
    all_dates = pd.concat([
        df_tasks["planned_start"].dropna(),
        df_tasks["planned_finish"].dropna(),
        df_tasks["actual_start"].dropna(),
    ])
    start = all_dates.min().date()
    end   = (df_tasks["planned_finish"].max() + timedelta(days=30)).date()
    today = date.today()
    dates = [start + timedelta(days=i) for i in range((end - start).days + 1)]

    total_tasks   = len(df_tasks)
    total_planned = df_costs["planned_cost"].sum() if "planned_cost" in df_costs.columns else 1
    total_earned  = df_costs["earned_value"].sum() if "earned_value" in df_costs.columns else 0
    total_actual  = df_costs["actual_cost"].sum() if "actual_cost" in df_costs.columns else 0

    planned_prog_curve, actual_prog_curve = [], []
    planned_cost_curve, actual_cost_curve, ev_curve = [], [], []

    for d in dates:
        dt = pd.Timestamp(d)
        # Planned progress: % of tasks that should be complete by day d
        plan_done = df_tasks.apply(
            lambda r: min(100, max(0, (dt - r["planned_start"]).days /
                          max(1, (r["planned_finish"] - r["planned_start"]).days) * 100))
            if pd.notna(r.get("planned_start")) and pd.notna(r.get("planned_finish")) else 0,
            axis=1
        )
        planned_prog_curve.append(round(plan_done.mean(), 2))

        # Actual progress (static – current snapshot)
        if d <= today:
            actual_prog_curve.append(round(df_tasks["actual_progress"].mean(), 2) if d == today else None)
        else:
            actual_prog_curve.append(None)

        # Cost curves (cumulative)
        if "planned_cost" in df_costs.columns:
            planned_cost_curve.append(round(total_planned * planned_prog_curve[-1] / 100, 0))
        if d <= today:
            actual_cost_curve.append(total_actual if d == today else None)
            ev_curve.append(total_earned if d == today else None)

    # Fill actual up to today as a straight line from 0 → current
    today_idx = min((today - start).days, len(dates) - 1)
    actual_cost_filled = [None] * len(dates)
    ev_filled = [None] * len(dates)
    if total_actual > 0:
        for i in range(today_idx + 1):
            frac = i / max(1, today_idx)
            actual_cost_filled[i] = round(total_actual * frac, 0)
            ev_filled[i] = round(total_earned * frac, 0)

    return {
        "dates": [str(d) for d in dates],
        "today": str(today),
        "planned_progress": planned_prog_curve,
        "actual_progress":  [df_tasks["actual_progress"].mean()] * (today_idx + 1) + [None] * (len(dates) - today_idx - 1),
        "planned_cost":     planned_cost_curve,
        "actual_cost":      actual_cost_filled,
        "earned_value":     ev_filled,
        "today_idx":        today_idx,
    }


# ═══════════════════════════════════════════════════════════════
#  DELAY ANALYSIS
# ═══════════════════════════════════════════════════════════════
def analyse_delays(df_tasks):
    today = pd.Timestamp(date.today())
    rows = []
    for _, r in df_tasks.iterrows():
        name  = r.get("task_name", r.get("task_id", "Unknown"))
        p_end = r.get("planned_finish")
        a_end = r.get("actual_finish")
        a_st  = r.get("actual_start")
        p_st  = r.get("planned_start")
        prog  = r.get("actual_progress", 0)
        status = r.get("status", "")

        start_delay = finish_delay = 0
        if pd.notna(a_st) and pd.notna(p_st):
            start_delay = max(0, (a_st - p_st).days)
        if pd.notna(a_end) and pd.notna(p_end):
            finish_delay = max(0, (a_end - p_end).days)
        elif pd.notna(p_end) and prog < 100 and p_end < today:
            finish_delay = (today - p_end).days  # overdue

        severity = "🟢 On Track" if finish_delay == 0 and start_delay == 0 else \
                   "🟡 Minor" if finish_delay <= 7 else \
                   "🟠 Moderate" if finish_delay <= 21 else "🔴 Critical"

        rows.append({
            "task": str(name),
            "start_delay": start_delay,
            "finish_delay": finish_delay,
            "progress": round(prog, 1),
            "severity": severity,
            "status": str(status),
        })
    return rows


# ═══════════════════════════════════════════════════════════════
#  KPIs
# ═══════════════════════════════════════════════════════════════
def compute_kpis(df_tasks, df_costs):
    total_planned = df_costs["planned_cost"].sum() if "planned_cost" in df_costs.columns else 0
    total_actual  = df_costs["actual_cost"].sum()  if "actual_cost"  in df_costs.columns else 0
    total_ev      = df_costs["earned_value"].sum()  if "earned_value" in df_costs.columns else 0

    cpi = round(total_ev / total_actual, 2)  if total_actual > 0 else None
    spi = round(total_ev / total_planned, 2) if total_planned > 0 else None
    avg_prog = round(df_tasks["actual_progress"].mean(), 1) if "actual_progress" in df_tasks.columns else 0
    delayed_count = df_tasks[df_tasks.get("status", pd.Series()).str.contains("Delay", na=False)].shape[0] \
                    if "status" in df_tasks.columns else 0
    overrun = round((total_actual - total_ev) / 1000, 1) if total_ev > 0 else 0

    return {
        "avg_progress":  avg_prog,
        "total_planned": f"{total_planned/1_000_000:.2f}M" if total_planned >= 1e6 else f"{total_planned:,.0f}",
        "total_actual":  f"{total_actual/1_000_000:.2f}M"  if total_actual  >= 1e6 else f"{total_actual:,.0f}",
        "total_ev":      f"{total_ev/1_000_000:.2f}M"      if total_ev      >= 1e6 else f"{total_ev:,.0f}",
        "cpi":           cpi,
        "spi":           spi,
        "delayed_tasks": delayed_count,
        "cost_overrun_k": overrun,
    }


# ═══════════════════════════════════════════════════════════════
#  HTML DASHBOARD GENERATOR
# ═══════════════════════════════════════════════════════════════
def generate_dashboard(df_tasks, df_costs):
    curves  = build_s_curves(df_tasks, df_costs)
    delays  = analyse_delays(df_tasks)
    kpis    = compute_kpis(df_tasks, df_costs)
    updated = datetime.now().strftime("%d %b %Y  %H:%M:%S")

    delay_rows = ""
    for d in delays:
        severity_bg = {"🟢 On Track": "#d4edda", "🟡 Minor": "#fff3cd",
                       "🟠 Moderate": "#ffe5b4", "🔴 Critical": "#f8d7da"}.get(d["severity"], "#fff")
        delay_rows += f"""
        <tr style="background:{severity_bg}">
          <td>{d['task']}</td>
          <td>{d['status']}</td>
          <td>{d['progress']}%</td>
          <td>{'+'+ str(d['start_delay'])+' days' if d['start_delay'] > 0 else '—'}</td>
          <td>{'+'+ str(d['finish_delay'])+' days' if d['finish_delay'] > 0 else '—'}</td>
          <td>{d['severity']}</td>
        </tr>"""

    cpi_color = "#28a745" if kpis["cpi"] and kpis["cpi"] >= 1 else "#dc3545"
    spi_color = "#28a745" if kpis["spi"] and kpis["spi"] >= 1 else "#dc3545"

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Project S-Curve Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.26.0.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; color: #1a1a2e; }}
  .header {{ background: linear-gradient(135deg, #1F4E79, #2E75B6); color: white;
             padding: 24px 32px; display: flex; justify-content: space-between; align-items: center; }}
  .header h1 {{ font-size: 1.6rem; font-weight: 700; }}
  .header .meta {{ font-size: 0.85rem; opacity: 0.85; text-align: right; }}
  .container {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 24px; }}
  .kpi {{ background: white; border-radius: 10px; padding: 18px 16px; text-align: center;
           box-shadow: 0 2px 8px rgba(0,0,0,.08); border-top: 4px solid #2E75B6; }}
  .kpi .val {{ font-size: 2rem; font-weight: 700; color: #1F4E79; }}
  .kpi .lbl {{ font-size: 0.75rem; color: #666; margin-top: 4px; text-transform: uppercase; letter-spacing: .5px; }}
  .card {{ background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,.08);
            margin-bottom: 24px; overflow: hidden; }}
  .card-title {{ background: #1F4E79; color: white; padding: 14px 20px; font-weight: 600; font-size: 1rem; }}
  .card-body {{ padding: 20px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.88rem; }}
  th {{ background: #2E75B6; color: white; padding: 10px 12px; text-align: left; }}
  td {{ padding: 9px 12px; border-bottom: 1px solid #eee; }}
  tr:hover td {{ background: #f5f8ff; }}
  .footer {{ text-align: center; color: #999; font-size: 0.8rem; padding: 20px; }}
  .legend-note {{ font-size:0.8rem; color:#555; padding:8px 20px 0; }}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>📊 Project Progress Dashboard</h1>
    <div style="opacity:.8;font-size:.9rem;margin-top:4px;">S-Curve · EVM · Delay Analysis</div>
  </div>
  <div class="meta">
    Last updated<br><strong>{updated}</strong><br>
    <span style="font-size:.75rem">Auto-refreshes when Excel changes</span>
  </div>
</div>

<div class="container">

  <!-- KPIs -->
  <div class="kpi-grid">
    <div class="kpi">
      <div class="val">{kpis['avg_progress']}%</div>
      <div class="lbl">Overall Progress</div>
    </div>
    <div class="kpi">
      <div class="val" style="font-size:1.4rem">${kpis['total_planned']}</div>
      <div class="lbl">Budget (PV)</div>
    </div>
    <div class="kpi">
      <div class="val" style="font-size:1.4rem">${kpis['total_actual']}</div>
      <div class="lbl">Actual Cost (AC)</div>
    </div>
    <div class="kpi">
      <div class="val" style="font-size:1.4rem">${kpis['total_ev']}</div>
      <div class="lbl">Earned Value (EV)</div>
    </div>
    <div class="kpi">
      <div class="val" style="color:{cpi_color}">{kpis['cpi'] or 'N/A'}</div>
      <div class="lbl">CPI (≥1 = under budget)</div>
    </div>
    <div class="kpi">
      <div class="val" style="color:{spi_color}">{kpis['spi'] or 'N/A'}</div>
      <div class="lbl">SPI (≥1 = on schedule)</div>
    </div>
    <div class="kpi">
      <div class="val" style="color:#dc3545">{kpis['delayed_tasks']}</div>
      <div class="lbl">Delayed Tasks</div>
    </div>
  </div>

  <!-- S-Curve: Progress -->
  <div class="card">
    <div class="card-title">📈 S-Curve — Schedule Progress (%)</div>
    <div class="card-body">
      <div id="chart_progress" style="height:380px"></div>
      <div class="legend-note">▸ Planned: what % should be done each day based on task durations &nbsp;|&nbsp; ▸ Actual: snapshot of real progress today</div>
    </div>
  </div>

  <!-- S-Curve: Cost / EV -->
  <div class="card">
    <div class="card-title">💰 S-Curve — Cost & Earned Value (USD)</div>
    <div class="card-body">
      <div id="chart_cost" style="height:380px"></div>
      <div class="legend-note">▸ PV = Planned Value &nbsp;|&nbsp; EV = Earned Value &nbsp;|&nbsp; AC = Actual Cost &nbsp;|&nbsp; EV &lt; AC → Cost overrun &nbsp;|&nbsp; EV &lt; PV → Behind schedule</div>
    </div>
  </div>

  <!-- Delay Table -->
  <div class="card">
    <div class="card-title">⚠️ Delay Analysis — Task by Task</div>
    <div class="card-body" style="overflow-x:auto">
      <table>
        <thead><tr>
          <th>Task</th><th>Status</th><th>Progress</th>
          <th>Start Delay</th><th>Finish Delay</th><th>Severity</th>
        </tr></thead>
        <tbody>{delay_rows}</tbody>
      </table>
    </div>
  </div>

</div>

<div class="footer">Auto-generated by Project Progress Pipeline · {updated}</div>

<script>
const dates = {json.dumps(curves['dates'])};
const today = "{curves['today']}";
const todayIdx = {curves['today_idx']};

// ── Progress Chart ────────────────────────────────────────────
const plannedProg = {json.dumps(curves['planned_progress'])};
const actualProg  = {json.dumps(curves['actual_progress'])};

const trace1 = {{
  x: dates, y: plannedProg, name: 'Planned Progress (%)',
  mode: 'lines', line: {{ color: '#2E75B6', width: 2.5, dash: 'dot' }},
  hovertemplate: '%{{x}}<br>Planned: %{{y:.1f}}%<extra></extra>'
}};
const trace2 = {{
  x: dates.slice(0, todayIdx+1),
  y: Array(todayIdx+1).fill(actualProg.find(v => v !== null) || 0),
  name: 'Actual Progress (%)',
  mode: 'lines', line: {{ color: '#ED7D31', width: 3 }},
  hovertemplate: '%{{x}}<br>Actual: %{{y:.1f}}%<extra></extra>'
}};
const todayLine1 = {{
  x: [today, today], y: [0, 100],
  mode: 'lines', name: 'Today',
  line: {{ color: '#dc3545', width: 1.5, dash: 'dash' }},
  hoverinfo: 'skip'
}};

Plotly.newPlot('chart_progress', [trace1, trace2, todayLine1], {{
  margin: {{t:20, b:50, l:55, r:30}},
  xaxis: {{ title: 'Date', showgrid: true, gridcolor: '#eee' }},
  yaxis: {{ title: 'Cumulative Progress (%)', range: [0,105], showgrid: true, gridcolor: '#eee' }},
  legend: {{ orientation: 'h', y: -0.15 }},
  paper_bgcolor: 'white', plot_bgcolor: 'white'
}}, {{responsive: true}});

// ── Cost / EV Chart ───────────────────────────────────────────
const pv = {json.dumps(curves['planned_cost'])};
const ac = {json.dumps(curves['actual_cost'])};
const ev = {json.dumps(curves['earned_value'])};

const tracePV = {{
  x: dates, y: pv, name: 'Planned Value (PV)',
  mode: 'lines', line: {{ color: '#2E75B6', width: 2.5, dash: 'dot' }},
  hovertemplate: '%{{x}}<br>PV: $%{{y:,.0f}}<extra></extra>'
}};
const traceAC = {{
  x: dates.slice(0, todayIdx+1), y: ac.slice(0, todayIdx+1), name: 'Actual Cost (AC)',
  mode: 'lines', line: {{ color: '#dc3545', width: 3 }},
  hovertemplate: '%{{x}}<br>AC: $%{{y:,.0f}}<extra></extra>'
}};
const traceEV = {{
  x: dates.slice(0, todayIdx+1), y: ev.slice(0, todayIdx+1), name: 'Earned Value (EV)',
  mode: 'lines', line: {{ color: '#28a745', width: 3 }},
  fill: 'tonexty', fillcolor: 'rgba(40,167,69,0.08)',
  hovertemplate: '%{{x}}<br>EV: $%{{y:,.0f}}<extra></extra>'
}};
const todayLine2 = {{
  x: [today, today], y: [0, Math.max(...pv.filter(Boolean))],
  mode: 'lines', name: 'Today',
  line: {{ color: '#dc3545', width: 1.5, dash: 'dash' }},
  hoverinfo: 'skip'
}};

Plotly.newPlot('chart_cost', [tracePV, traceAC, traceEV, todayLine2], {{
  margin: {{t:20, b:50, l:80, r:30}},
  xaxis: {{ title: 'Date', showgrid: true, gridcolor: '#eee' }},
  yaxis: {{ title: 'Cumulative Cost (USD)', showgrid: true, gridcolor: '#eee',
            tickformat: '$,.0f' }},
  legend: {{ orientation: 'h', y: -0.15 }},
  paper_bgcolor: 'white', plot_bgcolor: 'white'
}}, {{responsive: true}});
</script>
</body>
</html>"""
    return html


# ═══════════════════════════════════════════════════════════════
#  RUN PIPELINE
# ═══════════════════════════════════════════════════════════════
def run_pipeline():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Reading {EXCEL_PATH.name} ...")
    try:
        df_tasks, df_costs = load_data()
        html = generate_dashboard(df_tasks, df_costs)
        OUTPUT_PATH.write_text(html, encoding="utf-8")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ Dashboard saved → {OUTPUT_PATH}")
        # Auto-open on first run
        os.startfile(str(OUTPUT_PATH)) if sys.platform == "win32" else None
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ Error: {e}")


# ═══════════════════════════════════════════════════════════════
#  WATCHDOG FILE WATCHER
# ═══════════════════════════════════════════════════════════════
class ExcelHandler(FileSystemEventHandler):
    def __init__(self):
        self._last_run = 0

    def on_modified(self, event):
        if Path(event.src_path).name == EXCEL_FILENAME:
            now = time.time()
            if now - self._last_run > 3:   # debounce 3 s
                self._last_run = now
                time.sleep(1)              # wait for Excel to finish writing
                run_pipeline()

    on_created = on_modified


def main():
    print("=" * 60)
    print("  PROJECT PROGRESS PIPELINE")
    print(f"  Watching: {EXCEL_PATH}")
    print(f"  Output  : {OUTPUT_PATH}")
    print("=" * 60)

    if not EXCEL_PATH.exists():
        print(f"\n⚠️  File not found: {EXCEL_PATH}")
        print("   Place your Excel file there and restart.\n")
        return

    # Run once immediately
    run_pipeline()

    if HAS_WATCHDOG:
        print(f"\n👁️  Watching for changes … (Ctrl+C to stop)\n")
        handler  = ExcelHandler()
        observer = Observer()
        observer.schedule(handler, str(WATCH_FOLDER), recursive=False)
        observer.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            observer.stop()
        observer.join()
    else:
        print("\n⚠️  watchdog not installed. Running once only.")
        print("   Install it with:  pip install watchdog")


if __name__ == "__main__":
    main()
