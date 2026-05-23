"""
NSE1 Dashboard Updater — TVS Branded
=====================================
Run this script whenever you update your Excel file.

Usage:
    python update_dashboard.py

Requirements:
    pip install pandas openpyxl

Excel columns required:
    Country | Date | Event | Status | Description | Category of Event | Requires PR

Requires PR formula (in Excel):
    =IF(OR(F2="Brand Launch",F2="Product Launch",F2="CSR",F2="Partnership"),"Yes","No")
"""

EXCEL_FILE  = "NSE1 Important Dates.xlsx"
OUTPUT_FILE = "index.html"

import sys, os, json, base64
from datetime import datetime

def check_dependencies():
    missing = []
    try: import pandas
    except ImportError: missing.append("pandas")
    try: import openpyxl
    except ImportError: missing.append("openpyxl")
    if missing:
        print(f"\n❌  Missing packages: pip install {' '.join(missing)}\n")
        sys.exit(1)

check_dependencies()
import pandas as pd

def get_logo_b64():
    for name in ["tvs_logo.jpg", "tvs_logo.JPG", "tvs_logo_23.jpg", "tvs_logo.png"]:
        if os.path.exists(name):
            with open(name, "rb") as f:
                ext  = name.split(".")[-1].lower()
                mime = "image/png" if ext == "png" else "image/jpeg"
                return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    return None

def load_data(path):
    if not os.path.exists(path):
        print(f"\n❌  Excel file not found: {path}\n")
        sys.exit(1)
    df = pd.read_excel(path)
    required = {'Country', 'Date', 'Event', 'Status', 'Category of Event', 'Requires PR'}
    if not required.issubset(df.columns):
        print(f"\n❌  Excel must have columns: {required}\n    Found: {set(df.columns)}\n")
        sys.exit(1)
    df['Date']    = pd.to_datetime(df['Date'], errors='coerce')
    df            = df.dropna(subset=['Date'])
    df['Month']   = df['Date'].dt.strftime('%B %Y')
    df['DateStr'] = df['Date'].dt.strftime('%d %b %Y')
    df['Status']  = df['Status'].fillna('').str.strip()
    df['Category of Event'] = df['Category of Event'].fillna('').astype(str).str.strip()
    df['Requires PR'] = df['Requires PR'].fillna('No').astype(str).str.strip()
    if 'Description' not in df.columns:
        df['Description'] = ''
    df['Description'] = df['Description'].fillna('').astype(str).str.strip()
    return df

def build_json(df):
    today = pd.Timestamp.now().normalize()
    month_order, seen = [], set()
    for m in df.sort_values('Date')['Month']:
        if m not in seen:
            month_order.append(m)
            seen.add(m)
    data = {}
    for month in month_order:
        mdf = df[df['Month'] == month]
        data[month] = {}
        for country in sorted(mdf['Country'].unique()):
            cdf = mdf[mdf['Country'] == country].sort_values('Date')
            records = []
            for _, row in cdf.iterrows():
                date_passed = bool(row['Date'] <= today)
                desc = row['Description'] if date_passed else ''
                records.append({
                    'DateStr':         row['DateStr'],
                    'Event':           row['Event'],
                    'Status':          row['Status'],
                    'Description':     desc,
                    'DatePassed':      date_passed,
                    'CategoryOfEvent': row['Category of Event'],
                    'RequiresPR':      row['Requires PR'],
                    'Country':         country,
                })
            data[month][country] = records
    return {'months': month_order, 'data': data}

# ─────────────────────────────────────────────────────────────────────────────
# HTML TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────
def build_html(payload, updated_at, logo_src):
    raw_json  = json.dumps(payload, ensure_ascii=False)
    logo_html = (f'<div class="tvs-logo"><img src="{logo_src}" alt="TVS"/></div>'
                 if logo_src else '<div class="tvs-logo-text">TVS</div>')

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>TVS NSE1 Important Dates</title>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Poppins:wght@400;500;600&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  :root {{
    --bg: #f0f3f8; --surface: #fff; --surface2: #f5f7fa; --border: #dde2ec;
    --tvs-blue: #1B3F8B; --tvs-blue-dark: #142f6a; --tvs-blue-light: #e6ecf8;
    --tvs-red: #CC1313; --tvs-red-dark: #a80f0f;
    --executed: #1a7a4a; --executed-bg: #e8f5ee;
    --postponed: #CC1313; --postponed-bg: #fce8e8;
    --yet: #92650a; --yet-bg: #fffbe6;
    --text: #1a1f2e; --muted: #6b7280;
  }}
  body {{ background: var(--bg); color: var(--text); font-family: 'DM Sans', sans-serif; min-height: 100vh; }}
  .top-strip {{ height: 4px; background: linear-gradient(90deg, var(--tvs-red) 0%, var(--tvs-blue) 60%); }}

  header {{
    background: var(--tvs-blue); padding: 0 40px;
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100;
    gap: 16px; height: 72px;
    box-shadow: 0 3px 16px rgba(27,63,139,0.35);
  }}
  .header-left {{ display: flex; align-items: center; gap: 18px; flex-shrink: 0; }}
  .tvs-logo {{ height: 42px; background: white; border-radius: 7px; padding: 5px 12px; display: flex; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }}
  .tvs-logo img {{ height: 30px; width: auto; display: block; }}
  .tvs-logo-text {{ font-family: 'Montserrat', sans-serif; font-size: 22px; font-weight: 800; color: var(--tvs-blue); background: white; padding: 6px 14px; border-radius: 7px; }}
  .header-divider {{ width: 1px; height: 34px; background: rgba(255,255,255,0.2); flex-shrink: 0; }}
  .header-title h1 {{ font-family: 'Montserrat', sans-serif; font-size: 17px; font-weight: 700; color: white; white-space: nowrap; }}
  .header-title p {{ font-family: 'Poppins', sans-serif; font-size: 11px; color: rgba(255,255,255,0.55); margin-top: 2px; text-transform: uppercase; letter-spacing: 0.7px; font-weight: 400; white-space: nowrap; }}
  .header-right {{ display: flex; align-items: center; gap: 12px; flex-shrink: 0; }}
  .header-btns {{ display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
  .updated {{ font-size: 11px; color: rgba(255,255,255,0.5); white-space: nowrap; }}
  .updated strong {{ color: rgba(255,255,255,0.85); }}

  .dropdown-wrap {{ position: relative; flex-shrink: 0; }}
  .dropdown-btn {{
    background: var(--tvs-red); color: white; border: none;
    padding: 10px 14px; font-family: 'Montserrat', sans-serif; font-size: 12px; font-weight: 700;
    border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 8px;
    transition: background 0.15s, transform 0.1s; min-width: 140px; justify-content: space-between;
    box-shadow: 0 2px 8px rgba(204,19,19,0.4); white-space: nowrap;
  }}
  .dropdown-btn:hover {{ background: var(--tvs-red-dark); }}
  .dropdown-btn:active {{ transform: scale(0.98); }}
  .dropdown-btn.blue-btn {{ background: var(--tvs-blue-dark); box-shadow: 0 2px 8px rgba(27,63,139,0.4); }}
  .dropdown-btn.blue-btn:hover {{ background: #0e2152; }}
  .dropdown-btn.active {{ background: var(--tvs-blue-dark); box-shadow: 0 2px 8px rgba(27,63,139,0.4); }}
  .dropdown-btn .arrow {{ width: 15px; height: 15px; transition: transform 0.2s; flex-shrink: 0; }}
  .dropdown-btn.open .arrow {{ transform: rotate(180deg); }}
  .dropdown-menu {{
    position: absolute; top: calc(100% + 8px); right: 0;
    background: white; border: 1px solid var(--border); border-radius: 12px;
    overflow: hidden; overflow-y: auto; max-height: 360px;
    box-shadow: 0 16px 48px rgba(0,0,0,0.14);
    opacity: 0; transform: translateY(-8px); pointer-events: none;
    transition: opacity 0.18s, transform 0.18s; z-index: 200; min-width: 210px;
  }}
  .dropdown-menu.open {{ opacity: 1; transform: translateY(0); pointer-events: all; }}
  .dropdown-item {{
    padding: 11px 16px; font-size: 13px; font-weight: 500; cursor: pointer;
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
    transition: background 0.12s; border-bottom: 1px solid var(--border); color: var(--text);
  }}
  .dropdown-item:last-child {{ border-bottom: none; }}
  .dropdown-item:hover {{ background: var(--tvs-blue-light); color: var(--tvs-blue); }}
  .dropdown-item.selected {{ background: var(--tvs-blue-light); color: var(--tvs-blue); font-weight: 600; }}
  .event-count {{ font-size: 11px; background: var(--surface2); color: var(--muted); padding: 2px 8px; border-radius: 20px; font-weight: 400; flex-shrink: 0; }}
  .dropdown-item.selected .event-count {{ background: var(--tvs-blue); color: white; }}
  .dropdown-item .check {{ width:16px; height:16px; border:2px solid var(--border); border-radius:4px; flex-shrink:0; display:flex; align-items:center; justify-content:center; transition:all 0.15s; }}
  .dropdown-item.selected .check {{ background:var(--tvs-blue); border-color:var(--tvs-blue); }}
  .dropdown-item.selected .check::after {{ content:'✓'; color:white; font-size:10px; font-weight:700; }}
  .dropdown-divider {{ border:none; border-top:2px solid var(--tvs-red); margin:4px 0; }}
  .dropdown-action {{ padding:10px 16px; font-size:12px; font-weight:600; color:var(--tvs-blue); cursor:pointer; text-align:center; transition:background 0.12s; }}
  .dropdown-action:hover {{ background:var(--tvs-blue-light); }}

  .filter-bar {{
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
    padding: 10px 40px; background: var(--tvs-blue-light);
    border-bottom: 1px solid #c5d3ef; min-height: 44px;
  }}
  .filter-label {{ font-size: 11px; font-weight: 600; color: var(--tvs-blue); text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px; }}
  .filter-tag {{ display: flex; align-items: center; gap: 6px; background: var(--tvs-red); color: white; padding: 4px 10px 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }}
  .filter-tag.blue {{ background: var(--tvs-blue); }}
  .filter-tag.green {{ background: #1a7a4a; }}
  .filter-tag button {{ background: none; border: none; color: rgba(255,255,255,0.7); cursor: pointer; font-size: 14px; line-height: 1; padding: 0; display: flex; align-items: center; transition: color 0.1s; }}
  .filter-tag button:hover {{ color: white; }}
  .no-filters {{ font-size: 12px; color: var(--muted); font-style: italic; }}

  .stats-bar {{
    display: flex; gap: 10px; padding: 14px 40px;
    background: white; border-bottom: 1px solid var(--border);
    overflow-x: auto; flex-wrap: wrap; box-shadow: 0 1px 4px rgba(0,0,0,0.05);
  }}
  .stat-pill {{ display: flex; align-items: center; gap: 7px; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 500; white-space: nowrap; border: 1px solid transparent; }}
  .stat-pill .dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
  .stat-pill.total    {{ background: var(--tvs-blue-light); color: var(--tvs-blue); border-color: #c5d3ef; }}
  .stat-pill.total .dot {{ background: var(--tvs-blue); }}
  .stat-pill.executed {{ background: var(--executed-bg); color: var(--executed); border-color: #b8dfc9; }}
  .stat-pill.executed .dot {{ background: var(--executed); }}
  .stat-pill.postponed {{ background: var(--postponed-bg); color: var(--postponed); border-color: #f5b8b8; }}
  .stat-pill.postponed .dot {{ background: var(--postponed); }}
  .stat-pill.yet {{ background: var(--yet-bg); color: var(--yet); border-color: #e8d48a; }}
  .stat-pill.yet .dot {{ background: var(--yet); }}

  main {{ padding: 28px 40px 40px; }}
  .month-section {{ margin-bottom: 36px; }}
  .month-title {{
    font-family: 'Montserrat', sans-serif; font-size: 22px; font-weight: 800;
    letter-spacing: -0.5px; margin-bottom: 16px;
    display: flex; align-items: baseline; gap: 10px;
    color: var(--tvs-blue); border-left: 4px solid var(--tvs-red); padding-left: 14px;
    animation: fadeUp 0.3s ease;
  }}
  .month-title .year {{ color: var(--muted); font-size: 16px; font-weight: 400; }}
  @keyframes fadeUp {{ from {{ opacity:0; transform:translateY(10px); }} to {{ opacity:1; transform:translateY(0); }} }}

  .country-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 14px; }}
  .country-card {{ background: white; border: 1px solid var(--border); border-radius: 12px; overflow: hidden; transition: border-color 0.2s, box-shadow 0.2s, transform 0.2s; }}
  .country-card:hover {{ border-color: var(--tvs-blue); box-shadow: 0 6px 24px rgba(27,63,139,0.13); transform: translateY(-2px); }}
  .country-header {{ padding: 13px 16px; display: flex; align-items: center; gap: 10px; background: var(--tvs-blue); border-bottom: 3px solid var(--tvs-red); }}
  .country-flag {{ width: 30px; height: 30px; border-radius: 6px; background: rgba(255,255,255,0.12); display: flex; align-items: center; justify-content: center; font-size: 17px; flex-shrink: 0; }}
  .country-name {{ font-family: 'Montserrat', sans-serif; font-size: 14px; font-weight: 700; flex: 1; color: white; }}
  .country-event-count {{ font-size: 11px; color: rgba(255,255,255,0.65); background: rgba(255,255,255,0.12); padding: 2px 9px; border-radius: 20px; }}
  .event-list {{ padding: 6px 0; }}

  .event-row {{ display: flex; align-items: flex-start; gap: 10px; padding: 9px 16px; border-bottom: 1px solid rgba(0,0,0,0.04); transition: background 0.1s; cursor: default; }}
  .event-row:last-child {{ border-bottom: none; }}
  .event-row.has-desc {{ cursor: pointer; }}
  .event-row.has-desc:hover {{ background: var(--tvs-blue-light); }}
  .event-row.warn-desc:hover {{ background: #fff8f0; }}
  .event-row:not(.has-desc):not(.warn-desc):hover {{ background: var(--surface2); }}

  .event-date {{ font-size: 11px; font-weight: 600; color: var(--muted); white-space: nowrap; min-width: 78px; padding-top: 2px; font-variant-numeric: tabular-nums; }}
  .event-info {{ flex: 1; min-width: 0; }}
  .event-name {{ font-size: 12.5px; line-height: 1.4; color: var(--text); word-break: break-word; }}
  .event-category {{ font-size: 10px; color: var(--muted); margin-top: 2px; font-style: italic; }}
  .event-right {{ display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; }}

  .status-badge {{ font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 20px; white-space: nowrap; text-transform: uppercase; letter-spacing: 0.4px; }}
  .badge-executed  {{ background: var(--executed-bg);  color: var(--executed); }}
  .badge-postponed {{ background: var(--postponed-bg); color: var(--postponed); }}
  .badge-yet       {{ background: var(--yet-bg);       color: var(--yet); }}
  .badge-unknown   {{ background: var(--surface2);      color: var(--muted); }}

  .desc-indicator {{ font-size: 10px; font-weight: 600; padding: 2px 7px; border-radius: 20px; white-space: nowrap; display: flex; align-items: center; gap: 3px; }}
  .desc-indicator.has  {{ background: #e6ecf8; color: var(--tvs-blue); }}
  .desc-indicator.warn {{ background: #fce8e8; color: var(--postponed); }}

  /* PR TABLE VIEW */
  .pr-section {{ margin-bottom: 36px; }}
  .pr-section-title {{
    font-family: 'Montserrat', sans-serif; font-size: 20px; font-weight: 800;
    color: var(--tvs-blue); border-left: 4px solid var(--tvs-red);
    padding-left: 14px; margin-bottom: 16px; animation: fadeUp 0.3s ease;
    display: flex; align-items: center; gap: 10px;
  }}
  .pr-badge-count {{ font-size: 12px; background: var(--tvs-red); color: white; padding: 3px 10px; border-radius: 20px; font-weight: 700; }}
  .pr-table-wrap {{ background: white; border: 1px solid var(--border); border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05); overflow-x: auto; }}
  .pr-table {{ width: 100%; border-collapse: collapse; min-width: 500px; }}
  .pr-table thead tr {{ background: var(--tvs-blue); }}
  .pr-table thead th {{ padding: 12px 16px; text-align: left; font-family: 'Montserrat', sans-serif; font-size: 11px; font-weight: 700; color: white; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; }}
  .pr-table tbody tr {{ border-bottom: 1px solid var(--border); transition: background 0.1s; }}
  .pr-table tbody tr:last-child {{ border-bottom: none; }}
  .pr-table tbody tr:hover {{ background: var(--tvs-blue-light); }}
  .pr-table tbody td {{ padding: 11px 16px; font-size: 13px; color: var(--text); vertical-align: middle; }}
  .pr-table tbody td.date-col {{ font-size: 11px; font-weight: 600; color: var(--muted); white-space: nowrap; font-variant-numeric: tabular-nums; }}
  .pr-table tbody td.cat-col {{ font-size: 11px; }}
  .pr-table tbody td.country-col {{ font-size: 12px; font-weight: 600; }}
  .cat-chip {{ display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 700; background: var(--tvs-blue-light); color: var(--tvs-blue); white-space: nowrap; }}
  .export-bar {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; flex-wrap: wrap; gap: 8px; }}
  .export-btn {{
    background: #1a7a4a; color: white; border: none; padding: 9px 16px;
    border-radius: 8px; font-family: 'Montserrat', sans-serif; font-size: 12px; font-weight: 700;
    cursor: pointer; display: flex; align-items: center; gap: 7px; transition: background 0.15s;
    box-shadow: 0 2px 6px rgba(26,122,74,0.4);
  }}
  .export-btn:hover {{ background: #145c38; }}

  .empty {{ text-align: center; padding: 60px 20px; color: var(--muted); font-size: 14px; }}

  /* MODAL */
  .modal-overlay {{
    position: fixed; inset: 0; background: rgba(10,15,30,0.6);
    z-index: 1000; display: flex; align-items: center; justify-content: center;
    padding: 20px; opacity: 0; pointer-events: none;
    transition: opacity 0.2s; backdrop-filter: blur(3px);
  }}
  .modal-overlay.open {{ opacity: 1; pointer-events: all; }}
  .modal {{
    background: white; border-radius: 16px; width: 100%; max-width: 520px;
    box-shadow: 0 24px 80px rgba(0,0,0,0.25);
    transform: translateY(16px) scale(0.97);
    transition: transform 0.25s ease; overflow: hidden;
  }}
  .modal-overlay.open .modal {{ transform: translateY(0) scale(1); }}
  .modal-header {{
    background: var(--tvs-blue); padding: 18px 22px;
    border-bottom: 3px solid var(--tvs-red);
    display: flex; align-items: flex-start; justify-content: space-between; gap: 12px;
  }}
  .modal-header-info {{ flex: 1; min-width: 0; }}
  .modal-country {{ font-size: 11px; color: rgba(255,255,255,0.6); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 4px; }}
  .modal-event {{ font-family: 'Montserrat', sans-serif; font-size: 16px; font-weight: 700; color: white; line-height: 1.3; }}
  .modal-meta {{ display: flex; align-items: center; gap: 8px; margin-top: 8px; flex-wrap: wrap; }}
  .modal-date {{ font-size: 11px; color: rgba(255,255,255,0.65); }}
  .modal-close {{ background: rgba(255,255,255,0.15); border: none; color: white; width: 30px; height: 30px; border-radius: 50%; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: background 0.15s; flex-shrink: 0; }}
  .modal-close:hover {{ background: rgba(255,255,255,0.25); }}
  .modal-body {{ padding: 22px; }}
  .modal-section-label {{ font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 10px; }}
  .modal-description {{ font-size: 14px; line-height: 1.7; color: var(--text); background: var(--surface2); border-radius: 10px; padding: 14px 16px; border-left: 3px solid var(--tvs-blue); }}
  .modal-cat-row {{ display: flex; align-items: center; gap: 8px; margin-top: 14px; padding-top: 14px; border-top: 1px solid var(--border); }}
  .modal-cat-label {{ font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; }}

  .legend {{ display: flex; gap: 20px; flex-wrap: wrap; padding: 0 40px 32px; font-size: 12px; color: var(--muted); align-items: center; }}
  .legend span {{ font-weight: 600; color: var(--tvs-blue); margin-right: 4px; }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .legend-dot {{ width: 10px; height: 10px; border-radius: 50%; }}

  /* ── MOBILE RESPONSIVE ── */
  @media (max-width: 768px) {{
    header {{ height: auto; padding: 10px 14px 12px; flex-direction: column; align-items: stretch; gap: 0; }}
    .header-left {{ display: flex; align-items: center; gap: 10px; margin-bottom: 6px; }}
    .tvs-logo {{ height: 34px; padding: 4px 9px; flex-shrink: 0; }}
    .tvs-logo img {{ height: 24px; }}
    .header-divider {{ display: none; }}
    .header-title h1 {{ font-size: 13px; white-space: normal; line-height: 1.3; }}
    .header-title p {{ display: block; font-size: 9px; margin-top: 1px; }}
    .header-right {{ width: 100%; flex-direction: column; align-items: stretch; gap: 6px; }}
    .updated {{ display: block; font-size: 10px; text-align: right; }}
    .header-btns {{ display: grid; grid-template-columns: 1fr 1fr; width: 100%; gap: 6px; }}
    .dropdown-wrap {{ flex: 1; }}
    .dropdown-btn {{ width: 100%; min-width: 0; font-size: 10px; font-weight: 700; padding: 9px 8px; justify-content: center; gap: 4px; box-shadow: none; }}
    .dropdown-btn .arrow {{ width: 12px; height: 12px; }}
    .filter-bar {{ padding: 8px 16px; gap: 6px; }}
    .stats-bar {{ padding: 10px 16px; gap: 8px; flex-wrap: nowrap; overflow-x: auto; -webkit-overflow-scrolling: touch; }}
    .stat-pill {{ font-size: 11px; padding: 5px 10px; }}
    main {{ padding: 16px 16px 40px; }}
    .month-section {{ margin-bottom: 28px; }}
    .month-title {{ font-size: 18px; padding-left: 10px; margin-bottom: 12px; }}
    .month-title .year {{ font-size: 14px; }}
    .country-grid {{ grid-template-columns: 1fr; gap: 10px; }}
    .country-header {{ padding: 11px 14px; }}
    .country-name {{ font-size: 13px; }}
    .country-flag {{ width: 26px; height: 26px; font-size: 15px; }}
    .event-row {{ padding: 9px 14px; gap: 8px; }}
    .event-date {{ min-width: 68px; font-size: 10px; }}
    .event-name {{ font-size: 12px; }}
    .status-badge {{ font-size: 9px; padding: 2px 6px; }}
    .desc-indicator {{ font-size: 9px; padding: 2px 6px; }}
    .modal-overlay {{ padding: 0; align-items: flex-end; }}
    .modal {{ border-radius: 20px 20px 0 0; max-width: 100%; max-height: 85vh; overflow-y: auto; transform: translateY(100%); }}
    .modal-overlay.open .modal {{ transform: translateY(0); }}
    .modal-header {{ padding: 16px 18px; }}
    .modal-event {{ font-size: 15px; }}
    .modal-body {{ padding: 18px; }}
    .modal-description {{ font-size: 13px; }}
    .legend {{ padding: 0 16px 24px; gap: 12px; font-size: 11px; }}
    #monthMenu, #countryMenu, #requiresPRMenu, #timePeriodMenu {{ min-width: 180px; right: 0; left: auto; }}
    .pr-table tbody td {{ font-size: 12px; padding: 9px 12px; }}
    .pr-table thead th {{ font-size: 10px; padding: 10px 12px; }}
    .export-bar {{ flex-direction: column; align-items: stretch; }}
    .export-btn {{ justify-content: center; }}
    .pr-section-title {{ font-size: 17px; }}
  }}
  @media (max-width: 380px) {{
    .header-title h1 {{ font-size: 11px; }}
    .dropdown-btn {{ font-size: 9px; padding: 8px 5px; }}
    .month-title {{ font-size: 15px; }}
  }}
</style>
</head>
<body>

<div class="top-strip"></div>

<header>
  <div class="header-left">
    {logo_html}
    <div class="header-divider"></div>
    <div class="header-title">
      <h1>NSE1 Important Dates</h1>
      <p>Africa Region · 2026 Event Calendar</p>
    </div>
  </div>
  <div class="header-right">
    <div class="updated">Last updated: <strong>{updated_at}</strong></div>
    <div class="header-btns">
      <!-- Month filter -->
      <div class="dropdown-wrap">
        <button class="dropdown-btn" id="monthBtn">
          <span id="monthLabel">All Months</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="monthMenu"></div>
      </div>
      <!-- Country filter -->
      <div class="dropdown-wrap">
        <button class="dropdown-btn" id="countryBtn">
          <span id="countryLabel">All Countries</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="countryMenu"></div>
      </div>
      <!-- Requires PR filter -->
      <div class="dropdown-wrap">
        <button class="dropdown-btn blue-btn" id="requiresPRBtn">
          <span id="requiresPRLabel">Requires PR</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="requiresPRMenu">
          <div class="dropdown-item" data-value="all" onclick="setRequiresPR('all')">All Events</div>
          <div class="dropdown-item" data-value="Yes" onclick="setRequiresPR('Yes')">📣 Requires PR — Yes</div>
          <div class="dropdown-item" data-value="No" onclick="setRequiresPR('No')">No PR Required</div>
        </div>
      </div>
      <!-- Time period dropdown -->
      <div class="dropdown-wrap">
        <button class="dropdown-btn" id="timePeriodBtn">
          <span id="timePeriodLabel">Time Period</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="timePeriodMenu">
          <div class="dropdown-item" data-value="all" onclick="setTimePeriod('all')">All Dates</div>
          <div class="dropdown-item" data-value="45" onclick="setTimePeriod('45')">📅 Next 45 Days</div>
          <div class="dropdown-item" data-value="60" onclick="setTimePeriod('60')">📅 Next 60 Days</div>
          <div class="dropdown-item" data-value="90" onclick="setTimePeriod('90')">📅 Next 90 Days</div>
        </div>
      </div>
    </div>
  </div>
</header>

<div class="filter-bar" id="filterBar">
  <span class="filter-label">Filters:</span>
  <span class="no-filters" id="noFilters">None — showing all events</span>
</div>

<div class="stats-bar" id="statsBar"></div>
<main id="main"></main>

<div class="legend">
  <span>Legend:</span>
  <div class="legend-item"><div class="legend-dot" style="background:var(--executed)"></div> Executed</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--postponed)"></div> Postponed</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--yet)"></div> Upcoming</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--tvs-blue)"></div> 📋 Has report</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--postponed)"></div> ⚠️ Report pending</div>
</div>

<div class="modal-overlay" id="modalOverlay" onclick="closeModal(event)">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-header-info">
        <div class="modal-country" id="modalCountry"></div>
        <div class="modal-event"   id="modalEvent"></div>
        <div class="modal-meta">
          <span class="modal-date"  id="modalDate"></span>
          <span class="status-badge" id="modalBadge"></span>
        </div>
      </div>
      <button class="modal-close" onclick="closeModalDirect()">✕</button>
    </div>
    <div class="modal-body">
      <div class="modal-section-label">Event Report</div>
      <div class="modal-description" id="modalDesc"></div>
      <div class="modal-cat-row">
        <span class="modal-cat-label">Category:</span>
        <span class="cat-chip" id="modalCat"></span>
      </div>
    </div>
  </div>
</div>

<script>
const RAW = {raw_json};
const FLAGS = {{
  DRC:'🇨🇩', Egypt:'🇪🇬', Kenya:'🇰🇪', Madagascar:'🇲🇬', Morocco:'🇲🇦',
  Mozambique:'🇲🇿', Regional:'🌍', 'South Africa':'🇿🇦', Tanzania:'🇹🇿',
  Tunisia:'🇹🇳', Uganda:'🇺🇬', Zambia:'🇿🇲', Mauritius:'🇲🇺'
}};
const ALL_COUNTRIES = [...new Set(RAW.months.flatMap(m=>Object.keys(RAW.data[m])))].sort();

// State
let selectedMonths    = new Set();
let selectedCountries = new Set();
let requiresPRFilter  = 'all';  // 'all' | 'Yes' | 'No'
let timePeriodDays    = 'all';  // 'all' | '45' | '60' | '90'

const TODAY = new Date(); TODAY.setHours(0,0,0,0);

function getEndDate() {{
  if (timePeriodDays === 'all') return null;
  const d = new Date(TODAY);
  d.setDate(d.getDate() + parseInt(timePeriodDays));
  return d;
}}
function parseDate(str) {{ const d = new Date(str); return isNaN(d) ? null : d; }}
function countEventsInMonth(m) {{ return Object.values(RAW.data[m]).reduce((s,e)=>s+e.length,0); }}
function countEventsForCountry(c) {{ return RAW.months.reduce((s,m)=>s+(RAW.data[m][c]?.length||0),0); }}

function badgeClass(s) {{
  if(!s) return 'badge-unknown';
  const l=s.toLowerCase();
  if(l.includes('executed'))  return 'badge-executed';
  if(l.includes('postponed')) return 'badge-postponed';
  if(l.includes('yet'))       return 'badge-yet';
  return 'badge-unknown';
}}
function badgeLabel(s) {{
  if(!s) return '—';
  if(s.toLowerCase().includes('executed'))  return 'Executed';
  if(s.toLowerCase().includes('postponed')) return 'Postponed';
  if(s.toLowerCase().includes('yet'))       return 'Upcoming';
  return s;
}}

// ── Modal ─────────────────────────────────────────────────────────────────────
function openModal(country, event) {{
  document.getElementById('modalCountry').textContent = country;
  document.getElementById('modalEvent').textContent   = event.Event;
  document.getElementById('modalDate').textContent    = event.DateStr;
  const badge = document.getElementById('modalBadge');
  badge.textContent = badgeLabel(event.Status);
  badge.className   = 'status-badge ' + badgeClass(event.Status);
  document.getElementById('modalDesc').textContent = event.Description || 'No report added yet.';
  document.getElementById('modalCat').textContent  = event.CategoryOfEvent || '—';
  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}}
function closeModalDirect() {{
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}}
function closeModal(e) {{ if(e.target === document.getElementById('modalOverlay')) closeModalDirect(); }}
document.addEventListener('keydown', e => {{ if(e.key==='Escape') closeModalDirect(); }});

// ── Requires PR dropdown ───────────────────────────────────────────────────────
function setRequiresPR(val) {{
  requiresPRFilter = val;
  document.querySelectorAll('#requiresPRMenu .dropdown-item').forEach(i => {{
    i.classList.toggle('selected', i.dataset.value === val);
  }});
  closeAll();
  applyFilters();
}}

// ── Time period dropdown ──────────────────────────────────────────────────────
function setTimePeriod(val) {{
  timePeriodDays = val;
  if (val !== 'all') selectedMonths.clear();
  document.querySelectorAll('#timePeriodMenu .dropdown-item').forEach(i => {{
    i.classList.toggle('selected', i.dataset.value === val);
  }});
  closeAll();
  applyFilters();
}}

// ── Filter bar ────────────────────────────────────────────────────────────────
function renderFilterBar() {{
  const bar = document.getElementById('filterBar');
  bar.querySelectorAll('.filter-tag').forEach(t=>t.remove());
  const noF = document.getElementById('noFilters');
  const hasFilters = selectedMonths.size || selectedCountries.size || timePeriodDays !== 'all' || requiresPRFilter !== 'all';
  if (!hasFilters) {{ noF.style.display='inline'; return; }}
  noF.style.display='none';
  if (timePeriodDays !== 'all') {{
    const t=document.createElement('div'); t.className='filter-tag';
    t.innerHTML=`📅 Next ${{timePeriodDays}} Days <button onclick="setTimePeriod('all')">✕</button>`;
    bar.appendChild(t);
  }}
  if (requiresPRFilter !== 'all') {{
    const t=document.createElement('div'); t.className='filter-tag green';
    const label = requiresPRFilter === 'Yes' ? '📣 Requires PR: Yes' : 'No PR Required';
    t.innerHTML=`${{label}} <button onclick="setRequiresPR('all')">✕</button>`;
    bar.appendChild(t);
  }}
  selectedMonths.forEach(m=>{{
    const t=document.createElement('div'); t.className='filter-tag';
    t.innerHTML=`📅 ${{m}} <button onclick="removeMonth('${{m}}')">✕</button>`;
    bar.appendChild(t);
  }});
  selectedCountries.forEach(c=>{{
    const t=document.createElement('div'); t.className='filter-tag blue';
    t.innerHTML=`${{FLAGS[c]||'🌐'}} ${{c}} <button onclick="removeCountry('${{c.replace(/'/g,"\\'")}}')">✕</button>`;
    bar.appendChild(t);
  }});
}}

function removeMonth(m)   {{ selectedMonths.delete(m);   applyFilters(); }}
function removeCountry(c) {{ selectedCountries.delete(c); applyFilters(); }}

// ── Get filtered events ───────────────────────────────────────────────────────
function getFilteredEntries() {{
  const endDate = getEndDate();
  const result = [];
  const months = selectedMonths.size ? RAW.months.filter(m=>selectedMonths.has(m)) : RAW.months;
  months.forEach(m => {{
    let entries = Object.entries(RAW.data[m]);
    if (selectedCountries.size) entries = entries.filter(([c])=>selectedCountries.has(c));
    entries.forEach(([country, events]) => {{
      let evts = events;
      if (endDate) {{
        evts = events.filter(e => {{
          const d = parseDate(e.DateStr);
          return d && d >= TODAY && d <= endDate;
        }});
      }}
      if (requiresPRFilter !== 'all') {{
        evts = evts.filter(e => (e.RequiresPR || 'No') === requiresPRFilter);
      }}
      if (evts.length) result.push({{ month:m, country, events:evts }});
    }});
  }});
  return result;
}}

// ── Stats ─────────────────────────────────────────────────────────────────────
function renderStats() {{
  const filtered = getFilteredEntries();
  const c = {{executed:0,postponed:0,yet:0,unknown:0}};
  filtered.forEach(item => item.events.forEach(e=>{{
    const s=(e.Status||'').toLowerCase();
    if(s.includes('executed'))c.executed++;
    else if(s.includes('postponed'))c.postponed++;
    else if(s.includes('yet'))c.yet++;
    else c.unknown++;
  }}));
  const total=c.executed+c.postponed+c.yet+c.unknown;
  document.getElementById('statsBar').innerHTML=`
    <div class="stat-pill total"><div class="dot"></div>${{total}} Total Events</div>
    <div class="stat-pill executed"><div class="dot"></div>${{c.executed}} Executed</div>
    <div class="stat-pill postponed"><div class="dot"></div>${{c.postponed}} Postponed</div>
    <div class="stat-pill yet"><div class="dot"></div>${{c.yet}} Upcoming</div>`;
}}

// ── Export to CSV (opens in Excel) ────────────────────────────────────────────
function exportPRToExcel() {{
  const filtered = getFilteredEntries();
  const rows = [['Date of Event','Category of Event','Event Name','Country']];
  filtered.forEach(item => item.events.forEach(e => {{
    if ((e.RequiresPR||'No') === 'Yes') {{
      rows.push([e.DateStr, e.CategoryOfEvent||'', e.Event, item.country]);
    }}
  }}));
  if (rows.length === 1) {{ alert('No PR events to export.'); return; }}
  const csv = rows.map(r => r.map(c => `"${{String(c).replace(/"/g,'""')}}"`).join(',')).join('\\r\\n');
  const blob = new Blob(['\\uFEFF' + csv], {{type:'text/csv;charset=utf-8;'}});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = 'TVS_NSE1_Requires_PR.csv';
  a.click(); URL.revokeObjectURL(url);
}}

// ── Cards view ────────────────────────────────────────────────────────────────
function renderCardsView(filtered) {{
  const byMonth = {{}};
  filtered.forEach(item => {{
    if (!byMonth[item.month]) byMonth[item.month] = {{}};
    byMonth[item.month][item.country] = item.events;
  }});
  let html='';
  RAW.months.filter(m=>byMonth[m]).forEach(m => {{
    const [mon,yr]=m.split(' ');
    html+=`<div class="month-section">
      <div class="month-title">${{mon}} <span class="year">${{yr}}</span></div>
      <div class="country-grid">
        ${{Object.entries(byMonth[m]).map(([country,events])=>`
          <div class="country-card">
            <div class="country-header">
              <div class="country-flag">${{FLAGS[country]||'🌐'}}</div>
              <div class="country-name">${{country}}</div>
              <div class="country-event-count">${{events.length}} event${{events.length!==1?'s':''}}</div>
            </div>
            <div class="event-list">
              ${{events.map((e)=>{{
                const origIdx = RAW.data[m][country].indexOf(e);
                const hasDesc  = e.DatePassed && e.Description;
                const warnDesc = e.DatePassed && !e.Description && e.Status.toLowerCase().includes('executed');
                const rowClass = hasDesc?'event-row has-desc':warnDesc?'event-row warn-desc':'event-row';
                const onclick  = hasDesc?`onclick="openModal('${{country.replace(/'/g,"\\'")}}', RAW.data['${{m}}']['${{country}}'][${{origIdx}}])"` :'';
                const indicator= hasDesc?`<div class="desc-indicator has">📋 Report</div>`:warnDesc?`<div class="desc-indicator warn">⚠️ Add report</div>`:'';
                const catLabel = e.CategoryOfEvent ? `<div class="event-category">${{e.CategoryOfEvent}}</div>` : '';
                return `<div class="${{rowClass}}" ${{onclick}}>
                  <div class="event-date">${{e.DateStr||'—'}}</div>
                  <div class="event-info"><div class="event-name">${{e.Event}}</div>${{catLabel}}</div>
                  <div class="event-right">
                    <div class="status-badge ${{badgeClass(e.Status)}}">${{badgeLabel(e.Status)}}</div>
                    ${{indicator}}
                  </div>
                </div>`;
              }}).join('')}}
            </div>
          </div>`).join('')}}
      </div>
    </div>`;
  }});
  return html;
}}

// ── PR table view ─────────────────────────────────────────────────────────────
function renderPRTableView(filtered) {{
  const prEvents = [];
  filtered.forEach(item => item.events.forEach(e => {{
    prEvents.push({{ ...e, _country: item.country }});
  }}));
  if (!prEvents.length) return '<div class="empty">No PR events match your filters.</div>';
  const exportBar = `<div class="export-bar">
    <div class="pr-section-title">📣 Requires PR — Events <span class="pr-badge-count">${{prEvents.length}}</span></div>
    <button class="export-btn" onclick="exportPRToExcel()">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Export to Excel
    </button>
  </div>`;
  const rows = prEvents.map(e => `
    <tr>
      <td class="date-col">${{e.DateStr||'—'}}</td>
      <td class="cat-col"><span class="cat-chip">${{e.CategoryOfEvent||'—'}}</span></td>
      <td>${{e.Event}}</td>
      <td class="country-col">${{FLAGS[e._country]||'🌐'}} ${{e._country}}</td>
    </tr>`).join('');
  return `<div class="pr-section">
    ${{exportBar}}
    <div class="pr-table-wrap">
      <table class="pr-table">
        <thead><tr>
          <th>Date of Event</th><th>Category of Event</th><th>Event Name</th><th>Country</th>
        </tr></thead>
        <tbody>${{rows}}</tbody>
      </table>
    </div>
  </div>`;
}}

// ── Main render ────────────────────────────────────────────────────────────────
function renderMain() {{
  const filtered = getFilteredEntries();
  const mainEl = document.getElementById('main');
  if (!filtered.length) {{ mainEl.innerHTML='<div class="empty">No events match your filters.</div>'; return; }}
  mainEl.innerHTML = requiresPRFilter === 'Yes' ? renderPRTableView(filtered) : renderCardsView(filtered);
}}

// ── Apply all filters ─────────────────────────────────────────────────────────
function applyFilters() {{
  const mCount = selectedMonths.size;
  document.getElementById('monthLabel').textContent = mCount===0?'All Months':mCount===1?[...selectedMonths][0]:`${{mCount}} Months`;
  const cCount = selectedCountries.size;
  document.getElementById('countryLabel').textContent = cCount===0?'All Countries':cCount===1?[...selectedCountries][0]:`${{cCount}} Countries`;
  const prLabels = {{ all:'Requires PR', Yes:'📣 PR: Yes', No:'No PR' }};
  document.getElementById('requiresPRLabel').textContent = prLabels[requiresPRFilter] || 'Requires PR';
  document.getElementById('requiresPRBtn').classList.toggle('active', requiresPRFilter !== 'all');
  const tpLabels = {{ all:'Time Period', '45':'Next 45 Days', '60':'Next 60 Days', '90':'Next 90 Days' }};
  document.getElementById('timePeriodLabel').textContent = tpLabels[timePeriodDays] || 'Time Period';
  document.getElementById('timePeriodBtn').classList.toggle('active', timePeriodDays !== 'all');
  document.querySelectorAll('#monthMenu .dropdown-item[data-value]').forEach(i=>{{
    i.classList.toggle('selected', selectedMonths.has(i.dataset.value));
  }});
  document.querySelectorAll('#countryMenu .dropdown-item[data-value]').forEach(i=>{{
    i.classList.toggle('selected', selectedCountries.has(i.dataset.value));
  }});
  renderFilterBar(); renderStats(); renderMain();
}}

// ── Build dropdowns ───────────────────────────────────────────────────────────
function buildMonthDropdown() {{
  document.getElementById('monthMenu').innerHTML =
    `<div class="dropdown-action" onclick="selectedMonths.clear();applyFilters()">Clear selection</div>
     <hr class="dropdown-divider"/>`
    + RAW.months.map(m=>`
      <div class="dropdown-item" data-value="${{m}}">
        <div class="check"></div><span style="flex:1">${{m}}</span>
        <span class="event-count">${{countEventsInMonth(m)}}</span>
      </div>`).join('');
  document.querySelectorAll('#monthMenu .dropdown-item[data-value]').forEach(item=>
    item.addEventListener('click', e=>{{
      e.stopPropagation();
      const v = item.dataset.value;
      if (selectedMonths.has(v)) selectedMonths.delete(v);
      else {{ timePeriodDays='all'; document.querySelectorAll('#timePeriodMenu .dropdown-item').forEach(i=>i.classList.remove('selected')); selectedMonths.add(v); }}
      applyFilters();
    }})
  );
}}

function buildCountryDropdown() {{
  document.getElementById('countryMenu').innerHTML =
    `<div class="dropdown-action" onclick="selectedCountries.clear();applyFilters()">Clear selection</div>
     <hr class="dropdown-divider"/>`
    + ALL_COUNTRIES.map(c=>`
      <div class="dropdown-item" data-value="${{c}}">
        <div class="check"></div><span style="flex:1">${{FLAGS[c]||'🌐'}} ${{c}}</span>
        <span class="event-count">${{countEventsForCountry(c)}}</span>
      </div>`).join('');
  document.querySelectorAll('#countryMenu .dropdown-item[data-value]').forEach(item=>
    item.addEventListener('click', e=>{{
      e.stopPropagation();
      const v = item.dataset.value;
      if (selectedCountries.has(v)) selectedCountries.delete(v);
      else selectedCountries.add(v);
      applyFilters();
    }})
  );
}}

// ── Dropdown open/close ───────────────────────────────────────────────────────
function closeAll() {{
  document.querySelectorAll('.dropdown-btn').forEach(b=>b.classList.remove('open'));
  document.querySelectorAll('.dropdown-menu').forEach(m=>m.classList.remove('open'));
}}
function toggleDropdown(btnId,menuId) {{
  const isOpen=document.getElementById(menuId).classList.contains('open');
  closeAll();
  if(!isOpen){{document.getElementById(btnId).classList.add('open');document.getElementById(menuId).classList.add('open');}}
}}
document.getElementById('monthBtn').addEventListener('click',   e=>{{e.stopPropagation();toggleDropdown('monthBtn','monthMenu');}});
document.getElementById('countryBtn').addEventListener('click', e=>{{e.stopPropagation();toggleDropdown('countryBtn','countryMenu');}});
document.getElementById('requiresPRBtn').addEventListener('click', e=>{{e.stopPropagation();toggleDropdown('requiresPRBtn','requiresPRMenu');}});
document.getElementById('timePeriodBtn').addEventListener('click', e=>{{e.stopPropagation();toggleDropdown('timePeriodBtn','timePeriodMenu');}});
document.addEventListener('click', closeAll);

buildMonthDropdown();
buildCountryDropdown();
applyFilters();
</script>
</body>
</html>"""


def main():
    print("\n── TVS NSE1 Dashboard Updater ──────────────────────")
    print(f"  Reading:  {EXCEL_FILE}")
    df      = load_data(EXCEL_FILE)
    payload = build_json(df)
    logo    = get_logo_b64()
    updated = datetime.now().strftime("%d %b %Y, %H:%M")
    total   = sum(len(e) for m in payload['data'].values() for e in m.values())
    months  = len(payload['months'])
    html    = build_html(payload, updated, logo)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"  Logo:     {'✅ Embedded' if logo else '⚠️  Not found — add tvs_logo.jpg to folder'}")
    print(f"  Months:   {months}  |  Events: {total}")
    print(f"  Saved:    {OUTPUT_FILE}")
    print(f"  Updated:  {updated}")
    print("────────────────────────────────────────────────────")
    print("  ✅  Dashboard ready! Share the HTML with your team.")
    print("────────────────────────────────────────────────────\n")

if __name__ == "__main__":
    main()