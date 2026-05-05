"""
NSE1 Dashboard Updater — TVS Branded
=====================================
Run this script whenever you update your Excel file.

Usage:
    python update_dashboard.py

Requirements:
    pip install pandas openpyxl
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
    required = {'Country', 'Date', 'Event', 'Status'}
    if not required.issubset(df.columns):
        print(f"\n❌  Excel must have columns: {required}\n    Found: {set(df.columns)}\n")
        sys.exit(1)
    df['Date']    = pd.to_datetime(df['Date'], errors='coerce')
    df            = df.dropna(subset=['Date'])
    df['Month']   = df['Date'].dt.strftime('%B %Y')
    df['DateStr'] = df['Date'].dt.strftime('%d %b %Y')
    df['Status']  = df['Status'].fillna('').str.strip()
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
                    'DateStr':     row['DateStr'],
                    'Event':       row['Event'],
                    'Status':      row['Status'],
                    'Description': desc,
                    'DatePassed':  date_passed
                })
            data[month][country] = records
    return {'months': month_order, 'data': data}

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
<link rel="icon" type="image/png" href="data:image/png;base64,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"/>
<link rel="apple-touch-icon" href="data:image/png;base64,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"/>
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
  .header-btns {{ display: flex; align-items: center; gap: 12px; }}
  .updated {{ font-size: 11px; color: rgba(255,255,255,0.5); white-space: nowrap; }}
  .updated strong {{ color: rgba(255,255,255,0.85); }}

  .dropdown-wrap {{ position: relative; flex-shrink: 0; }}
  .dropdown-btn {{
    background: var(--tvs-red); color: white; border: none;
    padding: 11px 16px; font-family: 'Montserrat', sans-serif; font-size: 13px; font-weight: 700;
    border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 10px;
    transition: background 0.15s, transform 0.1s; min-width: 160px; justify-content: space-between;
    box-shadow: 0 2px 8px rgba(204,19,19,0.4); white-space: nowrap;
  }}
  .dropdown-btn:hover {{ background: var(--tvs-red-dark); }}
  .dropdown-btn:active {{ transform: scale(0.98); }}
  .dropdown-btn .arrow {{ width: 16px; height: 16px; transition: transform 0.2s; flex-shrink: 0; }}
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
  .dropdown-item.active {{ background: var(--tvs-blue-light); color: var(--tvs-blue); font-weight: 700; }}
  .event-count {{ font-size: 11px; background: var(--surface2); color: var(--muted); padding: 2px 8px; border-radius: 20px; font-weight: 400; flex-shrink: 0; }}
  .dropdown-item.active .event-count {{ background: var(--tvs-blue); color: white; }}

  .filter-bar {{
    display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
    padding: 10px 40px; background: var(--tvs-blue-light);
    border-bottom: 1px solid #c5d3ef; min-height: 44px;
  }}
  .filter-label {{ font-size: 11px; font-weight: 600; color: var(--tvs-blue); text-transform: uppercase; letter-spacing: 0.5px; margin-right: 4px; }}
  .filter-tag {{ display: flex; align-items: center; gap: 6px; background: var(--tvs-red); color: white; padding: 4px 10px 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }}
  .filter-tag.blue {{ background: var(--tvs-blue); }}
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
  .event-right {{ display: flex; flex-direction: column; align-items: flex-end; gap: 4px; flex-shrink: 0; }}

  .status-badge {{ font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 20px; white-space: nowrap; text-transform: uppercase; letter-spacing: 0.4px; }}
  .badge-executed  {{ background: var(--executed-bg);  color: var(--executed); }}
  .badge-postponed {{ background: var(--postponed-bg); color: var(--postponed); }}
  .badge-yet       {{ background: var(--yet-bg);       color: var(--yet); }}
  .badge-unknown   {{ background: var(--surface2);      color: var(--muted); }}

  .desc-indicator {{ font-size: 10px; font-weight: 600; padding: 2px 7px; border-radius: 20px; white-space: nowrap; display: flex; align-items: center; gap: 3px; }}
  .desc-indicator.has  {{ background: #e6ecf8; color: var(--tvs-blue); }}
  .desc-indicator.warn {{ background: #fce8e8; color: var(--postponed); }}

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

  .dropdown-btn.next60Btn {{ min-width: auto; padding: 11px 14px; white-space: nowrap; }}
  .dropdown-btn.next60Btn.active {{ background: var(--tvs-blue); box-shadow: 0 2px 8px rgba(27,63,139,0.4); }}

  .dropdown-item .check {{ width:16px; height:16px; border:2px solid var(--border); border-radius:4px; flex-shrink:0; display:flex; align-items:center; justify-content:center; transition:all 0.15s; }}
  .dropdown-item.selected .check {{ background:var(--tvs-blue); border-color:var(--tvs-blue); }}
  .dropdown-item.selected .check::after {{ content:'✓'; color:white; font-size:10px; font-weight:700; }}
  .dropdown-item.selected {{ background:var(--tvs-blue-light); color:var(--tvs-blue); font-weight:600; }}
  .dropdown-divider {{ border:none; border-top:2px solid var(--tvs-red); margin:4px 0; }}
  .dropdown-action {{ padding:10px 16px; font-size:12px; font-weight:600; color:var(--tvs-blue); cursor:pointer; text-align:center; transition:background 0.12s; }}
  .dropdown-action:hover {{ background:var(--tvs-blue-light); }}

  .legend {{ display: flex; gap: 20px; flex-wrap: wrap; padding: 0 40px 32px; font-size: 12px; color: var(--muted); align-items: center; }}
  .legend span {{ font-weight: 600; color: var(--tvs-blue); margin-right: 4px; }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .legend-dot {{ width: 10px; height: 10px; border-radius: 50%; }}

  /* ── MOBILE RESPONSIVE ── */
  @media (max-width: 768px) {{

    /* Header — 3-row stack on mobile:
       Row 1: logo + title
       Row 2: last updated
       Row 3: 3 filter buttons full-width */
    header {{
      height: auto;
      padding: 10px 14px 12px;
      flex-direction: column;
      align-items: stretch;
      gap: 0;
    }}

    /* Row 1: logo + title side by side */
    .header-left {{
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 6px;
    }}
    .tvs-logo {{ height: 34px; padding: 4px 9px; flex-shrink: 0; }}
    .tvs-logo img {{ height: 24px; }}
    .header-divider {{ display: none; }}
    .header-title h1 {{
      font-size: 13px;
      white-space: normal;
      line-height: 1.3;
    }}
    .header-title p {{ display: block; font-size: 9px; margin-top: 1px; }}

    /* Row 2: last updated — full width, right-aligned under title */
    .header-right {{
      width: 100%;
      flex-direction: column;
      align-items: stretch;
      gap: 8px;
    }}
    .updated {{
      display: block;
      font-size: 10px;
      color: rgba(255,255,255,0.5);
      text-align: right;
    }}
    .updated strong {{ color: rgba(255,255,255,0.8); }}

    /* Row 3: all three buttons equal width, full viewport */
    .header-btns {{
      display: flex;
      width: 100%;
      gap: 6px;
    }}
    .dropdown-wrap {{ flex: 1 1 0; }}
    .dropdown-btn {{
      width: 100%;
      min-width: 0;
      flex: 1 1 0;
      font-size: 11px;
      font-weight: 700;
      padding: 10px 6px;
      justify-content: center;
      gap: 0;
      box-shadow: none;
    }}
    .dropdown-btn .arrow {{ display: none; }}
    .dropdown-btn.next60Btn {{
      flex: 1 1 0;
      width: 100%;
      padding: 10px 6px;
      font-size: 11px;
      white-space: nowrap;
      min-width: 0;
      box-shadow: none;
    }}

    /* Filter bar */
    .filter-bar {{ padding: 8px 16px; gap: 6px; }}

    /* Stats bar — scrollable single line */
    .stats-bar {{
      padding: 10px 16px;
      gap: 8px;
      flex-wrap: nowrap;
      overflow-x: auto;
      -webkit-overflow-scrolling: touch;
    }}
    .stat-pill {{ font-size: 11px; padding: 5px 10px; }}

    /* Main content */
    main {{ padding: 16px 16px 40px; }}
    .month-section {{ margin-bottom: 28px; }}
    .month-title {{ font-size: 18px; padding-left: 10px; margin-bottom: 12px; }}
    .month-title .year {{ font-size: 14px; }}

    /* Country grid — single column on mobile */
    .country-grid {{
      grid-template-columns: 1fr;
      gap: 10px;
    }}

    /* Country cards */
    .country-header {{ padding: 11px 14px; }}
    .country-name {{ font-size: 13px; }}
    .country-flag {{ width: 26px; height: 26px; font-size: 15px; }}

    /* Event rows */
    .event-row {{ padding: 9px 14px; gap: 8px; }}
    .event-date {{ min-width: 68px; font-size: 10px; }}
    .event-name {{ font-size: 12px; }}
    .status-badge {{ font-size: 9px; padding: 2px 6px; }}
    .desc-indicator {{ font-size: 9px; padding: 2px 6px; }}

    /* Modal — slide up from bottom on mobile */
    .modal-overlay {{ padding: 0; align-items: flex-end; }}
    .modal {{
      border-radius: 20px 20px 0 0;
      max-width: 100%;
      max-height: 85vh;
      overflow-y: auto;
      transform: translateY(100%);
    }}
    .modal-overlay.open .modal {{ transform: translateY(0); }}
    .modal-header {{ padding: 16px 18px; }}
    .modal-event {{ font-size: 15px; }}
    .modal-body {{ padding: 18px; }}
    .modal-description {{ font-size: 13px; }}

    /* Legend */
    .legend {{ padding: 0 16px 24px; gap: 12px; font-size: 11px; }}

    /* Dropdown menus positioning */
    #monthMenu  {{ min-width: 160px; left: 0 !important; right: auto !important; }}
    #countryMenu {{ min-width: 160px; left: auto !important; right: 0 !important; }}
  }}

  @media (max-width: 380px) {{
    .header-title h1 {{ font-size: 12px; }}
    .dropdown-btn, .dropdown-btn.next60Btn {{ font-size: 10px; padding: 9px 4px; }}
    .month-title {{ font-size: 16px; }}
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
      <div class="dropdown-wrap">
        <button class="dropdown-btn" id="monthBtn">
          <span id="monthLabel">All Months</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="monthMenu"></div>
      </div>
      <div class="dropdown-wrap">
        <button class="dropdown-btn" id="countryBtn">
          <span id="countryLabel">All Countries</span>
          <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </button>
        <div class="dropdown-menu" id="countryMenu"></div>
      </div>
      <button class="dropdown-btn next60Btn" id="next60Btn" onclick="toggleNext60()">
        Next 60 Days
      </button>
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
    </div>
  </div>
</div>

<script>
const RAW = {raw_json};
const FLAGS = {{
  DRC:'🇨🇩', Egypt:'🇪🇬', Kenya:'🇰🇪', Madagascar:'🇲🇬', Morocco:'🇲🇦',
  Mozambique:'🇲🇿', Regional:'🌍', 'South Africa':'🇿🇦', Tanzania:'🇹🇿',
  Tunisia:'🇹🇳', Uganda:'🇺🇬', Zambia:'🇿🇲'
}};
const ALL_COUNTRIES = [...new Set(RAW.months.flatMap(m=>Object.keys(RAW.data[m])))].sort();

// State — now arrays for multi-select
let selectedMonths   = new Set();  // empty = all
let selectedCountries = new Set(); // empty = all
let next60Active     = false;

const TODAY = new Date(); TODAY.setHours(0,0,0,0);
const NEXT60 = new Date(TODAY); NEXT60.setDate(TODAY.getDate()+60);

function parseDate(str) {{
  const d = new Date(str); return isNaN(d) ? null : d;
}}

// ── Counts ───────────────────────────────────────────────────────────────────
function countEventsInMonth(m)    {{ return Object.values(RAW.data[m]).reduce((s,e)=>s+e.length,0); }}
function countEventsForCountry(c) {{ return RAW.months.reduce((s,m)=>s+(RAW.data[m][c]?.length||0),0); }}

// ── Badge helpers ─────────────────────────────────────────────────────────────
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
  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}}
function closeModalDirect() {{
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}}
function closeModal(e) {{
  if(e.target === document.getElementById('modalOverlay')) closeModalDirect();
}}
document.addEventListener('keydown', e => {{ if(e.key==='Escape') closeModalDirect(); }});

// ── Filter bar ────────────────────────────────────────────────────────────────
function renderFilterBar() {{
  const bar = document.getElementById('filterBar');
  bar.querySelectorAll('.filter-tag').forEach(t=>t.remove());
  const noF = document.getElementById('noFilters');
  const hasFilters = selectedMonths.size || selectedCountries.size || next60Active;
  if (!hasFilters) {{ noF.style.display='inline'; return; }}
  noF.style.display='none';

  if (next60Active) {{
    const t=document.createElement('div'); t.className='filter-tag';
    t.innerHTML=`📅 Next 60 Days <button onclick="clearNext60()">✕</button>`;
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
function clearNext60()    {{ next60Active=false; document.getElementById('next60Btn').classList.remove('active'); applyFilters(); }}
function clearAllFilters(){{ selectedMonths.clear(); selectedCountries.clear(); next60Active=false; document.getElementById('next60Btn').classList.remove('active'); applyFilters(); }}

// ── Next 60 toggle ────────────────────────────────────────────────────────────
function toggleNext60() {{
  next60Active = !next60Active;
  document.getElementById('next60Btn').classList.toggle('active', next60Active);
  // Only clear month selections (month + next60 conflict); country filters are kept
  if (next60Active) {{ selectedMonths.clear(); }}
  applyFilters();
}}

// ── Get filtered events ───────────────────────────────────────────────────────
function getFilteredEntries() {{
  // Returns array of {{month, country, events[]}}
  const result = [];
  const months = selectedMonths.size ? RAW.months.filter(m=>selectedMonths.has(m)) : RAW.months;

  months.forEach(m => {{
    let entries = Object.entries(RAW.data[m]);
    if (selectedCountries.size) entries = entries.filter(([c])=>selectedCountries.has(c));

    entries.forEach(([country, events]) => {{
      let evts = events;
      if (next60Active) {{
        evts = events.filter(e => {{
          const d = parseDate(e.DateStr);
          return d && d >= TODAY && d <= NEXT60;
        }});
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

// ── Main grid ─────────────────────────────────────────────────────────────────
function renderMain() {{
  const filtered = getFilteredEntries();
  if (!filtered.length) {{
    document.getElementById('main').innerHTML='<div class="empty">No events match your filters.</div>';
    return;
  }}

  // Group by month
  const byMonth = {{}};
  filtered.forEach(item => {{
    const m2=item.month, country2=item.country, events2=item.events;
    if (!byMonth[m2]) byMonth[m2] = {{}};
    byMonth[m2][country2] = events2;
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
              ${{events.map((e,i)=>{{
                const origIdx = RAW.data[m][country].indexOf(e);
                const hasDesc  = e.DatePassed && e.Description;
                const warnDesc = e.DatePassed && !e.Description && e.Status.toLowerCase().includes('executed');
                const rowClass = hasDesc?'event-row has-desc':warnDesc?'event-row warn-desc':'event-row';
                const onclick  = hasDesc?`onclick="openModal('${{country.replace(/'/g,"\\'")}}', RAW.data['${{m}}']['${{country}}'][${{origIdx}}])"` :'';
                const indicator= hasDesc
                  ?`<div class="desc-indicator has">📋 Report</div>`
                  :warnDesc?`<div class="desc-indicator warn">⚠️ Add report</div>`:'';
                return `<div class="${{rowClass}}" ${{onclick}}>
                  <div class="event-date">${{e.DateStr||'—'}}</div>
                  <div class="event-info"><div class="event-name">${{e.Event}}</div></div>
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
  document.getElementById('main').innerHTML=html;
}}

// ── Apply all filters ─────────────────────────────────────────────────────────
function applyFilters() {{
  // Update month button label
  const mCount = selectedMonths.size;
  document.getElementById('monthLabel').textContent = mCount===0?'All Months':mCount===1?[...selectedMonths][0]:`${{mCount}} Months`;
  // Update country button label
  const cCount = selectedCountries.size;
  document.getElementById('countryLabel').textContent = cCount===0?'All Countries':cCount===1?[...selectedCountries][0]:`${{cCount}} Countries`;
  // Sync checkmarks
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
  const total = RAW.months.reduce((s,m)=>s+countEventsInMonth(m),0);
  document.getElementById('monthMenu').innerHTML =
    `<div class="dropdown-action" onclick="selectedMonths.clear();applyFilters()">Clear selection</div>
     <hr class="dropdown-divider"/>`
    + RAW.months.map(m=>`
      <div class="dropdown-item" data-value="${{m}}">
        <div class="check"></div>
        <span style="flex:1">${{m}}</span>
        <span class="event-count">${{countEventsInMonth(m)}}</span>
      </div>`).join('');

  document.querySelectorAll('#monthMenu .dropdown-item[data-value]').forEach(item=>
    item.addEventListener('click', e=>{{
      e.stopPropagation();
      const v = item.dataset.value;
      if (selectedMonths.has(v)) selectedMonths.delete(v);
      else {{ next60Active=false; document.getElementById('next60Btn').classList.remove('active'); selectedMonths.add(v); }}
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
        <div class="check"></div>
        <span style="flex:1">${{FLAGS[c]||'🌐'}} ${{c}}</span>
        <span class="event-count">${{countEventsForCountry(c)}}</span>
      </div>`).join('');

  document.querySelectorAll('#countryMenu .dropdown-item[data-value]').forEach(item=>
    item.addEventListener('click', e=>{{
      e.stopPropagation();
      const v = item.dataset.value;
      if (selectedCountries.has(v)) selectedCountries.delete(v);
      else {{ selectedCountries.add(v); }}
      applyFilters();
    }})
  );
}}

// ── Dropdown open/close ───────────────────────────────────────────────────────
function closeAll() {{
  document.querySelectorAll('.dropdown-btn:not(.next60Btn)').forEach(b=>b.classList.remove('open'));
  document.querySelectorAll('.dropdown-menu').forEach(m=>m.classList.remove('open'));
}}
function toggleDropdown(btnId,menuId) {{
  const isOpen=document.getElementById(menuId).classList.contains('open');
  closeAll();
  if(!isOpen){{document.getElementById(btnId).classList.add('open');document.getElementById(menuId).classList.add('open');}}
}}
document.getElementById('monthBtn').addEventListener('click',   e=>{{e.stopPropagation();toggleDropdown('monthBtn','monthMenu');}});
document.getElementById('countryBtn').addEventListener('click', e=>{{e.stopPropagation();toggleDropdown('countryBtn','countryMenu');}});
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