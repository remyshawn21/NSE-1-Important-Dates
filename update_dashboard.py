"""
NSE1 Dashboard Updater — TVS Branded
=====================================
Run this script whenever you update your Excel file.
It will regenerate the TVS-branded dashboard HTML automatically.

Usage:
    python update_dashboard.py

Requirements:
    pip install pandas openpyxl

Configuration (edit the two lines below if filenames change):
"""

EXCEL_FILE  = "NSE1 Important Dates.xlsx"   # ← your Excel file
OUTPUT_FILE = "index.html"      # ← output dashboard

# ─────────────────────────────────────────────────────────────────────────────

import sys, os, json, base64
from datetime import datetime

def check_dependencies():
    missing = []
    try: import pandas
    except ImportError: missing.append("pandas")
    try: import openpyxl
    except ImportError: missing.append("openpyxl")
    if missing:
        print(f"\n❌  Missing packages: {', '.join(missing)}")
        print(f"    Fix it by running:  pip install {' '.join(missing)}\n")
        sys.exit(1)

check_dependencies()
import pandas as pd

# ── TVS Logo (embedded so it works offline) ──────────────────────────────────
# Place your TVS logo as "tvs_logo.jpg" in the same folder as this script.
# If not found, the header will show text only.

def get_logo_b64():
    for name in ["tvs_logo.jpg", "tvs_logo.JPG", "tvs_logo_23.jpg", "tvs_logo.png"]:
        if os.path.exists(name):
            with open(name, "rb") as f:
                ext = name.split(".")[-1].lower()
                mime = "image/png" if ext == "png" else "image/jpeg"
                return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    return None

# ── Data loading ─────────────────────────────────────────────────────────────

def load_data(path):
    if not os.path.exists(path):
        print(f"\n❌  Excel file not found: {path}")
        print("    Make sure the file is in the same folder as this script.\n")
        sys.exit(1)

    df = pd.read_excel(path)
    required = {'Country', 'Date', 'Event', 'Status'}
    if not required.issubset(df.columns):
        print(f"\n❌  Excel file must have columns: {required}")
        print(f"    Found: {set(df.columns)}\n")
        sys.exit(1)

    df['Date']    = pd.to_datetime(df['Date'], errors='coerce')
    df            = df.dropna(subset=['Date'])
    df['Month']   = df['Date'].dt.strftime('%B %Y')
    df['DateStr'] = df['Date'].dt.strftime('%d %b %Y')
    df['Status']  = df['Status'].fillna('').str.strip()
    return df

def build_json(df):
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
            data[month][country] = cdf[['DateStr', 'Event', 'Status']].to_dict('records')
    return {'months': month_order, 'data': data}

# ── HTML generation ───────────────────────────────────────────────────────────

def build_html(payload, updated_at, logo_src):
    raw_json = json.dumps(payload, ensure_ascii=False)

    logo_html = (
        f'<div class="tvs-logo"><img src="{logo_src}" alt="TVS Logo"/></div>'
        if logo_src else
        '<div class="tvs-logo-text">TVS</div>'
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>TVS NSE1 Important Dates</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  :root {{
    --bg:              #f0f3f8;
    --surface:         #ffffff;
    --surface2:        #f5f7fa;
    --border:          #dde2ec;
    --tvs-blue:        #1B3F8B;
    --tvs-blue-dark:   #142f6a;
    --tvs-blue-light:  #e6ecf8;
    --tvs-red:         #CC1313;
    --tvs-red-dark:    #a80f0f;
    --tvs-red-light:   #fceaea;
    --executed:        #1a7a4a;
    --executed-bg:     #e8f5ee;
    --postponed:       #b34000;
    --postponed-bg:    #fff3e8;
    --yet:             #1B3F8B;
    --yet-bg:          #e6ecf8;
    --text:            #1a1f2e;
    --muted:           #6b7280;
  }}

  body {{ background: var(--bg); color: var(--text); font-family: 'DM Sans', sans-serif; min-height: 100vh; }}

  .top-strip {{ height: 4px; background: linear-gradient(90deg, var(--tvs-red) 0%, var(--tvs-blue) 60%); }}

  header {{
    background: var(--tvs-blue); padding: 0 40px;
    display: flex; align-items: center; justify-content: space-between;
    position: sticky; top: 0; z-index: 100;
    gap: 24px; height: 72px;
    box-shadow: 0 3px 16px rgba(27,63,139,0.35); flex-wrap: wrap;
  }}
  .header-left {{ display: flex; align-items: center; gap: 18px; }}
  .tvs-logo {{
    height: 42px; background: white; border-radius: 7px;
    padding: 5px 12px; display: flex; align-items: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
  }}
  .tvs-logo img {{ height: 30px; width: auto; display: block; }}
  .tvs-logo-text {{
    font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 800;
    color: var(--tvs-blue); background: white; padding: 6px 14px; border-radius: 7px;
  }}
  .header-divider {{ width: 1px; height: 34px; background: rgba(255,255,255,0.2); }}
  .header-title h1 {{
    font-family: 'Syne', sans-serif; font-size: 17px; font-weight: 700;
    color: white; letter-spacing: -0.2px;
  }}
  .header-title p {{
    font-size: 11px; color: rgba(255,255,255,0.55); margin-top: 2px;
    text-transform: uppercase; letter-spacing: 0.7px; font-weight: 300;
  }}
  .header-right {{ display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }}
  .updated {{ font-size: 11px; color: rgba(255,255,255,0.5); white-space: nowrap; }}
  .updated strong {{ color: rgba(255,255,255,0.85); }}

  .dropdown-wrap {{ position: relative; flex-shrink: 0; }}
  #monthBtn {{
    background: var(--tvs-red); color: white; border: none;
    padding: 11px 16px; font-family: 'Syne', sans-serif; font-size: 13px;
    font-weight: 700; border-radius: 8px; cursor: pointer;
    display: flex; align-items: center; gap: 10px;
    transition: background 0.15s, transform 0.1s;
    min-width: 190px; justify-content: space-between;
    box-shadow: 0 2px 8px rgba(204,19,19,0.4);
  }}
  #monthBtn:hover {{ background: var(--tvs-red-dark); }}
  #monthBtn:active {{ transform: scale(0.98); }}
  #monthBtn .arrow {{ width: 16px; height: 16px; transition: transform 0.2s; flex-shrink: 0; }}
  #monthBtn.open .arrow {{ transform: rotate(180deg); }}

  .dropdown-menu {{
    position: absolute; top: calc(100% + 8px); left: 0; right: 0;
    background: white; border: 1px solid var(--border); border-radius: 12px;
    overflow: hidden; box-shadow: 0 16px 48px rgba(0,0,0,0.14);
    opacity: 0; transform: translateY(-8px); pointer-events: none;
    transition: opacity 0.18s, transform 0.18s; z-index: 200; min-width: 210px;
  }}
  .dropdown-menu.open {{ opacity: 1; transform: translateY(0); pointer-events: all; }}
  .dropdown-item {{
    padding: 12px 16px; font-size: 13px; font-weight: 500; cursor: pointer;
    display: flex; align-items: center; justify-content: space-between;
    transition: background 0.12s; border-bottom: 1px solid var(--border); color: var(--text);
  }}
  .dropdown-item:last-child {{ border-bottom: none; }}
  .dropdown-item:hover {{ background: var(--tvs-blue-light); color: var(--tvs-blue); }}
  .dropdown-item.active {{ background: var(--tvs-blue-light); color: var(--tvs-blue); font-weight: 700; }}
  .event-count {{
    font-size: 11px; background: var(--surface2); color: var(--muted);
    padding: 2px 8px; border-radius: 20px; font-weight: 400;
  }}
  .dropdown-item.active .event-count {{ background: var(--tvs-blue); color: white; }}

  .stats-bar {{
    display: flex; gap: 10px; padding: 14px 40px;
    background: white; border-bottom: 1px solid var(--border);
    overflow-x: auto; flex-wrap: wrap; box-shadow: 0 1px 4px rgba(0,0,0,0.05);
  }}
  .stat-pill {{
    display: flex; align-items: center; gap: 7px; padding: 6px 14px;
    border-radius: 20px; font-size: 12px; font-weight: 500; white-space: nowrap;
    border: 1px solid transparent;
  }}
  .stat-pill .dot {{ width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }}
  .stat-pill.total    {{ background: var(--tvs-blue-light); color: var(--tvs-blue);  border-color: #c5d3ef; }}
  .stat-pill.total .dot {{ background: var(--tvs-blue); }}
  .stat-pill.executed {{ background: var(--executed-bg); color: var(--executed);    border-color: #b8dfc9; }}
  .stat-pill.executed .dot  {{ background: var(--executed); }}
  .stat-pill.postponed {{ background: var(--postponed-bg); color: var(--postponed); border-color: #f5d5b8; }}
  .stat-pill.postponed .dot {{ background: var(--postponed); }}
  .stat-pill.yet      {{ background: var(--yet-bg); color: var(--yet);              border-color: #c5d3ef; }}
  .stat-pill.yet .dot {{ background: var(--yet); }}

  main {{ padding: 28px 40px 40px; }}
  .month-title {{
    font-family: 'Syne', sans-serif; font-size: 28px; font-weight: 800;
    letter-spacing: -0.8px; margin-bottom: 22px;
    display: flex; align-items: baseline; gap: 10px;
    color: var(--tvs-blue); animation: fadeUp 0.3s ease;
    border-left: 4px solid var(--tvs-red); padding-left: 14px;
  }}
  .month-title .year {{ color: var(--muted); font-size: 18px; font-weight: 400; }}

  @keyframes fadeUp {{
    from {{ opacity: 0; transform: translateY(10px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
  }}

  .country-grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
    gap: 14px; animation: fadeUp 0.35s ease;
  }}
  .country-card {{
    background: white; border: 1px solid var(--border); border-radius: 12px;
    overflow: hidden; transition: border-color 0.2s, box-shadow 0.2s, transform 0.2s;
  }}
  .country-card:hover {{
    border-color: var(--tvs-blue);
    box-shadow: 0 6px 24px rgba(27,63,139,0.13);
    transform: translateY(-2px);
  }}
  .country-header {{
    padding: 13px 16px; display: flex; align-items: center; gap: 10px;
    background: var(--tvs-blue); border-bottom: 3px solid var(--tvs-red);
  }}
  .country-flag {{
    width: 30px; height: 30px; border-radius: 6px;
    background: rgba(255,255,255,0.12);
    display: flex; align-items: center; justify-content: center;
    font-size: 17px; flex-shrink: 0;
  }}
  .country-name {{ font-family: 'Syne', sans-serif; font-size: 14px; font-weight: 700; flex: 1; color: white; }}
  .country-event-count {{
    font-size: 11px; color: rgba(255,255,255,0.65);
    background: rgba(255,255,255,0.12); padding: 2px 9px; border-radius: 20px;
  }}
  .event-list {{ padding: 6px 0; }}
  .event-row {{
    display: flex; align-items: flex-start; gap: 10px; padding: 9px 16px;
    border-bottom: 1px solid rgba(0,0,0,0.04); transition: background 0.1s;
  }}
  .event-row:last-child {{ border-bottom: none; }}
  .event-row:hover {{ background: var(--surface2); }}
  .event-date {{
    font-size: 11px; font-weight: 600; color: var(--muted);
    white-space: nowrap; min-width: 78px; padding-top: 2px;
    font-variant-numeric: tabular-nums;
  }}
  .event-info {{ flex: 1; min-width: 0; }}
  .event-name {{ font-size: 12.5px; line-height: 1.4; color: var(--text); word-break: break-word; }}
  .status-badge {{
    font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 20px;
    white-space: nowrap; flex-shrink: 0; text-transform: uppercase;
    letter-spacing: 0.4px; margin-top: 2px;
  }}
  .badge-executed  {{ background: var(--executed-bg);  color: var(--executed); }}
  .badge-postponed {{ background: var(--postponed-bg); color: var(--postponed); }}
  .badge-yet       {{ background: var(--yet-bg);       color: var(--yet); }}
  .badge-unknown   {{ background: var(--surface2);      color: var(--muted); }}
  .empty {{ text-align: center; padding: 80px 20px; color: var(--muted); }}

  .legend {{
    display: flex; gap: 20px; flex-wrap: wrap;
    padding: 0 40px 32px; font-size: 12px; color: var(--muted); align-items: center;
  }}
  .legend span {{ font-weight: 600; color: var(--tvs-blue); margin-right: 4px; }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .legend-dot {{ width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }}
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
    <div class="dropdown-wrap">
      <button id="monthBtn">
        <span id="btnLabel">All Months</span>
        <svg class="arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      <div class="dropdown-menu" id="dropdownMenu"></div>
    </div>
  </div>
</header>

<div class="stats-bar" id="statsBar"></div>

<main>
  <div class="month-title" id="monthTitle"></div>
  <div class="country-grid" id="countryGrid"></div>
</main>

<div class="legend">
  <span>Legend:</span>
  <div class="legend-item"><div class="legend-dot" style="background:var(--executed)"></div> Executed</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--postponed)"></div> Postponed</div>
  <div class="legend-item"><div class="legend-dot" style="background:var(--yet)"></div> Upcoming</div>
</div>

<script>
const RAW = {raw_json};
const FLAGS = {{
  DRC:'🇨🇩', Egypt:'🇪🇬', Kenya:'🇰🇪', Madagascar:'🇲🇬', Morocco:'🇲🇦',
  Mozambique:'🇲🇿', Regional:'🌍', 'South Africa':'🇿🇦', Tanzania:'🇹🇿',
  Tunisia:'🇹🇳', Uganda:'🇺🇬', Zambia:'🇿🇲'
}};
let currentMonth = null;

function countEvents(m) {{ return Object.values(RAW.data[m]).reduce((s,e)=>s+e.length,0); }}
function getStats(month) {{
  const c={{executed:0,postponed:0,yet:0,unknown:0}};
  (month?[month]:RAW.months).forEach(m=>Object.values(RAW.data[m]).forEach(evts=>evts.forEach(e=>{{
    const s=(e.Status||'').toLowerCase();
    if(s.includes('executed'))c.executed++;
    else if(s.includes('postponed'))c.postponed++;
    else if(s.includes('yet'))c.yet++;
    else c.unknown++;
  }})));
  return c;
}}
function badgeClass(s) {{
  if(!s)return 'badge-unknown';
  const l=s.toLowerCase();
  if(l.includes('executed'))return 'badge-executed';
  if(l.includes('postponed'))return 'badge-postponed';
  if(l.includes('yet'))return 'badge-yet';
  return 'badge-unknown';
}}
function badgeLabel(s) {{
  if(!s)return '—';
  if(s.toLowerCase().includes('executed'))return 'Executed';
  if(s.toLowerCase().includes('postponed'))return 'Postponed';
  if(s.toLowerCase().includes('yet'))return 'Upcoming';
  return s;
}}
function renderStats(month) {{
  const s=getStats(month),total=s.executed+s.postponed+s.yet+s.unknown;
  document.getElementById('statsBar').innerHTML=`
    <div class="stat-pill total"><div class="dot"></div>${{total}} Total Events</div>
    <div class="stat-pill executed"><div class="dot"></div>${{s.executed}} Executed</div>
    <div class="stat-pill postponed"><div class="dot"></div>${{s.postponed}} Postponed</div>
    <div class="stat-pill yet"><div class="dot"></div>${{s.yet}} Upcoming</div>`;
}}
function renderGrid(month) {{
  const months=month?[month]:RAW.months;
  const [mon,yr]=month?month.split(' '):['All','Months'];
  document.getElementById('monthTitle').innerHTML=month
    ?`${{mon}} <span class="year">${{yr}}</span>`
    :`All <span class="year">Months · 2026</span>`;
  const seen=new Map();
  months.forEach(m=>Object.entries(RAW.data[m]).forEach(([country,events])=>{{
    if(!seen.has(country))seen.set(country,[]);
    seen.get(country).push(...events);
  }}));
  if(!seen.size){{document.getElementById('countryGrid').innerHTML='<div class="empty"><p>No events found.</p></div>';return;}}
  document.getElementById('countryGrid').innerHTML=[...seen.entries()].map(([country,events])=>`
    <div class="country-card">
      <div class="country-header">
        <div class="country-flag">${{FLAGS[country]||'🌐'}}</div>
        <div class="country-name">${{country}}</div>
        <div class="country-event-count">${{events.length}} event${{events.length!==1?'s':''}}</div>
      </div>
      <div class="event-list">${{events.map(e=>`
        <div class="event-row">
          <div class="event-date">${{e.DateStr||'—'}}</div>
          <div class="event-info"><div class="event-name">${{e.Event}}</div></div>
          <div class="status-badge ${{badgeClass(e.Status)}}">${{badgeLabel(e.Status)}}</div>
        </div>`).join('')}}</div>
    </div>`).join('');
}}
function buildDropdown() {{
  const menu=document.getElementById('dropdownMenu');
  const total=RAW.months.reduce((s,m)=>s+countEvents(m),0);
  menu.innerHTML=`<div class="dropdown-item active" data-month="">All Months <span class="event-count">${{total}} events</span></div>`
    +RAW.months.map(m=>`<div class="dropdown-item" data-month="${{m}}">${{m}} <span class="event-count">${{countEvents(m)}} events</span></div>`).join('');
  menu.querySelectorAll('.dropdown-item').forEach(item=>
    item.addEventListener('click',()=>{{selectMonth(item.dataset.month||null);closeDropdown();}}));
}}
function selectMonth(month) {{
  currentMonth=month;
  document.getElementById('btnLabel').textContent=month||'All Months';
  document.querySelectorAll('.dropdown-item').forEach(i=>i.classList.toggle('active',(i.dataset.month||null)===month));
  renderStats(month);renderGrid(month);
}}
const btn=document.getElementById('monthBtn'),menu=document.getElementById('dropdownMenu');
function openDropdown(){{btn.classList.add('open');menu.classList.add('open');}}
function closeDropdown(){{btn.classList.remove('open');menu.classList.remove('open');}}
btn.addEventListener('click',e=>{{e.stopPropagation();menu.classList.contains('open')?closeDropdown():openDropdown();}});
document.addEventListener('click',e=>{{if(!menu.contains(e.target))closeDropdown();}});
buildDropdown();
selectMonth(null);
</script>
</body>
</html>"""

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("\n── TVS NSE1 Dashboard Updater ──────────────────────")
    print(f"  Reading:  {EXCEL_FILE}")

    df      = load_data(EXCEL_FILE)
    payload = build_json(df)
    logo    = get_logo_b64()
    updated = datetime.now().strftime("%d %b %Y, %H:%M")
    total   = sum(len(e) for m in payload['data'].values() for e in m.values())
    months  = len(payload['months'])

    html = build_html(payload, updated, logo)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  Logo:     {'✅ Embedded' if logo else '⚠️  Not found — add tvs_logo.jpg to the folder'}")
    print(f"  Months:   {months}  |  Events: {total}")
    print(f"  Saved:    {OUTPUT_FILE}")
    print(f"  Updated:  {updated}")
    print("────────────────────────────────────────────────────")
    print("  ✅  Dashboard ready! Share the HTML with your team.")
    print("────────────────────────────────────────────────────\n")

if __name__ == "__main__":
    main()
