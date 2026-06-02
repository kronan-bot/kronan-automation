# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py — patches month-scoped Yfirlit into the dashboard HTML.
Run after bootstrap, before update_dashboard_data.py and deploy.
"""
import os, re, sys

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')

if not os.path.exists(HTML):
    print('Dashboard not found, skipping JS fix')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

changed = False

# ── 1. Add monthLabel helper (idempotent) ────────────────────────────────────
if 'function monthLabel' not in src:
    helper = """
function monthLabel(m) {
  if (!m) return 'Allt';
  const parts = m.split('-');
  const islic = ['Jan','Feb','Mar','Apr','Maí','Jún','Júl','Ágú','Sep','Okt','Nóv','Des'];
  return (islic[parseInt(parts[1])-1] || parts[1]) + ' ' + parts[0];
}"""
    src = src.replace('const COMMISSION', helper + '\n\nconst COMMISSION', 1)
    changed = True
    print('Added monthLabel()')

# ── 2. Fix getStoresForDate for month-scoped all ─────────────────────────────
old_gsfd = "  if (d === 'all') return STORE_TOTALS;"
new_gsfd = """  if (d === 'all') {
    if (!currentMonth) return STORE_TOTALS;
    const mDates = DATES.filter(x => x.startsWith(currentMonth));
    const agg = {};
    mDates.forEach(date => {
      (STORES_BY_DATE[date]||[]).forEach(r => {
        if (!agg[r.store]) agg[r.store] = {sale:0,qty:0};
        agg[r.store].sale += r.sale; agg[r.store].qty +=r.qty;
      });
    });
    return Object.entries(agg).map(([s,v])=>[s,{sale:Math.round(v.sale),qty:v.qty}]).sort((a,b)=>b[1].sale-a[1].sale);
  }"""
if old_gsfd in src and new_gsfd not in src:
    src = src.replace(old_gsfd, new_gsfd, 1)
    changed = True
    print('Fixed getStoresForDate()')

# ── 3. Fix render() to use activeDates ───────────────────────────────────────
old_render = ("  const isAll = currentDate === 'all';\n"
              "  const totalSale = isAll ? Object.values(DAILY).reduce((s,v)=>s+v,0) : DAILY[currentDate];\n"
              "  const totalQty  = isAll ? Object.values(DAILY_QTY).reduce((s,v)=>s+v,0) : DAILY_QTY[currentDate];\n"
              "  const dayAvg    = Object.values(DAILY).reduce((s,v)=>s+v,0) / DATES.length;")
new_render = ("  const isAll = currentDate === 'all';\n"
              "  const activeDates = currentMonth ? DATES.filter(d => d.startsWith(currentMonth)) : DATES;\n"
              "  const totalSale = isAll ? activeDates.reduce((s,d)=>s+(DAILY[d]||0),0) : DAILY[currentDate];\n"
              "  const totalQty  = isAll ? activeDates.reduce((s,d)=>s+(DAILY_QTY[d]||0),0) : DAILY_QTY[currentDate];\n"
              "  const dayAvg    = activeDates.reduce((s,d)=>s+(DAILY[d]||0),0) / activeDates.length;")
if old_render in src and new_render not in src:
    src = src.replace(old_render, new_render, 1)
    changed = True
    print('Fixed render() aggregation')

# ── 4. Fix Fjöldi daga KPI ────────────────────────────────────────────────────
src2 = re.sub(
    r'<div class="kpi t-blue"><div class="kpi-label">Fj.ldi daga</div><div class="kpi-value">\$\{DATES\.length\}</div><div class="kpi-sub">[^<]*</div></div>',
    '<div class="kpi t-blue"><div class="kpi-label">Fjöldi daga</div><div class="kpi-value">${activeDates.length}</div><div class="kpi-sub">${monthLabel(currentMonth)}</div></div>',
    src, count=1
)
if src2 != src:
    src = src2
    changed = True
    print('Fixed Fjöldi daga KPI')

# ── 5. Fix main chart title ───────────────────────────────────────────────────
src2 = re.sub(
    r"isAll \? 'Dagleg sala [^']*' :",
    "isAll ? `Dagleg sala — ${monthLabel(currentMonth)}` :",
    src, count=1
)
if src2 != src:
    src = src2
    changed = True
    print('Fixed chart title')

# ── 6. Fix isAll chart to use activeDates ────────────────────────────────────
for old, new in [
    ("    const vals = DATES.map(d => DAILY[d]);",
     "    const vals = activeDates.map(d => DAILY[d]);"),
    ("labels:DATES.map(dayLabel)",
     "labels:activeDates.map(dayLabel)"),
]:
    if old in src and new not in src:
        src = src.replace(old, new, 1)
        changed = True

# ── 7. Fix renderStoreProducts ────────────────────────────────────────────────
old_sp = "  const dates = currentDate === 'all' ? DATES : [currentDate];"
new_sp = ("  const activeDates = currentMonth ? DATES.filter(d => d.startsWith(currentMonth)) : DATES;\n"
          "  const dates = currentDate === 'all' ? activeDates : [currentDate];")
if old_sp in src and new_sp not in src:
    src = src.replace(old_sp, new_sp, 1)
    changed = True
    print('Fixed renderStoreProducts()')

# ── 8. Fix store-products label ──────────────────────────────────────────────
old_lbl = "currentDate === 'all' ? '— Maí 2026'"
new_lbl = "currentDate === 'all' ? '— ' + monthLabel(currentMonth)"
if old_lbl in src and new_lbl not in src:
    src = src.replace(old_lbl, new_lbl)
    changed = True

if changed:
    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(src)
    print(f'✅ Dashboard JS patched ({len(src):,} bytes)')
else:
    print('Dashboard JS already patched, nothing to do')
