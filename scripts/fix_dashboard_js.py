# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py — REVERT: removes broken month-scoped patch, restores working JS.
"""
import os, re, sys

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')

if not os.path.exists(HTML):
    print('Dashboard not found, skipping')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

changed = False

# ── 1. Remove monthLabel helper ──────────────────────────────────────────────
if 'function monthLabel' in src:
    src = re.sub(
        r'\nfunction monthLabel\(m\) \{[^}]+\}\n\n',
        '\n',
        src, count=1
    )
    changed = True
    print('Removed monthLabel()')

# ── 2. Restore getStoresForDate('all') ───────────────────────────────────────
old = """  if (d === 'all') {
    if (!currentMonth) return STORE_TOTALS;
    const mDates = DATES.filter(x => x.startsWith(currentMonth));
    const agg = {};
    mDates.forEach(date => {
      (STORES_BY_DATE[date]||[]).forEach(r => {
        if (!agg[r.store]) agg[r.store] = {sale:0,qty:0};
        agg[r.store].sale += r.sale; agg[r.store].qty += r.qty;
      });
    });
    return Object.entries(agg).map(([s,v])=>[s,{sale:Math.round(v.sale),qty:v.qty}]).sort((a,b)=>b[1].sale-a[1].sale);
  }"""
new = "  if (d === 'all') return STORE_TOTALS;"
if old in src:
    src = src.replace(old, new, 1)
    changed = True
    print('Restored getStoresForDate()')

# ── 3. Restore render() to use DATES ─────────────────────────────────────────
old_r = ("  const isAll = currentDate === 'all';\n"
         "  const activeDates = currentMonth ? DATES.filter(d => d.startsWith(currentMonth)) : DATES;\n"
         "  const totalSale = isAll ? activeDates.reduce((s,d)=>s+(DAILY[d]||0),0) : DAILY[currentDate];\n"
         "  const totalQty  = isAll ? activeDates.reduce((s,d)=>s+(DAILY_QTY[d]||0),0) : DAILY_QTY[currentDate];\n"
         "  const dayAvg    = activeDates.reduce((s,d)=>s+(DAILY[d]||0),0) / activeDates.length;")
new_r = ("  const isAll = currentDate === 'all';\n"
         "  const totalSale = isAll ? Object.values(DAILY).reduce((s,v)=>s+v,0) : DAILY[currentDate];\n"
         "  const totalQty  = isAll ? Object.values(DAILY_QTY).reduce((s,v)=>s+v,0) : DAILY_QTY[currentDate];\n"
         "  const dayAvg    = Object.values(DAILY).reduce((s,v)=>s+v,0) / DATES.length;")
if old_r in src:
    src = src.replace(old_r, new_r, 1)
    changed = True
    print('Restored render() to use DATES')

# ── 4. Restore chart data to use DATES ───────────────────────────────────────
for old, new in [
    ("    const vals = activeDates.map(d => DAILY[d]);",
     "    const vals = DATES.map(d => DAILY[d]);"),
    ("labels:activeDates.map(dayLabel)",
     "labels:DATES.map(dayLabel)"),
]:
    if old in src:
        src = src.replace(old, new, 1)
        changed = True

# ── 5. Restore renderStoreProducts ───────────────────────────────────────────
old_sp = ("  const activeDates = currentMonth ? DATES.filter(d => d.startsWith(currentMonth)) : DATES;\n"
          "  const dates = currentDate === 'all' ? activeDates : [currentDate];")
new_sp = "  const dates = currentDate === 'all' ? DATES : [currentDate];"
if old_sp in src:
    src = src.replace(old_sp, new_sp, 1)
    changed = True
    print('Restored renderStoreProducts()')

# ── 6. Fix any remaining activeDates references in other functions ────────────
if 'activeDates' in src:
    # Replace any remaining activeDates with DATES as fallback
    src = src.replace('activeDates.length', 'DATES.length')
    src = src.replace('activeDates.map', 'DATES.map')
    src = src.replace('activeDates.reduce', 'DATES.reduce')
    changed = True
    print('Cleaned up remaining activeDates references')

if changed:
    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(src)
    print(f'✅ Reverted broken JS patch ({len(src):,} bytes)')
else:
    print('Nothing to revert — HTML looks clean')
