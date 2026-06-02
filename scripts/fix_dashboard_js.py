# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py
Injects month-scoped Yfirlit KPI override into the dashboard HTML.
When viewing Yfirlit (all) in a month tab, shows only that month's totals.
Safe: appends AFTER existing code, never modifies original functions.
Idempotent: checks for marker before injecting.
"""
import os, sys

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')
MARKER = '/* MONTH_SCOPED_YFIRLIT */'

if not os.path.exists(HTML):
    print('Dashboard not found, skipping')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

if MARKER in src:
    print('Month-scoped Yfirlit already injected, nothing to do')
    sys.exit(0)

# Find the last </script> tag to inject before it
inject_point = src.rfind('</script>')
if inject_point < 0:
    print('No </script> found, skipping')
    sys.exit(1)

JS_PATCH = """
// """ + MARKER + """
(function() {
  var _baseR = render;
  window.render = function monthScopedRender() {
    _baseR();
    if (currentDate !== 'all' || !currentMonth) return;
    var mDates = DATES.filter(function(d) { return d.startsWith(currentMonth); });
    if (!mDates.length) return;
    var total = mDates.reduce(function(s,d){return s+(DAILY[d]||0);},0);
    var avg   = total / mDates.length;
    var net   = total * (1 - COMMISSION);
    var vals  = mDates.map(function(d){return DAILY[d]||0;});
    var bestIdx = vals.indexOf(Math.max.apply(null,vals));
    var kpiGrid = document.getElementById('kpi-grid');
    if (!kpiGrid) return;
    kpiGrid.innerHTML =
      '<div class="kpi t-red">' +
        '<div class="kpi-label">Heildarsala</div>' +
        '<div class="kpi-value">' + fmt(total) + '</div>' +
        '<div class="kpi-sub">' + fmtKr(total) + '</div>' +
        '<div class="kpi-divider"></div>' +
        '<div class="kpi-commission-label">Eftir þóknun Krónan (19,75%)</div>' +
        '<div class="kpi-commission">' + fmt(net) + ' &nbsp;<span style="font-size:12px;font-weight:500;color:var(--muted)">' + fmtKr(net) + '</span></div>' +
      '</div>' +
      '<div class="kpi t-blue"><div class="kpi-label">Fjöldi daga</div><div class="kpi-value">' + mDates.length + '</div><div class="kpi-sub">' + (typeof monthLabel==='function'?monthLabel(currentMonth):currentMonth) + '</div></div>' +
      '<div class="kpi t-green"><div class="kpi-label">Meðaltal á dag</div><div class="kpi-value">' + fmt(avg) + '</div><div class="kpi-sub">' + fmtKr(avg) + '</div></div>' +
      '<div class="kpi t-amber"><div class="kpi-label">Besti dagur</div><div class="kpi-value" style="font-size:18px;padding-top:6px">' + (mDates[bestIdx]?dayLabel(mDates[bestIdx]):'-') + '</div><div class="kpi-sub">' + fmtKr(Math.max.apply(null,vals)) + '</div></div>';
  };
})();
"""

patched = src[:inject_point] + JS_PATCH + src[inject_point:]

with open(HTML, 'w', encoding='utf-8') as f:
    f.write(patched)

print(f'✅ Month-scoped Yfirlit injected ({len(patched):,} bytes)')
