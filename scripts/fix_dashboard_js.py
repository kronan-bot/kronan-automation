# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py
Injects month-scoped Yfirlit overrides into the dashboard HTML.
1. KPI override  — shows only current-month totals in Yfirlit
2. Heatmap override — buildHeatmap always filtered to current month tab
Safe: appends AFTER existing code, never modifies original functions.
Idempotent: checks for marker before injecting.
"""
import os, sys

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')
MARKER = '/* MONTH_SCOPED_YFIRLIT_V3 */'
OLD_MARKER = '/* MONTH_SCOPED_YFIRLIT */'

if not os.path.exists(HTML):
    print('Dashboard not found, skipping')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

if MARKER in src:
    print('Month-scoped Yfirlit V2 already injected, nothing to do')
    sys.exit(0)

# Remove old V1 injection if present
if OLD_MARKER in src:
    # Find and remove the old IIFE block
    old_start = src.rfind('\n// /* MONTH_SCOPED_YFIRLIT */')
    old_end = src.find('\n// /* MONTH_SCOPED_YFIRLIT */', old_start)
    if old_start >= 0 and old_end < 0:
        # V1 is the last injection — remove from its start to just before </script>
        close_tag = src.rfind('</script>')
        src = src[:old_start] + src[close_tag:]
        print('Removed old V1 injection')

inject_point = src.rfind('</script>')
if inject_point < 0:
    print('No </script> found, skipping')
    sys.exit(1)

JS_PATCH = """
// """ + MARKER + """
(function() {

  // ── 1. KPI override: month-scoped totals in Yfirlit ──────────────────────
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

  // ── 2. Heatmap override: always filtered to current month tab ─────────────
  //    When a month tab is selected, the heatmap shows ONLY that month's dates
  //    regardless of whether Yfirlit or a specific day is active.
  var _baseHeat = buildHeatmap;
  window.buildHeatmap = function monthScopedHeatmap() {
    if (!currentMonth) { _baseHeat(); return; }
    var mDates = DATES.filter(function(d) { return d.startsWith(currentMonth); });
    if (!mDates.length) { _baseHeat(); return; }

    // Top-12 stores ranked by month sales
    var mTotals = {};
    mDates.forEach(function(d) {
      (STORES_BY_DATE[d]||[]).forEach(function(r) {
        mTotals[r.store] = (mTotals[r.store]||0) + r.sale;
      });
    });
    var storeNames = Object.keys(mTotals)
      .sort(function(a,b){ return mTotals[b]-mTotals[a]; })
      .slice(0,12);

    var dayLabels = mDates.map(function(d) {
      var p = d.split('-').map(Number);
      return DOW[new Date(p[0],p[1]-1,p[2]).getDay()] + '<br>' + p[2] + '.' + p[1];
    });
    var matrix = storeNames.map(function(store) {
      return mDates.map(function(d) {
        var row = (STORES_BY_DATE[d]||[]).find(function(r){return r.store===store;});
        return row ? row.sale : 0;
      });
    });
    var storeMax = matrix.map(function(row){ return Math.max.apply(null,row.concat([0])); });
    var storeMin = matrix.map(function(row){ return Math.min.apply(null,row.concat([0])); });

    function heatColor(v,mn,mx) {
      var t = mx===mn ? 0.5 : (v-mn)/(mx-mn);
      return 'rgb('+Math.round(240-t*218)+','+Math.round(253-t*152)+','+Math.round(244-t*192)+')';
    }
    function textColor(v,mn,mx) {
      return ((mx===mn?0.5:(v-mn)/(mx-mn))>0.5)?'#fff':'#374151';
    }

    // Highlight selected day column
    var selIdx = (currentDate !== 'all') ? mDates.indexOf(currentDate) : -1;

    var html = '<table class="heatmap"><thead><tr><th class="store-col">Verslun</th>';
    dayLabels.forEach(function(l,i){
      var style = (i===selIdx) ? ' style="background:#1d4ed8;color:#fff;border-radius:6px"' : '';
      html += '<th'+style+'>'+l+'</th>';
    });
    html += '<th>Heild</th><th style="color:var(--green)">Eftir þókn.</th></tr></thead><tbody>';

    storeNames.forEach(function(name,si) {
      var total = matrix[si].reduce(function(s,v){return s+v;},0);
      var net = total*(1-COMMISSION);
      html += '<tr><td class="store-name">'+name.replace('Krónan ','')+' </td>';
      matrix[si].forEach(function(val,di) {
        var bg = heatColor(val,storeMin[si],storeMax[si]);
        var tc = textColor(val,storeMin[si],storeMax[si]);
        var border = (di===selIdx) ? ';outline:2px solid #1d4ed8' : '';
        html += '<td><div class="hm-cell" style="background:'+bg+';color:'+tc+border+'" title="'+name+' \xb7 '+dayLabels[di].replace('<br>',' ')+' \xb7 '+fmtKr(val)+'">'+fmt(val)+'</div></td>';
      });
      html += '<td style="font-weight:700;color:var(--text)">'+fmt(total)+'</td>';
      html += '<td style="font-weight:700;color:var(--green)">'+fmt(net)+'</td></tr>';
    });

    html += '</tbody></table>';
    var wrap = document.getElementById('heatmap-wrap');
    if (wrap) wrap.innerHTML = html;
  };

})();
"""

patched = src[:inject_point] + JS_PATCH + src[inject_point:]

with open(HTML, 'w', encoding='utf-8') as f:
    f.write(patched)

print(f'✅ Month-scoped Yfirlit + Heatmap (always month-scoped) injected ({len(patched):,} bytes)')
