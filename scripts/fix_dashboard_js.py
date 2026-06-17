# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py
Injects month-scoped overrides into the dashboard HTML.
1. KPI override            — shows only current-month totals in Yfirlit
2. Heatmap override        — buildHeatmap always filtered to current month tab
3. Store-products override — "Vörur eftir verslunum" scoped to selected month/day
4. Store-sales override    — "Vara sala eftir verslun" scoped to selected month
5. Date dropdown rebuild   — only shows dates of the selected month
6. Store breakdown fix     — "Verslanir — sundurliðun" aggregates by month
7. Month tab default       — clicking month tab defaults to month-total view
Safe: appends AFTER existing code, never modifies original functions.
Idempotent: checks for marker before injecting.
"""
import os, sys, re

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')
MARKER = '/* MONTH_SCOPED_YFIRLIT_V6 */'
OLD_MARKERS = [
    '/* MONTH_SCOPED_YFIRLIT */',
    '/* MONTH_SCOPED_YFIRLIT_V3 */',
    '/* MONTH_SCOPED_YFIRLIT_V4 */',
    '/* MONTH_SCOPED_YFIRLIT_V5 */',
]

if not os.path.exists(HTML):
    print('Dashboard not found, skipping')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

if MARKER in src:
    print('Month-scoped Yfirlit V6 already injected, nothing to do')
    sys.exit(0)

# Remove any older injection (each was appended as the last block before </script>)
for om in OLD_MARKERS:
    if om in src:
        old_start = src.rfind('\n// ' + om)
        if old_start >= 0:
            close_tag = src.rfind('</script>')
            src = src[:old_start] + src[close_tag:]
            print('Removed old injection ' + om)

inject_point = src.rfind('</script>')
if inject_point < 0:
    print('No </script> found, skipping')
    sys.exit(1)

JS_PATCH = """
// """ + MARKER + """
(function() {

  function monthDates() {
    if (!currentMonth) return DATES;
    var out = [];
    for (var i = 0; i < DATES.length; i++) {
      if (DATES[i].indexOf(currentMonth) === 0) out.push(DATES[i]);
    }
    return out.length ? out : DATES;
  }

  var MONTH_FULL_IS = ['','Janúar','Febrúar','Mars','Apríl','Maí','Júní','Júlí','Ágúst','September','Október','Nóvember','Desember'];
  function monthTitle(m) {
    var p = m.split('-');
    return MONTH_FULL_IS[parseInt(p[1],10)] + ' ' + p[0];
  }

  // ── 0. getStoresForDate: month-aggregated when currentDate==='all' & currentMonth ──
  var _baseGetStores = getStoresForDate;
  window.getStoresForDate = function(d) {
    if (d === 'all' && currentMonth) {
      var mDates = DATES.filter(function(date) { return date.indexOf(currentMonth) === 0; });
      var agg = {};
      mDates.forEach(function(date) {
        (STORES_BY_DATE[date] || []).forEach(function(r) {
          if (!agg[r.store]) agg[r.store] = {sale: 0, qty: 0};
          agg[r.store].sale += r.sale;
          agg[r.store].qty  += r.qty;
        });
      });
      return Object.keys(agg).map(function(name) {
        return [name, agg[name]];
      }).sort(function(a, b) { return b[1].sale - a[1].sale; });
    }
    return _baseGetStores(d);
  };

  // ── 0b. updateDayTabs: after rebuilding day tabs, default to month-total view ──
  var _baseUpdateDayTabs = updateDayTabs;
  window.updateDayTabs = function(m) {
    _baseUpdateDayTabs(m);
    // Default selection after tab rebuild is 'all' so month-total view shows immediately
    currentDate = 'all';
    // Activate the 'all' tab button visually
    var allBtn = document.querySelector('#date-tabs .tab-all');
    if (allBtn) {
      document.querySelectorAll('.tab,.tab-all').forEach(function(b){b.classList.remove('active');});
      allBtn.classList.add('active');
    }
  };

  // ── 1. KPI override + month-scoped date dropdown ──────────────────────────
  var _baseR = render;
  window.render = function monthScopedRender() {
    rebuildStoreSaleDropdown();
    _baseR();
    // Fix hardcoded chart title
    var chartTitle = document.getElementById('main-chart-title');
    if (chartTitle && currentDate === 'all') {
      if (currentMonth) {
        chartTitle.textContent = 'Dagleg sala — ' + monthTitle(currentMonth);
      } else {
        chartTitle.textContent = 'Dagleg sala — Allt tímabilið';
      }
    }
    if (currentDate !== 'all' || !currentMonth) return;
    var mDates = monthDates();
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
      '<div class="kpi t-blue"><div class="kpi-label">Fjöldi daga</div><div class="kpi-value">' + mDates.length + '</div><div class="kpi-sub">' + monthTitle(currentMonth) + '</div></div>' +
      '<div class="kpi t-green"><div class="kpi-label">Meðaltal á dag</div><div class="kpi-value">' + fmt(avg) + '</div><div class="kpi-sub">' + fmtKr(avg) + '</div></div>' +
      '<div class="kpi t-amber"><div class="kpi-label">Besti dagur</div><div class="kpi-value" style="font-size:18px;padding-top:6px">' + (mDates[bestIdx]?dayLabel(mDates[bestIdx]):'-') + '</div><div class="kpi-sub">' + fmtKr(Math.max.apply(null,vals)) + '</div></div>';
  };

  // Date dropdown of "Vara sala eftir verslun": only current month's dates
  function rebuildStoreSaleDropdown() {
    var sel = document.getElementById('store-sale-date');
    if (!sel) return;
    var mDates = monthDates();
    var prev = sel.value;
    while (sel.options.length > 1) sel.remove(1);
    mDates.forEach(function(d) {
      var p = d.split('-').map(Number);
      var opt = document.createElement('option');
      opt.value = d;
      opt.textContent = DOW[new Date(p[0],p[1]-1,p[2]).getDay()] + ' ' + p[2] + '.' + p[1];
      sel.appendChild(opt);
    });
    sel.value = (mDates.indexOf(prev) >= 0) ? prev : 'all';
  }

  // ── 2. Heatmap override: always filtered to current month tab ─────────────
  var _baseHeat = buildHeatmap;
  window.buildHeatmap = function monthScopedHeatmap() {
    if (!currentMonth) { _baseHeat(); return; }
    var mDates = monthDates();
    if (!mDates.length) { _baseHeat(); return; }

    // Top stores ranked by month sales
    var mTotals = {};
    mDates.forEach(function(d) {
      (STORES_BY_DATE[d]||[]).forEach(function(r) {
        mTotals[r.store] = (mTotals[r.store]||0) + r.sale;
      });
    });
    var storeNames = Object.keys(mTotals)
      .sort(function(a,b){ return mTotals[b]-mTotals[a]; });

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
      var r,g,b;
      if(t<=0.5){var s=t*2;r=Math.round(239+s*(234-239));g=Math.round(68+s*(179-68));b=Math.round(68+s*(8-68));}
      else{var s=(t-0.5)*2;r=Math.round(234-s*(234-22));g=Math.round(179-s*(179-163));b=Math.round(8+s*(74-8));}
      return 'rgb('+r+','+g+','+b+')';
    }
    function textColor(v,mn,mx) {
      var t = mx===mn?0.5:(v-mn)/(mx-mn);
      return (t>0.25&&t<0.75)?'#1a1a1a':'#fff';
    }

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
        html += '<td><div class="hm-cell" style="background:'+bg+';color:'+tc+border+'" title="'+name+' \\xb7 '+dayLabels[di].replace('<br>',' ')+' \\xb7 '+fmtKr(val)+'">'+fmt(val)+'</div></td>';
      });
      html += '<td style="font-weight:700;color:var(--text)">'+fmt(total)+'</td>';
      html += '<td style="font-weight:700;color:var(--green)">'+fmt(net)+'</td></tr>';
    });

    html += '</tbody></table>';
    var wrap = document.getElementById('heatmap-wrap');
    if (wrap) wrap.innerHTML = html;
  };

  // ── 3. "Vörur eftir verslunum": scoped to selected month (or day) ─────────
  window.renderStoreProducts = function monthScopedStoreProducts() {
    var wrap = document.getElementById('store-prods-wrap');
    if (!wrap) return;

    var dates = (currentDate !== 'all') ? [currentDate] : monthDates();

    var storeMap = {};
    dates.forEach(function(d) {
      var day = STORE_PRODS[d] || {};
      Object.keys(day).forEach(function(store) {
        if (!storeMap[store]) storeMap[store] = {};
        day[store].forEach(function(p) {
          var pnr = p[0], name = p[1], sale = p[2], qty = p[3];
          if (!storeMap[store][name]) storeMap[store][name] = [0, 0, pnr];
          storeMap[store][name][0] += sale;
          storeMap[store][name][1] += qty;
        });
      });
    });

    var lbl = document.getElementById('sp-date-label');
    if (lbl) {
      if (currentDate !== 'all') lbl.textContent = '— ' + dayLabel(currentDate);
      else if (currentMonth)     lbl.textContent = '— ' + monthTitle(currentMonth);
      else                       lbl.textContent = '— Allt tímabilið';
    }

    var stores = Object.keys(storeMap).map(function(s){ return [s, storeMap[s]]; })
      .sort(function(a, b) {
        var sa = 0, sb = 0, k;
        for (k in a[1]) sa += a[1][k][0];
        for (k in b[1]) sb += b[1][k][0];
        return sb - sa;
      });

    if (!stores.length) {
      wrap.innerHTML = '<div class="sp-empty">Engar vörugögn fyrir þennan dag</div>';
      return;
    }

    var html = '';
    stores.forEach(function(entry) {
      var store = entry[0], prods = entry[1];
      var shortName = store.replace('Krónan ', '');
      var totalQty = 0, k;
      for (k in prods) totalQty += prods[k][1];
      var sortedProds = Object.keys(prods).map(function(n){ return [n, prods[n]]; })
        .sort(function(a,b) { return b[1][1] - a[1][1]; });
      html += '<div class="sp-store-card">' +
        '<div class="sp-store-hdr">' +
          '<span>' + shortName + '</span>' +
          '<span style="font-size:11px;font-weight:600;color:var(--muted)">' + totalQty + ' stk</span>' +
        '</div>' +
        '<div class="sp-store-body">';
      sortedProds.forEach(function(pe) {
        html += '<div class="sp-item">' +
          '<span class="sp-item-name" title="' + pe[0] + '">' + pe[0] + '</span>' +
          '<span class="sp-item-qty">' + pe[1][1] + '</span>' +
        '</div>';
      });
      html += '</div></div>';
    });
    wrap.innerHTML = html;
  };

  // ── 4. "Vara sala eftir verslun": "Allur mánuður" = selected month only ───
  window.renderStoreSales = function monthScopedStoreSales() {
    var wrap = document.getElementById('store-sale-wrap');
    if (!wrap) return;

    var selEl = document.getElementById('store-sale-date');
    var srchEl = document.getElementById('store-sale-search');
    var selDate = selEl ? selEl.value : 'all';
    var srch = srchEl ? srchEl.value.toLowerCase() : '';

    var datesToShow = (selDate === 'all') ? monthDates() : [selDate];

    var storeMap = {};
    for (var di = 0; di < datesToShow.length; di++) {
      var d = datesToShow[di];
      var dayData = STORE_PRODS[d];
      if (!dayData) continue;
      var storeNames = Object.keys(dayData);
      for (var si = 0; si < storeNames.length; si++) {
        var sname = storeNames[si];
        var prods = dayData[sname];
        if (!storeMap[sname]) storeMap[sname] = {};
        for (var pi = 0; pi < prods.length; pi++) {
          var ppnr  = prods[pi][0];
          var pname = prods[pi][1];
          var psale = prods[pi][2];
          var pqty  = prods[pi][3];
          if (!storeMap[sname][pname]) storeMap[sname][pname] = {sale:0, qty:0, pnr: ppnr};
          storeMap[sname][pname].sale += psale;
          storeMap[sname][pname].qty  += pqty;
        }
      }
    }

    var storeList = Object.keys(storeMap).sort(function(a, b) {
      var sa = 0, sb = 0;
      var pa = Object.values(storeMap[a]); for (var i=0;i<pa.length;i++) sa += pa[i].sale;
      var pb = Object.values(storeMap[b]); for (var i=0;i<pb.length;i++) sb += pb[i].sale;
      return sb - sa;
    });

    var out = '<table style="width:100%;border-collapse:collapse">';
    out += '<thead><tr style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:0.5px;border-bottom:2px solid var(--border)">';
    out += '<th style="padding:7px 10px;text-align:left;width:80px">Nr.</th>';
    out += '<th style="padding:7px 10px;text-align:left">Verslun / Vara</th>';
    out += '<th style="padding:7px 10px;text-align:right">Magn</th>';
    out += '<th style="padding:7px 10px;text-align:right">Sala (kr)</th>';
    out += '</tr></thead><tbody>';

    var totalRows = 0;
    var grandTotalSale = 0, grandTotalQty = 0;

    for (var si = 0; si < storeList.length; si++) {
      var sname = storeList[si];
      var shortName = sname.replace('Krónan ', '');
      var prods = storeMap[sname];
      var prodNames = Object.keys(prods);

      var matchStore = srch === '' || shortName.toLowerCase().indexOf(srch) >= 0;
      var filteredProds = [];
      for (var pi = 0; pi < prodNames.length; pi++) {
        var pn = prodNames[pi];
        var ppnr = prods[pn].pnr || '';
        if (matchStore || pn.toLowerCase().indexOf(srch) >= 0 || ppnr.toLowerCase().indexOf(srch) >= 0) {
          filteredProds.push(pn);
        }
      }
      if (filteredProds.length === 0) continue;

      filteredProds.sort(function(a, b) { return prods[b].qty - prods[a].qty; });

      var storeTotalSale = 0, storeTotalQty = 0;
      for (var pi = 0; pi < filteredProds.length; pi++) {
        storeTotalSale += prods[filteredProds[pi]].sale;
        storeTotalQty  += prods[filteredProds[pi]].qty;
      }
      grandTotalSale += storeTotalSale;
      grandTotalQty  += storeTotalQty;

      out += '<tr style="background:#f4f6fa;border-top:2px solid var(--border)">';
      out += '<td style="padding:8px 10px;font-weight:800;font-size:12px;color:var(--muted)"></td>';
      out += '<td style="padding:8px 10px;font-weight:800;font-size:12px;color:var(--text)">' + shortName + '</td>';
      out += '<td style="padding:8px 10px;text-align:right;font-weight:700;font-size:12px;color:var(--muted)">' + storeTotalQty.toLocaleString('is-IS') + ' stk</td>';
      out += '<td style="padding:8px 10px;text-align:right;font-weight:700;font-size:12px;color:var(--red)">' + fmtKr(storeTotalSale) + '</td>';
      out += '</tr>';

      for (var pi = 0; pi < filteredProds.length; pi++) {
        var pn = filteredProds[pi];
        var psale = prods[pn].sale;
        var pqty  = prods[pn].qty;
        var ppnr  = prods[pn].pnr || '';
        out += '<tr style="border-bottom:1px solid #f3f4f6">';
        out += '<td style="padding:5px 10px 5px 22px;font-size:11px;color:var(--muted);font-family:monospace;white-space:nowrap">' + ppnr + '</td>';
        out += '<td style="padding:5px 10px;font-size:12px;color:var(--text)">' + pn + '</td>';
        out += '<td style="padding:5px 10px;text-align:right;font-size:12px;font-weight:600">' + pqty + '</td>';
        out += '<td style="padding:5px 10px;text-align:right;font-size:12px;color:var(--muted)">' + fmtKr(psale) + '</td>';
        out += '</tr>';
        totalRows++;
      }
    }

    out += '</tbody>';
    out += '<tfoot><tr style="border-top:3px solid var(--text);background:var(--text)">';
    out += '<td style="padding:10px 10px;font-weight:800;font-size:13px;color:#fff"></td>';
    out += '<td style="padding:10px 10px;font-weight:800;font-size:13px;color:#fff;text-transform:uppercase;letter-spacing:0.5px">Samtals</td>';
    out += '<td style="padding:10px 10px;text-align:right;font-weight:800;font-size:13px;color:#fff">' + grandTotalQty.toLocaleString('is-IS') + ' stk</td>';
    out += '<td style="padding:10px 10px;text-align:right;font-weight:800;font-size:13px;color:#fbbf24">' + fmtKr(grandTotalSale) + '</td>';
    out += '</tr></tfoot>';
    out += '</table>';

    if (totalRows === 0) {
      wrap.innerHTML = '<div style="padding:20px;text-align:center;color:var(--muted);font-size:12px">Engar niðurstöður</div>';
    } else {
      wrap.innerHTML = out;
    }

    var countEl = document.getElementById('store-sale-count');
    if (countEl) countEl.textContent = storeList.length + ' verslanir, ' + totalRows + ' vörur';
  };

  // ── 5. Initial render: apply month-total view on page load ────────────────
  currentDate = 'all';
  var allBtn = document.querySelector('#date-tabs .tab-all');
  if (allBtn) {
    document.querySelectorAll('.tab,.tab-all').forEach(function(b){b.classList.remove('active');});
    allBtn.classList.add('active');
  }
  render();

})();
"""

patched = src[:inject_point] + JS_PATCH + src[inject_point:]

with open(HTML, 'w', encoding='utf-8') as f:
    f.write(patched)

print(f'V6 injected: month-scoped stores breakdown + KPI + heatmap + products ({len(patched):,} bytes)')
