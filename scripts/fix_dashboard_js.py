# -*- coding: utf-8 -*-
"""
fix_dashboard_js.py — ensures monthLabel() helper exists (fixes ReferenceError).
The render() function calls monthLabel() in the Yfirlit KPI sub-label.
If the function is missing, all charts go blank due to a ReferenceError.
"""
import os, sys

BASE = os.environ.get('KRONAN_BASE', 'data')
HTML = os.path.join(BASE, 'Krónan_Dashboard.html')

if not os.path.exists(HTML):
    print('Dashboard not found, skipping')
    sys.exit(0)

with open(HTML, 'r', encoding='utf-8') as f:
    src = f.read()

if 'function monthLabel' not in src:
    # Add the helper before const COMMISSION
    helper = """
function monthLabel(m) {
  if (!m) return '';
  const parts = m.split('-');
  const names = ['Jan','Feb','Mar','Apr','Maí','Jún','Júl','Ágú','Sep','Okt','Nóv','Des'];
  return (names[parseInt(parts[1])-1] || parts[1]) + ' ' + parts[0];
}"""
    src = src.replace('const COMMISSION', helper + '\n\nconst COMMISSION', 1)
    with open(HTML, 'w', encoding='utf-8') as f:
        f.write(src)
    print('✅ Added monthLabel() helper to prevent ReferenceError')
else:
    print('monthLabel() already present, nothing to do')
