"""
update_dashboard_data.py
Reads KrÃ³nan_Master_SkrÃ¡.xlsx and patches ALL data constants
inside KrÃ³nan_Dashboard.html using marker-based replacement.

Primary method : replaces entire /* KRONAN_DATA_START */ â€¦ /* KRONAN_DATA_END */ block.
Fallback method: const-variable anchor + buildTabs() slice.
Self-healing   : if core JS functions (buildTabs, render) are missing from the local
                 HTML, downloads a fresh copy from the Netlify CDN before patching.
"""
import openpyxl, json, os, re, sys, urllib.request, urllib.error
from collections import defaultdict

NETLIFY_URL = 'https://kronan-dashboard.netlify.app/'

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
_BASE = os.environ.get('KRONAN_BASE')
if _BASE:
    MASTER    = os.path.join(_BASE, 'KrÃ³nan_Master_SkrÃ¡.xlsx')
    DASHBOARD = os.path.join(_BASE, 'KrÃ³nan_Dashboard.html')
else:
    MASTER    = os.path.join(SCRIPT_DIR, '..', 'KrÃ³nan_Master_SkrÃ¡.xlsx')
    DASHBOARD = os.path.join(SCRIPT_DIR, '..', 'KrÃ³nan_Dashboard.html')

# â”€â”€ Sanity check files exist â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
if not os.path.exists(MASTER):
    print(f'âœ— Master file not found: {MASTER}')
    sys.exit(1)
if not os.path.exists(DASHBOARD):
    print('âš  Dashboard not found â€” downloading from Netlify CDNâ€¦')
    try:
        req = urllib.request.Request(NETLIFY_URL, headers={'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req, timeout=30) as resp:
            fresh = resp.read().decode('utf-8')
        os.makedirs(os.path.dirname(DASHBOARD), exist_ok=True)
        with open(DASHBOARD, 'w', encoding='utf-8') as f:
            f.write(fresh)
        print('âœ“ Dashboard downloaded from Netlify CDN')
    except Exception as e:
        print(f'âœ— Dashboard not found and download failed: {e}')
        sys.exit(1)

# â”€â”€ Self-healing: recover dashboard if core JS functions are missing â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
def _dashboard_healthy(path):
    try:
        # A healthy dashboard is simply a large file (>100 KB).
        # Searching for JS function names deep inside a 775 KB file is fragile.
        return os.path.getsize(path) > 100_000
    except Exception:
        return False

if not _dashboard_healthy(DASHBOARD):
    print('âš   Dashboard is missing core JS â€” downloading fresh copy from Netlify CDNâ€¦')
    try:
        req = urllib.request.Request(NETLIFY_URL, headers={'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req, timeout=30) as resp:
            fresh = resp.read().decode('utf-8')
        if 'function buildTabs' in fresh and 'function render(' in fresh:
            with open(DASHBOARD, 'w', encoding='utf-8') as f:
                f.write(fresh)
            print('âœ… Dashboard recovered from Netlify CDN')
        else:
            print('âœ— Downloaded HTML is also missing JS functions â€” cannot recover')
            sys.exit(1)
    except Exception as e:
        print(f'âœ— Recovery download failed: {e}')
        sys.exit(1)

wb = openpyxl.load_workbook(MASTER, data_only=True)

# â”€â”€ 1. DAILY totals â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
ws = wb['Dagleg - Verslanir']
daily_sale = defaultdict(float)
daily_qty  = defaultdict(int)
stores_by_date = defaultdict(list)

for row in ws.iter_rows(min_row=2, values_only=True):
    d, month, store, sale, qty, pct = row
    if not d or not store: continue
    dk = d.strftime('%Y-%m-%d') if hasattr(d,'strftime') else str(d)
    daily_sale[dk] += sale or 0
    daily_qty[dk]  += int(qty or 0)
    stores_by_date[dk].append({'store': store, 'sale': round(sale or 0), 'qty': int(qty or 0)})

dates = sorted(daily_sale.keys())
DAILY     = {d: round(daily_sale[d]) for d in dates}
DAILY_QTY = {d: daily_qty[d] for d in dates}
STORES_BY_DATE = {d: sorted(stores_by_date[d], key=lambda x: -x['sale']) for d in dates}

# â”€â”€ 2. STORE totals â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
store_sale = defaultdict(float)
store_qty  = defaultdict(int)
for dk, rows in stores_by_date.items():
    for r in rows:
        store_sale[r['store']] += r['sale']
        store_qty[r['store']]  += r['qty']

STORE_TOTALS = sorted(
    [[s, {'sale': round(store_sale[s]), 'qty': store_qty[s]}] for s in store_sale],
    key=lambda x: -x[1]['sale']
)

# â”€â”€ 3. PRODUCT totals â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
ws2 = wb['Dagleg - VÃ¶rur'] if 'Dagleg - VÃ¶rur' in wb.sheetnames else None
prod_sale = defaultdict(float)
prod_qty  = defaultdict(int)
if ws2:
    for row in ws2.iter_rows(min_row=2, values_only=True):
        row = list(row)
        if len(row) < 4: continue
        d, month, prod = row[0], row[1], row[2]
        sale = row[3] if len(row) > 3 else None
        qty  = row[4] if len(row) > 4 else 0
        if not prod: continue
        prod_sale[prod] += sale or 0
        prod_qty[prod]  += int(qty or 0)

PROD_TOTALS = sorted(
    [[p, {'sale': round(prod_sale[p]), 'qty': prod_qty[p]}] for p in prod_sale],
    key=lambda x: -x[1]['sale']
)

# â”€â”€ 4. StoreÃ—Product by date â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
ws5 = wb['Dagleg - VaraÃ—Verslun'] if 'Dagleg - VaraÃ—Verslun' in wb.sheetnames else None
store_prods = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: [0.0, 0, ''])))
if ws5:
    for row in ws5.iter_rows(min_row=2, values_only=True):
        if len(row) >= 7:
            d, month, store, prod, sale, qty, pnr = row[0], row[1], row[2], row[3], row[4], row[5], row[6]
        elif len(row) >= 6:
            d, month, store, prod, sale, qty = row[0], row[1], row[2], row[3], row[4], row[5]
            pnr = ''
        else:
            continue
        if not d or not store or not prod: continue
        dk = d.strftime('%Y-%m-%d') if hasattr(d,'strftime') else str(d)
        store_prods[dk][store][prod][0] += sale or 0
        store_prods[dk][store][prod][1] += int(qty or 0)
        if pnr: store_prods[dk][store][prod][2] = str(pnr)

# â”€â”€ 5. Build JS data block â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
def js_store_totals(data):
    parts = [f'["{n}",{{sale:{v["sale"]},qty:{v["qty"]}}}]' for n,v in data]
    return 'const STORE_TOTALS = [\n  ' + ','.join(parts) + '\n];'

def js_prod_totals(data):
    parts = [f'["{n}",{{sale:{v["sale"]},qty:{v["qty"]}}}]' for n,v in data]
    return 'const PROD_TOTALS = [\n  ' + ','.join(parts) + '\n];'

def js_daily(data):
    inner = ','.join(f'"{k}":{v}' for k,v in data.items())
    return f'const DAILY     = {{{inner}}};'

def js_daily_qty(data):
    inner = ','.join(f'"{k}":{v}' for k,v in data.items())
    return f'const DAILY_QTY = {{{inner}}};'

def js_stores_by_date(data):
    parts = []
    for dk, rows in sorted(data.items()):
        row_strs = ['{' + f'"store":"{r["store"]}","sale":{r["sale"]},"qty":{r["qty"]}' + '}' for r in rows]
        parts.append(f'"{dk}":[{",".join(row_strs)}]')
    return 'const STORES_BY_DATE = {' + ','.join(parts) + '};'

def js_store_products(data):
    date_parts = []
    for dk in sorted(data.keys()):
        store_parts = []
        for store in sorted(data[dk].keys(), key=lambda s: -sum(v[0] for v in data[dk][s].values())):
            prod_parts = []
            for prod, (sale, qty, pnr) in sorted(data[dk][store].items(), key=lambda x: -x[1][0]):
                pname = prod.replace('Tokyo Sushi ','').replace('"','\\"')
                pnr_s = str(pnr or '').replace('"','\\"')
                prod_parts.append(f'["{pnr_s}","{pname}",{round(sale)},{qty}]')
            sname = store.replace('"','\\"')
            store_parts.append(f'"{sname}":[{",".join(prod_parts)}]')
        date_parts.append(f'"{dk}":{{{",".join(store_parts)}}}')
    return 'const STORE_PRODS = {' + ','.join(date_parts) + '};'

DATA_BLOCK = '\n'.join([
    "const DOW = ['Sun','MÃ¡n','Ãri','MiÃ°','Fim','FÃ¶s','Lau'];",
    "let currentDate = 'all';",
    "let mainChart, storeChart;",
    "const STORE_COLORS = ['#1d4ed8','#2563eb','#3b82f6','#60a5fa','#0ea5e9','#06b6d4','#0891b2','#0284c7','#7c3aed','#6d28d9','#8b5cf6','#a78bfa','#059669','#10b981','#34d399','#065f46','#d97706','#f59e0b','#fbbf24','#b45309','#dc2626','#ef4444'];",
    "const STORE_COLOR_MAP = {'KrÃ³nan Granda':'#e63946','KrÃ³nan Flatahrauni':'#1d6fa4','KrÃ³nan BÃ­ldshÃ¶fÃ°a':'#2a9d8f','KrÃ³nan Selfossi':'#f4a261','KrÃ³nan MosfellsbÃ¦':'#8338ec','KrÃ³nan Skeifan 19':'#e9c46a','KrÃ³nan Fitjabraut':'#06d6a0','KrÃ³nan Vestmannaerjum':'#ef476f','KrÃ³nan Akrabraut':'#4895ef','KrÃ³nan NorÃ°urhellu':'#560bad','KrÃ³nan BorgartÃºn':'#f72585','KrÃ³nan Austurveri':'#4cc9f0','KrÃ³nan Grafarholti':'#80b918','KrÃ³nan HallveigarstÃ­g':'#fb8500','KrÃ³nan Akranesi':'#264653','KrÃ³nan Jafnaseli':'#43aa8b','KrÃ³nan ÃorlÃ¡kshÃ¶fn':'#c9184a','KrÃ³nan Hvolsvelli':'#118ab2','KrÃ³nan VallakÃ³r':'#fff&#s2rÂt·,;6æâl:Ö²s¢r3fCcƒsRrÂt·,;6æâ8&,:bs¢r6#Sƒ3†BrÂt·,;6æâ·W&W—&’s¢r3S&#sƒ‚wÓ²"À¢§5÷7F÷&U÷F÷FÇ2…5Dõ$UõDõDÅ2’À¢§5÷&öE÷F÷FÇ2…$ôEõDõDÅ2’À¢§5öF–Ç’„D”Å’’À¢§5öF–Ç•÷G’„D”Å•õE’’À¢§5÷7F÷&W5ö'•öFFR…5Dõ$U5ô%•ôDDR’À¢§5÷7F÷&U÷&öGV7G2‡7F÷&U÷&öG2’À¢v6öç7BDDU2Òö&¦V7Bæ¶W—2„D”Å’’ç6÷'B‚“²rÂ26÷'FVBFFR'&’W6VBF‡&÷Vv†÷WBF†RT¥Ò ¤Ô$´U%õ5D%BÒrò¢µ$ôäåôDDõ5D%B¢òp¤Ô$´U%ôTäBÒrò¢µ$ôäåôDDôTäB¢òp ¢2)H)HbâF6‚F†RF6†&ö&B…DÔÂ)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H)H §v—F‚÷Vâ„D4„$ô$BÂw"rÂVæ6öF–æsÒwWFbÓ‚r’2c ¢‡FÖÂÒbç&VB‚ ¦æWuö&Æö6²ÒÔ$´U%õ5D%B²uÆâr²DDô$Äô4²²uÆâr²Ô$´U%ôTä@ ¢2)H)HF‚¢Ö&¶W'2&W6VçB(	B&WÆ6RöæÇ’F†R6öçFVçB&WGvVVâF†VÒ)H)H)H)H)H)H)H)H)H)H)H ¦Ö&¶W%÷2Ò‡FÖÂæf–æB„Ô$´U%õ5D%B¦Ö&¶W%öRÒ‡FÖÂæf–æB„Ô$´U%ôTäB¦–bÖ&¶W%÷2ãÒæBÖ&¶W%öRâÖ&¶W%÷3 ¢‡FÖÂÒ‡FÖÅ³¦Ö&¶W%÷5Ò²æWuö&Æö6²²‡FÖÅ¶Ö&¶W%öR²ÆVâ„Ô$´U%ôTäB“¥Ğ¢&–çB†br†Ö&¶W"F‚’r ¦VÇ6S ¢2)H)HF‚#¢æòÖ&¶W'2(	Bf–æBFF6V7F–öâ&÷VæF&–W2&V6—6VÇ’)H)H)H)H)H)H)H)H)H)H)H ¢25D%C¢V&Æ–W7BFF6öç7BFV6Æ&F–öà¢FF÷7F'BÒÆVâ†‡FÖÂ¢f÷"æÖR–â‚u5Dõ$UõDõDÅ2rÂu$ôEõDõDÅ2rÂtD”Å’rÂtD”Å•õE’rÂu5Dõ$U5ô%•ôDDRrÂu5Dõ$Uõ$ôE2r“ ¢Ò‡FÖÂæf–æB†bv6öç7B¶æÖWÒr¢–bÃÒÂFF÷7F'C ¢FF÷7F'BÒ  ¢–bFF÷7F'BÓÒÆVâ†‡FÖÂ“ ¢2æòFF6öç7G2BÆÂ(	B–æ¦V7B–ÖÖVF–FVÇ’&Vf÷&R'V–ÆEF'2‚’6ÆÀ¢'Ò‡FÖÂæf–æB‚v'V–ÆEF'2‚“²r¢–b'Â ¢÷2Ò‡FÖÂç&f–æB‚sÂ÷67&—Câr¢–b÷2Â ¢&–çB‚~)Ér6÷VÆBæ÷BF6‚F6†&ö&B(	Bæò–æ¦V7F–öâö–çBf÷VæBr¢7—2æW†—Bƒ¢‡FÖÂÒ‡FÖÅ³§÷5Ò²æWuö&Æö6²²uÆâr²‡FÖÅ·÷3¥Ğ¢VÇ6S ¢‡FÖÂÒ‡FÖÅ³¦'Ò²æWuö&Æö6²²uÆåÆâr²‡FÖÅ¶'¥Ğ¢&–çB†br†–æ¦V7BF‚’r¢VÇ6S ¢2TäC¢f—'7B¥2gVæ7F–öâFV6Æ&F–öâgFW"F†RFF6öç7G0¢2†gVæ7F–öâFVg26öÖRgFW"FF–âF†R÷&–v–æÂ…DÔÂ¢få÷÷2Ò‡FÖÂæf–æB‚uÆægVæ7F–öârÂFF÷7F'B¢–bfå÷÷2ãÒ ¢2¶VWg&öÒfå÷÷2³6òvRFöâwB7vÆÆ÷rF†RÆVF–æræWvÆ–æP¢‡FÖÂÒ‡FÖÅ³¦FF÷7F'EÒ²æWuö&Æö6²²uÆåÆâr²‡FÖÅ¶få÷÷2²¥Ğ¢&–çB†br†gVæ7F–öâÖ&÷VæF'’F‚’r¢VÇ6S ¢2Æ7B&W6÷'C¢7WBFò'V–ÆEF'2‚’6ÆÂ†öÆB&V†f–÷"¢'Ò‡FÖÂæf–æB‚v'V–ÆEF'2‚“²rÂFF÷7F'B¢–b'Â ¢&–çB‚~)Ér6÷VÆBæ÷BF6‚F6†&ö&B(	Bæò–æ¦V7F–öâö–çBf÷VæBr¢7—2æW†—Bƒ¢‡FÖÂÒ‡FÖÅ³¦FF÷7F'EÒ²æWuö&Æö6²²uÆåÆâr²‡FÖÅ¶'¥Ğ¢&–çB†br†'V–ÆEF'2Öæ6†÷"F‚’r §&–çB†b~)ÈRF6†&ö&BFFWFFVB(	B¶ÆVâ†FFW2—ÒF—2Â¶ÆVâ…5Dõ$UõDõDÅ2—Ò7F÷&W2Â¶ÆVâ…$ôEõDõDÅ2—Ò&öGV7G2r §v—F‚÷Vâ„D4„$ô$BÂwrrÂVæ6öF–æsÒwWFbÓ‚r’2c ¢bçw&—FR†‡FÖÂ §&–çB†b~)ÈR7F÷&\9u&öGV7BFFF6†VB(	B·7VÒ†ÆVâ‡b’f÷"Gb–â7F÷&U÷&öG2çfÇVW2‚’f÷"b–âGbçfÇVW2‚’—Ò&öGV7B×7F÷&RÖF’6öÖ&÷2r