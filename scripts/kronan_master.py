"""
KrÃ³nan Master Sales Tracker
Run: python3 kronan_master.py <new_report.xlsx>
- Appends new daily data to master file (no duplicates)
- Regenerates monthly summaries automatically
"""

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import defaultdict
import sys, os
from openpyxl.cell.cell import MergedCel

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE = os.environ.get('KRONAN_BASE')
MASTER = os.path.join(_BASE, 'KrÃ³nan_Master_SkrÃ¡.xlsx') if _BASE else os.path.join(os.path.expanduser("~"), "Documents", "KrÃ³nan", "KrÃ³nan_Master_SkrÃ¡.xlsx")

# ââ Styles ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
BLUE   = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
DARK   = PatternFill("solid", start_color="1A5276", end_color="1A5276")
GREEN  = PatternFill("solid", start_color="1E8449", end_color="1E8449")
MHDR   = PatternFill("solid", start_color="117A65", end_color="117A65")
ALT    = PatternFill("solid", start_color="EBF5FB", end_color="EBF5FB")
ALT2   = PatternFill("solid", start_color="E9F7EF", end_color="E9F7EF")
WHITE  = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
TITL   = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
thin   = Side(style="thin", color="BDC3C7")
brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
hf     = Font(name="Arial", bold=True, color="FFFFFF", size=10)
df     = Font(name="Arial", size=10)
tf     = Font(name="Arial", bold=True, color="FFFFFF", size=10)
titf   = Font(name="Arial", bold=True, color="1F4E79", size=13)
secf   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
ctr    = Alignment(horizontal="center", vertical="center")
lft    = Alignment(horizontal="left",   vertical="center")
rgt    = Alignment(horizontal="right",  vertical="center")

def style_header(cell, text, fill=BLUE):
    cell.value = cell.value if text is None else text
    cell.fill = fill; cell.font = hf; cell.alignment = ctr; cell.border = brd

def style_data(cell, val, fmt=None, aln=None, fill=WHITE):
    cell.value = val
    cell.font = df; cell.fill = fill; cell.border = brd
    cell.alignment = aln or lft
    if fmt: cell.number_format = fmt

def style_total(cell, val, fmt=None, aln=None):
    cell.value = val; cell.fill = DARK; cell.font = tf; cell.border = brd
    cell.alignment = aln or lft
    if fmt: cell.number_format = fmt

def style_month_hdr(cell, val):
    cell.value = val; cell.fill = GREEN; cell.font = secf
    cell.alignment = lft; cell.border = brd

# ââ Read incoming report âââââââââââââââââââââââââââââââââââââââââââââââââââââ
def read_report(path):
    wb = openpyxl.load_workbook(path)
    # Stores â try canonical name, fall back to first sheet
    try:
        ws_s = wb['Verslanir']
    except KeyError:
        # Sheet name may differ (e.g. localised or renamed) â use first sheet
        ws_s = wb.worksheets[0]
        print(f"  â  Sheet 'Verslanir' not found, using first sheet: {ws_s.title}")
    store_rows = defaultdict(lambda: [0.0, 0])
    store_product_rows = defaultdict(lambda: defaultdict(lambda: [0.0, 0, '']))
    date = None
    for row in ws_s.iter_rows(min_row=2, values_only=True):
        if len(row) < 9: continue
        d, chain, ean, store, pnr, prod, spnr, sale, qty = row[:9]
        if store and sale:
            store_rows[store][0] += float(str(sale).replace(',','').strip() or 0); store_rows[store][1] += int(qty or 0)
            if d and not date: date = d
        if store and prod and sale:
            store_product_rows[store][prod][0] += float(str(sale).replace(',','').strip() or 0)
            store_product_rows[store][prod][1] += int(qty or 0)
            # Use spnr (supplier item number, e.g. R0173) if available, else pnr
            item_num = str(spnr or pnr or '')
            store_product_rows[store][prod][2] = item_num
    # Products â try canonical name, fall back to second sheet
    try:
        ws_p = wb['Heild']
    except KeyError:
        ws_p = wb.worksheets[1] if len(wb.worksheets) > 1 else None
        if ws_p:
            print(f"  â  Sheet 'Heild' not found, using sheet: {ws_p.title}")
    item_rows = {}
    if ws_p:
        for row in ws_p.iter_rows(min_row=2, values_only=True):
            if len(row) < 4: continue
            pnr, prod, spnr, sale = row[:4]
            qty = row[4] if len(row) > 4 else 0
            if prod and sale: item_rows[prod] = [sale, int(qty or 0)]
    if not date:
        raise ValueError(f"Could not extract a date from report: {path}")
    return date, store_rows, item_rows, store_product_rows

# ââ Load or create master ââââââââââââââââââââââââââââââââââââââââââââââââââââ
def load_or_create():
    if os.path.exists(MASTER):
        return load_workbook(MASTER)
    wb = Workbook()
    # Sheet 1: Daily stores
    ws = wb.active; ws.title = "Dagleg - Verslanir"
    for col, (h, w) in enumerate(zip(
        ["Dags","MÃ¡nuÃ°ur","Verslun","Sala (kr)","Magn","% Dagsins"], 
        [14, 12, 32, 18, 10, 12]), 1):
        c = ws.cell(1, col); style_header(c, h)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"
    # Sheet 2: Monthly stores
    ws2 = wb.create_sheet("MÃ¡naÃ°arleg - Verslanir")
    for col, (h, w) in enumerate(zip(
        ["MÃ¡nuÃ°ur","Verslun","Sala (kr)","Magn","% MÃ¡naÃ°arins"],
        [14, 32, 18, 10, 16]), 1):
        c = ws2.cell(1, col); style_header(c, h, MHDR)
        ws2.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws2.row_dimensions[1].height = 20
    ws2.freeze_panes = "A2"
    # Sheet 3: Daily products
    ws3 = wb.create_sheet("Dagleg - VÃ¶rur")
    for col, (h, w) in enumerate(zip(
        ["Dags","MÃ¡nuÃ°ur","Vara","Sala (kr)","Magn","% Dagsins"],
        [14, 12, 44, 18, 10, 12]), 1):
        c = ws3.cell(1, col); style_header(c, h)
        ws3.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws3.row_dimensions[1].height = 20
    ws3.freeze_panes = "A2"
    # Sheet 4: Monthly products
    ws4 = wb.create_sheet("MÃ¡naÃ°arleg - VÃ¶rur")
    for col, (h, w) in enumerate(zip(
        ["MÃ¡nuÃ°ur","Vara","Sala (kr)","Magn","% MÃ¡naÃ°arins"],
        [14, 44, 18, 10, 16]), 1):
        c = ws4.cell(1, col); style_header(c, h, MHDR)
        ws4.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws4.row_dimensions[1].height = 20
    ws4.freeze_panes = "A2"
    # Sheet 5: Daily storeÃproduct breakdown
    ws5 = wb.create_sheet("Dagleg - VaraÃVerslun")
    for col, (h, w) in enumerate(zip(
        ["Dags","MÃ¡nuÃ°ur","Verslun","Vara","Sala (kr)","Magn"],
        [14, 12, 32, 44, 18, 10]), 1):
        c = ws5.cell(1, col); style_header(c, h)
        ws5.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
    ws5.row_dimensions[1].height = 20
    ws5.freeze_panes = "A2"
    return wb

# ââ Append daily data ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def append_daily(wb, date, store_rows, item_rows, store_product_rows=None):
    date_val = date.date() if hasattr(date, 'date') else date
    month_str = date.strftime("%Y-%m") if hasattr(date, 'strftime') else str(date)[:7]

    # --- Stores sheet ---
    ws = wb["Dagleg - Verslanir"]
    # Check for duplicates
    existing_dates = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]: 
            d = row[0].date() if hasattr(row[0], 'date') else row[0]
            existing_dates.add(d)
    if date_val in existing_dates:
        print(f"  â  {date_val} already in Dagleg - Verslanir, skipping.")
    else:
        store_total = sum(v[0] for v in store_rows.values())
        stores_sorted = sorted(store_rows.items(), key=lambda x: -x[1][0])
        r = ws.max_row + 1
        for store, (sale, qty) in stores_sorted:
            fill = ALT if (r % 2 == 0) else WHITE
            pct = sale / store_total if store_total else 0
            for col, (val, fmt, aln) in enumerate([
                (date,    "DD.MM.YYYY", ctr),
                (month_str, None,       ctr),
                (store,   None,         lft),
                (sale,    '#,##0.00',   rgt),
                (qty,     '#,##0',      ctr),
                (pct,     '0.0%',       ctr),
            ], 1):
                style_data(ws.cell(r, col), val, fmt, aln, fill)
            ws.row_dimensions[r].height = 17
            r += 1
        print(f"  â Appended {len(stores_sorted)} store rows for {date_val}")

    # --- Products sheet ---
    ws3 = wb["Dagleg - VÃ¶rur"]
    existing_dates3 = set()
    for row in ws3.iter_rows(min_row=2, values_only=True):
        if row[0]:
            d = row[0].date() if hasattr(row[0], 'date') else row[0]
            existing_dates3.add(d)
    if date_val in existing_dates3:
        print(f"  â  {date_val} already in Dagleg - VÃ¶rur, skipping.")
    else:
        item_total = sum(v[0] for v in item_rows.values())
        items_sorted = sorted(item_rows.items(), key=lambda x: -x[1][0])
        r = ws3.max_row + 1
        for prod, (sale, qty) in items_sorted:
            fill = ALT if (r % 2 == 0) else WHITE
            pct = sale / item_total if item_total else 0
            for col, (val, fmt, aln) in enumerate([
                (date,      "DD.MM.YYYY", ctr),
                (month_str, None,         ctr),
                (prod,      None,         lft),
                (sale,      '#,##0.00',   rgt),
                (qty,       '#,##0',      ctr),
                (pct,       '0.0%',       ctr),
            ], 1):
                style_data(ws3.cell(r, col), val, fmt, aln, fill)
            ws3.row_dimensions[r].height = 17
            r += 1
        print(f"  â Appended {len(items_sorted)} product rows for {date_val}")

    # --- StoreÃProduct sheet ---
    if "Dagleg - VaraÃVerslun" not in wb.sheetnames:
        ws5 = wb.create_sheet("Dagleg - VaraÃVerslun")
        for col, (h, w) in enumerate(zip(
            ["Dags","MÃ¡nuÃ°ur","Verslun","Vara","Sala (kr)","Magn","VÃ¶runÃºmer"],
            [14, 12, 32, 44, 18, 10, 14]), 1):
            c = ws5.cell(1, col); style_header(c, h)
            ws5.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w
        ws5.row_dimensions[1].height = 20
        ws5.freeze_panes = "A2"
    ws5 = wb["Dagleg - VaraÃVerslun"]
    existing_dates5 = set()
    for row in ws5.iter_rows(min_row=2, values_only=True):
        if row[0]:
            d = row[0].date() if hasattr(row[0], 'date') else row[0]
            existing_dates5.add(d)
    if date_val in existing_dates5:
        print(f"  â  {date_val} already in Dagleg - VaraÃVerslun, skipping.")
    else:
        r = ws5.max_row + 1
        count = 0
        for store in sorted(store_product_rows.keys()):
            prods = sorted(store_product_rows[store].items(), key=lambda x: -x[1][0])
            for prod, (sale, qty, pnr) in prods:
                fill = ALT if (r % 2 == 0) else WHITE
                for col, (val, fmt, aln) in enumerate([
                    (date,      "DD.MM.YYYY", ctr),
                    (month_str, None,         ctr),
                    (store,     None,         lft),
                    (prod,      None,         lft),
                    (sale,      '#,##0.00',   rgt),
                    (qty,       '#,##0',      ctr),
                    (pnr,       None,         ctr),
                ], 1):
                    style_data(ws5.cell(r, col), val, fmt, aln, fill)
                ws5.row_dimensions[r].height = 17
                r += 1; count += 1
        print(f"  â Appended {count} storeÃproduct rows for {date_val}")

# ââ Rebuild monthly summaries ââââââââââââââââââââââââââââââââââââââââââââââââ
def rebuild_monthly(wb):
    # --- Monthly Stores ---
    ws_daily = wb["Dagleg - Verslanir"]
    monthly_stores = defaultdict(lambda: defaultdict(lambda: [0.0, 0]))
    for row in ws_daily.iter_rows(min_row=2, values_only=True):
        d, month, store, sale, qty, pct = row
        if month and store and sale:
            monthly_stores[month][store][0] += sale
            monthly_stores[month][store][1] += int(qty or 0)

    ws_m = wb["MÃ¡naÃ°arleg - Verslanir"]
    # Unmerge all, then clear
    for merge in list(ws_m.merged_cells.ranges):
        ws_m.unmerge_cells(str(merge))
    for row in ws_m.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell, MergedCell): continue
            cell.value = None; cell.fill = WHITE; cell.border = Border()

    r = 2
    for month in sorted(monthly_stores.keys()):
        stores = sorted(monthly_stores[month].items(), key=lambda x: -x[1][0])
        month_total = sum(v[0] for _, v in stores)
        # Month header
        ws_m.merge_cells(f"A{r}:E{r}")
        c = ws_m.cell(r, 1)
        style_month_hdr(c, f"  {month}")
        ws_m.row_dimensions[r].height = 18
        r += 1
        for store, (sale, qty) in stores:
            fill = ALT if (r % 2 == 0) else WHITE
            pct = sale / month_total if month_total else 0
            for col, (val, fmt, aln) in enumerate([
                (month, None, ctr), (store, None, lft),
                (sale, '#,##0.00', rgt), (qty, '#,##0', ctr), (pct, '0.0%', ctr)
            ], 1):
                style_data(ws_m.cell(r, col), val, fmt, aln, fill)
            ws_m.row_dimensions[r].height = 17
            r += 1
        # Month total row
        for col, (val, fmt, aln) in enumerate([
            (month, None, ctr), (f"HEILD {month}", None, lft),
            (month_total, '#,##0.00', rgt), (sum(v[1] for _,v in stores), '#,##0', ctr),
            ("100.0%", '0.0%', ctr)
        ], 1):
            style_total(ws_m.cell(r, col), val, fmt, aln)
        ws_m.row_dimensions[r].height = 18
        r += 1
    print(f"  â Monthly stores summary rebuilt ({len(monthly_stores)} months)")

    # --- Monthly Products ---
    ws_daily3 = wb["Dagleg - VÃ¶rur"]
    monthly_items = defaultdict(lambda: defaultdict(lambda: [0.0, 0]))
    for row in ws_daily3.iter_rows(min_row=2, values_only=True):
        d, month, prod, sale, qty, pct = row
        if month and prod and sale:
            monthly_items[month][prod][0] += sale
            monthly_items[month][prod][1] += int(qty or 0)

    ws_m4 = wb["MÃ¡naÃ°arleg - VÃ¶rur"]
    # Unmerge all, then clear
    for merge in list(ws_m4.merged_cells.ranges):
        ws_m4.unmerge_cells(str(merge))
    for row in ws_m4.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell, MergedCell): continue
            cell.value = None; cell.fill = WHITE; cell.border = Border()

    r = 2
    for month in sorted(monthly_items.keys()):
        items = sorted(monthly_items[month].items(), key=lambda x: -x[1][0])
        month_total = sum(v[0] for _, v in items)
        ws_m4.merge_cells(f"A{r}:E{r}")
        c = ws_m4.cell(r, 1)
        style_month_hdr(c, f"  {month}")
        ws_m4.row_dimensions[r].height = 18
        r += 1
        for prod, (sale, qty) in items:
            fill = ALT2 if (r % 2 == 0) else WHITE
            pct = sale / month_total if month_total else 0
            for col, (val, fmt, aln) in enumerate([
                (month, None, ctr), (prod, None, lft),
                (sale, '#,##0.00', rgt), (qty, '#,##0', ctr), (pct, '0.0%', ctr)
            ], 1):
                style_data(ws_m4.cell(r, col), val, fmt, aln, fill)
            ws_m4.row_dimensions[r].height = 17
            r += 1
        for col, (val, fmt, aln) in enumerate([
            (month, None, ctr), (f"HEILD {month}", None, lft),
            (month_total, '#,##0.00', rgt), (sum(v[1] for _,v in items), '#,##0', ctr),
            ("100.0%", '0.0%', ctr)
        ], 1):
            style_total(ws_m4.cell(r, col), val, fmt, aln)
        ws_m4.row_dimensions[r].height = 18
        r += 1
    print(f"  â Monthly products summary rebuilt ({len(monthly_items)} months)")

# ââ Main âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
def run(report_path):
    print(f"\nð Reading report: {report_path}")
    date, store_rows, item_rows, store_product_rows = read_report(report_path)
    print(f"  Date: {date}, Stores: {len(store_rows)}, Products: {len(item_rows)}")
    print(f"\nð Loading master file...")
    wb = load_or_create()
    print(f"\nâ Appending daily data...")
    append_daily(wb, date, store_rows, item_rows, store_product_rows)
    print(f"\nð Rebuilding monthly summaries...")
    rebuild_monthly(wb)
    wb.save(MASTER)
    print(f"\nâ Master file saved: {MASTER}")

if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SCRIPT_DIR, 'KrÃ³nan sÃ¶luskÃrsla.xlsx')
    run(path)
