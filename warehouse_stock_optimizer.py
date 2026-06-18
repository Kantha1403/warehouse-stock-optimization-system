import requests
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from datetime import date
import urllib3
import math
import time
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import sys

# Try to reconfigure stdout encoding (fails on some systems, safe to ignore)
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_URL    = ""
AUTH_TOKEN  = ""
OUTPUT_FILE = f"Warehouse_Stock_Optimizer_{date.today().strftime('%Y_%m')}.xlsx"

LEAD_TIME_MONTHS = 3   # reorder_qty = ceil(monthly_avg × 3)
MIN_REORDER_QTY  = 1   # minimum reorder when item has demand
CUTOFF_DATE = (pd.Timestamp.today() - pd.DateOffset(months=12)).strftime("%Y-%m-%d")

auth_headers = {"Authorization": AUTH_TOKEN}



# Full-access recipients (get complete report with all regions)
FULL_ACCESS_RECIPIENTS = [
    
]

# ASM email map
ASM_EMAIL_MAP = {
    
}

# Valid sub-regions
VALID_SUB_REGIONS = {
    
}

# ASM -> sub-region ownership (locked)
ASM_REGION_MAP = {
    
}

# Warehouse → region mapping
WAREHOUSE_REGION_MAP = {
    
}

SEND_EMAIL = True

# ── SESSION WITH RETRY ────────────────────────────────────────────────────────
session = requests.Session()
retry = Retry(total=5, backoff_factor=0.3,
              status_forcelist=[500, 502, 503, 504],
              allowed_methods=["GET"])
session.mount("https://", HTTPAdapter(max_retries=retry))

# ── FETCH ALL (paginated) ─────────────────────────────────────────────────────
def fetch_all(endpoint, fields, filters=None, page_size=200):
    records = []
    start = 0
    params = {
        "fields": fields,
        "limit_page_length": page_size,
    }
    if filters:
        params["filters"] = filters

    while True:
        params["limit_start"] = start
        r = session.get(
            f"{BASE_URL}{endpoint}",
            headers=auth_headers,
            params=params,
            timeout=60,
            verify=False,
        )
        r.raise_for_status()
        batch = r.json().get("data", [])
        if not batch:
            break
        records.extend(batch)
        if len(batch) < page_size:
            break
        start += page_size

    return records

# ── STYLE CONSTANTS ───────────────────────────────────────────────────────────
RED_FILL    = PatternFill("solid", fgColor="FF9999")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF99")
GREEN_FILL  = PatternFill("solid", fgColor="99FF99")
GREY_FILL   = PatternFill("solid", fgColor="D3D3D3")

_thin  = Side(style="thin")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)

def hdr(cell, value):
    cell.value     = value
    cell.font      = Font(bold=True, color="FFFFFF", size=10)
    cell.fill      = PatternFill("solid", fgColor="4472C4")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = BORDER

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — FETCH ITEM MASTER
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[1/6] Fetching item master...")
item_records = fetch_all(
    "",
    fields='["item_code","item_name","item_group","is_stock_item","stock_uom"]'
)
item_df = pd.DataFrame(item_records)
item_name_lookup  = item_df.set_index("item_code")["item_name"].to_dict()
item_group_lookup = item_df.set_index("item_code")["item_group"].to_dict()
print(f"  Items: {len(item_df)}")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — FETCH SERIAL NO → derive Covered Under per item_code
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[2/6] Fetching Serial No data...")
serial_records = fetch_all(
    "",
    fields='["name","item_code","territory","warranty_expiry_date","amc_expiry_date","maintenance_status","status"]',
    filters='[["disabled","=",0]]',
    page_size=500
)
serial_df = pd.DataFrame(serial_records)
print(f"  Serial records: {len(serial_df)}")

today_ts = pd.Timestamp(date.today())
serial_df["warranty_expiry_date"] = pd.to_datetime(serial_df["warranty_expiry_date"], errors="coerce")
serial_df["amc_expiry_date"]      = pd.to_datetime(serial_df["amc_expiry_date"],      errors="coerce")
serial_df["has_warranty"] = serial_df["warranty_expiry_date"] > today_ts
serial_df["has_amc"]      = serial_df["amc_expiry_date"]      > today_ts


def _covered(row):
    w, a = row["has_warranty"], row["has_amc"]
    if w and a:   return "Warranty + AMC"
    elif w:       return "Warranty"
    elif a:       return "AMC"
    else:         return "None"


serial_df["covered_under"] = serial_df.apply(_covered, axis=1)
_prio = {"Warranty + AMC": 4, "Warranty": 3, "AMC": 2, "None": 1}
serial_df["_prio"] = serial_df["covered_under"].map(_prio)
item_coverage = (serial_df.sort_values("_prio", ascending=False)
                 .drop_duplicates("item_code")
                 .set_index("item_code")["covered_under"]
                 .to_dict())
print(f"  Coverage map: {len(item_coverage)} items")

# Commitment For per warehouse (which plan types are active at that location)
location_coverage = {
    
}


def get_commitment(warehouse):
    loc = warehouse.replace("Engg - ", "").replace(" - EIPL", "").strip()
    return location_coverage.get(loc, "None")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — FETCH SERVICE COMMON SLE + MAP TO ENGG WAREHOUSES VIA DN
# Service Common consumption is attributed to an Engg warehouse using the
# Refurbish target warehouse field on the Delivery Note.
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[3/6] Fetching Service Common consumption...")

sc_records = []
start = 0
while True:
    params = {
        "filters": f'[["warehouse","=",""],["actual_qty","<",0],'
                   f'["is_cancelled","=",0],["posting_date",">=","{CUTOFF_DATE}"]]',
        "fields": '["item_code","warehouse","posting_date","actual_qty",'
                  '"voucher_type","voucher_no"]',
        "limit_start": start, "limit_page_length": 500,
    }
    r = session.get(f"{BASE_URL}",
                    headers=auth_headers, params=params, timeout=120, verify=False)
    batch = r.json().get("data", [])
    if not batch:
        break
    sc_records.extend(batch)
    if len(batch) < 500:
        break
    start += 500

sc_sle      = pd.DataFrame(sc_records) if sc_records else pd.DataFrame(
    columns=["item_code", "warehouse", "posting_date", "actual_qty", "voucher_no"])
unique_dns  = sc_sle["voucher_no"].unique().tolist() if len(sc_sle) else []
print(f"  SC records: {len(sc_sle)} | Unique DNs: {len(unique_dns)}")

# Look up each DN's set_target_warehouse to find which Engg location it belongs to
dn_target_map = {}
for i, dn in enumerate(unique_dns):
    try:
        r = session.get(f"{BASE_URL}",
                        headers=auth_headers, timeout=15, verify=False)
        if r.status_code == 200:
            dn_target_map[dn] = r.json().get("data", {}).get("set_target_warehouse")
        if i % 100 == 0:
            print(f"  DN lookup: {i}/{len(unique_dns)}...", end="\r")
        time.sleep(0.05)
    except Exception:
        pass
print(f"  DN lookup: {len(unique_dns)}/{len(unique_dns)} done          ")


def _extract_city(target):
    if target and "Refurbish" in str(target):
        city = str(target).replace("Refurbish - ", "").replace(" - EIPL", "").strip()
        return city if city != "Service Common" else None
    return None


if len(sc_sle):
    sc_sle["city"]             = sc_sle["voucher_no"].map(dn_target_map).apply(_extract_city)
    sc_sle["mapped_warehouse"] = sc_sle["city"].apply(lambda c: f"Engg - {c}" if c else None)
    sc_mapped = sc_sle[sc_sle["mapped_warehouse"].notna()].copy()
    sc_mapped = sc_mapped.drop(columns=["warehouse"]).rename(columns={"mapped_warehouse": "warehouse"})
    sc_mapped["warehouse"] = sc_mapped["warehouse"].str.replace(
        "Engg - Ahemdabad", "Engg - Ahmedabad")   # fix ERP spelling error
    sc_mapped["warehouse"] = sc_mapped["warehouse"].apply(
        lambda w: w + " - EIPL" if not w.endswith("- EIPL") else w)
    sc_mapped = sc_mapped[["item_code", "warehouse", "posting_date", "actual_qty"]].copy()
    sc_mapped["actual_qty"] = pd.to_numeric(sc_mapped["actual_qty"], errors="coerce").fillna(0)
    print(f"  SC mapped: {len(sc_mapped)} records → {sc_mapped['warehouse'].nunique()} warehouses")
else:
    sc_mapped = pd.DataFrame(columns=["item_code", "warehouse", "posting_date", "actual_qty"])
    print("  No SC records found")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — FETCH ENGG + REPAIRS SLE (Delivery Notes, last 12 months)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[4/6] Fetching Engg + Repairs SLE...")

engg_records = []
start = 0
while True:
    params = {
        "filters": f'[["voucher_type","=","Delivery Note"],["actual_qty","<",0],'
                   f'["is_cancelled","=",0],["posting_date",">=","{CUTOFF_DATE}"]]',
        "fields": '["item_code","warehouse","posting_date","actual_qty"]',
        "limit_start": start, "limit_page_length": 500,
    }
    r = session.get(f"{BASE_URL}",
                    headers=auth_headers, params=params, timeout=180, verify=False)
    batch = r.json().get("data", [])
    if not batch:
        break
    engg_records.extend(batch)
    print(f"  {len(engg_records)} SLE records...", end="\r")
    if len(batch) < 500:
        break
    start += 500

engg_sle_df = pd.DataFrame(engg_records)
engg_sle_df["actual_qty"]    = pd.to_numeric(engg_sle_df["actual_qty"], errors="coerce").fillna(0)
engg_sle_df["posting_date"]  = pd.to_datetime(engg_sle_df["posting_date"])

engg_repairs_sle = engg_sle_df[
    engg_sle_df["warehouse"].str.startswith("Engg -") |
    (engg_sle_df["warehouse"] == "Repairs - EIPL")
].copy()
print(f"  Engg+Repairs SLE: {len(engg_repairs_sle)} records")

# Combine direct Engg/Repairs + SC-mapped
sc_mapped["posting_date"] = pd.to_datetime(sc_mapped["posting_date"])
combined_sle = pd.concat([engg_repairs_sle, sc_mapped], ignore_index=True)
combined_sle["actual_qty"] = pd.to_numeric(combined_sle["actual_qty"], errors="coerce").fillna(0)

date_min = combined_sle["posting_date"].min()
date_max = combined_sle["posting_date"].max()
months   = max((date_max - date_min).days / 30.44, 1)
print(f"  Date range: {date_min.date()} → {date_max.date()} ({months:.1f} months)")

demand = (combined_sle.groupby(["item_code", "warehouse"])["actual_qty"]
          .sum().abs().reset_index())
demand.columns = ["item_code", "warehouse", "total_consumption"]
demand["monthly_avg"]  = (demand["total_consumption"] / months).round(3)
demand["reorder_qty"]  = demand["monthly_avg"].apply(
    lambda x: max(math.ceil(x * LEAD_TIME_MONTHS), MIN_REORDER_QTY) if x > 0 else 0)
print(f"  Demand: {len(demand)} item-warehouse pairs | {demand['item_code'].nunique()} unique items")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — FETCH ALL SPARE PARTS + CURRENT STOCK (Bin)
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[5/6] Fetching spare parts and current stock...")

spare_records = fetch_all(
    "",
    fields='["name","item_name","item_group","stock_uom"]',
    filters='[["item_group","like","%Spares%"],["disabled","=",0]]',
    page_size=500
)
spares_all  = pd.DataFrame(spare_records).rename(columns={"name": "item_code"})
spare_codes = set(spares_all["item_code"].tolist())

for _, row in spares_all.iterrows():
    item_name_lookup.setdefault(row["item_code"],  row.get("item_name",  ""))
    item_group_lookup.setdefault(row["item_code"], row.get("item_group", ""))

accessories = spares_all[spares_all["item_group"].str.contains("Accessories", case=False, na=False)]
breakdown   = spares_all[spares_all["item_group"].str.contains("Breakdown",   case=False, na=False)]
print(f"  Spare parts: {len(spares_all)} | Accessories: {len(accessories)} | Breakdown: {len(breakdown)}")

# Fetch current stock for all spare items in batches of 100
stock_records = []
spare_list    = list(spare_codes)
for i in range(0, len(spare_list), 100):
    batch       = spare_list[i:i + 100]
    item_filter = '[\"item_code\",\"in\",' + str(batch).replace("'", '"') + ']'
    params = {
        "filters": f'[{item_filter},["actual_qty",">",0]]',
        "fields":  '["item_code","warehouse","actual_qty"]',
        "limit_page_length": 500,
    }
    r = session.get(f"{BASE_URL}",
                    headers=auth_headers, params=params, verify=False, timeout=60)
    stock_records.extend(r.json().get("data", []))
    if i % 500 == 0:
        print(f"  Stock: {i}/{len(spare_list)} items...", end="\r")
print(f"  Stock: {len(spare_list)}/{len(spare_list)} items done")

stock_df = pd.DataFrame(stock_records) if stock_records else pd.DataFrame(
    columns=["item_code", "warehouse", "actual_qty"])
stock_df["actual_qty"] = pd.to_numeric(stock_df["actual_qty"], errors="coerce").fillna(0)

# engg_stock with location column for Accessories/Breakdown sheets
engg_stock = stock_df[stock_df["warehouse"].str.startswith("Engg -")].copy()
engg_stock["location"] = (engg_stock["warehouse"]
                          .str.replace(" - EIPL", "", regex=False)
                          .str.replace("Engg - ",  "", regex=False))
locations = sorted(engg_stock["location"].unique().tolist())
print(f"  Locations ({len(locations)}): {locations}")

stock_lookup = {(r["item_code"], r["warehouse"]): r["actual_qty"]
                for _, r in stock_df.iterrows()}

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — BUILD EXCEL REPORT
# ═══════════════════════════════════════════════════════════════════════════════
print("\n[6/6] Building Excel report...")

wb_main = openpyxl.Workbook()   # Warehouse Report + Breakdown
wb_acc  = openpyxl.Workbook()   # Accessories only

# ════════════════════════════════════════════════════════════════════════════════
# SHEET 1: Warehouse Report
# One row per (spare item, warehouse) that has demand OR stock
# ════════════════════════════════════════════════════════════════════════════════
ws1 = wb_main.active
ws1.title = "Warehouse Report"

WR_LEGEND = [
    ("Negative Gap",             "Stock BELOW required — replenishment needed", RED_FILL),
    ("Positive Gap",             "Stock ABOVE required — excess inventory",      YELLOW_FILL),
    ("Zero Gap",                 "Stock matches required — optimal",             GREEN_FILL),
    ("No Stock / No Requirement","No consumption and no current stock",          GREY_FILL),
]
for i, (label, desc, fill) in enumerate(WR_LEGEND, 1):
    for c in [11, 12]:
        ws1.cell(i, c).fill      = fill
        ws1.cell(i, c).border    = BORDER
        ws1.cell(i, c).alignment = LEFT
    ws1.cell(i, 11).value = label
    ws1.cell(i, 11).font  = Font(bold=True, size=9)
    ws1.cell(i, 12).value = desc
    ws1.cell(i, 12).font  = Font(size=9)
ws1.column_dimensions[get_column_letter(11)].width = 30
ws1.column_dimensions[get_column_letter(12)].width = 50

HROW = 5
WR_HEADERS    = ["Warehouse", "Model", "Part Number", "Part Name",
                 "Avg Monthly Demand", "Current Stock", "Required Stock", "Gap", "Commitment For"]
WR_COL_WIDTHS = [26, 28, 18, 50, 16, 13, 13, 8, 16]
for c, (h, w) in enumerate(zip(WR_HEADERS, WR_COL_WIDTHS), 1):
    hdr(ws1.cell(HROW, c), h)
    ws1.column_dimensions[get_column_letter(c)].width = w
ws1.row_dimensions[HROW].height = 30

# Build union of demand combos + stock combos
demand_combos = set(zip(demand["item_code"], demand["warehouse"]))
stock_spare   = stock_df[
    stock_df["warehouse"].str.startswith("Engg -") &
    (stock_df["actual_qty"] > 0)
]
all_combos    = demand_combos | set(zip(stock_spare["item_code"], stock_spare["warehouse"]))
demand_lookup = {(r["item_code"], r["warehouse"]): r for _, r in demand.iterrows()}

# breakdown item codes (for commitment filtering)
breakdown_codes = set(
    spares_all[spares_all["item_group"].str.contains("Breakdown", case=False, na=False)]["item_code"]
)

wr_rows = []
for (ic, wh) in all_combos:
    curr = stock_lookup.get((ic, wh), 0)
    d    = demand_lookup.get((ic, wh), None)

    if d is not None:
        monthly_avg = round(d["monthly_avg"], 3)
        req         = int(d["reorder_qty"])
    else:
        monthly_avg = 0
        req         = MIN_REORDER_QTY if curr > 0 else 0

    warehouse_coverage = get_commitment(wh)

    if ic in breakdown_codes:
        if warehouse_coverage == "None" and round(monthly_avg) == 0:
            continue

    gap   = round(curr) - req
    grp   = item_group_lookup.get(ic, "")
    model = grp.split(" - Spares")[0].strip() if " - Spares" in str(grp) else grp
    wr_rows.append({
        "Warehouse":          wh,
        "Model":              model,
        "Part Number":        ic,
        "Part Name":          item_name_lookup.get(ic, ""),
        "Avg Monthly Demand": round(monthly_avg),
        "Current Stock":      round(curr),
        "Required Stock":     req,
        "Gap":                gap,
        "Commitment For":     get_commitment(wh),
    })

wr_df = pd.DataFrame(wr_rows).sort_values(["Warehouse", "Gap"])

for _, row in wr_df.iterrows():
    curr = row["Current Stock"]
    req  = row["Required Stock"]
    gap  = row["Gap"]
    if   gap < 0:                  fill = RED_FILL
    elif gap > 0:                  fill = YELLOW_FILL
    elif gap == 0 and curr > 0:    fill = GREEN_FILL
    elif curr == 0 and req == 0:   fill = GREY_FILL
    else:                          fill = YELLOW_FILL
    ws1.append(list(row))
    rn = ws1.max_row
    for c in range(1, 10):
        ws1.cell(rn, c).fill      = fill
        ws1.cell(rn, c).border    = BORDER
        ws1.cell(rn, c).alignment = LEFT

ws1.auto_filter.ref = f"A{HROW}:I{ws1.max_row}"
ws1.freeze_panes    = f"A{HROW + 1}"
print(f"  Warehouse Report: {len(wr_df)} rows | {wr_df['Warehouse'].nunique()} warehouses")


# ════════════════════════════════════════════════════════════════════════════════
# HELPER: build Accessories sheet
# ════════════════════════════════════════════════════════════════════════════════
def build_accessories_sheet(ws, group_items):
    demand_by_item = demand.groupby("item_code")["total_consumption"].sum().to_dict()

    SPARE_LEGEND = [
        ("Consumed, Zero Stock", "Item has demand but no stock at this location", RED_FILL),
        ("Has Stock",            "Item has stock at this location",               GREEN_FILL),
        ("No Demand, No Stock",  "No consumption and no stock at this location",  YELLOW_FILL),
    ]
    for i, (label, desc, fill) in enumerate(SPARE_LEGEND, 1):
        for c in [8, 9]:
            ws.cell(i, c).fill      = fill
            ws.cell(i, c).border    = BORDER
            ws.cell(i, c).alignment = LEFT
        ws.cell(i, 8).value = label
        ws.cell(i, 8).font  = Font(bold=True, size=9)
        ws.cell(i, 9).value = desc
        ws.cell(i, 9).font  = Font(size=9)
    ws.column_dimensions[get_column_letter(8)].width = 28
    ws.column_dimensions[get_column_letter(9)].width = 50

    HROW_L = 5
    SP_HEADERS    = ["Warehouse", "Model", "Part Number", "Part Name", "12M Consumption", "Current Stock"]
    SP_COL_WIDTHS = [26, 35, 18, 50, 14, 13]
    for c, (h, w) in enumerate(zip(SP_HEADERS, SP_COL_WIDTHS), 1):
        hdr(ws.cell(HROW_L, c), h)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[HROW_L].height = 30

    df = group_items.copy()
    df["consumption"] = df["item_code"].map(demand_by_item).fillna(0)
    df = df.sort_values(["item_group", "consumption"], ascending=[True, False])

    for _, item in df.iterrows():
        ic    = item["item_code"]
        name  = item.get("item_name", "") or ""
        grp   = item["item_group"]
        model = grp.split(" - Spares")[0].strip() if " - Spares" in str(grp) else grp
        cons  = round(item["consumption"], 1)
        for loc in locations:
            wh    = f"Engg - {loc} - EIPL"
            stock = stock_lookup.get((ic, wh), 0)
            if   cons > 0 and stock == 0: fill = RED_FILL
            elif stock > 0:               fill = GREEN_FILL
            else:                         continue
            ws.append([wh, model, ic, name, cons, round(stock)])
            rn = ws.max_row
            for c in range(1, 7):
                ws.cell(rn, c).fill      = fill
                ws.cell(rn, c).border    = BORDER
                ws.cell(rn, c).alignment = LEFT

    ws.auto_filter.ref = f"A{HROW_L}:F{ws.max_row}"
    ws.freeze_panes    = f"A{HROW_L + 1}"
    return len(df)


# ════════════════════════════════════════════════════════════════════════════════
# HELPER: build Breakdown sheet
# ════════════════════════════════════════════════════════════════════════════════
def build_breakdown_sheet(ws, group_items):
    demand_lookup_local = {(r["item_code"], r["warehouse"]): r for _, r in demand.iterrows()}
    stock_lookup_local  = {(r["item_code"], r["warehouse"]): r["actual_qty"] for _, r in stock_df.iterrows()}

    SPARE_LEGEND = [
        ("Below Required Stock", "Stock is less than required",           RED_FILL),
        ("Above Required Stock", "Stock is more than required",           YELLOW_FILL),
        ("Exact Required Stock", "Stock exactly matches requirement",     GREEN_FILL),
    ]
    for i, (label, desc, fill) in enumerate(SPARE_LEGEND, 1):
        for c in [8, 9]:
            ws.cell(i, c).fill      = fill
            ws.cell(i, c).border    = BORDER
            ws.cell(i, c).alignment = LEFT
        ws.cell(i, 8).value = label
        ws.cell(i, 8).font  = Font(bold=True, size=9)
        ws.cell(i, 9).value = desc
        ws.cell(i, 9).font  = Font(size=9)
    ws.column_dimensions[get_column_letter(8)].width = 38
    ws.column_dimensions[get_column_letter(9)].width = 55

    HROW_L = 5
    SP_HEADERS    = ["Warehouse", "Model", "Part Number", "Part Name", "12M Consumption", "Current Stock"]
    SP_COL_WIDTHS = [26, 35, 18, 50, 14, 13]
    for c, (h, w) in enumerate(zip(SP_HEADERS, SP_COL_WIDTHS), 1):
        hdr(ws.cell(HROW_L, c), h)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[HROW_L].height = 30

    df = group_items.copy().sort_values(["item_group"], ascending=[True])

    for _, item in df.iterrows():
        ic    = item["item_code"]
        name  = item.get("item_name", "") or ""
        grp   = item["item_group"]
        model = grp.split(" - Spares")[0].strip() if " - Spares" in str(grp) else grp

        for loc in locations:
            wh = f"Engg - {loc} - EIPL"
            d  = demand_lookup_local.get((ic, wh), None)
            if d is not None:
                req  = int(d["reorder_qty"])
                cons = d["total_consumption"]
            else:
                req  = 0
                cons = 0
            stock = stock_lookup_local.get((ic, wh), 0)
            gap   = stock - req

            if   gap < 0:                fill = RED_FILL
            elif gap > 0:                fill = YELLOW_FILL
            elif gap == 0 and stock > 0: fill = GREEN_FILL
            else:                        continue

            ws.append([wh, model, ic, name, round(cons, 1), round(stock)])
            rn = ws.max_row
            for c in range(1, 7):
                ws.cell(rn, c).fill      = fill
                ws.cell(rn, c).border    = BORDER
                ws.cell(rn, c).alignment = LEFT

    ws.auto_filter.ref = f"A{HROW_L}:F{ws.max_row}"
    ws.freeze_panes    = f"A{HROW_L + 1}"
    return ws.max_row - HROW_L


# ── Sheet 2: Accessories ──────────────────────────────────────────────────────
print("  Building Accessories sheet...")
ws2      = wb_acc.active
ws2.title = "Accessories"
acc_items = spares_all[spares_all["item_group"].str.contains("Accessories", case=False, na=False)].copy()
n_acc     = build_accessories_sheet(ws2, acc_items)
print(f"  Accessories: {ws2.max_row - 5} rows")

# ── Sheet 3: Breakdown ────────────────────────────────────────────────────────
print("  Building Breakdown sheet...")
ws3      = wb_main.create_sheet("Breakdown")
bkd_items = spares_all[spares_all["item_group"].str.contains("Breakdown", case=False, na=False)].copy()
n_bkd    = build_breakdown_sheet(ws3, bkd_items)
print(f"  Breakdown:   {ws3.max_row - 5} rows")


# ═══════════════════════════════════════════════════════════════════════════════
# REGION-WISE FILTERING & EMAIL DISTRIBUTION
# ═══════════════════════════════════════════════════════════════════════════════

def filter_df_by_region(df, region):
    if region is None:
        return df
    region_warehouses = [wh for wh, r in WAREHOUSE_REGION_MAP.items() if r == region]
    return df[df["Warehouse"].isin(region_warehouses)].copy()


def filter_accessories_by_region(df, region):
    if region is None:
        return df
    region_warehouses = [wh for wh, r in WAREHOUSE_REGION_MAP.items() if r == region]
    return df[df["Warehouse"].isin(region_warehouses)].copy()


def build_filtered_workbook(wr_df, region=None, include_accessories=False):
    wb = openpyxl.Workbook()

    filtered_wr = filter_df_by_region(wr_df, region)

    # ── Sheet 1: Warehouse Report ─────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Warehouse Report"

    for i, (label, desc, fill) in enumerate([
        ("Negative Gap",              "Stock BELOW required — replenishment needed", RED_FILL),
        ("Positive Gap",              "Stock ABOVE required — excess inventory",      YELLOW_FILL),
        ("Zero Gap",                  "Stock matches required — optimal",             GREEN_FILL),
        ("No Stock / No Requirement", "No consumption and no current stock",          GREY_FILL),
    ], 1):
        for c in [11, 12]:
            ws1.cell(i, c).fill      = fill
            ws1.cell(i, c).border    = BORDER
            ws1.cell(i, c).alignment = LEFT
        ws1.cell(i, 11).value = label
        ws1.cell(i, 11).font  = Font(bold=True, size=9)
        ws1.cell(i, 12).value = desc
        ws1.cell(i, 12).font  = Font(size=9)
    ws1.column_dimensions[get_column_letter(11)].width = 30
    ws1.column_dimensions[get_column_letter(12)].width = 50

    HROW = 5
    for c, (h, w) in enumerate(zip(
        ["Warehouse", "Model", "Part Number", "Part Name",
         "Avg Monthly Demand", "Current Stock", "Required Stock", "Gap", "Commitment For"],
        [26, 28, 18, 50, 16, 13, 13, 8, 16]
    ), 1):
        hdr(ws1.cell(HROW, c), h)
        ws1.column_dimensions[get_column_letter(c)].width = w
    ws1.row_dimensions[HROW].height = 30

    for _, row in filtered_wr.iterrows():
        curr = row["Current Stock"]
        req  = row["Required Stock"]
        gap  = row["Gap"]
        if   gap < 0:                fill = RED_FILL
        elif gap > 0:                fill = YELLOW_FILL
        elif gap == 0 and curr > 0:  fill = GREEN_FILL
        elif curr == 0 and req == 0: fill = GREY_FILL
        else:                        fill = YELLOW_FILL
        ws1.append(list(row))
        rn = ws1.max_row
        for c in range(1, 10):
            ws1.cell(rn, c).fill      = fill
            ws1.cell(rn, c).border    = BORDER
            ws1.cell(rn, c).alignment = LEFT

    ws1.auto_filter.ref = (f"A{HROW}:I{ws1.max_row}" if ws1.max_row > HROW else f"A{HROW}:I{HROW}")
    ws1.freeze_panes    = f"A{HROW + 1}"

    # ── Sheet 2: Accessories (optional) ──────────────────────────────────────
    if include_accessories:
        ws2 = wb.create_sheet("Accessories")
        demand_by_item = demand.groupby("item_code")["total_consumption"].sum().to_dict()

        for i, (label, desc, fill) in enumerate([
            ("Consumed, Zero Stock", "Item has demand but no stock at this location", RED_FILL),
            ("Has Stock",            "Item has stock at this location",               GREEN_FILL),
            ("No Demand, No Stock",  "No consumption and no stock at this location",  YELLOW_FILL),
        ], 1):
            for c in [8, 9]:
                ws2.cell(i, c).fill      = fill
                ws2.cell(i, c).border    = BORDER
                ws2.cell(i, c).alignment = LEFT
            ws2.cell(i, 8).value = label
            ws2.cell(i, 8).font  = Font(bold=True, size=9)
            ws2.cell(i, 9).value = desc
            ws2.cell(i, 9).font  = Font(size=9)
        ws2.column_dimensions[get_column_letter(8)].width = 28
        ws2.column_dimensions[get_column_letter(9)].width = 50

        HROW_ACC = 5
        for c, (h, w) in enumerate(zip(
            ["Warehouse", "Model", "Part Number", "Part Name", "12M Consumption", "Current Stock"],
            [26, 35, 18, 50, 14, 13]
        ), 1):
            hdr(ws2.cell(HROW_ACC, c), h)
            ws2.column_dimensions[get_column_letter(c)].width = w
        ws2.row_dimensions[HROW_ACC].height = 30

        df_acc = acc_items.copy()
        df_acc["consumption"] = df_acc["item_code"].map(demand_by_item).fillna(0)
        df_acc = df_acc.sort_values(["item_group", "consumption"], ascending=[True, False])

        for _, item in df_acc.iterrows():
            ic    = item["item_code"]
            name  = item.get("item_name", "") or ""
            grp   = item["item_group"]
            model = grp.split(" - Spares")[0].strip() if " - Spares" in str(grp) else grp
            cons  = round(item["consumption"], 1)
            for loc in locations:
                wh = f"Engg - {loc} - EIPL"
                if region and WAREHOUSE_REGION_MAP.get(wh) != region:
                    continue
                stock = stock_lookup.get((ic, wh), 0)
                if   cons > 0 and stock == 0: fill = RED_FILL
                elif stock > 0:               fill = GREEN_FILL
                else:                         continue
                ws2.append([wh, model, ic, name, cons, round(stock)])
                rn = ws2.max_row
                for c in range(1, 7):
                    ws2.cell(rn, c).fill      = fill
                    ws2.cell(rn, c).border    = BORDER
                    ws2.cell(rn, c).alignment = LEFT

        ws2.auto_filter.ref = (f"A{HROW_ACC}:F{ws2.max_row}" if ws2.max_row > HROW_ACC else f"A{HROW_ACC}:F{HROW_ACC}")
        ws2.freeze_panes    = f"A{HROW_ACC + 1}"

    # ── Sheet 3: Breakdown ────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Breakdown")
    demand_lookup_local = {(r["item_code"], r["warehouse"]): r for _, r in demand.iterrows()}
    stock_lookup_local  = {(r["item_code"], r["warehouse"]): r["actual_qty"] for _, r in stock_df.iterrows()}

    for i, (label, desc, fill) in enumerate([
        ("Below Required Stock", "Stock is less than required",       RED_FILL),
        ("Above Required Stock", "Stock is more than required",       YELLOW_FILL),
        ("Exact Required Stock", "Stock exactly matches requirement", GREEN_FILL),
    ], 1):
        for c in [8, 9]:
            ws3.cell(i, c).fill      = fill
            ws3.cell(i, c).border    = BORDER
            ws3.cell(i, c).alignment = LEFT
        ws3.cell(i, 8).value = label
        ws3.cell(i, 8).font  = Font(bold=True, size=9)
        ws3.cell(i, 9).value = desc
        ws3.cell(i, 9).font  = Font(size=9)
    ws3.column_dimensions[get_column_letter(8)].width = 38
    ws3.column_dimensions[get_column_letter(9)].width = 55

    HROW = 5
    for c, (h, w) in enumerate(zip(
        ["Warehouse", "Model", "Part Number", "Part Name", "12M Consumption", "Current Stock"],
        [26, 35, 18, 50, 14, 13]
    ), 1):
        hdr(ws3.cell(HROW, c), h)
        ws3.column_dimensions[get_column_letter(c)].width = w
    ws3.row_dimensions[HROW].height = 30

    df_bkd = bkd_items.copy().sort_values(["item_group"], ascending=[True])

    for _, item in df_bkd.iterrows():
        ic    = item["item_code"]
        name  = item.get("item_name", "") or ""
        grp   = item["item_group"]
        model = grp.split(" - Spares")[0].strip() if " - Spares" in str(grp) else grp

        for loc in locations:
            wh = f"Engg - {loc} - EIPL"
            if region and WAREHOUSE_REGION_MAP.get(wh) != region:
                continue
            d = demand_lookup_local.get((ic, wh), None)
            if d is not None:
                req  = int(d["reorder_qty"])
                cons = d["total_consumption"]
            else:
                req  = 0
                cons = 0
            stock = stock_lookup_local.get((ic, wh), 0)
            gap   = stock - req

            if   gap < 0:                fill = RED_FILL
            elif gap > 0:                fill = YELLOW_FILL
            elif gap == 0 and stock > 0: fill = GREEN_FILL
            else:                        continue

            ws3.append([wh, model, ic, name, round(cons, 1), round(stock)])
            rn = ws3.max_row
            for c in range(1, 7):
                ws3.cell(rn, c).fill      = fill
                ws3.cell(rn, c).border    = BORDER
                ws3.cell(rn, c).alignment = LEFT

    ws3.auto_filter.ref = (f"A{HROW}:F{ws3.max_row}" if ws3.max_row > HROW else f"A{HROW}:F{HROW}")
    ws3.freeze_panes    = f"A{HROW + 1}"

    return wb


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE FILES & SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
main_file = OUTPUT_FILE
acc_file  = f"Accessories_Report_{date.today().strftime('%Y_%m')}.xlsx"

wb_main.save(main_file)
wb_acc.save(acc_file)

neg        = (wr_df["Gap"] < 0).sum()
pos        = (wr_df["Gap"] > 0).sum()
zero       = (wr_df["Gap"] == 0).sum()
total_rows = len(wr_df)

