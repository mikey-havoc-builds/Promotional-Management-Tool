"""Load an MM_Promo_Tool.xlsx workbook into PostgreSQL.

    python -m app.importer excel/MM_Promo_Tool.xlsx [--reset]

Maps: Product Master, Hotspots, Location Details, Replen Settings, the Sales/Inv/Forecast sheets,
Lists, Alerts and the Database (promotions) sheet.
"""
import re
import sys
import datetime as dt
import openpyxl
from .db import connect

MONTHS = {m: i + 1 for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split())}


def month_from_header(h):
    """'Oct 25 Sales Qty' / 'Nov-26' / 'Sept-27' -> date(2025-10-01) etc."""
    m = re.match(r"\s*([A-Za-z]{3})[A-Za-z]*[\s\-]+(\d{2,4})", str(h or ""))
    if not m or m.group(1).lower() not in MONTHS:
        return None
    y = int(m.group(2))
    y += 2000 if y < 100 else 0
    return dt.date(y, MONTHS[m.group(1).lower()], 1)


def num(v):
    if isinstance(v, bool) or v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    return None


def txt(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if s and not s.startswith("#") else None


def table_rows(ws, header_row, first_col=1):
    """Yield (headers, row values) for a sheet whose headers are in header_row."""
    hdr = [c.value for c in ws[header_row]]
    for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if any(v is not None for v in r):
            yield hdr, r


def find_header(ws, label):
    for r in range(1, 8):
        for c in ws[r]:
            if c.value == label:
                return r
    raise ValueError(f"header {label!r} not found in {ws.title}")


def load(path, reset=False):
    wb = openpyxl.load_workbook(path, data_only=True)
    with connect() as con, con.cursor() as cur:
        if reset:
            cur.execute("""TRUNCATE promotions, sales_monthly, inventory_monthly, forecast_monthly, location_stock,
                           replen_settings, hotspots, locations, product_master, lists, alert_definitions RESTART IDENTITY CASCADE""")
        # --- product master
        ws = wb["Product Master"]
        hr = find_header(ws, "Product Code2")
        hdr = [c.value for c in ws[hr]]
        ix = {h: i for i, h in enumerate(hdr) if h}
        for r in ws.iter_rows(min_row=hr + 1, values_only=True):
            code = txt(r[ix["Product Code2"]])
            if not code:
                continue
            g = lambda h: r[ix[h]] if h in ix else None
            cur.execute("""INSERT INTO product_master VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON CONFLICT (product_code) DO UPDATE SET description=EXCLUDED.description,l1=EXCLUDED.l1,l2=EXCLUDED.l2,
                l3=EXCLUDED.l3,vendor=EXCLUDED.vendor,product_owner=EXCLUDED.product_owner,primary_sector=EXCLUDED.primary_sector,
                classification=EXCLUDED.classification,national_price=EXCLUDED.national_price,cost_price=EXCLUDED.cost_price,
                moq=EXCLUDED.moq,lifecycle_state=EXCLUDED.lifecycle_state""",
                        (code, txt(g("Description")), txt(g("L1")), txt(g("L2")), txt(g("L3")), txt(g("Vendor")),
                         txt(g("Product Owner")), txt(g("Primary Sector")), txt(g("Classification")),
                         num(g("National Price")), num(g("Cost Price")), num(g("MOQ")), txt(g("Lifecycle State"))))

        def ensure_loc(code, wtype=None, name=None, region=None):
            code = txt(code)
            if not code:
                return None
            cur.execute("""INSERT INTO locations(code,name,region,warehouse_type) VALUES (%s,%s,%s,COALESCE(%s::text,'Retail'))
                ON CONFLICT (code) DO UPDATE SET name=COALESCE(EXCLUDED.name,locations.name),
                region=COALESCE(EXCLUDED.region,locations.region),
                warehouse_type=COALESCE(%s::text, locations.warehouse_type)""",
                        (code, name, region, wtype, wtype))
            return code

        def ensure_prod(code):
            cur.execute("INSERT INTO product_master(product_code) VALUES (%s) ON CONFLICT DO NOTHING", (code,))

        # --- hotspots
        for r in wb["Hotspots"].iter_rows(min_row=2, values_only=True):
            loc = ensure_loc(r[2], name=txt(r[1]), region=txt(r[0]))
            if loc and txt(r[3]) in ("Small", "Medium", "Large"):
                cur.execute("UPDATE locations SET branch_size=%s WHERE code=%s", (txt(r[3]), loc))
            if loc and txt(r[4]):
                cur.execute("""INSERT INTO hotspots VALUES (%s,%s,%s) ON CONFLICT (location_code,hotspot_name)
                               DO UPDATE SET product_spaces=EXCLUDED.product_spaces""", (loc, txt(r[4]), num(r[5]) or 0))

        # --- history sheets (month taken from the column header text, e.g. 'Oct 25 Sales Qty')
        def history(sheet, code_col, loc_col, type_col, kind_cols):
            ws = wb[sheet]
            hr = find_header(ws, "Product Code")
            hdr = [c.value for c in ws[hr]]
            p, l = hdr.index("Product Code"), hdr.index("Warehouse Name")
            t = hdr.index("Warehouse Type") if "Warehouse Type" in hdr else None
            out = {}
            for r in ws.iter_rows(min_row=hr + 1, values_only=True):
                prod, loc = txt(r[p]), txt(r[l])
                if not prod or not loc:
                    continue
                ensure_prod(prod)
                ensure_loc(loc, txt(r[t]) if t is not None else None)
                for ci, h in enumerate(hdr):
                    m = month_from_header(h)
                    if m and num(r[ci]) is not None:
                        out[(prod, loc, m)] = num(r[ci])
            return out

        def merge(sheets):
            keys = {}
            for field, sheet in sheets.items():
                for k, v in history(sheet, None, None, None, None).items():
                    keys.setdefault(k, {})[field] = v
            return keys

        for (prod, loc, m), d in merge({"qty": "Sales Qty", "revenue": "Sales Revenue", "gm": "Sales GM",
                                       "transactions": "Sales Txns", "customers": "Sales Customers"}).items():
            cur.execute("INSERT INTO sales_monthly VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                        (prod, loc, m, d.get("qty"), d.get("revenue"), d.get("gm"), d.get("transactions"), d.get("customers")))
        for (prod, loc, m), d in merge({"avg_qty": "Inv Qty", "avg_value": "Inv Value", "osa_days": "OSA-Days"}).items():
            cur.execute("INSERT INTO inventory_monthly VALUES (%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                        (prod, loc, m, d.get("avg_qty"), d.get("avg_value"), d.get("osa_days")))
        for (prod, loc, m), d in merge({"qty": "Retained Forecast Qty", "revenue": "Retained Fcst Revenue",
                                       "gm": "Retained Forecast GM"}).items():
            cur.execute("INSERT INTO forecast_monthly VALUES ('retained',%s,%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                        (prod, loc, m, d.get("qty"), d.get("revenue"), d.get("gm")))
        for (prod, loc, m), d in merge({"qty": "Current Forecast Qty"}).items():
            cur.execute("INSERT INTO forecast_monthly(kind,product_code,location_code,month,qty) VALUES ('current',%s,%s,%s,%s) ON CONFLICT DO NOTHING",
                        (prod, loc, m, d["qty"]))

        # --- replen settings and current stock
        ws = wb["Replen Settings"]
        hr = find_header(ws, "Product Code")
        for _, r in table_rows(ws, hr):
            prod, loc = txt(r[2]), txt(r[3])
            if prod and loc:
                ensure_prod(prod); ensure_loc(loc)
                cur.execute("""INSERT INTO replen_settings VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (product_code,location_code) DO UPDATE SET range_type=EXCLUDED.range_type,
                    replen_method=EXCLUDED.replen_method,stopped_for_purchase=EXCLUDED.stopped_for_purchase,
                    open_for_replen=EXCLUDED.open_for_replen,planning_floor=EXCLUDED.planning_floor,
                    main_warehouse=EXCLUDED.main_warehouse,supply_tier=EXCLUDED.supply_tier""",
                            (prod, loc, txt(r[4]), txt(r[5]), txt(r[6]), txt(r[7]), num(r[8]), txt(r[9]), txt(r[10])))
        ws = wb["Location Details"]
        hr = find_header(ws, "Product Code")
        for _, r in table_rows(ws, hr):
            prod, loc = txt(r[2]), txt(r[3])
            if prod and loc and (num(r[4]) is not None or num(r[5]) is not None):
                ensure_prod(prod); ensure_loc(loc, region=txt(r[7]))
                cur.execute("""INSERT INTO location_stock VALUES (%s,%s,%s,%s)
                    ON CONFLICT (product_code,location_code) DO UPDATE SET on_hand_qty=EXCLUDED.on_hand_qty,on_hand_value=EXCLUDED.on_hand_value""",
                            (prod, loc, num(r[4]), num(r[5])))

        # --- lists (header cells are in row 1 for the column groups on the Lists sheet)
        ws = wb["Lists"]
        for hdr_text, name in (("Promo Type", "promo_type"), ("Intent of Promotion", "intent"), ("Hotspot", "hotspot"),
                               ("Category Team", "owner")):
            for row in ws.iter_rows():
                hit = next((c for c in row if c.value == hdr_text), None)
                if hit:
                    i = 0
                    for r in range(hit.row + 1, ws.max_row + 1):
                        v = ws.cell(r, hit.column).value
                        if v is None or txt(v) is None:
                            break
                        if txt(v) != "All":
                            cur.execute("INSERT INTO lists VALUES (%s,%s,%s) ON CONFLICT DO NOTHING", (name, txt(v), i)); i += 1
                    break
        # market-in list: not a clean block in the workbook, seed the defaults used by the form
        for i, v in enumerate(["Mailer", "EDM", "Farmlands Pro", "In-store", "Social"]):
            cur.execute("INSERT INTO lists VALUES ('market_in',%s,%s) ON CONFLICT DO NOTHING", (v, i))

        # --- alerts
        stage = None
        for r in wb["Alerts"].iter_rows(min_row=1, values_only=True):
            name, desc = txt(r[1]), txt(r[2])
            if name and name.lower().endswith("alerts / warnings") or name in ("Mid Promo Alerts",):
                stage = "pre" if name.startswith("Pre") else "mid" if name.startswith("Mid") else "post"
            elif name and stage:
                cur.execute("INSERT INTO alert_definitions(stage,name,description) VALUES (%s,%s,%s) ON CONFLICT DO NOTHING", (stage, name, desc))

        # --- promotions (Database sheet): header row 5, data from row 6, columns C..BL
        ws = wb["Database"]
        cols = ["promo_month", None, "product_code", "product_name", "l1", "l2", "l3", "vendor", "classification",
                "product_owner", "national_price", "vendor_cost", "moq", "ranged_locations", "in_cdc", "promo_type",
                "promo_intent", "hotspot", "promo_cost_override", "uplift_pct", "promo_qty_input", "bdf_applied",
                "pres_qty_small", "pres_qty_medium", "pres_qty_large", "market_in_1", "market_in_2", "market_in_3",
                "discount_pct", "discount_amt", "promo_price", "x4y_qty_1", "x4y_qty_2", "x4y_qty_3", "x4y_amt_1",
                "x4y_amt_2", "x4y_amt_3", "xy_qty_1", "xy_qty_2", "xy_qty_3", "xy_pct_1", "xy_pct_2", "xy_pct_3",
                "ly_qty", "ly_rev", "ly_gm", None, "promo_qty", "promo_rev", "promo_gm", None, None, None, None, None,
                "ly_soh_qty_retail", "promo_soh_qty_retail", "promo_soh_val_retail", "promo_soh_qty_cdc", "promo_soh_val_cdc"]
        textcols = {"product_code", "product_name", "l1", "l2", "l3", "vendor", "classification", "product_owner", "in_cdc",
                    "promo_type", "promo_intent", "hotspot", "market_in_1", "market_in_2", "market_in_3"}
        n = 0
        for r in ws.iter_rows(min_row=6, min_col=3, values_only=True):
            rec = {c: v for c, v in zip(cols, r) if c}
            if not txt(rec.get("product_code")) or not isinstance(rec.get("promo_month"), (dt.date, dt.datetime)):
                continue
            ensure_prod(txt(rec["product_code"]))
            vals = {k: (txt(v) if k in textcols else num(v)) for k, v in rec.items() if k != "promo_month"}
            vals["promo_month"] = rec["promo_month"].date().replace(day=1) if isinstance(rec["promo_month"], dt.datetime) else rec["promo_month"]
            if not (vals.get("promo_type") and vals.get("promo_intent") and vals.get("hotspot")):
                continue    # incomplete rows (e.g. half-filled sample row) are not importable
            ks = list(vals)
            cur.execute(f"INSERT INTO promotions ({','.join(ks)}) VALUES ({','.join(['%s']*len(ks))}) ON CONFLICT (promo_month,product_code) DO NOTHING",
                        [vals[k] for k in ks]); n += 1
        con.commit()
        cur.execute("""SELECT (SELECT count(*) FROM product_master) products,(SELECT count(*) FROM locations) locations,
              (SELECT count(*) FROM sales_monthly) sales,(SELECT count(*) FROM inventory_monthly) inv,
              (SELECT count(*) FROM forecast_monthly) fcst,(SELECT count(*) FROM promotions) promos""")
        return cur.fetchone()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    print(load(args[0] if args else "excel/MM_Promo_Tool.xlsx", reset="--reset" in sys.argv))
