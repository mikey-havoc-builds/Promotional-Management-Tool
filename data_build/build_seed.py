"""Build the compressed seed data (data/seed/*.csv.gz) from the source extracts.

Source files (all in one folder, e.g. PromoMgmt/Data_Tables):
    BIS.xlsb                      current stock / range / replen snapshot, one row per product variant x warehouse
    SNO_Qty / SNO_Revenue / SNO_GM / SNO_Cust .xlsb    monthly sales pivots by product x region
    Promo_Hist.xlsx               the promotions "Database" sheet

Usage:  python data_build/build_seed.py --src <Data_Tables folder> [--out data/seed] [--cache /tmp/xlsb_cache]
The big .xlsb files are streamed to gzip CSV first (cached in --cache so re-runs are fast).
"""
import argparse, csv, gzip, pathlib, re, sys
import pandas as pd

CDC_CODES = {"National 3PL DC Hamilton": "155-4008",       # North Island CDC
             "National 3PL DC Christchurch": "155-4007"}   # South Island CDC
RANGE_RANK = {"Core Range": 4, "Local Range": 3, "Planned Range": 2, "Not Ranged": 1}


def xlsb_to_csv(src, dst):
    from pyxlsb import open_workbook
    with open_workbook(str(src)) as wb, wb.get_sheet(1) as sh, gzip.open(dst, "wt", newline="") as f:
        w = csv.writer(f)
        for row in sh.rows(sparse=False):
            w.writerow([c.v for c in row])


def cached_csv(src, cache):
    dst = pathlib.Path(cache) / (pathlib.Path(src).stem + ".csv.gz")
    if not dst.exists():
        print("converting", src, flush=True)
        xlsb_to_csv(src, dst)
    return dst


def clean_code(s):
    s = s.astype("string").str.strip()
    return s.str.replace(r"\.0$", "", regex=True)


def is_cdc(name, ltype):
    if name in CDC_CODES:
        return True
    if ltype != "CDC / Mill / Other":
        return False
    n = name or ""
    if re.search(r"Mill|Bulk|Blending|Offsite|Port\b|Online|Store|Trading|Procurement|Central Stock|Pallet|Fuel|Support|Courier|Inghams|Pure Oil|Izone DC|Haydn", n):
        return False
    return bool(re.search(r"\bDC\b|CDC|Hub|3PL", n))


def build_bis(src_csv, out):
    b = pd.read_csv(src_csv, dtype=str, low_memory=False)
    b = b.iloc[:, 1:]                                       # first column is a blank spacer
    b = b[b["Product Number"].notna() & b["Warehouse"].notna()]
    b["Product Number"] = clean_code(b["Product Number"])
    for c in ["AVAILABLE_ON_HAND_QUANTITY_EA", "AVAILABLE_ON_HAND_VALUE", "National Price", "WAC", "PLANNING_FLOOR",
              "Min Order Qty", "Preferred Vendor Replacement Cost"]:
        b[c] = pd.to_numeric(b[c], errors="coerce")
    # ---- locations: retail branches and the distribution centres; "DO NOT USE" and non-stocking sites dropped
    wh = b.drop_duplicates("Warehouse")[["Warehouse", "Location Type", "Region Name"]].copy()
    wh = wh[~wh["Warehouse"].str.match(r"^zz?[ _]?DO NOT USE|^zDO", na=False)]
    wh["warehouse_type"] = None
    wh.loc[wh["Location Type"] == "Retail", "warehouse_type"] = "Retail"
    wh.loc[[is_cdc(n, t) for n, t in zip(wh["Warehouse"], wh["Location Type"])], "warehouse_type"] = "CDC"
    wh = wh[wh["warehouse_type"].notna()]
    wh["code"] = wh["Warehouse"].map(lambda n: CDC_CODES.get(n, n))
    wh["name"] = wh["Warehouse"].str.replace(r"^Retail ", "", regex=True)
    wh["region"] = wh["Region Name"]
    # a CDC with no region of its own takes the region of the branches it supplies
    locs = wh[["code", "name", "region", "warehouse_type"]].sort_values("code")
    locs.to_csv(out / "locations.csv.gz", index=False)
    keep_wh = set(wh["Warehouse"])
    code_of = dict(zip(wh["Warehouse"], wh["code"]))
    b = b[b["Warehouse"].isin(keep_wh)].copy()
    b["location_code"] = b["Warehouse"].map(code_of)

    # ---- product master (one row per product number)
    def first(s):
        s = s.dropna()
        return s.iloc[0] if len(s) else None
    g = b.groupby("Product Number")
    pm = pd.DataFrame({
        "product_code": g.size().index,
        "description": g["Product Name"].agg(first).values,
        "l1": g["L1 Category Name"].agg(first).values, "l2": g["L2 Category Name"].agg(first).values,
        "l3": g["L3 Category Name"].agg(first).values,
        "vendor": g["Preferred Vendor"].agg(first).values,
        "national_price": g["National Price"].median().values,
        "cost_price": g["Preferred Vendor Replacement Cost"].median().fillna(g["WAC"].median()).values,
        "moq": g["Min Order Qty"].max().values,
        "variant_count": g["Product Variant Number"].nunique().values})
    pm["moq"] = pm["moq"].round().astype("Int64")
    pm.to_csv(out / "products_bis.csv.gz", index=False)

    # ---- replen settings + stock, one row per product x location (variants combined)
    b["rank"] = b["CORE_RANGE_TYPE"].map(RANGE_RANK).fillna(0)
    b["tier"] = b["DEMAND_TIER"]
    b["method"] = b["tier"].map(lambda t: None if pd.isna(t) else ("Kanban" if str(t).startswith("K") else "Forecast"))
    b["stopped"] = b["BRANCH_STOPPED_FOR_PURCHASE"]
    ranged_mask = b["rank"] > 1
    b = b[ranged_mask | (b["AVAILABLE_ON_HAND_QUANTITY_EA"].fillna(0) != 0)].copy()      # drop never-ranged, no-stock rows
    b = b.sort_values("rank", ascending=False)
    g = b.groupby(["Product Number", "location_code"], sort=False)
    rs = pd.DataFrame({
        "range_type": g["CORE_RANGE_TYPE"].first(),
        "replen_method": g["method"].first(),
        "stopped_for_purchase": g["stopped"].agg(lambda s: "Yes" if (s == "Yes").all() else "No"),
        "open_for_replen": g["OPEN_FOR_REPLENISHMENT"].first(),
        "planning_floor": g["PLANNING_FLOOR"].sum(min_count=1),
        "main_warehouse": g["MAIN_WAREHOUSE"].first(),
        "supply_tier": g["tier"].first(),
        "on_hand_qty": g["AVAILABLE_ON_HAND_QUANTITY_EA"].sum(min_count=1),
        "on_hand_value": g["AVAILABLE_ON_HAND_VALUE"].sum(min_count=1)}).reset_index()
    rs = rs.rename(columns={"Product Number": "product_code", "location_code": "location_code"})
    rs.loc[rs["range_type"] == "Not Ranged", ["range_type"]] = None
    rs.drop(columns=["on_hand_qty", "on_hand_value"]).to_csv(out / "replen_settings.csv.gz", index=False)
    st = rs[rs["on_hand_qty"].notna() & (rs["on_hand_qty"] != 0)][["product_code", "location_code", "on_hand_qty", "on_hand_value"]]
    st.to_csv(out / "location_stock.csv.gz", index=False)
    print("BIS: locations", len(locs), "products", len(pm), "replen rows", len(rs), "stock rows", len(st))
    return pm


def read_pivot(path, dim_cols, first_month_col):
    raw = pd.read_csv(path, dtype=str, header=None, low_memory=False)
    hdr = raw.index[raw[4] == "L1_CATEGORY_NAME"][0]
    months = {}
    for c in range(first_month_col, raw.shape[1]):
        v = pd.to_numeric(raw.iat[hdr, c], errors="coerce")
        if pd.notna(v):
            months[c] = (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v))).date()
    df = raw.iloc[hdr + 1:, 4:first_month_col].copy()
    df.columns = dim_cols
    vals = raw.iloc[hdr + 1:, list(months)].apply(pd.to_numeric, errors="coerce")
    vals.columns = list(months.values())
    df = pd.concat([df, vals], axis=1)
    df = df[df["region"].notna() & df["name"].notna() & ~df["l1"].fillna("").str.contains("Grand Total|Total", regex=True)]
    return df, list(months.values())


def melt(df, months, val):
    keys = [c for c in ("l1", "l2", "l3", "vendor", "region", "name", "product_code") if c in df.columns]
    m = df.melt(id_vars=keys, value_vars=months, var_name="month", value_name=val).dropna(subset=[val])
    return m


def build_sales(src_dir, cache, out, pm_bis):
    dims = ["l1", "l2", "l3", "vendor", "region", "name", "product_code"]
    parts = {}
    meta = []
    for val, f in (("qty", "SNO_Qty"), ("revenue", "SNO_Revenue"), ("gm", "SNO_GM")):
        df, months = read_pivot(cached_csv(pathlib.Path(src_dir) / f"{f}.xlsb", cache), dims, 11)
        df = df[df["product_code"].notna()]
        df["product_code"] = clean_code(df["product_code"])
        meta.append(df[["product_code", "name", "l1", "l2", "l3", "vendor"]].drop_duplicates("product_code"))
        m = melt(df, months, val)
        parts[val] = m.groupby(["product_code", "region", "month"], as_index=False)[val].sum()
        print(f, "rows", len(parts[val]), "months", min(months), "->", max(months))
    s = parts["qty"]
    for k in ("revenue", "gm"):
        s = s.merge(parts[k], on=["product_code", "region", "month"], how="outer")
    meta = pd.concat(meta).drop_duplicates("product_code")
    # customers: the extract has no product number, so match on product name + vendor + region
    cdims = ["l1", "l2", "l3", "vendor", "region", "name"]
    cdf, cmonths = read_pivot(cached_csv(pathlib.Path(src_dir) / "SNO_Cust.xlsb", cache), cdims, 10)
    cm = melt(cdf, cmonths, "customers").groupby(["vendor", "region", "name", "month"], as_index=False)["customers"].sum()
    key = meta.assign(n=meta.groupby(["vendor", "name"]).product_code.transform("count"))
    uniq = key[key.n == 1][["vendor", "name", "product_code"]]
    cm = cm.merge(uniq, on=["vendor", "name"], how="left")
    print("customer rows matched to a product:", int(cm.product_code.notna().sum()), "of", len(cm))
    cm = cm.dropna(subset=["product_code"]).groupby(["product_code", "region", "month"], as_index=False)["customers"].sum()
    s = s.merge(cm, on=["product_code", "region", "month"], how="left")
    s = s[(s[["qty", "revenue", "gm"]].fillna(0) != 0).any(axis=1)]
    s["customers"] = s["customers"].round().astype("Int64")
    s["qty"] = s["qty"].round(2); s["revenue"] = s["revenue"].round(2); s["gm"] = s["gm"].round(2)
    s.to_csv(out / "sales_monthly.csv.gz", index=False)
    meta.rename(columns={"name": "description"}).to_csv(out / "products_sales.csv.gz", index=False)
    print("sales rows", len(s))


def build_hist(src, out):
    raw = pd.read_excel(src, sheet_name="Database", header=None)
    hr = next(i for i in range(min(15, len(raw))) if (raw.iloc[i] == "Promo Month").any())
    h = raw.iloc[hr + 1:].copy()
    h.columns = [str(c).strip() for c in raw.iloc[hr]]
    h = h.dropna(subset=["Promo Month", "Product Code"])
    cols = {"Promo Month": "promo_month", "Product Code": "product_code", "Product Name": "product_name", "L1": "l1", "L2": "l2",
            "L3": "l3", "Vendor": "vendor", "Classification": "classification", "Product Owner": "product_owner",
            "National Price excl GST": "national_price", "Vendor replacement cost": "vendor_cost", "MOQ": "moq",
            "Ranged Locations": "ranged_locations", "In CDC": "in_cdc", "Promotional Type": "promo_type",
            "Promotional Intent": "promo_intent", "Hotspot": "hotspot", "Promotional Cost": "promo_cost_override",
            "Uplift %": "uplift_pct", "Uplift Qty": "promo_qty_input", "BDF Applied": "bdf_applied",
            "Presentation Qty SM": "pres_qty_small", "Presentation Qty MD": "pres_qty_medium", "Presentation Qty LG": "pres_qty_large",
            "Market in 1": "market_in_1", "Market in 2": "market_in_2", "Market in 3": "market_in_3",
            "Discount %": "discount_pct", "Discount $": "discount_amt", "Promotional Price": "promo_price",
            "x4y Qty 1": "x4y_qty_1", "x4y Qty 2": "x4y_qty_2", "x4y Qty 3": "x4y_qty_3",
            "x4y $ 1": "x4y_amt_1", "x4y $ 2": "x4y_amt_2", "x4y $ 3": "x4y_amt_3",
            "x&y Qty 1": "xy_qty_1", "x&y Qty 2": "xy_qty_2", "x&y Qty 3": "xy_qty_3",
            "x&y % off 1": "xy_pct_1", "x&y % off 2": "xy_pct_2", "x&y % off 3": "xy_pct_3",
            "LY Sls Qty": "ly_qty", "LY Sls Rev": "ly_rev", "LY Sls GM": "ly_gm",
            "Promo Sls qty": "promo_qty", "Promo Sls Rev": "promo_rev", "Promo Sls GM": "promo_gm",
            "LY SOH Qty Retail": "ly_soh_qty_retail", "Promo SOH Qty Retail": "promo_soh_qty_retail",
            "Promo SOH $ Retail": "promo_soh_val_retail", "Promo SOH Qty CDC": "promo_soh_qty_cdc", "Promo SOH $ CDC": "promo_soh_val_cdc"}
    h = h[list(cols)].rename(columns=cols)
    h["promo_month"] = pd.to_datetime(h["promo_month"]).dt.to_period("M").dt.to_timestamp().dt.date
    h["product_code"] = clean_code(h["product_code"])
    for c in ("moq", "ranged_locations", "pres_qty_small", "pres_qty_medium", "pres_qty_large"):
        h[c] = pd.to_numeric(h[c], errors="coerce").round().astype("Int64")
    h["source"] = "history"
    h = h.sort_values(["promo_month", "product_code"])
    h.to_csv(out / "promotions_history.csv.gz", index=False)
    h[["product_code", "product_name", "l1", "l2", "l3", "vendor", "classification", "product_owner", "national_price", "vendor_cost", "promo_month"]] \
        .sort_values("promo_month").drop_duplicates("product_code", keep="last") \
        .to_csv(out / "products_history.csv.gz", index=False)
    print("promotion history rows", len(h), "products", h.product_code.nunique())


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", default=str(pathlib.Path(__file__).resolve().parents[1] / "data" / "seed"))
    ap.add_argument("--cache", default="/tmp/xlsb_cache")
    a = ap.parse_args()
    out = pathlib.Path(a.out); out.mkdir(parents=True, exist_ok=True)
    pathlib.Path(a.cache).mkdir(parents=True, exist_ok=True)
    src = pathlib.Path(a.src)
    pm = build_bis(cached_csv(src / "BIS.xlsb", a.cache), out)
    build_sales(src, a.cache, out, pm)
    build_hist(src / "Promo_Hist.xlsx", out)
