"""Promotion calculations - a direct port of the Excel logic (Input2 + SOH Form Calculator)."""
import math
import datetime as dt
from zoneinfo import ZoneInfo

TZ = ZoneInfo("Pacific/Auckland")
CDC_PLANNED_WAREHOUSES = ("155-4007", "155-4008")


def xl_round(x, nd=0):
    """Excel ROUND (half away from zero)."""
    f = 10 ** nd
    return math.copysign(math.floor(abs(x) * f + 0.5) / f, x)


def n(v):
    try:
        return float(v) if v not in (None, "") else 0.0
    except (TypeError, ValueError):
        return 0.0


def first_of_month(d):
    if isinstance(d, str):
        d = dt.date.fromisoformat(d[:10])
    return d.replace(day=1)


def add_months(d, k):
    m = d.year * 12 + d.month - 1 + k
    return dt.date(m // 12, m % 12 + 1, 1)


def is_locked(month, today=None):
    """The current month and any past month cannot be created, changed or deleted."""
    today = today or dt.datetime.now(TZ).date()
    return month.year * 12 + month.month <= today.year * 12 + today.month


def soh_for_location(min_pres, floor, method, fc_uplift, ranged, stopped):
    """Calculated SOH for one product-location (Kanban / Forecast rule)."""
    if not ranged or str(stopped or "").strip().upper() == "YES":
        return 0.0
    if method == "Kanban":
        if n(min_pres) > n(floor):
            return n(min_pres) + 1 if n(fc_uplift) >= n(min_pres) else n(min_pres)
        return n(floor)
    if method == "Forecast":
        if n(min_pres) >= n(fc_uplift):
            return n(min_pres) + 1
        return n(min_pres) + n(fc_uplift) / 4
    return n(floor)


def hotspot_type(name):
    return (name or "").split(" - ")[0].strip().lower()


def promo_price(national, discount_pct, discount_amt, explicit=None):
    if explicit not in (None, ""):
        return float(explicit)
    if discount_amt not in (None, "") and n(discount_amt) > 0:
        return n(national) - n(discount_amt)
    if discount_pct not in (None, "") and n(discount_pct) > 0:
        return n(national) * (1 - n(discount_pct))
    return None


def validate(f, product_exists=True):
    """Same rules as the Excel macro's ValidateForm. Returns a list of messages."""
    e = []
    if not f.get("promo_month"):
        e.append("Month must be selected")
    if not f.get("product_code") or not product_exists:
        e.append("Select a valid product")
    t = (f.get("promo_type") or "").lower()
    h = (f.get("hotspot") or "").lower()
    if not t:
        e.append("Promotional Type must be filled")
    if not f.get("promo_intent"):
        e.append("Promotional Intent must be filled")
    if not h:
        e.append("Hotspot must be filled")
    up, pq = f.get("uplift_pct"), f.get("promo_qty_input")
    blank = lambda v: v in (None, "")
    if not blank(up) and not blank(pq):
        e.append("Uplift % and Promo Qty are both filled - clear one of them")
    elif blank(up) and blank(pq):
        e.append("Either Uplift % or Promo Qty must be filled")
    else:
        for v in (up, pq):
            if not blank(v):
                try:
                    float(v)
                except (TypeError, ValueError):
                    e.append("Uplift % / Promo Qty must be a number")
    if h and h not in ("bargain bin", "shelf only"):
        if any(blank(f.get(k)) for k in ("pres_qty_small", "pres_qty_medium", "pres_qty_large")):
            e.append("Presentation Qty (Small, Medium and Large) must be filled for this Hotspot")
    pos = lambda v: not blank(v) and n(v) > 0
    if t == "discount":
        if not (pos(f.get("discount_amt")) or pos(f.get("discount_pct"))):
            e.append("Discount: enter Discount ($) or % Off")
        if pos(f.get("discount_pct")) and n(f["discount_pct"]) > 1:
            e.append("% Off must be between 0% and 100%")
    elif t.startswith("buy 2"):
        if not any(pos(f.get(f"x4y_amt_{i}")) for i in (1, 2, 3)):
            e.append("Buy 2 or more for $XX: fill at least one $Y row")
    elif t.startswith("buy x"):
        pcts = [f.get(f"xy_pct_{i}") for i in (1, 2, 3)]
        if not any(pos(p) for p in pcts):
            e.append("Buy X get XX% discount: fill at least one Y% row")
        if any(pos(p) and n(p) > 1 for p in pcts):
            e.append("Y% must be between 0% and 100%")
    return e


def preview(cur, f):
    """Everything the form shows for a product + month + hotspot + inputs."""
    month = first_of_month(f["promo_month"])
    ly_month = add_months(month, -12)
    code = f["product_code"]
    cur.execute("SELECT * FROM product_master WHERE product_code=%s", (code,))
    prod = cur.fetchone()
    if not prod:
        return None
    cost = n(f.get("promo_cost_override")) if f.get("promo_cost_override") not in (None, "") else n(prod["cost_price"])
    national = n(prod["national_price"])

    cur.execute("""SELECT COALESCE(SUM(qty),0) q, COALESCE(SUM(revenue),0) r, COALESCE(SUM(gm),0) g
                   FROM sales_monthly WHERE product_code=%s AND month=%s""", (code, ly_month))
    ly = cur.fetchone()
    cur.execute("""SELECT COALESCE(SUM(i.avg_qty) FILTER (WHERE l.warehouse_type='Retail'),0) rq,
                          COALESCE(SUM(i.avg_value) FILTER (WHERE l.warehouse_type='Retail'),0) rv,
                          COALESCE(SUM(i.avg_qty) FILTER (WHERE l.warehouse_type='CDC'),0) cq,
                          COALESCE(SUM(i.avg_value) FILTER (WHERE l.warehouse_type='CDC'),0) cv
                   FROM inventory_monthly i JOIN locations l ON l.code=i.location_code
                   WHERE i.product_code=%s AND i.month=%s""", (code, ly_month))
    inv = cur.fetchone()

    pq_in, up_in = f.get("promo_qty_input"), f.get("uplift_pct")
    has_pq = pq_in not in (None, "")
    cur.execute("""SELECT COALESCE(SUM(f.qty),0) z FROM forecast_monthly f JOIN locations l ON l.code=f.location_code
                   WHERE f.kind='current' AND f.product_code=%s AND f.month=%s AND l.warehouse_type='Retail'""", (code, month))
    retail_fc = n(cur.fetchone()["z"])
    if has_pq:
        eff_uplift = (n(pq_in) / retail_fc - 1) if retail_fc > 0 else 0.0
    else:
        eff_uplift = n(up_in)

    promo_qty = n(pq_in) if has_pq else xl_round(n(ly["q"]) * (1 + n(up_in)))
    price = promo_price(national, f.get("discount_pct"), f.get("discount_amt"), f.get("promo_price"))
    promo_rev = promo_qty * (price if price is not None else national)
    promo_gm = promo_rev - promo_qty * cost
    bdf = n(f.get("bdf_applied"))

    # --- locations to be stocked: retail branches that have this hotspot type; if the hotspot type has no
    # spaces defined (e.g. Bargain Bin), every ranged retail location for the product is used.
    htype = hotspot_type(f.get("hotspot"))
    cur.execute("""SELECT bs.code, bs.size, l.name, l.region FROM branch_size bs JOIN locations l ON l.code=bs.code
                   WHERE l.warehouse_type='Retail' AND EXISTS (
                      SELECT 1 FROM hotspots h WHERE h.location_code=bs.code AND h.product_spaces>0
                      AND lower(split_part(h.hotspot_name,' - ',1))=%s)
                   ORDER BY bs.code""", (htype,))
    branches = cur.fetchall()
    if not branches:
        cur.execute("""SELECT bs.code, bs.size, l.name, l.region FROM branch_size bs JOIN locations l ON l.code=bs.code
                       JOIN replen_settings r ON r.location_code=l.code AND r.product_code=%s
                       WHERE l.warehouse_type='Retail' ORDER BY bs.code""", (code,))
        branches = cur.fetchall()
    pres = {"Small": f.get("pres_qty_small"), "Medium": f.get("pres_qty_medium"), "Large": f.get("pres_qty_large")}
    cur.execute("SELECT location_code, range_type, replen_method, stopped_for_purchase, planning_floor FROM replen_settings WHERE product_code=%s", (code,))
    rs = {r["location_code"]: r for r in cur.fetchall()}
    cur.execute("SELECT location_code, on_hand_qty FROM location_stock WHERE product_code=%s", (code,))
    onhand = {r["location_code"]: n(r["on_hand_qty"]) for r in cur.fetchall()}
    cur.execute("SELECT location_code, qty FROM forecast_monthly WHERE kind='current' AND product_code=%s AND month=%s", (code, month))
    fcs = {r["location_code"]: n(r["qty"]) for r in cur.fetchall()}

    locs, soh_total, added_qty = [], 0.0, 0.0
    for b in branches:
        r = rs.get(b["code"], {})
        ranged = bool(r.get("range_type"))
        floor = n(r.get("planning_floor"))
        minp = n(pres.get(b["size"]))
        fcu = xl_round(fcs.get(b["code"], 0) * (1 + eff_uplift))
        soh = soh_for_location(minp, floor, r.get("replen_method"), fcu, ranged, r.get("stopped_for_purchase"))
        cur_soh = onhand.get(b["code"], 0.0)
        inc = max(0.0, soh - cur_soh)      # increase over current actual SOH (not over the planning floor)
        soh_total += soh
        added_qty += inc
        locs.append({"code": b["code"], "name": b["name"], "size": b["size"], "min_pres": minp,
                     "replen_method": r.get("replen_method"), "floor": floor, "uplift_forecast": fcu,
                     "current_soh": cur_soh, "soh": soh, "increase_qty": inc, "increase_value": inc * n(prod["cost_price"]),
                     "range": r.get("range_type"), "stopped": r.get("stopped_for_purchase")})

    cur.execute("""SELECT COALESCE(SUM(s.on_hand_qty),0) q FROM location_stock s JOIN locations l ON l.code=s.location_code
                   WHERE s.product_code=%s AND l.warehouse_type='CDC'""", (code,))
    cdc_soh = n(cur.fetchone()["q"])

    # Ranged locations: Core Range or Local Range rows on Replen Settings
    cur.execute("""SELECT count(*) c FROM replen_settings WHERE product_code=%s
                   AND lower(trim(range_type)) IN ('core range','local range')""", (code,))
    ranged_locs = cur.fetchone()["c"]
    # In CDC: the product has a Planned range at one of the CDC warehouses 155-4007 / 155-4008
    cur.execute("""SELECT EXISTS(SELECT 1 FROM replen_settings WHERE product_code=%s
                   AND location_code = ANY(%s) AND lower(range_type) LIKE '%%planned%%') e""", (code, list(CDC_PLANNED_WAREHOUSES)))
    in_cdc = "Yes" if cur.fetchone()["e"] else "No"
    # Variants: unique product (variant) codes on Replen Settings for the selected product
    cur.execute("SELECT count(DISTINCT product_code) c FROM replen_settings WHERE product_code=%s", (code,))
    variants = cur.fetchone()["c"]

    # --- promo history
    cur.execute("SELECT promo_type, discount_pct, promo_qty FROM promotions WHERE product_code=%s AND promo_month=%s", (code, ly_month))
    same_ly = cur.fetchone()
    cur.execute("""SELECT promo_month, promo_type, discount_pct, promo_qty FROM promotions
                   WHERE product_code=%s AND promo_month<>%s ORDER BY promo_month DESC LIMIT 1""", (code, month))
    last = cur.fetchone()
    cur.execute("SELECT count(*) c FROM promotions WHERE promo_month=%s AND lower(hotspot)=lower(%s)", (month, f.get("hotspot") or ""))
    in_hotspot = cur.fetchone()["c"]

    # --- alerts (pre-promotion, one line each, wording from the alert_definitions table)
    cur.execute("SELECT name FROM alert_definitions WHERE stage='pre' ORDER BY id")
    names = [r["name"] for r in cur.fetchall()]
    nm = lambda i, d: names[i] if i < len(names) else d
    money = lambda v: ("-" if v < 0 else "") + "${:,.0f}".format(abs(v))
    gm_bdf = promo_gm + bdf
    alerts = []
    if gm_bdf < 0:
        alerts.append(f"{nm(0,'Negative GM')}: promo GM + BDF = {money(gm_bdf)}")
    elif gm_bdf < n(ly["g"]):
        alerts.append(f"{nm(1,'Lower GM than last year')}: {money(gm_bdf)} vs {money(n(ly['g']))} LY")
    if added_qty > cdc_soh:
        alerts.append(f"{nm(2,'Not enough inventory in CDC')}: need {added_qty:,.0f} vs {cdc_soh:,.0f} on hand")
    if added_qty > promo_qty:
        alerts.append(f"{nm(3,'Sales less than increased inventory')} (+{added_qty:,.0f} vs {promo_qty:,.0f})")
    if retail_fc < promo_qty:
        alerts.append(f"{nm(4,'Forecast less than Promo Qty')}: {retail_fc:,.0f} vs {promo_qty:,.0f}")

    return {
        "product": {k: prod[k] for k in ("product_code", "description", "l1", "l2", "l3", "vendor", "product_owner",
                                         "primary_sector", "classification", "national_price", "cost_price", "moq", "lifecycle_state")},
        "ranged_locations": ranged_locs, "in_cdc": in_cdc, "variants": variants,
        "ly": {"qty": n(ly["q"]), "rev": n(ly["r"]), "gm": n(ly["g"]), "soh_qty_retail": n(inv["rq"]),
               "soh_val_retail": n(inv["rv"]), "soh_qty_cdc": n(inv["cq"]), "soh_val_cdc": n(inv["cv"])},
        "promo": {"qty": promo_qty, "rev": promo_rev, "gm": promo_gm, "price": price,
                  "gm_pct": (promo_gm / promo_rev) if promo_rev else None,
                  "soh_qty_retail": soh_total, "soh_val_retail": soh_total * n(prod["cost_price"]),
                  "effective_uplift": eff_uplift},
        "change": {"qty": promo_qty - n(ly["q"]), "rev": promo_rev - n(ly["r"]), "gm": promo_gm - n(ly["g"]),
                   "soh_qty_retail": added_qty, "soh_val_retail": added_qty * n(prod["cost_price"])},
        "retail_forecast": retail_fc, "cdc_soh": cdc_soh, "added_inventory_qty": added_qty,
        "added_inventory_value": added_qty * n(prod["cost_price"]),
        "locations": locs, "alerts": alerts, "skus_in_hotspot": in_hotspot,
        "promo_same_month_ly": {"yes": bool(same_ly), "type": same_ly and same_ly["promo_type"],
                                "discount_pct": same_ly and same_ly["discount_pct"],
                                "qty": same_ly and same_ly["promo_qty"]},
        "last_promo": last and {"month": last["promo_month"].isoformat(), "type": last["promo_type"],
                                "discount_pct": last["discount_pct"], "qty": last["promo_qty"]},
    }
