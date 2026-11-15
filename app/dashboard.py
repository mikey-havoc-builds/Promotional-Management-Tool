"""Dashboard aggregation (port of Dashboard2 / Dash2_Calc)."""
from collections import defaultdict
from .calc import add_months, first_of_month, n

LEVELS = ("l1", "vendor", "owner", "hotspot", "promo_type", "product", "product_warehouse")
METRICS = ("qty", "rev", "gm", "soh")


def _blank():
    return {k: 0.0 for k in ("ly_qty", "ly_rev", "ly_gm", "promo_qty", "promo_rev", "promo_gm", "act_qty", "act_rev",
                             "act_gm", "ly_soh", "promo_soh", "act_soh")}


def build(cur, month, window=1, l1=None, vendor=None, owner=None, hotspot=None, level="product"):
    month = first_of_month(month)
    months = [month] + ([add_months(month, 1)] if window == 2 else [])
    if level not in LEVELS:
        level = "product"

    cur.execute("""SELECT DISTINCT l1, vendor, product_owner AS owner, hotspot FROM promotions WHERE promo_month=%s""", (month,))
    opts = {"l1": set(), "vendor": set(), "owner": set(), "hotspot": set()}
    for r in cur.fetchall():
        for k in opts:
            if r[k]:
                opts[k].add(r[k])

    where, args = ["promo_month=%s"], [month]
    for col, val in (("l1", l1), ("vendor", vendor), ("product_owner", owner), ("hotspot", hotspot)):
        if val and val != "All":
            where.append(f"{col}=%s"); args.append(val)
    cur.execute(f"SELECT * FROM promotions WHERE {' AND '.join(where)} ORDER BY product_code", args)
    promos = cur.fetchall()
    codes = sorted({p["product_code"] for p in promos})
    empty = {"month": month.isoformat(), "months": [m.isoformat() for m in months], "level": level,
             "options": {k: sorted(v) for k, v in opts.items()}, "rows": [], "totals": _blank(), "promotions": 0}
    if not promos:
        return empty

    ly_months = [add_months(m, -12) for m in months]
    cur.execute("""SELECT product_code, location_code, month, qty, revenue, gm FROM sales_monthly
                   WHERE product_code = ANY(%s) AND month = ANY(%s)""", (codes, ly_months + months))
    sales = defaultdict(lambda: {"qty": 0.0, "rev": 0.0, "gm": 0.0})
    for r in cur.fetchall():
        k = (r["product_code"], r["location_code"], r["month"])
        sales[k] = {"qty": n(r["qty"]), "rev": n(r["revenue"]), "gm": n(r["gm"])}
    cur.execute("""SELECT i.product_code, i.location_code, i.month, i.avg_qty FROM inventory_monthly i
                   JOIN locations l ON l.code=i.location_code WHERE l.warehouse_type='Retail'
                   AND i.product_code = ANY(%s) AND i.month = ANY(%s)""", (codes, ly_months + months))
    inv = defaultdict(float)
    for r in cur.fetchall():
        inv[(r["product_code"], r["location_code"], r["month"])] = n(r["avg_qty"])

    sale_locs = defaultdict(set)
    for (p, l, m) in list(sales) + list(inv):
        sale_locs[p].add(l)

    groups = defaultdict(lambda: {"v": _blank(), "n": 0, "meta": {}})

    def key_for(p, loc=None):
        if level == "l1": return p["l1"] or "(none)", p["l1"] or "(none)"
        if level == "vendor": return p["vendor"] or "(none)", p["vendor"] or "(none)"
        if level == "owner": return p["product_owner"] or "(none)", p["product_owner"] or "(none)"
        if level == "hotspot": return p["hotspot"], p["hotspot"]
        if level == "promo_type": return p["promo_type"], p["promo_type"]
        if level == "product": return p["product_code"], f"{p['product_code']} - {p['product_name'] or ''}".strip(" -")
        return f"{p['product_code']}|{loc}", f"{p['product_code']} - {p['product_name'] or ''} @ {loc}".replace(" - @", " @")

    for p in promos:
        code = p["product_code"]
        locs = sorted(sale_locs[code])
        tot = _blank()
        per_loc = {l: _blank() for l in locs}
        for idx, (m, lm) in enumerate(zip(months, ly_months)):
            for l in locs:
                s_ly, s_act = sales[(code, l, lm)], sales[(code, l, m)]
                pl = per_loc[l]
                pl["ly_qty"] += s_ly["qty"]; pl["ly_rev"] += s_ly["rev"]; pl["ly_gm"] += s_ly["gm"]
                pl["act_qty"] += s_act["qty"]; pl["act_rev"] += s_act["rev"]; pl["act_gm"] += s_act["gm"]
                pl["ly_soh"] += inv[(code, l, lm)] / len(months)
                pl["act_soh"] += inv[(code, l, m)] / len(months)
        # promo forecast: month 1 = saved promo figures, month 2 = LY of month 2
        ly_tot = {k: sum(per_loc[l][k] for l in locs) for k in ("ly_qty", "ly_rev", "ly_gm", "ly_soh")}
        m2 = {"qty": 0.0, "rev": 0.0, "gm": 0.0}
        if window == 2:
            for l in locs:
                s = sales[(code, l, ly_months[1])]
                for a in m2: m2[a] += s[a]
        promo = {"promo_qty": n(p["promo_qty"]) + m2["qty"], "promo_rev": n(p["promo_rev"]) + m2["rev"],
                 "promo_gm": n(p["promo_gm"]) + m2["gm"], "promo_soh": n(p["promo_soh_qty_retail"])}
        for l in locs:
            pl = per_loc[l]
            # promo forecast per warehouse is allocated by the warehouse's share of last year
            sq = pl["ly_qty"] / ly_tot["ly_qty"] if ly_tot["ly_qty"] else (1 / len(locs) if locs else 0)
            ss = pl["ly_soh"] / ly_tot["ly_soh"] if ly_tot["ly_soh"] else (1 / len(locs) if locs else 0)
            pl["promo_qty"], pl["promo_rev"], pl["promo_gm"] = (promo["promo_qty"] * sq, promo["promo_rev"] * sq, promo["promo_gm"] * sq)
            pl["promo_soh"] = promo["promo_soh"] * ss
        targets = [(None, tot, True)] if level != "product_warehouse" else [(l, per_loc[l], False) for l in locs]
        if level != "product_warehouse":
            for l in locs:
                for k, v in per_loc[l].items(): tot[k] += v
            tot.update(promo)
            tot["ly_soh"] = ly_tot["ly_soh"]
        for l, vals, _ in targets:
            k, label = key_for(p, l)
            g = groups[k]
            g["meta"]["label"] = label
            for kk, vv in vals.items(): g["v"][kk] += vv
            g["n"] += 1 if level != "product_warehouse" else 0

    rows, totals = [], _blank()
    for k, g in groups.items():
        v = g["v"]
        row = {"key": k, "label": g["meta"]["label"], "promotions": g["n"], **v,
               "ly_gm_pct": v["ly_gm"] / v["ly_rev"] if v["ly_rev"] else None,
               "promo_gm_pct": v["promo_gm"] / v["promo_rev"] if v["promo_rev"] else None,
               "act_gm_pct": v["act_gm"] / v["act_rev"] if v["act_rev"] else None}
        rows.append(row)
        for kk in totals: totals[kk] += v[kk]
    rows.sort(key=lambda r: r["label"])
    totals.update({"ly_gm_pct": totals["ly_gm"] / totals["ly_rev"] if totals["ly_rev"] else None,
                   "promo_gm_pct": totals["promo_gm"] / totals["promo_rev"] if totals["promo_rev"] else None,
                   "act_gm_pct": totals["act_gm"] / totals["act_rev"] if totals["act_rev"] else None})
    empty.update({"rows": rows, "totals": totals, "promotions": len(promos)})
    return empty
