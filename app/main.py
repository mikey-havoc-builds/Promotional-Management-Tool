"""Promo Management Tool - FastAPI backend.   uvicorn app.main:app --reload"""
import datetime as dt
from pathlib import Path
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict

from . import calc, dashboard
from .db import connect

app = FastAPI(title="Promo Management Tool")
STATIC = Path(__file__).parent / "static"

FIELDS = ["promo_month", "product_code", "promo_type", "promo_intent", "hotspot", "promo_cost_override", "uplift_pct",
          "promo_qty_input", "bdf_applied", "pres_qty_small", "pres_qty_medium", "pres_qty_large", "market_in_1",
          "market_in_2", "market_in_3", "discount_pct", "discount_amt", "promo_price",
          "x4y_qty_1", "x4y_qty_2", "x4y_qty_3", "x4y_amt_1", "x4y_amt_2", "x4y_amt_3",
          "xy_qty_1", "xy_qty_2", "xy_qty_3", "xy_pct_1", "xy_pct_2", "xy_pct_3"]


class PromoIn(BaseModel):
    model_config = ConfigDict(extra="ignore")
    promo_month: str | None = None
    product_code: str | None = None
    promo_type: str | None = None
    promo_intent: str | None = None
    hotspot: str | None = None
    promo_cost_override: float | None = None
    uplift_pct: float | None = None
    promo_qty_input: float | None = None
    bdf_applied: float | None = None
    pres_qty_small: float | None = None
    pres_qty_medium: float | None = None
    pres_qty_large: float | None = None
    market_in_1: str | None = None
    market_in_2: str | None = None
    market_in_3: str | None = None
    discount_pct: float | None = None
    discount_amt: float | None = None
    promo_price: float | None = None
    x4y_qty_1: float | None = None
    x4y_qty_2: float | None = None
    x4y_qty_3: float | None = None
    x4y_amt_1: float | None = None
    x4y_amt_2: float | None = None
    x4y_amt_3: float | None = None
    xy_qty_1: float | None = None
    xy_qty_2: float | None = None
    xy_qty_3: float | None = None
    xy_pct_1: float | None = None
    xy_pct_2: float | None = None
    xy_pct_3: float | None = None
    confirm_alerts: bool = False


def form_dict(p: PromoIn):
    d = {k: getattr(p, k) for k in FIELDS}
    d["promo_month"] = calc.first_of_month(d["promo_month"]) if d["promo_month"] else None
    return d


def j(o):
    """JSON-safe copy (Decimal/date)."""
    from fastapi.encoders import jsonable_encoder
    return jsonable_encoder(o, custom_encoder={__import__("decimal").Decimal: float})


@app.get("/api/health")
def health():
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT 1 ok")
        return cur.fetchone()


@app.get("/api/lists")
def lists():
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT list_name, value FROM lists ORDER BY list_name, sort_order")
        out = {}
        for r in cur.fetchall():
            out.setdefault(r["list_name"], []).append(r["value"])
        cur.execute("""SELECT product_code, description, l1, vendor FROM product_master
                       WHERE description IS NOT NULL ORDER BY product_code""")
        out["products"] = cur.fetchall()
        cur.execute("SELECT DISTINCT vendor FROM product_master WHERE vendor IS NOT NULL ORDER BY 1")
        out["vendors"] = [r["vendor"] for r in cur.fetchall()]
        out["today"] = dt.datetime.now(calc.TZ).date().isoformat()
        return j(out)


@app.post("/api/preview")
def preview(p: PromoIn):
    f = form_dict(p)
    if not f["promo_month"] or not f["product_code"]:
        raise HTTPException(422, "promo_month and product_code are required")
    with connect() as con, con.cursor() as cur:
        res = calc.preview(cur, f)
        if res is None:
            raise HTTPException(404, "Unknown product")
        res["locked"] = calc.is_locked(f["promo_month"])
        return j(res)


@app.get("/api/promotions")
def list_promotions(month: str):
    m = calc.first_of_month(month)
    with connect() as con, con.cursor() as cur:
        cur.execute("""SELECT id, promo_month, product_code, product_name, l1, product_owner, vendor, promo_type,
                       promo_intent, hotspot FROM promotions WHERE promo_month=%s ORDER BY product_code""", (m,))
        return j({"locked": calc.is_locked(m), "rows": cur.fetchall()})


@app.get("/api/promotions/{pid}")
def get_promotion(pid: int):
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT * FROM promotions WHERE id=%s", (pid,))
        r = cur.fetchone()
        if not r:
            raise HTTPException(404, "Not found")
        r["locked"] = calc.is_locked(r["promo_month"])
        return j(r)


def _save(p: PromoIn, pid: int | None):
    f = form_dict(p)
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT 1 FROM product_master WHERE product_code=%s", (f["product_code"],)) if f["product_code"] else None
        exists = bool(f["product_code"]) and cur.fetchone() is not None
        errs = calc.validate(f, exists)
        if f["promo_month"] and calc.is_locked(f["promo_month"]):
            errs.append("The current month and past months are locked and cannot be saved")
        old = None
        if pid is not None:
            cur.execute("SELECT promo_month, product_code FROM promotions WHERE id=%s", (pid,))
            old = cur.fetchone()
            if not old:
                raise HTTPException(404, "Promotion not found")
            if calc.is_locked(old["promo_month"]):
                errs.append("This promotion is in a locked month and cannot be changed")
            if f["product_code"] and (old["product_code"] != f["product_code"] or old["promo_month"] != f["promo_month"]):
                errs.append("Product and month cannot be changed on an existing promotion - add a new one instead")
        if errs:
            return JSONResponse({"saved": False, "errors": errs}, status_code=422)
        if pid is None:
            cur.execute("SELECT id FROM promotions WHERE promo_month=%s AND product_code=%s", (f["promo_month"], f["product_code"]))
            if cur.fetchone():
                return JSONResponse({"saved": False, "errors": ["This product already has a promotion in that month - use Update"]}, status_code=409)
        pv = calc.preview(cur, f)
        if pv["alerts"] and not p.confirm_alerts:
            return {"saved": False, "needs_confirmation": True, "alerts": pv["alerts"]}
        prod, ly, pr = pv["product"], pv["ly"], pv["promo"]
        rec = {k: f[k] for k in FIELDS if k not in ("promo_price", "promo_qty_input") or True}
        rec.pop("promo_price"); rec.pop("promo_qty_input")
        rec.update({
            "promo_qty_input": f["promo_qty_input"], "promo_price": pr["price"],
            "product_name": prod["description"], "l1": prod["l1"], "l2": prod["l2"], "l3": prod["l3"],
            "vendor": prod["vendor"], "classification": prod["classification"], "product_owner": prod["product_owner"],
            "national_price": prod["national_price"], "vendor_cost": prod["cost_price"], "moq": prod["moq"],
            "ranged_locations": pv["ranged_locations"], "in_cdc": pv["in_cdc"],
            "ly_qty": ly["qty"], "ly_rev": ly["rev"], "ly_gm": ly["gm"],
            "promo_qty": pr["qty"], "promo_rev": pr["rev"], "promo_gm": pr["gm"],
            "ly_soh_qty_retail": ly["soh_qty_retail"], "promo_soh_qty_retail": pr["soh_qty_retail"],
            "promo_soh_val_retail": pr["soh_val_retail"],
        })
        if pid is None:
            ks = list(rec)
            cur.execute(f"INSERT INTO promotions ({','.join(ks)}) VALUES ({','.join(['%s']*len(ks))}) RETURNING id", [rec[k] for k in ks])
            new_id = cur.fetchone()["id"]
        else:
            rec.pop("promo_month"); rec.pop("product_code")
            ks = list(rec)
            cur.execute(f"UPDATE promotions SET {','.join(k+'=%s' for k in ks)}, updated_at=now() WHERE id=%s", [rec[k] for k in ks] + [pid])
            new_id = pid
        con.commit()
        return j({"saved": True, "id": new_id, "alerts": pv["alerts"], "preview": pv})


@app.post("/api/promotions")
def add_promotion(p: PromoIn):
    return _save(p, None)


@app.put("/api/promotions/{pid}")
def update_promotion(pid: int, p: PromoIn):
    return _save(p, pid)


@app.delete("/api/promotions/{pid}")
def delete_promotion(pid: int):
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT promo_month FROM promotions WHERE id=%s", (pid,))
        r = cur.fetchone()
        if not r:
            raise HTTPException(404, "Not found")
        if calc.is_locked(r["promo_month"]):
            raise HTTPException(422, "Locked month")
        cur.execute("DELETE FROM promotions WHERE id=%s", (pid,))
        con.commit()
        return {"deleted": True}


@app.get("/api/dashboard")
def get_dashboard(month: str, window: int = 1, l1: str | None = None, vendor: str | None = None,
                  owner: str | None = None, hotspot: str | None = None, level: str = "product"):
    with connect() as con, con.cursor() as cur:
        return j(dashboard.build(cur, month, 2 if window == 2 else 1, l1, vendor, owner, hotspot, level))


@app.get("/api/months")
def months():
    with connect() as con, con.cursor() as cur:
        cur.execute("SELECT DISTINCT promo_month FROM promotions ORDER BY 1")
        return j({"months": [r["promo_month"] for r in cur.fetchall()]})


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/dashboard")
def dash_page():
    return FileResponse(STATIC / "dashboard.html")


app.mount("/static", StaticFiles(directory=STATIC), name="static")
