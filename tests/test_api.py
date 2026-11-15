BASE = {"promo_month": "2026-12-01", "product_code": "abc", "promo_type": "Discount", "promo_intent": "Increase market share",
        "hotspot": "Aisle End", "uplift_pct": 0.2, "pres_qty_small": 2, "pres_qty_medium": 4, "pres_qty_large": 4,
        "discount_pct": 0.1, "confirm_alerts": True}


def test_preview_matches_excel(client):
    r = client.post("/api/preview", json=BASE).json()
    assert r["ly"]["qty"] == 37 and r["promo"]["qty"] == 44          # Excel Input2 L11 / M11
    assert r["promo"]["price"] == 9 and r["promo"]["rev"] == 396     # E26 / M12
    assert r["ly"]["soh_qty_retail"] == 50 and r["ly"]["soh_qty_cdc"] == 20
    assert r["promo"]["soh_qty_retail"] == 2                         # Kanban, min pres 2 = floor 2 -> keep floor
    assert r["alerts"] == ["Forecast less than Promo Qty: 4 vs 44"]


def test_promo_qty_overrides_uplift(client):
    r = client.post("/api/preview", json={**BASE, "uplift_pct": None, "promo_qty_input": 4}).json()
    assert r["promo"]["qty"] == 4 and r["promo"]["effective_uplift"] == 0.0   # 4 / forecast 4 - 1
    assert r["alerts"] == []


def test_negative_gm_suppresses_lower_gm(client):
    r = client.post("/api/preview", json={**BASE, "discount_pct": 0.5}).json()
    assert r["alerts"][0].startswith("Negative GM")
    assert not any(a.startswith("Lower GM") for a in r["alerts"])


def test_save_requires_confirmation_then_saves(client):
    body = {**BASE}; body.pop("confirm_alerts")
    r = client.post("/api/promotions", json=body).json()
    assert r["needs_confirmation"] and not r["saved"]
    r = client.post("/api/promotions", json={**body, "confirm_alerts": True}).json()
    assert r["saved"]
    # a second promotion for the same product/month is rejected
    assert client.post("/api/promotions", json=BASE).status_code == 409
    # hotspot count now sees one promotion
    assert client.post("/api/preview", json=BASE).json()["skus_in_hotspot"] == 1


def test_update_and_delete(client):
    pid = client.post("/api/promotions", json=BASE).json()["id"]
    r = client.put(f"/api/promotions/{pid}", json={**BASE, "uplift_pct": 0.5}).json()
    assert r["saved"] and r["preview"]["promo"]["qty"] == 56          # round(37 * 1.5)
    assert client.get(f"/api/promotions/{pid}").json()["promo_qty"] == 56
    assert client.put(f"/api/promotions/{pid}", json={**BASE, "product_code": "def"}).status_code == 422
    assert client.delete(f"/api/promotions/{pid}").json()["deleted"]


def test_validation_and_locked_month(client):
    r = client.post("/api/promotions", json={**BASE, "hotspot": None, "uplift_pct": None})
    assert r.status_code == 422 and len(r.json()["errors"]) >= 2
    r = client.post("/api/promotions", json={**BASE, "promo_month": "2020-01-01"})
    assert r.status_code == 422 and any("locked" in e for e in r.json()["errors"])


def test_dashboard(client):
    client.post("/api/promotions", json=BASE)
    d = client.get("/api/dashboard", params={"month": "2026-12-01"}).json()
    assert d["promotions"] == 1 and d["totals"]["ly_qty"] == 37 and d["totals"]["promo_qty"] == 44
    assert d["options"]["l1"] == ["pet"]
    # filter that excludes the promotion
    assert client.get("/api/dashboard", params={"month": "2026-12-01", "vendor": "nobody"}).json()["promotions"] == 0
    # product-warehouse: promo forecast is allocated by last-year share and sums back to the total
    pw = client.get("/api/dashboard", params={"month": "2026-12-01", "level": "product_warehouse"}).json()
    assert round(sum(r["promo_qty"] for r in pw["rows"]), 6) == 44
    # 2-month window: month 2 promo = LY of month 2 (Jan-26 sales for abc: 9 + 39)
    w2 = client.get("/api/dashboard", params={"month": "2026-12-01", "window": 2}).json()
    assert w2["totals"]["promo_qty"] == 44 + 48 and w2["totals"]["ly_qty"] == 37 + 48
    for lvl in ("l1", "vendor", "owner", "hotspot", "promo_type", "product"):
        assert client.get("/api/dashboard", params={"month": "2026-12-01", "level": lvl}).json()["rows"]


def test_pages_and_lists(client):
    assert client.get("/").status_code == 200 and client.get("/dashboard").status_code == 200
    l = client.get("/api/lists").json()
    assert "Discount" in l["promo_type"] and any(p["product_code"] == "abc" for p in l["products"])
