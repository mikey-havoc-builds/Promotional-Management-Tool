import datetime as dt
import pytest
from app import calc


@pytest.mark.parametrize("min_pres,floor,method,fc,ranged,stopped,expect", [
    (4, 2, "Kanban", 5, True, "No", 5),      # min pres > floor, forecast >= min pres -> min pres + 1
    (4, 2, "Kanban", 3, True, "No", 4),      # min pres > floor, forecast < min pres  -> min pres
    (2, 4, "Kanban", 9, True, "No", 4),      # min pres < floor -> keep floor
    (4, 0, "Forecast", 4, True, "No", 5),    # min pres >= forecast -> min pres + 1
    (4, 0, "Forecast", 12, True, "No", 7),   # min pres < forecast -> min pres + forecast/4
    (4, 2, "Kanban", 9, False, "No", 0),     # not ranged
    (4, 2, "Kanban", 9, True, "Yes", 0),     # stopped for purchase
])
def test_soh_rule(min_pres, floor, method, fc, ranged, stopped, expect):
    assert calc.soh_for_location(min_pres, floor, method, fc, ranged, stopped) == expect


def test_excel_round_half_up():
    assert calc.xl_round(2.5) == 3 and calc.xl_round(44.4) == 44 and calc.xl_round(-2.5) == -3


def test_month_lock():
    today = dt.date(2026, 11, 15)
    assert calc.is_locked(dt.date(2026, 11, 1), today) and calc.is_locked(dt.date(2026, 1, 1), today)
    assert not calc.is_locked(dt.date(2026, 12, 1), today)


def test_validation_rules():
    ok = dict(promo_month=dt.date(2026, 12, 1), product_code="abc", promo_type="Discount", promo_intent="Defend sales",
              hotspot="Aisle End", uplift_pct=0.2, pres_qty_small=1, pres_qty_medium=2, pres_qty_large=3, discount_pct=0.1)
    assert calc.validate(ok) == []
    assert any("both filled" in e for e in calc.validate({**ok, "promo_qty_input": 5}))
    assert any("Either" in e for e in calc.validate({**ok, "uplift_pct": None}))
    assert any("Presentation" in e for e in calc.validate({**ok, "pres_qty_small": None}))
    assert calc.validate({**ok, "pres_qty_small": None, "hotspot": "Bargain Bin"}) == []
    assert any("Discount" in e for e in calc.validate({**ok, "discount_pct": None}))
    assert any("% Off" in e for e in calc.validate({**ok, "discount_pct": 10}))
    assert any("Buy 2" in e for e in calc.validate({**ok, "promo_type": "Buy 2 or more for $XX"}))
    assert any("Buy X" in e for e in calc.validate({**ok, "promo_type": "Buy X get XX% discount"}))
