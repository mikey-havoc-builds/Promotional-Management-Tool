# Promo Management Tool

A promotional planning tool in two forms that share the same logic:

| | |
|---|---|
| `excel/MM_Promo_Tool.xlsx` + `excel/PromoMacros.bas` | The original Excel tool (Input form, SOH Form Calculator, Dashboard2) and the VBA module for its Add / Update / Load buttons (import into an `.xlsm`). |
| `app/`, `db/` | The online version: PostgreSQL database, FastAPI backend, and an HTML input form and dashboard that read from and write to the database. |

## Run it

```bash
docker compose up --build                       # Postgres + web app on http://localhost:8000
docker compose run --rm web python -m app.importer excel/MM_Promo_Tool.xlsx --reset
```

Without Docker (needs PostgreSQL 14+):

```bash
createdb promo && psql promo -f db/schema.sql
pip install -r requirements.txt
export DATABASE_URL=postgresql://postgres:promo@localhost:5432/promo
python -m app.importer excel/MM_Promo_Tool.xlsx --reset
uvicorn app.main:app --reload
```

* Input form: `http://localhost:8000/`   * Dashboard: `http://localhost:8000/dashboard`   * API docs: `/docs`

Tests: `pip install -r requirements-dev.txt && pytest` (uses a throw-away `promo_test` database; set
`TEST_DATABASE_BASE` to a connection template such as `postgresql://postgres:promo@localhost:5432/{db}`).

## Database (`db/schema.sql`)

`product_master`, `locations` (+ `branch_size` view), `hotspots`, `replen_settings`, `location_stock`,
`sales_monthly`, `inventory_monthly`, `forecast_monthly` (retained / current), `lists`, `alert_definitions`
and `promotions` (one row per product per promo month - the Excel `Database` sheet).
Monthly tables store the month the figures describe (the workbook's `Oct 25 Sales Qty` column becomes `2025-10-01`).
`python -m app.importer` loads all of this from the workbook, so the sheets can still be used to seed or refresh data.

## Business rules (same as the Excel tool)

* **LY** = sales / inventory for the promo month minus 12 months. **Promo qty** = Promo Qty, or `ROUND(LY qty x (1 + uplift))`.
  Promo revenue = qty x promo price; GM = revenue - qty x cost (or the promo cost override).
* **Branch size**: Large if total hotspot spaces > 50, Small if < 33, else Medium (a size set on `locations.branch_size` overrides this).
* **Promo SOH per product-location** (retail branches that have the chosen hotspot type; if the hotspot type has no spaces
  defined, e.g. Bargain Bin, every ranged retail location for the product):
  not ranged or stopped for purchase -> 0; Kanban: min pres > floor -> (forecast >= min pres ? min pres + 1 : min pres), else keep floor;
  Forecast: min pres >= forecast -> min pres + 1, else min pres + forecast / 4. "Forecast" is the current system forecast x (1 + uplift).
* **Pre-promotion alerts** (one line each, wording from `alert_definitions`): Negative GM; Lower GM than last year (hidden while
  Negative GM shows); Not enough CDC inventory; Sales less than increased inventory; Forecast less than Promo Qty.
* **Saving**: type, intent and hotspot required; exactly one of Uplift % / Promo Qty; presentation qty unless Bargain Bin / Shelf only;
  discount / multi-buy rules by type; the current and past months are locked (New Zealand time); alerts ask for confirmation.
* **Dashboard**: filters L1 / Vendor / Owner / Hotspot, detail level L1 -> Product - Warehouse, 1- or 2-month window
  (month 2 promo = LY of month 2), warehouse promo allocated by last-year share.

## Notes

* No login yet - put the app behind your network / SSO or add authentication before exposing it.
* Promo SOH $ is quantity x vendor cost; CDC promo SOH is not forecast (blank in the Excel tool too).
