# Promo Management Tool

A promotional planning tool in two forms that share the same logic:

| | |
|---|---|
| `excel/MM_Promo_Tool.xlsx` + `excel/PromoMacros.bas` | The original Excel tool (Input form, SOH Form Calculator, Dashboard2) and the VBA module for its Add / Update / Load buttons (import into an `.xlsm`). |
| `app/`, `db/` | The online version: PostgreSQL database, FastAPI backend, and an HTML input form and dashboard that read from and write to the database. |

## Run it

```bash
docker compose up --build                       # Postgres + web app on http://localhost:8000
docker compose run --rm web python -m app.seed                 # real data (see "Real data" below)
# or the small sample workbook:  docker compose run --rm web python -m app.importer excel/MM_Promo_Tool.xlsx --reset
```

Without Docker (needs PostgreSQL 14+):

```bash
createdb promo && psql promo -f db/schema.sql
pip install -r requirements.txt
export DATABASE_URL=postgresql://postgres:promo@localhost:5432/promo
python -m app.seed                 # real data (or: python -m app.importer excel/MM_Promo_Tool.xlsx --reset for the sample)
uvicorn app.main:app --reload
```

* Input form: `http://localhost:8000/`   * Dashboard: `http://localhost:8000/dashboard`   * API docs: `/docs`

Tests: `pip install -r requirements-dev.txt && pytest` (uses a throw-away `promo_test` database; set
`TEST_DATABASE_BASE` to a connection template such as `postgresql://postgres:promo@localhost:5432/{db}`).

## Database (`db/schema.sql`)

`product_master`, `locations` (+ `branch_size` view), `hotspots`, `replen_settings`, `location_stock`,
`sales_monthly` (by region), `inventory_monthly`, `forecast_monthly` (retained / current), `lists`, `alert_definitions`
and `promotions` (one row per product per promo month - the Excel `Database` sheet).
Monthly tables store the month the figures describe (the workbook's `Oct 25 Sales Qty` column becomes `2025-10-01`).
`python -m app.importer` loads all of this from the workbook, so the sheets can still be used to seed or refresh data.

## Real data (`data_build/`, `app/seed.py`)

The demo database is loaded from the extracts in `PromoMgmt/Data_Tables`:

| Source file | Goes into |
|---|---|
| `BIS.xlsb` (stock / range / replen snapshot, variant x warehouse) | `product_master` (one row per product number, with variant count), `locations`, `replen_settings`, `location_stock` |
| `SNO_Qty`, `SNO_Revenue`, `SNO_GM`, `SNO_Cust` `.xlsb` (monthly pivots, product x region) | `sales_monthly` (qty, revenue, GM, customers **by region**) |
| `Promo_Hist.xlsx` (Database sheet) | `promotions` with `source='history'` (duplicates and blank type/intent kept as-is) |

```bash
pip install -r requirements-data.txt
python data_build/build_seed.py --src "C:/path/to/PromoMgmt/Data_Tables"      # writes data/seed/*.csv.gz (about 1-2 min)
python -m app.seed                                                          # dropdown lists / HotSpots panel from the workbook, then all data tables replaced from data/seed
```

Notes on how the extracts are mapped:
* Locations: retail branches plus distribution centres. `National 3PL DC Hamilton` / `Christchurch` are coded `155-4008` / `155-4007` (the CDC planned-range check); "DO NOT USE", real-estate, port, mill and bulk-store sites are dropped. Other locations use the BIS warehouse name as their code.
* `replen_settings` is combined to product x location (best range across variants; floors and stock summed). `replen_method` is derived from the demand tier: `K*` = Kanban, everything else = Forecast. Never-ranged locations with no stock are not stored.
* Sales have no warehouse, so `sales_monthly` is keyed by region and the dashboard's "product / warehouse" level shows product / region. Customers are matched to products by name + vendor (unique names only).
* **Not in the extracts, so empty:** monthly inventory history (`inventory_monthly`), forecasts (`forecast_monthly` - the "forecast less than promo qty" alert fires until loaded), branch hotspot spaces (`hotspots` - branch size falls back to Small and the form uses all ranged branches).
* `data/seed/` is git-ignored: it holds real cost, stock and sales figures. Rebuild it from the source files rather than committing it to a public repository.

## Business rules (same as the Excel tool)

* **LY** = sales / inventory for the promo month minus 12 months. **Promo qty** = Promo Qty, or `ROUND(LY qty x (1 + uplift))`.
  Promo revenue = qty x promo price; GM = revenue - qty x cost (or the promo cost override).
* **Branch size**: Large if total hotspot spaces > 50, Small if < 33, else Medium (a size set on `locations.branch_size` overrides this).
* **Promo SOH per product-location** (retail branches that have the chosen hotspot type; if the hotspot type has no spaces
  defined, e.g. Bargain Bin, every ranged retail location for the product):
  not ranged or stopped for purchase -> 0; Kanban: min pres > floor -> (forecast >= min pres ? min pres + 1 : min pres), else keep floor;
  Forecast: min pres >= forecast -> min pres + 1, else min pres + forecast / 4. "Forecast" is the current system forecast x (1 + uplift).
* **Selection counts** (from `replen_settings`): *# Ranged locations* = locations with Core Range or Local Range; *In CDC* = Yes if the product has a
  Planned range at warehouse 155-4007 or 155-4008, else No; *# Variants* = unique product codes for the selected product.
* **SOH change / added inventory** = new SOH minus the location's current on-hand (`location_stock`), never below 0
  (e.g. new 5 vs on-hand 2 -> +3; new 5 vs on-hand 8 -> 0). The Excel "Change" column for Ave SOH shows the total.
* **Pre-promotion alerts** (one line each, wording from `alert_definitions`): Negative GM; Lower GM than last year (hidden while
  Negative GM shows); Not enough CDC inventory; Sales less than increased inventory; Forecast less than Promo Qty.
* **Saving**: type, intent and hotspot required; exactly one of Uplift % / Promo Qty; presentation qty unless Bargain Bin / Shelf only;
  discount / multi-buy rules by type; the current and past months are locked (New Zealand time); alerts ask for confirmation.
* **HotSpots panel** (right of the form, as on Input2): every hotspot position with its number of Spots. The form's Hotspot dropdown lists these positions
  (e.g. `Aisle End - Hotspot 3`), so each promotion is assigned to a spot. *Used* = number of products assigned to that spot in the selected month, and turns red
  when Used is greater than Spots. SOH branches and "# SKUs already in Hotspot" still work at hotspot-type level (the text before ` - `).
* **Dashboard**: filters L1 / Vendor / Owner / Hotspot, detail level L1 -> Product - Warehouse, 1- or 2-month window
  (month 2 promo = LY of month 2), warehouse promo allocated by last-year share.

## Notes

* No login yet - put the app behind your network / SSO or add authentication before exposing it.
* Promo SOH $ is quantity x vendor cost; CDC promo SOH is not forecast (blank in the Excel tool too).
