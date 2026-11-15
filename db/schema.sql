-- Promo Management Tool - PostgreSQL schema
-- Mirrors the Excel workbook: product master, branches/hotspots, monthly Sales / Inventory / Forecast
-- history, replenishment settings and the promotions table (the Excel "Database" sheet).

CREATE TABLE IF NOT EXISTS product_master (
    product_code    text PRIMARY KEY,
    description     text,
    l1              text, l2 text, l3 text,
    vendor          text,
    product_owner   text,
    primary_sector  text,
    classification  text,
    national_price  numeric(12,4),
    cost_price      numeric(12,4),
    moq             integer,
    lifecycle_state text
);

-- Branches (stores / warehouses). warehouse_type: Retail or CDC
CREATE TABLE IF NOT EXISTS locations (
    code            text PRIMARY KEY,           -- warehouse / branch code
    name            text,
    region          text,
    warehouse_type  text NOT NULL DEFAULT 'Retail' CHECK (warehouse_type IN ('Retail','CDC')),
    branch_size     text CHECK (branch_size IN ('Small','Medium','Large'))   -- optional manual override
);

-- Hotspot spaces per branch (the Excel "Hotspots" sheet)
CREATE TABLE IF NOT EXISTS hotspots (
    location_code   text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    hotspot_name    text NOT NULL,   -- e.g. 'Aisle End - Hotspot 1'; hotspot type = text before ' - '
    product_spaces  integer NOT NULL DEFAULT 0,
    PRIMARY KEY (location_code, hotspot_name)
);

-- Branch size: the manual override on locations if set, otherwise by total spaces:
-- Large > 50, Small < 33, otherwise Medium (same rule as the Excel Hotspots(fml) sheet)
CREATE OR REPLACE VIEW branch_size AS
SELECT l.code,
       COALESCE(SUM(h.product_spaces), 0) AS total_spaces,
       COALESCE(l.branch_size,
                CASE WHEN COALESCE(SUM(h.product_spaces),0) > 50 THEN 'Large'
                     WHEN COALESCE(SUM(h.product_spaces),0) < 33 THEN 'Small'
                     ELSE 'Medium' END) AS size
FROM locations l LEFT JOIN hotspots h ON h.location_code = l.code
GROUP BY l.code, l.branch_size;

CREATE TABLE IF NOT EXISTS replen_settings (
    product_code        text NOT NULL REFERENCES product_master(product_code) ON DELETE CASCADE,
    location_code       text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    range_type          text,
    replen_method       text,                    -- Kanban / Forecast
    stopped_for_purchase text,                   -- Yes / No
    open_for_replen     text,
    planning_floor      numeric(12,2),
    main_warehouse      text,
    supply_tier         text,
    PRIMARY KEY (product_code, location_code)
);

CREATE TABLE IF NOT EXISTS location_stock (      -- current on hand (Excel "Location Details")
    product_code  text NOT NULL REFERENCES product_master(product_code) ON DELETE CASCADE,
    location_code text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    on_hand_qty   numeric(14,2), on_hand_value numeric(14,2),
    PRIMARY KEY (product_code, location_code)
);

-- Monthly history. "month" is the first day of the month the figures describe.
CREATE TABLE IF NOT EXISTS sales_monthly (
    product_code  text NOT NULL REFERENCES product_master(product_code) ON DELETE CASCADE,
    location_code text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    month         date NOT NULL,
    qty numeric(14,2), revenue numeric(14,2), gm numeric(14,2),
    transactions integer, customers integer,
    PRIMARY KEY (product_code, location_code, month)
);
CREATE INDEX IF NOT EXISTS sales_monthly_month_idx ON sales_monthly(month);

CREATE TABLE IF NOT EXISTS inventory_monthly (
    product_code  text NOT NULL REFERENCES product_master(product_code) ON DELETE CASCADE,
    location_code text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    month         date NOT NULL,
    avg_qty numeric(14,2), avg_value numeric(14,2), osa_days numeric(8,2),
    PRIMARY KEY (product_code, location_code, month)
);

-- kind: 'retained' = forecast that was retained at the time (history), 'current' = current system forecast
CREATE TABLE IF NOT EXISTS forecast_monthly (
    kind          text NOT NULL CHECK (kind IN ('retained','current')),
    product_code  text NOT NULL REFERENCES product_master(product_code) ON DELETE CASCADE,
    location_code text NOT NULL REFERENCES locations(code) ON DELETE CASCADE,
    month         date NOT NULL,
    qty numeric(14,2), revenue numeric(14,2), gm numeric(14,2),
    PRIMARY KEY (kind, product_code, location_code, month)
);

CREATE TABLE IF NOT EXISTS lists (               -- dropdown values: promo_type, intent, hotspot, market_in, owner
    list_name text NOT NULL, value text NOT NULL, sort_order integer NOT NULL DEFAULT 0,
    PRIMARY KEY (list_name, value)
);

CREATE TABLE IF NOT EXISTS alert_definitions (
    id serial PRIMARY KEY, stage text NOT NULL, name text NOT NULL, description text,
    UNIQUE (stage, name)
);

-- One row per product per promo month (the Excel "Database" sheet)
CREATE TABLE IF NOT EXISTS promotions (
    id               serial PRIMARY KEY,
    promo_month      date NOT NULL CHECK (promo_month = date_trunc('month', promo_month)::date),
    product_code     text NOT NULL REFERENCES product_master(product_code),
    product_name     text,
    l1 text, l2 text, l3 text, vendor text, classification text, product_owner text,
    national_price   numeric(12,4), vendor_cost numeric(12,4), moq integer,
    ranged_locations integer, in_cdc text,
    promo_type       text NOT NULL,
    promo_intent     text NOT NULL,
    hotspot          text NOT NULL,
    promo_cost_override numeric(12,4),
    uplift_pct       numeric(8,4),
    promo_qty_input  numeric(14,2),
    bdf_applied      numeric(14,2),
    pres_qty_small integer, pres_qty_medium integer, pres_qty_large integer,
    market_in_1 text, market_in_2 text, market_in_3 text,
    discount_pct numeric(8,4), discount_amt numeric(12,4), promo_price numeric(12,4),
    x4y_qty_1 numeric, x4y_qty_2 numeric, x4y_qty_3 numeric,
    x4y_amt_1 numeric, x4y_amt_2 numeric, x4y_amt_3 numeric,
    xy_qty_1 numeric, xy_qty_2 numeric, xy_qty_3 numeric,
    xy_pct_1 numeric, xy_pct_2 numeric, xy_pct_3 numeric,
    -- snapshot of the calculated figures when saved
    ly_qty numeric(14,2), ly_rev numeric(14,2), ly_gm numeric(14,2),
    promo_qty numeric(14,2), promo_rev numeric(14,2), promo_gm numeric(14,2),
    ly_soh_qty_retail numeric(14,2), promo_soh_qty_retail numeric(14,2),
    promo_soh_val_retail numeric(14,2), promo_soh_qty_cdc numeric(14,2), promo_soh_val_cdc numeric(14,2),
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (promo_month, product_code)
);
CREATE INDEX IF NOT EXISTS promotions_month_idx ON promotions(promo_month);
CREATE INDEX IF NOT EXISTS promotions_hotspot_idx ON promotions(promo_month, lower(hotspot));
