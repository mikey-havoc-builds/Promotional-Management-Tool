"""Load the real data (data/seed/*.csv.gz, built by data_build/build_seed.py) into PostgreSQL.

    python -m app.seed [--seed data/seed] [--workbook excel/MM_Promo_Tool.xlsx]

--workbook first loads the settings from the workbook (dropdown lists, alert definitions, the HotSpots panel),
then every data table (products, locations, replen settings, stock, sales, promotions) is replaced with the seed data.
"""
import argparse, gzip, pathlib
from .db import connect
from . import importer

ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA_TABLES = ("promotions", "sales_monthly", "inventory_monthly", "forecast_monthly", "location_stock",
               "replen_settings", "hotspots", "locations", "product_master")


def copy_csv(cur, path, target, columns=None):
    cols = f"({','.join(columns)})" if columns else ""
    with gzip.open(path, "rb") as f, cur.copy(f"COPY {target} {cols} FROM STDIN WITH (FORMAT csv, HEADER true)") as cp:
        while chunk := f.read(1 << 20):
            cp.write(chunk)


def load(seed_dir, workbook=None):
    seed = pathlib.Path(seed_dir)
    if workbook:
        importer.load(str(workbook), reset=True)
    with connect() as con, con.cursor() as cur:
        cur.execute(importer.SCHEMA.read_text())
        cur.execute("TRUNCATE " + ",".join(DATA_TABLES) + " RESTART IDENTITY CASCADE")

        # ---- products: BIS gives the full master; sales and Promo_Hist add any product BIS does not carry
        cur.execute("""CREATE TEMP TABLE p_bis (product_code text, description text, l1 text, l2 text, l3 text, vendor text,
                       national_price numeric, cost_price numeric, moq numeric, variant_count int) ON COMMIT DROP""")
        copy_csv(cur, seed / "products_bis.csv.gz", "p_bis")
        cur.execute("""CREATE TEMP TABLE p_sales (product_code text, description text, l1 text, l2 text, l3 text, vendor text) ON COMMIT DROP""")
        copy_csv(cur, seed / "products_sales.csv.gz", "p_sales")
        cur.execute("""CREATE TEMP TABLE p_hist (product_code text, product_name text, l1 text, l2 text, l3 text, vendor text,
                       classification text, product_owner text, national_price numeric, vendor_cost numeric, promo_month date) ON COMMIT DROP""")
        copy_csv(cur, seed / "products_history.csv.gz", "p_hist")
        cur.execute("""INSERT INTO product_master(product_code,description,l1,l2,l3,vendor,national_price,cost_price,moq,variant_count)
                       SELECT product_code,description,l1,l2,l3,vendor,national_price,cost_price,moq::int,variant_count FROM p_bis""")
        cur.execute("""INSERT INTO product_master(product_code,description,l1,l2,l3,vendor)
                       SELECT product_code,description,l1,l2,l3,vendor FROM p_sales ON CONFLICT DO NOTHING""")
        cur.execute("""INSERT INTO product_master(product_code,description,l1,l2,l3,vendor,national_price,cost_price)
                       SELECT product_code,product_name,l1,l2,l3,vendor,national_price,vendor_cost FROM p_hist ON CONFLICT DO NOTHING""")
        # owner / classification only exist in the promotion history (latest value per product)
        cur.execute("""UPDATE product_master pm SET product_owner=h.product_owner, classification=h.classification
                       FROM p_hist h WHERE h.product_code=pm.product_code""")
        cur.execute("""UPDATE product_master pm SET national_price=h.national_price FROM p_hist h
                       WHERE h.product_code=pm.product_code AND pm.national_price IS NULL""")
        cur.execute("""UPDATE product_master pm SET cost_price=h.vendor_cost FROM p_hist h
                       WHERE h.product_code=pm.product_code AND pm.cost_price IS NULL""")

        copy_csv(cur, seed / "locations.csv.gz", "locations", ["code", "name", "region", "warehouse_type"])
        copy_csv(cur, seed / "replen_settings.csv.gz", "replen_settings",
                 ["product_code", "location_code", "range_type", "replen_method", "stopped_for_purchase", "open_for_replen",
                  "planning_floor", "main_warehouse", "supply_tier"])
        copy_csv(cur, seed / "location_stock.csv.gz", "location_stock", ["product_code", "location_code", "on_hand_qty", "on_hand_value"])
        copy_csv(cur, seed / "sales_monthly.csv.gz", "sales_monthly", ["product_code", "region", "month", "qty", "revenue", "gm", "customers"])
        with gzip.open(seed / "promotions_history.csv.gz", "rt") as f:
            cols = f.readline().strip().split(",")
        copy_csv(cur, seed / "promotions_history.csv.gz", "promotions", cols)
        cur.execute("""SELECT (SELECT count(*) FROM product_master) products, (SELECT count(*) FROM locations) locations,
                       (SELECT count(*) FROM replen_settings) replen, (SELECT count(*) FROM location_stock) stock,
                       (SELECT count(*) FROM sales_monthly) sales, (SELECT count(*) FROM promotions) promotions""")
        counts = cur.fetchone()
        con.commit()
    return counts


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--seed", default=str(ROOT / "data" / "seed"))
    ap.add_argument("--workbook", default=str(ROOT / "excel" / "MM_Promo_Tool.xlsx"))
    a = ap.parse_args()
    print(load(a.seed, a.workbook))
