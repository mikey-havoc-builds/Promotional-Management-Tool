import os, subprocess, sys, pathlib
import psycopg
import pytest

ROOT = pathlib.Path(__file__).resolve().parents[1]
BASE = os.environ.get("TEST_DATABASE_BASE", "postgresql://postgres@/{db}?host=/tmp&port=5433")
ADMIN = BASE.format(db="postgres")
TESTDB = "promo_test"
os.environ["DATABASE_URL"] = BASE.format(db=TESTDB)
sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session", autouse=True)
def database():
    with psycopg.connect(ADMIN, autocommit=True) as c:
        c.execute(f"DROP DATABASE IF EXISTS {TESTDB}")
        c.execute(f"CREATE DATABASE {TESTDB}")
    with psycopg.connect(os.environ["DATABASE_URL"]) as c:
        c.execute((ROOT / "db" / "schema.sql").read_text())
    from app import importer
    importer.load(str(ROOT / "excel" / "MM_Promo_Tool.xlsx"), reset=True)
    with psycopg.connect(os.environ["DATABASE_URL"]) as c:       # the sample workbook has no on-hand stock: seed some
        c.execute("INSERT INTO location_stock VALUES ('abc','123',2,18),('abc','234',20,180)")
    yield


@pytest.fixture()
def client():
    from fastapi.testclient import TestClient
    from app.main import app
    with psycopg.connect(os.environ["DATABASE_URL"]) as c:
        c.execute("DELETE FROM promotions")
    return TestClient(app)
