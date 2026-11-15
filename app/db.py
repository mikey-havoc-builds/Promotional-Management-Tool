import os
import psycopg
from psycopg.rows import dict_row

DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://postgres:postgres@localhost:5432/promo")


def connect():
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)
