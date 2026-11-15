#!/usr/bin/env bash
# Drop and recreate all tables:  DATABASE_URL=postgresql://... ./db/reset.sh
set -e
psql "${DATABASE_URL:?set DATABASE_URL}" -c "DROP SCHEMA public CASCADE; CREATE SCHEMA public;"
psql "$DATABASE_URL" -f "$(dirname "$0")/schema.sql"
