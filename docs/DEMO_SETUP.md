# Running the demo on your laptop

The easiest route is **Docker Desktop**: it runs PostgreSQL and the web app for you, so you don't install Postgres or Python yourself.
(Not yet run end-to-end on a real Docker install - if a step fails, send me the error text.)

## 1. One-time setup (about 15 minutes)

1. **Install Docker Desktop** from https://www.docker.com/products/docker-desktop
   * Windows: accept the "WSL 2" option if asked, restart if prompted.
   * Mac: pick the Apple Silicon or Intel build to match your Mac.
   * Open Docker Desktop and wait until it says it is running (whale icon steady).
2. **Get the code.** Either:
   * with git: `git clone https://github.com/mikey-havoc-builds/Promotional-Management-Tool.git`, or
   * without git: on the repo page click **Code > Download ZIP** (sign in to GitHub first, the repo may be private) and unzip it.
3. **Open a terminal in the project folder** (Windows: PowerShell; Mac: Terminal), e.g. `cd Promotional-Management-Tool`.

## 2. Start it

```bash
docker compose up --build
```

The first run downloads Postgres and builds the app (a few minutes). Leave this window open - it is the running app.
It is ready when you see `Uvicorn running on http://0.0.0.0:8000`.

## 3. Load the sample data from the Excel workbook

Open a **second** terminal in the same folder and run:

```bash
docker compose run --rm web python -m app.importer excel/MM_Promo_Tool.xlsx --reset
```

It prints counts of products, locations, sales rows and so on. (`--reset` clears the tables first, so it also wipes anything entered in the web app.)

## 4. Use it

* Input form: http://localhost:8000/
* Dashboard: http://localhost:8000/dashboard
* Database and API explorer: http://localhost:8000/docs

### A quick demo walk-through
1. On the form pick **Dec 2026**, type `abc` in Product and choose `abc - kasdjfiksadf`.
2. Set Promotional type = Discount, an intent, Hotspot = Aisle End, Uplift % = 20, Presentation qty = 2 / 4 / 4, % off = 10.
3. The Historical data, Alerts and "Calculated SOH by location" panels update as you type.
4. Click **Add promotion** (confirm the alerts), then open the **Dashboard** and pick Dec 2026 to see it.

Notes: the current and past months are locked, so from 1 December pick **Jan 2027** or later. The sample workbook has no on-hand
stock or actual sales for future months, so Current SOH is 0 and "Actual" figures are 0 until you load real data.

## 5. Stop, restart, reset

* Stop: press `Ctrl+C` in the first window (or `docker compose down`). Your data is kept.
* Start again: `docker compose up` (no need to re-import).
* Wipe everything and start clean: `docker compose down -v`, then repeat steps 2 and 3.
* After pulling new code: `docker compose up --build`.

## Troubleshooting

| Problem | Fix |
|---|---|
| `Cannot connect to the Docker daemon` | Open Docker Desktop and wait until it is running. |
| `port is already allocated` (5432 or 8000) | Another Postgres/app is using it. Stop it, or change the left-hand port numbers in `docker-compose.yml` (e.g. `"5433:5432"`, `"8001:8000"`) and use that port in the URL. |
| Page loads but dropdowns are empty | You skipped step 3 - run the import. |
| Windows: `docker` not recognised | Reopen PowerShell after installing Docker Desktop. |

## Alternative: no Docker

1. Install PostgreSQL 16 (Mac: `brew install postgresql@16 && brew services start postgresql@16`; Windows: the EnterpriseDB installer, remember the password).
2. Install Python 3.12+ from python.org.
3. In the project folder:
   ```bash
   createdb promo                       # Windows: use pgAdmin or  psql -U postgres -c "create database promo"
   psql promo -f db/schema.sql          # Windows: psql -U postgres -d promo -f db/schema.sql
   pip install -r requirements.txt
   export DATABASE_URL=postgresql://postgres:YOURPASSWORD@localhost:5432/promo     # Windows PowerShell: $env:DATABASE_URL="..."
   python -m app.importer excel/MM_Promo_Tool.xlsx --reset
   uvicorn app.main:app
   ```
4. Open http://localhost:8000/
