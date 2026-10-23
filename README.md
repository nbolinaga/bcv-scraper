# BCV compra/venta API

A free, serverless JSON API for the BCV's compra/venta rates. A GitHub Action
scrapes the BCV page every weekday, reads the newest Excel file (compra = column F,
venta = column G), and commits JSON files. GitHub Pages serves them.

## Endpoints

Base URL: `https://<your-user>.github.io/<repo>/api`

| Endpoint | Returns |
|---|---|
| `/latest.json` | latest compra/venta for every currency |
| `/USD/latest.json` | latest USD rate |
| `/USD/history.json` | every USD date, newest first |
| `/USD/2026-10-26.json` | USD for one fecha valor |

Swap `USD` for `EUR`, `CNY`, etc.

Example `/USD/latest.json`:
```json
{
  "fecha_valor": "2026-10-26",
  "fecha_operacion": "2026-10-23",
  "moneda": "USD",
  "compra": 217.1234,
  "venta": 217.9876,
  "updated_at": "2026-10-23T21:31:02+00:00"
}
```

GitHub Pages sends `Access-Control-Allow-Origin: *`, so you can call it straight
from a browser:
```js
const r = await fetch("https://<your-user>.github.io/<repo>/api/USD/latest.json");
const { compra, venta } = await r.json();
```

## Setup

1. Create a **public** GitHub repo (Pages on private repos needs a paid plan) and push
   this folder, keeping `.github/workflows/bcv-daily.yml`.
2. Repo → Settings → Pages → *Deploy from a branch* → `main` / `(root)` → Save.
3. Repo → Actions → "BCV daily scrape" → Run workflow → tick **backfill**.
4. After a minute or two, open `https://<your-user>.github.io/<repo>/api/latest.json`.

After that it updates itself Mon–Fri at 5:30 PM and 7:30 PM Venezuela time.

## Run locally

```fish
python3 -m venv .venv
source .venv/bin/activate.fish     # bash/zsh: source .venv/bin/activate
pip install -r requirements.txt
python bcv_scraper.py --dry-run --currency USD   # print only
python bcv_scraper.py                            # writes ./api/
python -m http.server 8000                       # then open localhost:8000/api/latest.json
```

## If it breaks

- **"No Excel links found"**: the BCV changed the page. Check `BCV_PAGE` in the script.
- **"no rows parsed"** or wrong numbers: the Excel layout changed. Adjust `COMPRA_COL` /
  `VENTA_COL` near the top of the script (0-based, so F = 5, G = 6).
- **TLS errors**: bcv.org.ve often has a broken certificate chain. The script retries
  without verification and logs a warning (set `BCV_STRICT_SSL=1` to forbid that).
- **Timeouts in GitHub Actions**: bcv.org.ve sometimes blocks foreign IPs. If that
  happens, run the script on a machine in Venezuela and push the `api/` folder from there.
