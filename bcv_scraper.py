"""
BCV compra/venta scraper -> static JSON API

1. Opens the BCV "Tipo de cambio de referencia (SMC)" page
2. Finds the link to the newest .xls file (or all files with --all)
3. Downloads it and reads every sheet (one sheet per business day)
4. Extracts compra / venta (Bs per unit of foreign currency) for each currency
5. Merges them into JSON files under ./api (served by GitHub Pages):

   api/latest.json                 latest rate for every currency
   api/<MONEDA>/latest.json        e.g. api/USD/latest.json
   api/<MONEDA>/history.json       every date we have, newest first
   api/<MONEDA>/<YYYY-MM-DD>.json  one specific fecha valor

Usage:
  python bcv_scraper.py              # newest file only (daily run)
  python bcv_scraper.py --all        # every file on the page (backfill)
  python bcv_scraper.py --dry-run    # parse and print, write nothing
  python bcv_scraper.py --file x.xls # parse a local file (for testing)
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from datetime import date, datetime
from urllib.parse import urljoin

import requests
import urllib3
import xlrd
from bs4 import BeautifulSoup

BCV_PAGE = "https://www.bcv.org.ve/estadisticas/tipo-cambio-de-referencia-smc"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; bcv-scraper/1.0)"}
TIMEOUT = 60

# Excel columns holding the Bs/ME rates (0-based: A=0 ... F=5, G=6)
COMPRA_COL = 5  # column F
VENTA_COL = 6   # column G


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #
def http_get(url: str) -> requests.Response:
    """GET with TLS verification first. bcv.org.ve often serves an incomplete
    certificate chain, so if verification fails we retry unverified and warn.
    Set BCV_STRICT_SSL=1 to disable that fallback."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    except requests.exceptions.SSLError:
        if os.getenv("BCV_STRICT_SSL") == "1":
            raise
        print(f"WARN: TLS verification failed for {url}; retrying without it", file=sys.stderr)
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, verify=False)
    r.raise_for_status()
    return r


# --------------------------------------------------------------------------- #
# Step 1: find the Excel links
# --------------------------------------------------------------------------- #
def find_excel_links(page_url: str = BCV_PAGE, max_pages: int = 1) -> list[str]:
    """Return .xls/.xlsx links in page order (the BCV lists the newest first)."""
    links: list[str] = []
    for page in range(max_pages):
        url = page_url if page == 0 else f"{page_url}?page={page}"
        soup = BeautifulSoup(http_get(url).text, "html.parser")
        found = 0
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if re.search(r"\.xlsx?(\?|$)", href, re.I):
                full = urljoin(url, href)
                if full not in links:
                    links.append(full)
                    found += 1
        if found == 0:
            break  # no more pages
    if not links:
        raise RuntimeError(f"No Excel links found on {page_url} - the page layout may have changed")
    return links


# --------------------------------------------------------------------------- #
# Step 2: parse the Excel
# --------------------------------------------------------------------------- #
DATE_RE = re.compile(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})")


def _to_float(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if v > 0 else None
    if isinstance(v, str):
        s = v.strip().replace(" ", "")
        if not s:
            return None
        # Venezuelan format "1.234,5678" -> 1234.5678
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            f = float(s)
            return f if f > 0 else None
        except ValueError:
            return None
    return None


def _find_date(rows: list[list], label: str) -> date | None:
    """Find 'Fecha Valor: dd/mm/yyyy' (or similar) anywhere in the sheet."""
    for row in rows:
        for i, cell in enumerate(row):
            if isinstance(cell, str) and label in cell.lower():
                m = DATE_RE.search(cell)
                if not m and i + 1 < len(row):  # date may sit in the next cell
                    m = DATE_RE.search(str(row[i + 1]))
                if m:
                    d, mth, y = map(int, m.groups())
                    return date(y, mth, d)
    return None


def _date_from_sheet_name(name: str) -> date | None:
    m = re.fullmatch(r"(\d{2})(\d{2})(\d{4})", name.strip())
    if m:
        d, mth, y = map(int, m.groups())
        try:
            return date(y, mth, d)
        except ValueError:
            return None
    return None


def parse_sheet(sheet: "xlrd.sheet.Sheet") -> list[dict]:
    rows = [sheet.row_values(r) for r in range(sheet.nrows)]

    fecha_valor = _find_date(rows, "fecha valor") or _date_from_sheet_name(sheet.name)
    fecha_operacion = _find_date(rows, "fecha operaci")
    if not fecha_valor:
        return []

    out = []
    for row in rows:
        # A currency row starts with a 3-letter ISO code (USD, EUR, CNY...)
        code_idx = next(
            (i for i, c in enumerate(row[:3]) if isinstance(c, str) and re.fullmatch(r"[A-Z]{3}", c.strip())),
            None,
        )
        if code_idx is None:
            continue
        if len(row) <= max(COMPRA_COL, VENTA_COL):
            continue
        compra = _to_float(row[COMPRA_COL])
        venta = _to_float(row[VENTA_COL])
        if compra is None or venta is None:
            continue
        out.append({
            "fecha_valor": fecha_valor.isoformat(),
            "fecha_operacion": fecha_operacion.isoformat() if fecha_operacion else None,
            "moneda": row[code_idx].strip(),
            "compra": round(compra, 8),
            "venta": round(venta, 8),
            "sheet": sheet.name,
        })
    return out


def parse_excel(content: bytes, source_url: str) -> list[dict]:
    book = xlrd.open_workbook(file_contents=content)
    rows: dict[tuple, dict] = {}
    for sheet in book.sheets():
        for r in parse_sheet(sheet):
            r["source_url"] = source_url
            rows[(r["fecha_valor"], r["moneda"])] = r  # dedupe within file
    return list(rows.values())


# --------------------------------------------------------------------------- #
# Step 3: write the static JSON API
# --------------------------------------------------------------------------- #
PUBLIC_FIELDS = ("fecha_valor", "fecha_operacion", "moneda", "compra", "venta")


def _public(r: dict) -> dict:
    return {k: r.get(k) for k in PUBLIC_FIELDS}


def _write(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def write_json_api(rows: list[dict], out_dir: Path) -> int:
    """Merge new rows into api/<MONEDA>/history.json and regenerate the rest.
    History is the source of truth, so old dates are kept across runs."""
    now = datetime.now().astimezone().isoformat(timespec="seconds")
    by_currency: dict[str, list[dict]] = {}
    for r in rows:
        by_currency.setdefault(r["moneda"], []).append(_public(r))

    # include currencies we already have on disk, even if not in this run
    if out_dir.exists():
        for p in out_dir.glob("*/history.json"):
            by_currency.setdefault(p.parent.name, [])

    latest_all = {}
    for moneda, new_rows in by_currency.items():
        hist_path = out_dir / moneda / "history.json"
        merged: dict[str, dict] = {}
        if hist_path.exists():
            for r in json.loads(hist_path.read_text(encoding="utf-8")).get("data", []):
                merged[r["fecha_valor"]] = r
        for r in new_rows:
            merged[r["fecha_valor"]] = r  # new data wins
        history = sorted(merged.values(), key=lambda r: r["fecha_valor"], reverse=True)
        if not history:
            continue

        _write(hist_path, {"moneda": moneda, "updated_at": now, "count": len(history), "data": history})
        _write(out_dir / moneda / "latest.json", {**history[0], "updated_at": now})
        for r in new_rows:
            _write(out_dir / moneda / f"{r['fecha_valor']}.json", r)
        latest_all[moneda] = history[0]

    _write(out_dir / "latest.json", {"updated_at": now, "source": BCV_PAGE, "monedas": latest_all})
    return len(latest_all)


# --------------------------------------------------------------------------- #
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="process every file on the page (backfill)")
    ap.add_argument("--pages", type=int, default=10, help="max listing pages to scan with --all")
    ap.add_argument("--dry-run", action="store_true", help="parse and print, write nothing")
    ap.add_argument("--out", default="api", help="output folder for the JSON API (default: api)")
    ap.add_argument("--file", help="parse a local .xls instead of downloading")
    ap.add_argument("--currency", default=None, help="only keep one currency, e.g. USD")
    args = ap.parse_args()

    if args.file:
        with open(args.file, "rb") as f:
            sources = [(args.file, f.read())]
    else:
        links = find_excel_links(max_pages=args.pages if args.all else 1)
        links = links if args.all else links[:1]
        sources = []
        for link in links:
            print(f"Downloading {link}")
            sources.append((link, http_get(link).content))

    all_rows = []
    for src, content in sources:
        rows = parse_excel(content, src)
        if args.currency:
            rows = [r for r in rows if r["moneda"] == args.currency.upper()]
        print(f"  {src}: {len(rows)} rows")
        all_rows.extend(rows)

    if not all_rows:
        print("ERROR: no rows parsed - the Excel layout may have changed", file=sys.stderr)
        return 1

    latest = max((r for r in all_rows if r["moneda"] == "USD"), key=lambda r: r["fecha_valor"], default=None)
    if latest:
        print(f"Latest USD -> fecha valor {latest['fecha_valor']}: compra {latest['compra']}, venta {latest['venta']}")

    if args.dry_run:
        for r in sorted(all_rows, key=lambda r: (r["fecha_valor"], r["moneda"]))[-10:]:
            print(r)
        return 0

    n = write_json_api(all_rows, Path(args.out))
    print(f"Wrote JSON for {n} currencies to {args.out}/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
