import json
import os
import random
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, Page

# ── Config ────────────────────────────────────────────────────────────────────
OUTPUT_FILE = "liege_private_sellers.xlsx"
COLUMNS     = ["url", "postal_code", "locality", "price", "done"]

DELAY_MIN   = 2.5
DELAY_MAX   = 5.0

# ── Target postal code range ──────────────────────────────────────────────────
TARGET_RANGES = [
    (4000, 4999),
]

def is_target(postal_code: str) -> bool:
    try:
        code = int(postal_code.strip())
        return any(lo <= code <= hi for lo, hi in TARGET_RANGES)
    except ValueError:
        return False


# ── Search URL builder ────────────────────────────────────────────────────────
BASE_URL     = "https://www.immoweb.be/en/search/{type}/for-sale/liege/province?countries=BE&page={page}&orderBy=newest"
SEARCH_TYPES = ("house", "apartment")


# ── Parse a single search result page ─────────────────────────────────────────
def parse_page(html: str) -> list[dict]:
    soup     = BeautifulSoup(html, "html.parser")
    listings = soup.find_all("article", class_="card--result")
    results  = []

    for listing in listings:
        # ── Filter: skip agency listings ──────────────────────────────────────
        if listing.find("img", src=lambda s: s and "/customers/" in s):
            continue

        # ── Extract URL ────────────────────────────────────────────────────────
        link = listing.find("a", class_="card__title-link")
        if not link or not link.get("href"):
            continue

        # ── Extract locality, filter by target range ───────────────────────────
        locality_tag  = listing.find("p", class_="card--results__information--locality")
        locality_text = locality_tag.get_text(strip=True) if locality_tag else ""
        parts         = locality_text.split(maxsplit=1)
        postal_code   = parts[0] if len(parts) >= 1 else ""
        locality      = parts[1] if len(parts) >= 2 else ""

        if not is_target(postal_code):
            continue

        # ── Extract price ──────────────────────────────────────────────────────
        price_tag = listing.find("p", class_="card--result__price")
        price = price_tag.get_text(strip=True) if price_tag else ""

        results.append({
            "url":         link["href"],
            "postal_code": postal_code,
            "locality":    locality,
            "price":       price,
        })

    return results


# ── Helpers ───────────────────────────────────────────────────────────────────
def ts():
    return datetime.now().strftime("%H:%M:%S")


def fetch(page: Page, url: str) -> str | None:
    """Navigate to url, return HTML string or None on failure."""
    for attempt in range(1, 4):
        try:
            response = page.goto(url, wait_until="domcontentloaded", timeout=30_000)
            status   = response.status if response else 0

            if status == 200:
                return page.content()
            elif status == 429:
                wait = random.uniform(45, 90)
                print(f"  [{ts()}] WARN  Rate limited — sleeping {wait:.0f}s")
                time.sleep(wait)
            elif status == 403:
                wait = random.uniform(15, 30) * attempt
                print(f"  [{ts()}] WARN  403 (attempt {attempt}/3) — sleeping {wait:.0f}s")
                time.sleep(wait)
            else:
                print(f"  [{ts()}] FAIL  HTTP {status} — gave up on {url}")
                return None

        except Exception as e:
            wait = 10 * attempt
            print(f"  [{ts()}] WARN  Error attempt {attempt}/3: {e} — sleeping {wait}s")
            if attempt == 3:
                return None
            time.sleep(wait)

    print(f"  [{ts()}] FAIL  gave up on {url}")
    return None


def is_private_owner(page: Page, listing_url: str) -> bool:
    print(f"    [{ts()}] CHECK {listing_url}")
    html = fetch(page, listing_url)
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    if html is None:
        return False

    match = re.search(r'window\.classified\s*=\s*(\{.*?\})\s*;', html, re.DOTALL)
    if match:
        try:
            data      = json.loads(match.group(1))
            customers = data.get("customers") or []
            result    = bool(customers) and all(c.get("type") == "PRIVATE" for c in customers)
            print(f"    [{ts()}] {'  OK ✓ private owner' if result else '  SKIP agency/unknown'}")
            return result
        except json.JSONDecodeError:
            pass
    print(f"    [{ts()}]   SKIP could not parse window.classified")
    return False


# ── Excel helpers ──────────────────────────────────────────────────────────────
def load_existing_urls(filepath: str) -> set[str]:
    if not os.path.exists(filepath):
        return set()
    wb     = load_workbook(filepath)
    ws     = wb.active
    header = [cell.value for cell in ws[1]]
    if "url" not in header:
        return set()
    url_col = header.index("url") + 1
    return {ws.cell(row=r, column=url_col).value for r in range(2, ws.max_row + 1)}


def append_to_xlsx(results: list[dict], filepath: str):
    if not results:
        return
    if not os.path.exists(filepath):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(COLUMNS)
    else:
        wb = load_workbook(filepath)
        ws = wb.active
    for r in results:
        ws.append([r.get(c, "") for c in COLUMNS])
    wb.save(filepath)
    print(f"  [{ts()}] SAVED {len(results)} new row(s) → '{filepath}'")


# ── Main scraper ───────────────────────────────────────────────────────────────
def scrape():
    seen_urls = load_existing_urls(OUTPUT_FILE)
    if seen_urls:
        print(f"[{ts()}] Loaded {len(seen_urls)} existing URLs — will skip these\n")

    total_new = 0
    print(f"[{ts()}] Starting scrape — {len(SEARCH_TYPES)} types, page 1 onwards until empty\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="fr-BE",
            viewport={"width": 1280, "height": 800},
        )
        page = context.new_page()

        for search_type in SEARCH_TYPES:
            print(f"[{ts()}] ── {search_type.upper()} ──────────────────────────────────")
            pg = 0
            while True:
                pg += 1
                url = BASE_URL.format(type=search_type, page=pg)
                print(f"[{ts()}] Page {pg:>3}  fetching search results...")

                try:
                    html = fetch(page, url)
                    if html is None:
                        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
                        continue

                    soup = BeautifulSoup(html, "html.parser")
                    if not soup.find_all("article", class_="card--result"):
                        print(f"[{ts()}] Page {pg:>3}  no listings found — done")
                        break

                    candidates     = parse_page(html)
                    new_candidates = [r for r in candidates if r["url"] not in seen_urls]
                    seen_urls.update(r["url"] for r in new_candidates)

                    print(f"[{ts()}] Page {pg:>3}  {len(candidates)} Liège candidates ({len(new_candidates)} new) — verifying each...")

                    confirmed = []
                    for r in new_candidates:
                        if is_private_owner(page, r["url"]):
                            confirmed.append(r)

                    append_to_xlsx(confirmed, OUTPUT_FILE)
                    total_new += len(confirmed)
                    print(f"[{ts()}] Page {pg:>3}  confirmed {len(confirmed)}/{len(new_candidates)} private owners  (running total: {total_new})\n")

                except Exception as e:
                    print(f"[{ts()}] ERROR {url} → {e}")

                time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        context.close()
        browser.close()

    print(f"[{ts()}] Done — {total_new} new private owner listings appended to '{OUTPUT_FILE}'")


# ── Run ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    scrape()
