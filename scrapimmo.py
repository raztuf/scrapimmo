import json
import os
import random
import re
import sys
import time
import openpyxl
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, Page

# ── Config ────────────────────────────────────────────────────────────────────
COLUMNS   = ["url", "postal_code", "locality", "price", "done"]

DELAY_MIN = 2.5
DELAY_MAX = 5.0

PROVINCES = {
    "1": {
        "name":          "Liège",
        "url_slug":      "liege/province",
        "postal_ranges": [(4000, 4999)],
        "output_file":   "liege_private_sellers.xlsx",
    },
    "2": {
        "name":          "Namur",
        "url_slug":      "namur/province",
        "postal_ranges": [(5000, 5999)],
        "output_file":   "namur_private_sellers.xlsx",
    },
}

def choose_province() -> dict:
    print("Select province to scrape:")
    for key, p in PROVINCES.items():
        print(f"  {key}. {p['name']}")
    while True:
        choice = input("Enter number: ").strip()
        if choice in PROVINCES:
            return PROVINCES[choice]
        print("Invalid choice, try again.")


def is_target(postal_code: str, ranges: list[tuple[int, int]]) -> bool:
    try:
        code = int(postal_code.strip())
        return any(lo <= code <= hi for lo, hi in ranges)
    except ValueError:
        return False


# ── Search URL builder ────────────────────────────────────────────────────────
BASE_URL     = "https://www.immoweb.be/en/search/{type}/for-sale/{slug}?countries=BE&page={page}&orderBy=newest"
SEARCH_TYPES = ("house", "apartment")


# ── Parse a single search result page ─────────────────────────────────────────
def parse_page(html: str, postal_ranges: list[tuple[int, int]]) -> list[dict]:
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

        if not is_target(postal_code, postal_ranges):
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


def sleep_progress(seconds: float, label: str = "cooldown"):
    bar_width = 30
    start = time.time()
    end   = start + seconds
    while True:
        elapsed = time.time() - start
        ratio   = min(elapsed / seconds, 1.0)
        filled  = int(bar_width * ratio)
        bar     = "█" * filled + "░" * (bar_width - filled)
        remaining = max(0, seconds - elapsed)
        sys.stdout.write(f"\r  [{ts()}] {label}  [{bar}] {remaining:.0f}s remaining  ")
        sys.stdout.flush()
        if time.time() >= end:
            break
        time.sleep(0.5)
    sys.stdout.write("\n")


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
                sleep_progress(wait, f"WARN  Rate limited")
            elif status == 403:
                wait = random.uniform(60, 120) * attempt
                sleep_progress(wait, f"WARN  403 attempt {attempt}/3")
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
    province    = choose_province()

    while True:
        raw = input("Start at page (default 1): ").strip()
        if raw == "":
            start_page = 1
            break
        if raw.isdigit() and int(raw) >= 1:
            start_page = int(raw)
            break
        print("Invalid, enter a number >= 1.")

    output_file = province["output_file"]
    slug        = province["url_slug"]
    postal_ranges = province["postal_ranges"]
    prov_name   = province["name"]

    seen_urls = load_existing_urls(output_file)
    if seen_urls:
        print(f"[{ts()}] Loaded {len(seen_urls)} existing URLs — will skip these\n")

    total_new = 0
    print(f"[{ts()}] Scraping {prov_name} — {len(SEARCH_TYPES)} types, page 1 onwards until empty\n")

    with sync_playwright() as p:
        print(f"[{ts()}] Launching Chrome...")
        browser = p.chromium.launch(
            channel="chrome",
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-infobars",
            ],
        )
        print(f"[{ts()}] Chrome launched, opening page...")
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="fr-BE",
            viewport={"width": 1280, "height": 800},
        )
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        page = context.new_page()
        print(f"[{ts()}] Page ready, starting scrape...")

        for search_type in SEARCH_TYPES:
            print(f"[{ts()}] ── {search_type.upper()} ──────────────────────────────────")
            pg = start_page - 1
            while True:
                pg += 1
                url = BASE_URL.format(type=search_type, slug=slug, page=pg)
                print(f"[{ts()}] Page {pg:>3}  fetching search results...")

                # ── Cooldown every 10 pages ───────────────────────────────────
                if pg % 10 == 0:
                    cooldown = random.uniform(30, 60)
                    sleep_progress(cooldown, f"[{ts()}] Page {pg:>3}  cooldown")

                try:
                    html = fetch(page, url)
                    if html is None:
                        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
                        continue

                    # Random scroll to appear human
                    page.evaluate("window.scrollTo(0, Math.random() * document.body.scrollHeight)")
                    time.sleep(random.uniform(0.5, 1.5))

                    soup = BeautifulSoup(html, "html.parser")
                    if not soup.find_all("article", class_="card--result"):
                        print(f"[{ts()}] Page {pg:>3}  no listings found — done")
                        break

                    candidates     = parse_page(html, postal_ranges)
                    new_candidates = [r for r in candidates if r["url"] not in seen_urls]
                    seen_urls.update(r["url"] for r in new_candidates)

                    print(f"[{ts()}] Page {pg:>3}  {len(candidates)} {prov_name} candidates ({len(new_candidates)} new) — verifying each...")

                    confirmed = []
                    for r in new_candidates:
                        if is_private_owner(page, r["url"]):
                            confirmed.append(r)

                    append_to_xlsx(confirmed, output_file)
                    total_new += len(confirmed)
                    print(f"[{ts()}] Page {pg:>3}  confirmed {len(confirmed)}/{len(new_candidates)} private owners  (running total: {total_new})\n")

                except Exception as e:
                    print(f"[{ts()}] ERROR {url} → {e}")

                time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

        context.close()
        browser.close()

    print(f"[{ts()}] Done — {total_new} new private owner listings appended to '{output_file}'")


# ── Run ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    scrape()
