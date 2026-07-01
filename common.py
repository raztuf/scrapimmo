"""Shared config and helpers for the immoweb scraper and contact sender."""
import json
import os
import random
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Page


# ── .env loader ───────────────────────────────────────────────────────────────
def load_env(path: str | Path | None = None) -> None:
    """Load KEY=VALUE pairs from a .env file into os.environ.

    Dependency-free (no python-dotenv). Existing environment variables win, so
    you can still override a value from the shell. Lines that are blank, start
    with '#', or lack '=' are ignored; surrounding quotes are stripped.
    """
    env_path = Path(path) if path else Path(__file__).with_name(".env")
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


# Load .env as soon as common is imported so any script gets the sender vars.
load_env()

# ── Timing ──────────────────────────────────────────────────────────────────
DELAY_MIN = 2.5
DELAY_MAX = 5.0

# ── Browser config ──────────────────────────────────────────────────────────
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)
LOCALE = "fr-BE"
VIEWPORT = {"width": 1280, "height": 800}
BROWSER_ARGS = [
    "--disable-blink-features=AutomationControlled",
    "--no-sandbox",
    "--disable-dev-shm-usage",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-infobars",
]
WEBDRIVER_INIT_SCRIPT = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"

# ── Province registry ───────────────────────────────────────────────────────
# Hainaut spans two postal ranges (6000-6599 Charleroi/Thuin, 7000-7999
# Mons/Tournai); 6600-6999 belongs to Luxembourg province and is excluded.
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
    "3": {
        "name":          "Hainaut",
        "url_slug":      "hainaut/province",
        "postal_ranges": [(6000, 6599), (7000, 7999)],
        "output_file":   "hainaut_private_sellers.xlsx",
    },
}


# ── Helpers ─────────────────────────────────────────────────────────────────
def ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def sleep(lo=DELAY_MIN, hi=DELAY_MAX):
    time.sleep(random.uniform(lo, hi))


def sleep_progress(seconds: float, label: str = "cooldown"):
    bar_width = 30
    start = time.time()
    end   = start + seconds
    while True:
        elapsed   = time.time() - start
        ratio     = min(elapsed / seconds, 1.0)
        filled    = int(bar_width * ratio)
        bar       = "█" * filled + "░" * (bar_width - filled)
        remaining = max(0, seconds - elapsed)
        sys.stdout.write(f"\r  [{ts()}] {label}  [{bar}] {remaining:.0f}s remaining  ")
        sys.stdout.flush()
        if time.time() >= end:
            break
        time.sleep(0.5)
    sys.stdout.write("\n")


def fetch(page: Page, url: str, timeout: int = 30_000) -> str | None:
    """Navigate to url with 403/429 retry. Returns page HTML or None on failure."""
    for attempt in range(1, 4):
        try:
            response = page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            status   = response.status if response else 0

            if status == 200:
                return page.content()
            elif status == 429:
                sleep_progress(random.uniform(45, 90), "WARN  Rate limited (429)")
            elif status == 403:
                sleep_progress(random.uniform(60, 120) * attempt, f"WARN  403 attempt {attempt}/3")
            else:
                print(f"  [{ts()}] FAIL  HTTP {status} — {url}")
                return None

        except Exception as e:
            wait = 10 * attempt
            print(f"  [{ts()}] WARN  Error attempt {attempt}/3: {e} — sleeping {wait}s")
            if attempt == 3:
                return None
            time.sleep(wait)

    print(f"  [{ts()}] FAIL  gave up on {url}")
    return None


def extract_json_object(html: str, marker: str) -> dict | None:
    """Extract the JSON object assigned after `marker` (e.g. 'window.classified').

    Walks the source counting braces — and ignoring braces inside strings — so it
    survives nested objects, unlike a non-greedy regex that stops at the first '}'.
    """
    i = html.find(marker)
    if i == -1:
        return None
    start = html.find("{", i)
    if start == -1:
        return None

    depth   = 0
    in_str  = False
    escaped = False
    for j in range(start, len(html)):
        c = html[j]
        if in_str:
            if escaped:
                escaped = False
            elif c == "\\":
                escaped = True
            elif c == '"':
                in_str = False
        elif c == '"':
            in_str = True
        elif c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0:
                try:
                    return json.loads(html[start:j + 1])
                except json.JSONDecodeError:
                    return None
    return None
