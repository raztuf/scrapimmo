import os
import random
import openpyxl
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

from common import (
    PROVINCES, BROWSER_ARGS, USER_AGENT, LOCALE, VIEWPORT, WEBDRIVER_INIT_SCRIPT,
    ts, sleep, sleep_progress,
)

# ── Sender info (from environment — see .env) ──────────────────────────────────
FIRST_NAME = os.environ.get("SENDER_FIRST_NAME", "")
LAST_NAME  = os.environ.get("SENDER_LAST_NAME", "")
EMAIL      = os.environ.get("SENDER_EMAIL", "")
PHONE      = os.environ.get("SENDER_PHONE", "")

CONTACT_FILE = "contact.txt"
DEFAULT_XLSX = PROVINCES["2"]["output_file"]  # Namur


def validate_sender() -> bool:
    """Ensure every sender field is set so we never submit a blank contact form."""
    missing = [
        name for name, val in (
            ("SENDER_FIRST_NAME", FIRST_NAME),
            ("SENDER_LAST_NAME",  LAST_NAME),
            ("SENDER_EMAIL",      EMAIL),
            ("SENDER_PHONE",      PHONE),
        ) if not val.strip()
    ]
    if missing:
        print(f"[{ts()}] FAIL  missing sender env vars: {', '.join(missing)}")
        print(f"[{ts()}]       set them (e.g. in .env) before sending.")
        return False
    return True


def resolve_xlsx(province: str | None, file: str | None) -> str:
    if file:
        return file
    if province:
        for p in PROVINCES.values():
            if p["url_slug"].split("/")[0] == province.lower():
                return p["output_file"]
    return DEFAULT_XLSX


# ── Form interaction ──────────────────────────────────────────────────────────
def accept_cookies(page: Page):
    try:
        btn = page.locator(
            "button#didomi-notice-agree-button, "
            "button:has-text('OK'), "
            "button:has-text('Accepter'), "
            "button:has-text('Accept all'), "
            "button:has-text('Tout accepter')"
        ).first
        btn.wait_for(timeout=8_000)
        btn.click()
        sleep(1.5, 1.5)
    except PWTimeout:
        pass


def fill_contact_form(page: Page, url: str, message: str) -> bool:
    """Navigate to a listing, fill and submit the contact form. Returns True on success."""
    response = None
    try:
        response = page.goto(url, wait_until="domcontentloaded", timeout=40_000)
    except Exception as e:
        print(f"  [{ts()}] FAIL  navigation error: {e}")
        return False
    if not response or response.status != 200:
        print(f"  [{ts()}] FAIL  HTTP {response.status if response else '?'} — {url}")
        return False

    accept_cookies(page)

    # Zoom out so the full dialog fits on screen
    page.evaluate("document.body.style.zoom = '0.75'")

    # Scroll gently after cookies dismissed to appear human
    page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.4)")
    sleep(1.5, 2.5)

    # ── Click "Get in touch" to open the modal ────────────────────────────────
    page.evaluate("window.scrollTo(0, document.body.scrollHeight * 0.6)")
    sleep(0.8, 1.2)
    try:
        get_in_touch = page.locator(
            "button:has-text('Get in touch'), "
            "button:has-text('Prendre contact'), "
            "button:has-text('Contacter')"
        ).first
        get_in_touch.wait_for(timeout=10_000)
        get_in_touch.scroll_into_view_if_needed()
        get_in_touch.click()
        sleep(1.0, 2.0)
    except PWTimeout:
        print(f"  [{ts()}] FAIL  'Get in touch' button not found")
        return False

    # ── Wait for dialog then fill fields scoped inside it ────────────────────
    dialog = page.get_by_role("dialog", name="Get in touch")
    try:
        dialog.wait_for(state="visible", timeout=8_000)
    except PWTimeout:
        print(f"  [{ts()}] FAIL  dialog did not open")
        return False

    try:
        dialog.locator("input[name='firstName']").fill(FIRST_NAME)
        dialog.locator("input[name='lastName']").fill(LAST_NAME)
        dialog.locator("input[name='email']").fill(EMAIL)
        dialog.locator("input[name='phone']").fill(PHONE)
        # Select "No" for "I already own a property" — click the label
        no_id = dialog.locator("input[id*='sellingProperty-no']").get_attribute("id")
        dialog.locator(f"label[for='{no_id}']").click()
        # Check "Ask for more info"
        ask_id = dialog.locator("input[id*='askMoreInfo']").get_attribute("id")
        dialog.locator(f"label[for='{ask_id}']").click()
        dialog.locator("textarea[name='message']").fill(message)
        sleep(1.0, 2.0)
    except PWTimeout as e:
        print(f"  [{ts()}] FAIL  form field not found: {e}")
        return False

    # ── Click submit and wait for dialog to close ────────────────────────────
    try:
        submit_btn = dialog.locator("button.sendMyRequestInModalButton")
        submit_btn.wait_for(state="attached", timeout=5_000)
        submit_btn.scroll_into_view_if_needed()
        sleep(0.5, 1.0)
        submit_btn.click()
    except PWTimeout:
        print(f"  [{ts()}] FAIL  submit button not found")
        return False

    # Success = dialog closes OR shows a success message inside
    try:
        page.locator(
            "text=Request successfully sent, "
            "text=successfully sent, "
            "text=Demande envoyée, "
            "text=envoyée avec succès, "
            "text=Votre message a été envoyé"
        ).first.wait_for(timeout=10_000)
        return True
    except PWTimeout:
        pass
    try:
        dialog.wait_for(state="hidden", timeout=5_000)
        return True
    except PWTimeout:
        print(f"  [{ts()}] FAIL  no success confirmation detected")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────
def main(xlsx_file=DEFAULT_XLSX, limit=None):
    if not validate_sender():
        return
    if not os.path.exists(xlsx_file):
        print(f"[{ts()}] FAIL  file not found: {xlsx_file}")
        return
    if not os.path.exists(CONTACT_FILE):
        print(f"[{ts()}] FAIL  message template not found: {CONTACT_FILE}")
        return

    with open(CONTACT_FILE, encoding="utf-8") as f:
        message = f.read().strip()

    wb = openpyxl.load_workbook(xlsx_file)
    ws = wb.active

    header = [cell.value for cell in ws[1]]
    url_col  = header.index("url")  + 1
    done_col = header.index("done") + 1

    rows = [
        (r, ws.cell(r, url_col).value)
        for r in range(2, ws.max_row + 1)
        if not ws.cell(r, done_col).value
    ]
    if limit:
        rows = rows[:limit]

    total   = len(rows)
    success = 0
    failed  = 0

    print(f"[{ts()}] {total} listings to contact from {xlsx_file}\n")

    PROFILE_DIR = os.path.join(os.path.dirname(__file__), "chrome_profile")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=BROWSER_ARGS,
            user_agent=USER_AGENT,
            locale=LOCALE,
            viewport=VIEWPORT,
        )
        context.add_init_script(WEBDRIVER_INIT_SCRIPT)
        page = context.new_page()

        for idx, (row_num, url) in enumerate(rows, 1):
            print(f"  [{ts()}] [{idx}/{total}] {url}")
            ok = fill_contact_form(page, url, message)

            # Save after every listing so a crash never re-contacts the same seller.
            if ok:
                ws.cell(row_num, done_col).value = "done"
                wb.save(xlsx_file)
                success += 1
                print(f"  [{ts()}]   ✓ sent — progress saved")
            else:
                ws.cell(row_num, done_col).value = "failed"
                wb.save(xlsx_file)
                failed += 1
                print(f"  [{ts()}]   ✗ failed — marked in xlsx")

            # Cooldown every 15 listings
            if idx % 15 == 0:
                wait = random.uniform(60, 120)
                sleep_progress(wait, f"cooldown after {idx} listings")
            else:
                sleep()

        try:
            context.close()
        except Exception:
            pass

    print(f"\n[{ts()}] Done — {success} sent, {failed} failed out of {total}")


if __name__ == "__main__":
    import argparse
    prov_choices = [p["url_slug"].split("/")[0] for p in PROVINCES.values()]
    ap = argparse.ArgumentParser()
    ap.add_argument("--province", choices=prov_choices,
                    help="province to contact; resolves to its xlsx file")
    ap.add_argument("--file", help="explicit xlsx file (overrides --province)")
    ap.add_argument("--test", action="store_true", help="Process only the first listing")
    args = ap.parse_args()
    main(xlsx_file=resolve_xlsx(args.province, args.file),
         limit=1 if args.test else None)
