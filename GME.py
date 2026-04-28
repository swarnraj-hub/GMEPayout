"""
GME Remit Payments Portal — Full Automation
URL  : https://payments.gmeremit.com/
Flow : Login → KRW → Transactions → Transaction Detail
       → Custom Date Range → Filter → Export Excel (All Pages) → S3 Upload
"""

import asyncio
import os
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import openpyxl
except ImportError:
    sys.exit("Missing dependency: pip install openpyxl")

from playwright.async_api import async_playwright, Page

# ── Config ────────────────────────────────────────────────────────────
_headless = os.environ.get("GME_HEADLESS", "false").lower() == "true"

CONFIG = {
    "USERNAME":     os.environ.get("GME_USERNAME", "Adarsh_T"),
    "PASSWORD":     os.environ.get("GME_PASSWORD", "Adarsh_T@321"),
    "HEADLESS":     _headless,
    "SLOW_MO":      0 if _headless else 120,
    "DATE_SHEET":   os.environ.get("GME_DATE_SHEET", "date_range.xlsx"),
    "DOWNLOAD_DIR": os.path.join(os.path.dirname(os.path.abspath(__file__)), "gmeremit_downloads"),
}

# S3 config (used when UPLOAD_S3=true)
S3_BUCKET         = os.environ.get("S3_BUCKET", "payout-recon")
S3_PREFIX         = os.environ.get("S3_PREFIX", "gme/raw_xlsx")
AWS_REGION        = os.environ.get("AWS_REGION") or "ap-southeast-1"
AWS_ACCESS_KEY_ID = os.environ.get("AWS_ACCESS_KEY_ID", "")
AWS_SECRET_KEY    = os.environ.get("AWS_SECRET_ACCESS_KEY", "")

LOGIN_URL      = "https://payments.gmeremit.com/"
SCREENSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gmeremit_screenshots")

STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
if (!window.chrome) {
    window.chrome = { runtime:{}, loadTimes:function(){}, csi:function(){}, app:{} };
}
const _origQuery = window.navigator.permissions.query;
window.navigator.permissions.query = p =>
    p.name === 'notifications'
        ? Promise.resolve({ state: Notification.permission })
        : _origQuery(p);
"""


# ── Helpers ───────────────────────────────────────────────────────────
async def shot(page: Page, name: str):
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    try:
        await page.screenshot(path=path, full_page=False)
        print(f"    [screenshot] {path}")
    except Exception:
        pass


def read_date_ranges(sheet_path: str) -> list[tuple[str, str]]:
    # Allow direct env-var override — used by GitHub Actions
    start = os.environ.get("START_DATE")
    end   = os.environ.get("END_DATE")
    if start and end:
        print(f"[SHEET] Using env vars: {start} → {end}")
        return [(start, end)]

    wb = openpyxl.load_workbook(sheet_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    start_idx = next(
        (i for i, h in enumerate(header) if h in ("start_date", "start date", "from", "from date")), 0
    )
    end_idx = next(
        (i for i, h in enumerate(header) if h in ("end_date", "end date", "to", "to date")), 1
    )

    pairs = []
    for row in rows[1:]:
        if not row or all(v is None for v in row):
            continue
        start_raw = row[start_idx] if len(row) > start_idx else None
        end_raw   = row[end_idx]   if len(row) > end_idx   else None
        if start_raw is None or end_raw is None:
            continue

        def fmt(v):
            if hasattr(v, "strftime"):
                return v.strftime("%m/%d/%Y")
            return str(v).strip()

        pairs.append((fmt(start_raw), fmt(end_raw)))

    print(f"[SHEET] Loaded {len(pairs)} date range(s) from {sheet_path}")
    return pairs


def upload_to_s3(local_dir: str, end_date: str):
    try:
        import boto3
    except ImportError:
        print("[S3] boto3 not installed — skipping upload.")
        return

    print(f"[S3] Config → bucket={S3_BUCKET!r}  prefix={S3_PREFIX!r}  region={AWS_REGION!r}")

    if not S3_BUCKET:
        print("[S3] S3_BUCKET is empty — skipping upload.")
        return

    files = [f for f in Path(local_dir).glob("*") if f.is_file()]
    print(f"[S3] Files to upload: {[f.name for f in files]}")
    if not files:
        print("[S3] No files found in download directory — nothing to upload.")
        return

    s3_key_prefix = f"{S3_PREFIX}/{end_date}" if S3_PREFIX else f"gme/raw_xlsx/{end_date}"
    client_kwargs = {"region_name": AWS_REGION or "ap-southeast-1"}
    if AWS_ACCESS_KEY_ID and AWS_SECRET_KEY:
        client_kwargs["aws_access_key_id"]     = AWS_ACCESS_KEY_ID
        client_kwargs["aws_secret_access_key"] = AWS_SECRET_KEY

    try:
        s3 = boto3.client("s3", **client_kwargs)
        uploaded = 0
        for fpath in files:
            key = f"{s3_key_prefix}/{fpath.name}"
            print(f"[S3] Uploading {fpath.name} → s3://{S3_BUCKET}/{key}")
            s3.upload_file(str(fpath), S3_BUCKET, key)
            print(f"[S3] ✓ Uploaded: s3://{S3_BUCKET}/{key}")
            uploaded += 1
        print(f"[S3] Done — {uploaded} file(s) uploaded to s3://{S3_BUCKET}/{s3_key_prefix}/")
    except Exception as e:
        print(f"[S3] ERROR: {e}")
        raise


# ── Step: Login ───────────────────────────────────────────────────────
async def do_login(page: Page) -> bool:
    print("\n[LOGIN] Navigating to GME Remit portal...")
    await page.goto(LOGIN_URL, wait_until="domcontentloaded")
    await page.wait_for_timeout(3000)
    await shot(page, "01_landing")

    user_sel = (
        'input[name="username"], input[name="email"], input[name="id"], '
        'input[placeholder*="username" i], input[placeholder*="email" i], '
        'input[placeholder*="id" i], input[type="text"]:visible'
    )
    try:
        await page.wait_for_selector(user_sel, timeout=15000)
        await page.fill(user_sel, CONFIG["USERNAME"])
    except Exception as e:
        print(f"[LOGIN] ERROR: Username input not found — {e}")
        await shot(page, "fail_username")
        return False

    await page.wait_for_timeout(400)

    try:
        await page.wait_for_selector('input[type="password"]', timeout=10000)
        await page.fill('input[type="password"]', CONFIG["PASSWORD"])
    except Exception as e:
        print(f"[LOGIN] ERROR: Password input not found — {e}")
        await shot(page, "fail_password")
        return False

    await shot(page, "02_credentials_filled")

    try:
        await page.click(
            'button[type="submit"], input[type="submit"], '
            'button:has-text("Login"), button:has-text("Sign In"), button:has-text("Log In")',
            timeout=5000,
        )
    except Exception:
        await page.keyboard.press("Enter")

    await page.wait_for_timeout(5000)
    await shot(page, "03_after_login")

    page_text = (await page.inner_text("body")).lower()
    for err in ("invalid", "incorrect", "wrong password", "login failed"):
        if err in page_text:
            print(f"[LOGIN] ERROR: '{err}' detected on page.")
            return False

    print(f"[LOGIN] SUCCESS — URL: {page.url}")
    return True


# ── Step: Click KRW ──────────────────────────────────────────────────
async def click_krw(page: Page) -> bool:
    print("\n[KRW] Looking for KRW selector...")
    await page.wait_for_timeout(2000)
    await shot(page, "04_before_krw")

    try:
        await page.click("text=KRW", timeout=8000)
        print("[KRW] Clicked KRW.")
        await page.wait_for_timeout(2000)
        await shot(page, "05_after_krw")
        return True
    except Exception:
        pass

    for sel in ['a:has-text("KRW")', 'button:has-text("KRW")', 'li:has-text("KRW")',
                'span:has-text("KRW")', '[class*="tab"]:has-text("KRW")']:
        try:
            await page.click(sel, timeout=4000)
            print(f"[KRW] Clicked KRW via: {sel}")
            await page.wait_for_timeout(2000)
            await shot(page, "05_after_krw")
            return True
        except Exception:
            continue

    print("[KRW] ERROR: Could not find KRW element.")
    await shot(page, "fail_krw")
    return False


# ── Step: Navigate to Transaction Detail ─────────────────────────────
async def go_to_transaction_detail(page: Page) -> bool:
    print("\n[NAV] Clicking Transactions in top bar...")
    await shot(page, "06_before_transactions")

    for sel in ['nav a:has-text("Transaction")', 'a:has-text("Transaction")',
                'button:has-text("Transaction")', 'li:has-text("Transaction")']:
        try:
            await page.click(sel, timeout=5000)
            print(f"[NAV] Clicked Transactions via: {sel}")
            await page.wait_for_timeout(2000)
            break
        except Exception:
            continue
    else:
        print("[NAV] ERROR: Transactions menu not found.")
        await shot(page, "fail_transactions")
        return False

    await shot(page, "07_after_transactions_click")

    for sel in ['a:has-text("Transaction Detail")', 'li:has-text("Transaction Detail")',
                'button:has-text("Transaction Detail")', 'a:has-text("Detail")']:
        try:
            await page.click(sel, timeout=5000)
            print(f"[NAV] Clicked Transaction Detail via: {sel}")
            await page.wait_for_timeout(3000)
            break
        except Exception:
            continue
    else:
        print("[NAV] ERROR: Transaction Detail not found.")
        await shot(page, "fail_transaction_detail")
        return False

    await shot(page, "08_transaction_detail_page")
    print(f"[NAV] On Transaction Detail. URL: {page.url}")
    return True


# ── Step: Set Date Range & Export ─────────────────────────────────────
def to_portal_date(date_str: str) -> str:
    from datetime import datetime
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return date_str


async def set_datepicker(page: Page, field_id: str, date_value: str, label: str) -> bool:
    result = await page.evaluate(f"""
        () => {{
            const el = document.getElementById('{field_id}');
            if (!el) return 'not_found';
            el.removeAttribute('readonly');
            el.value = '{date_value}';
            el.dispatchEvent(new Event('input',  {{bubbles: true}}));
            el.dispatchEvent(new Event('change', {{bubbles: true}}));
            return el.value;
        }}
    """)
    print(f"[DATE] {label} → set to: {result}")
    return result not in (None, "not_found", "")


async def set_date_and_export(page: Page, start_date: str, end_date: str) -> bool:
    print(f"\n[DATE] Setting date range: {start_date} → {end_date}")
    await shot(page, f"09_before_date_{start_date.replace('/', '-')}")

    from_val = to_portal_date(start_date)
    to_val   = to_portal_date(end_date)
    print(f"[DATE] Converted → From: {from_val}  To: {to_val}")

    await set_datepicker(page, "AgentTransactionPG_fromDate", from_val, "From Date")
    await page.wait_for_timeout(400)
    await set_datepicker(page, "AgentTransactionPG_toDate",   to_val,   "To Date")
    await page.wait_for_timeout(400)
    await shot(page, "10_dates_entered")

    print("[DATE] Clicking Filter...")
    for sel in ['button:has-text("Filter")', 'input[value="Filter"]',
                'button:has-text("Search")', 'button[type="submit"]']:
        try:
            await page.click(sel, timeout=5000)
            print(f"[DATE] Filter clicked via: {sel}")
            await page.wait_for_timeout(5000)
            break
        except Exception:
            continue

    await shot(page, "11_filtered_results")

    print("[EXPORT] Clicking Excel button...")
    excel_clicked = False
    for sel in ['button:has-text("Excel")', 'a:has-text("Excel")', '[class*="excel"]']:
        try:
            await page.click(sel, timeout=6000)
            print(f"[EXPORT] Excel button clicked via: {sel}")
            excel_clicked = True
            break
        except Exception:
            continue

    if not excel_clicked:
        print("[EXPORT] ERROR: Excel button not found.")
        await shot(page, "fail_excel_btn")
        return False

    await page.wait_for_timeout(2500)
    await shot(page, "12_excel_dropdown")

    print("[EXPORT] Selecting All Pages...")
    for sel in ['text=All Pages', 'span:has-text("All Pages")', 'li:has-text("All Pages")',
                'a:has-text("All Pages")', 'button:has-text("All Pages")']:
        try:
            async with page.expect_download(timeout=30000) as dl_info:
                await page.click(sel, timeout=4000)
            download = await dl_info.value
            filename = download.suggested_filename or f"gme_{from_val}_{to_val}.xlsx"
            save_path = os.path.join(CONFIG["DOWNLOAD_DIR"], filename)
            await download.save_as(save_path)
            print(f"[EXPORT] Downloaded: {save_path}")
            await shot(page, f"13_exported_{start_date.replace('/', '-')}_{end_date.replace('/', '-')}")
            return True
        except Exception:
            continue

    print("[EXPORT] ERROR: 'All Pages' not found or download failed.")
    await shot(page, "fail_all_pages")
    return False


# ── Main ──────────────────────────────────────────────────────────────
async def main():
    os.makedirs(CONFIG["DOWNLOAD_DIR"], exist_ok=True)

    date_ranges = read_date_ranges(CONFIG["DATE_SHEET"])
    if not date_ranges:
        sys.exit(
            "[ERROR] No dates found. Set START_DATE & END_DATE env vars, "
            f"or create '{CONFIG['DATE_SHEET']}' with columns: start_date | end_date"
        )

    # Add --no-sandbox for GitHub Actions
    browser_args = ["--disable-blink-features=AutomationControlled"]
    if CONFIG["HEADLESS"]:
        browser_args += ["--no-sandbox", "--disable-dev-shm-usage"]

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=CONFIG["HEADLESS"],
            slow_mo=CONFIG["SLOW_MO"],
            args=browser_args,
            downloads_path=CONFIG["DOWNLOAD_DIR"],
        )
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/121.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1440, "height": 900},
            locale="en-US",
            accept_downloads=True,
        )
        await context.add_init_script(STEALTH_JS)
        page = await context.new_page()

        if not await do_login(page):
            await browser.close(); sys.exit(1)

        if not await click_krw(page):
            await browser.close(); sys.exit(1)

        if not await go_to_transaction_detail(page):
            await browser.close(); sys.exit(1)

        last_end = None
        for i, (start_date, end_date) in enumerate(date_ranges, 1):
            print(f"\n{'='*55}")
            print(f"  Processing range {i}/{len(date_ranges)}: {start_date} → {end_date}")
            print(f"{'='*55}")
            success = await set_date_and_export(page, start_date, end_date)
            if not success:
                print(f"[WARNING] Export failed for {start_date} → {end_date}. Continuing...")
            last_end = end_date
            await page.wait_for_timeout(3000)

        await browser.close()

    print(f"\n[DONE] All ranges processed. Downloads → {CONFIG['DOWNLOAD_DIR']}")

    if os.environ.get("UPLOAD_S3", "false").lower() == "true":
        print("\n[S3] Uploading downloads...")
        upload_to_s3(CONFIG["DOWNLOAD_DIR"], to_portal_date(last_end or "unknown"))


if __name__ == "__main__":
    asyncio.run(main())
