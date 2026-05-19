"""
Avito parser for SUP rental competitor analysis.

Reads search URLs from urls.txt, collects all listing links from each search
page (with pagination), then visits every listing and extracts:
  - title, price, location/district
  - full description (no truncation)
  - seller name, type, rating, reviews count, "on Avito since"
  - photos count, posted date if visible
  - listing category and URL

Outputs to data/competitors.json and data/competitors.csv.

Usage:
    pip install -r requirements.txt
    playwright install chromium
    python avito_parser.py

The script runs in HEADED mode (visible browser) by default - this is
intentional, headless mode triggers Avito's bot detection much more often.
You can keep using your computer while it runs; just don't close the
browser window it opens.
"""

import asyncio
import csv
import json
import random
import re
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeout,
    async_playwright,
)

ROOT = Path(__file__).parent
URLS_FILE = ROOT / "urls.txt"
OUTPUT_DIR = ROOT / "data"

# Be polite. Avito will captcha you fast if you go faster than this.
DELAY_BETWEEN_LISTINGS = (4.0, 9.0)
DELAY_BETWEEN_SEARCH_PAGES = (3.0, 6.0)
PAGE_TIMEOUT_MS = 60_000

# Only keep listings whose URL path contains one of these city slugs.
# Avito mixes results from neighbouring cities when local matches are few.
# Add more slugs (e.g. "ryazanskaya_oblast_*") if you want regional results.
ALLOWED_CITY_SLUGS = ["/ryazan/"]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


def read_urls() -> list[str]:
    if not URLS_FILE.exists():
        print(f"ERROR: {URLS_FILE} not found.", file=sys.stderr)
        sys.exit(1)
    urls = []
    for line in URLS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        urls.append(line)
    if not urls:
        print(f"ERROR: no URLs in {URLS_FILE}", file=sys.stderr)
        sys.exit(1)
    return urls


async def sleep_random(low: float, high: float) -> None:
    await asyncio.sleep(random.uniform(low, high))


async def safe_text(page: Page, selector: str) -> str | None:
    try:
        el = await page.query_selector(selector)
        if not el:
            return None
        text = await el.text_content()
        return text.strip() if text else None
    except Exception:
        return None


async def safe_attr(page: Page, selector: str, attr: str) -> str | None:
    try:
        el = await page.query_selector(selector)
        if not el:
            return None
        return await el.get_attribute(attr)
    except Exception:
        return None


async def wait_for_no_captcha(page: Page) -> bool:
    """If we hit a captcha/firewall page, pause for the human to solve it."""
    body_text = ""
    try:
        body_text = (await page.locator("body").inner_text(timeout=5_000))[:500].lower()
    except Exception:
        pass
    if "доступ ограничен" in body_text or "captcha" in body_text or "робот" in body_text:
        print("\n  >>> Captcha or block detected. Solve it in the browser window.")
        print("  >>> Press ENTER here when you're done...")
        await asyncio.get_event_loop().run_in_executor(None, input)
        return True
    return False


def is_allowed_city(url: str) -> bool:
    """Keep only listings located in our target city/cities."""
    return any(slug in url for slug in ALLOWED_CITY_SLUGS)


async def collect_listing_urls(page: Page, search_url: str) -> list[str]:
    """Open a search results page, paginate, collect every listing URL.

    Stops paginating once Avito starts showing "from other cities" results.
    Filters URLs by ALLOWED_CITY_SLUGS.
    """
    print(f"\n=== Search: {search_url}")
    await page.goto(search_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
    await sleep_random(2, 4)
    await wait_for_no_captcha(page)

    urls: list[str] = []
    seen: set[str] = set()
    page_num = 1
    max_pages = 15  # safety cap

    while page_num <= max_pages:
        # Scroll to make sure all items are loaded.
        for _ in range(4):
            await page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
            await sleep_random(0.4, 0.9)

        # Detect "results from other cities" marker - everything after it is junk.
        try:
            page_text = await page.locator("body").inner_text(timeout=5_000)
        except Exception:
            page_text = ""
        if "встречаются объявления из других городов" in page_text.lower() \
                or "из других городов" in page_text.lower():
            other_cities_marker = True
        else:
            other_cities_marker = False

        # Try multiple selectors for listing items - Avito changes them.
        item_selectors = [
            'div[data-marker="item"] a[data-marker="item-title"]',
            'div[data-marker="item"] a[itemprop="url"]',
            'div[class*="iva-item-root"] a[href*="/items/"]',
        ]
        anchors = []
        for sel in item_selectors:
            anchors = await page.query_selector_all(sel)
            if anchors:
                break

        if not anchors:
            print(f"  page {page_num}: no listings found (selector miss?)")
            break

        added = 0
        skipped_city = 0
        for a in anchors:
            href = await a.get_attribute("href")
            if not href:
                continue
            full = urljoin("https://www.avito.ru", href.split("?")[0])
            if not is_allowed_city(full):
                skipped_city += 1
                continue
            if full not in seen:
                seen.add(full)
                urls.append(full)
                added += 1

        msg = f"  page {page_num}: +{added} listings (total: {len(urls)})"
        if skipped_city:
            msg += f" | skipped {skipped_city} from other cities"
        print(msg)

        # If Avito is now showing other-cities results, no point paginating further.
        if other_cities_marker and added == 0:
            print("  -> reached 'other cities' section, stopping pagination")
            break

        next_btn = await page.query_selector('a[data-marker="pagination-button/nextPage"]')
        if not next_btn:
            break
        href = await next_btn.get_attribute("href")
        if not href:
            break
        next_url = urljoin("https://www.avito.ru", href)
        await sleep_random(*DELAY_BETWEEN_SEARCH_PAGES)
        await page.goto(next_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
        await wait_for_no_captcha(page)
        page_num += 1

    return urls


async def expand_description(page: Page) -> None:
    """Click 'Читать полностью' / 'Показать полностью' if present, so the
    description isn't truncated."""
    expand_selectors = [
        'button:has-text("Читать полностью")',
        'a:has-text("Читать полностью")',
        'button:has-text("Показать полностью")',
        'a:has-text("Показать полностью")',
        'button:has-text("Развернуть")',
        '[data-marker="item-view/expand-description"]',
    ]
    for sel in expand_selectors:
        try:
            loc = page.locator(sel).first
            if await loc.is_visible(timeout=1_000):
                await loc.click(timeout=2_000)
                await asyncio.sleep(0.4)
                return
        except Exception:
            continue


async def parse_listing(page: Page, url: str) -> dict:
    """Open one listing page and extract structured info."""
    record: dict = {"url": url, "scraped_at": datetime.now().isoformat(timespec="seconds")}
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)
        await sleep_random(1.5, 3.0)
        await wait_for_no_captcha(page)
    except PlaywrightTimeout:
        record["error"] = "timeout"
        return record
    except Exception as e:
        record["error"] = str(e)
        return record

    # Expand "Читать полностью" if visible.
    await expand_description(page)

    record["title"] = await safe_text(page, 'h1[data-marker="item-view/title-info"]') \
        or await safe_text(page, "h1")

    record["price"] = await safe_text(page, '[data-marker="item-view/item-price"]') \
        or await safe_attr(page, '[itemprop="price"]', "content")
    record["price_period"] = await safe_text(page, '[data-marker="item-view/item-price-period"]')

    record["location"] = await safe_text(page, '[itemprop="address"]') \
        or await safe_text(page, '[data-marker="item-view/item-address"]')

    # Address block (often more specific than location: street, building).
    record["address"] = await safe_text(page, '[itemprop="streetAddress"]') \
        or await safe_text(page, '[data-marker="delivery-options/address"]')

    record["description"] = await safe_text(page, '[data-marker="item-view/item-description"]') \
        or await safe_text(page, 'div[itemprop="description"]')

    # Category breadcrumbs
    crumbs = []
    crumb_els = await page.query_selector_all('[data-marker="breadcrumbs/link"]')
    for el in crumb_els:
        t = await el.text_content()
        if t:
            crumbs.append(t.strip())
    record["breadcrumbs"] = " > ".join(crumbs) if crumbs else None
    record["category"] = crumbs[-1] if crumbs else None

    record["seller_name"] = await safe_text(page, '[data-marker="seller-info/name"]') \
        or await safe_text(page, 'div[data-marker="seller-info"] a')
    record["seller_type"] = await safe_text(page, '[data-marker="seller-info/label"]')
    record["seller_rating"] = await safe_text(page, '[data-marker="seller-info/score"]') \
        or await safe_text(page, 'span[itemprop="ratingValue"]')

    summary = await safe_text(page, '[data-marker="seller-info/summary"]')
    record["seller_summary"] = summary
    if summary:
        m = re.search(r"(\d[\d\s]*)\s*отзыв", summary)
        if m:
            record["seller_reviews_count"] = int(m.group(1).replace(" ", ""))

    record["seller_registered"] = await safe_text(page, '[data-marker="seller-info/registered"]')

    # Photo count
    photo_els = await page.query_selector_all('[data-marker="image-frame/image-wrapper"]')
    if not photo_els:
        photo_els = await page.query_selector_all('div[class*="gallery-img-frame"]')
    record["photos_count"] = len(photo_els) if photo_els else None

    # When the ad was posted
    record["posted"] = await safe_text(page, '[data-marker="item-view/item-date"]')

    # Item id from URL (e.g. _1234567890)
    m = re.search(r"_(\d{8,})$", url)
    record["item_id"] = m.group(1) if m else None

    # "Ctrl+A on the listing" - capture every visible piece of text in the
    # main listing block. Useful as a safety net so we don't miss anything
    # (extra params, address, schedule, contacts, etc.).
    full_text = None
    for sel in [
        '[data-marker="item-view"]',
        'div[class*="item-view"]',
        "main",
    ]:
        try:
            el = await page.query_selector(sel)
            if el:
                full_text = (await el.inner_text()).strip()
                if full_text:
                    break
        except Exception:
            continue
    record["full_text"] = full_text

    return record


def _sort_key(r: dict) -> tuple:
    """Sort by city slug, then by seller rating desc, then by reviews desc."""
    url = r.get("url") or ""
    city = ""
    m = re.search(r"avito\.ru/([^/]+)/", url)
    if m:
        city = m.group(1)
    rating = 0.0
    rating_str = (r.get("seller_rating") or "").replace(",", ".")
    try:
        rating = float(rating_str)
    except Exception:
        pass
    reviews = -int(r.get("seller_reviews_count") or 0)
    # city asc, rating desc (so negative), reviews desc
    return (city, -rating, reviews, url)


def save_results(results: list[dict]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    json_path = OUTPUT_DIR / "competitors.json"
    csv_path = OUTPUT_DIR / "competitors.csv"

    sorted_results = sorted(results, key=_sort_key)

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(sorted_results, f, ensure_ascii=False, indent=2)

    if sorted_results:
        # Order columns sensibly for human reading in Excel.
        preferred_order = [
            "title", "price", "price_period",
            "location", "address",
            "category", "breadcrumbs",
            "seller_name", "seller_type", "seller_rating",
            "seller_reviews_count", "seller_summary", "seller_registered",
            "photos_count", "posted",
            "description",
            "full_text",
            "url", "item_id", "scraped_at", "error",
        ]
        all_keys = {k for r in sorted_results for k in r.keys()}
        keys = [k for k in preferred_order if k in all_keys]
        keys += sorted(all_keys - set(keys))

        with csv_path.open("w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            for r in sorted_results:
                writer.writerow({k: r.get(k, "") for k in keys})

    print(f"  saved -> {json_path.relative_to(ROOT)} | {csv_path.relative_to(ROOT)}")


def load_existing_results() -> list[dict]:
    """Resume support: load already-scraped listings from previous run."""
    json_path = OUTPUT_DIR / "competitors.json"
    if not json_path.exists():
        return []
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


async def run(playwright: Playwright) -> None:
    urls = read_urls()
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Resume: load what's already been scraped, skip those URLs.
    existing = load_existing_results()
    # Only keep records that match current city filter and have no error.
    existing = [r for r in existing if is_allowed_city(r.get("url", "")) and not r.get("error")]
    already_done: set[str] = {r["url"] for r in existing if r.get("url")}
    if already_done:
        print(f"=== Resume: {len(already_done)} listings already scraped, will skip them.")

    browser: Browser = await playwright.chromium.launch(headless=False)
    context: BrowserContext = await browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 900},
        locale="ru-RU",
    )
    page = await context.new_page()

    all_listings: list[str] = []
    seen: set[str] = set()
    for search_url in urls:
        listings = await collect_listing_urls(page, search_url)
        for lurl in listings:
            if lurl not in seen:
                seen.add(lurl)
                all_listings.append(lurl)

    todo = [u for u in all_listings if u not in already_done]
    print(f"\n=== Found {len(all_listings)} listings in target city(s).")
    print(f"=== {len(todo)} new to scrape, {len(all_listings) - len(todo)} already done.\n")

    results: list[dict] = list(existing)
    for i, lurl in enumerate(todo, 1):
        print(f"[{i}/{len(todo)}] {lurl}")
        rec = await parse_listing(page, lurl)
        results.append(rec)
        if i % 5 == 0 or i == len(todo):
            save_results(results)
        await sleep_random(*DELAY_BETWEEN_LISTINGS)

    save_results(results)
    await browser.close()
    print(f"\nDone. Total records: {len(results)}")


async def main() -> None:
    async with async_playwright() as playwright:
        await run(playwright)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
