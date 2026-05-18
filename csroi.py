"""
csroi.com Scraper (Firefox Version + Cookie Bypass)
---------------------------------------------------
Selects: Cases, Armory, Stickers filters
Sorts by: Profit Chance
Scrapes: Item Name, Price, Profit Chance, Investing ROI
Saves to: csroi_data.xlsx
Identifies: Top 3 items by highest Investing ROI
 Handles: Cookie consent overlays
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import time
import re


# ─── Config ───────────────────────────────────────────────────────────────────
URL = "https://csroi.com"
OUTPUT_FILE = "csroi_data.xlsx"
SCROLL_PAUSE = 1.5        # seconds between scrolls
MAX_SCROLL_ATTEMPTS = 60  # safety cap for infinite-scroll pages


# ─── Helpers ──────────────────────────────────────────────────────────────────
def accept_cookies(page):
    """Detects and clicks cookie consent acceptance buttons to unblock the UI."""
    print("  → Checking for cookie consent banner…")

    # List of common text and selectors used by popular cookie consent managers
    cookie_selectors = [
        "text=Accept All",
        "text=Accept cookies",
        "text=Accept",
        "text=Allow All",
        "text=Agree",
        "button:has-text('Accept')",
        "button:has-text('Agree')",
        "div[role='button']:has-text('Accept')",
        "[id*='cookie'] button",
        "[class*='cookie'] button"
    ]

    for sel in cookie_selectors:
        try:
            # Look for the element with a short 3-second timeout to prevent stalling
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                page.wait_for_timeout(1500)  # Wait for animation overlay to disappear
                print(f"  ✓ Cookie banner dismissed via selector: {sel}")
                return
        except Exception:
            continue
    print("  ✓ No active blocking cookie banner detected.")


def click_filter(page, label: str):
    """Click a filter tab/button by its visible text (case-insensitive)."""
    selector = f"text={label}"
    try:
        page.click(selector, timeout=8000)
        page.wait_for_timeout(1200)
        print(f"  ✓ Clicked filter: {label}")
    except PlaywrightTimeout:
        # Fall back: find any element whose inner text matches
        els = page.query_selector_all("button, div[role='button'], span")
        for el in els:
            if el.inner_text().strip().upper() == label.upper():
                el.click()
                page.wait_for_timeout(1200)
                print(f"  ✓ Clicked filter (fallback): {label}")
                return
        print(f"  ✗ Could not find filter: {label}")


def select_sort_profit_chance(page):
    """Open the Sort By dropdown and pick 'Profit Chance'."""
    sort_selectors = [
        "text=UNBOX ROI",         # default label shown in the screenshot
        "[aria-label*='Sort']",
        "text=Sort By",
        ".sort-select",
        "select[name*='sort']",
    ]
    opened = False
    for sel in sort_selectors:
        try:
            page.click(sel, timeout=4000)
            page.wait_for_timeout(800)
            opened = True
            print(f"  ✓ Opened Sort By dropdown via: {sel}")
            break
        except Exception:
            continue

    if not opened:
        print("  ✗ Could not open Sort By dropdown — trying generic <select>")
        selects = page.query_selector_all("select")
        for s in selects:
            opts = s.query_selector_all("option")
            for o in opts:
                if "profit" in o.inner_text().lower():
                    s.select_option(value=o.get_attribute("value") or o.inner_text())
                    page.wait_for_timeout(1000)
                    print(f"  ✓ Selected via <select>: {o.inner_text()}")
                    return

    # Now click the "Profit Chance" option in the open dropdown
    profit_selectors = [
        "text=Profit Chance",
        "[data-value*='profit']",
        "li:has-text('Profit Chance')",
        "div[role='option']:has-text('Profit Chance')",
    ]
    for sel in profit_selectors:
        try:
            page.click(sel, timeout=4000)
            page.wait_for_timeout(1200)
            print(f"  ✓ Selected 'Profit Chance' via: {sel}")
            return
        except Exception:
            continue
    print("  ✗ Could not select 'Profit Chance' — sort may remain default")


def parse_roi_to_float(roi_str: str) -> float:
    """Extract ROI as a float value for accurate sorting, handling negative numbers."""
    if not roi_str:
        return -float('inf')
    cleaned = roi_str.replace(",", "").strip()
    match = re.search(r"-?[\d]+\.?\d*", cleaned)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return -float('inf')
    return -float('inf')


def scroll_to_bottom(page):
    """Scroll until no new items load (handles virtual / infinite scroll)."""
    prev_height = 0
    attempts = 0
    while attempts < MAX_SCROLL_ATTEMPTS:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(SCROLL_PAUSE)
        new_height = page.evaluate("document.body.scrollHeight")
        if new_height == prev_height:
            break
        prev_height = new_height
        attempts += 1
    print(f"  ✓ Scrolled to bottom ({attempts} scroll steps)")


def scrape_items(page) -> list[dict]:
    """Extract item data from the loaded page safely and efficiently."""
    items = []

    # ── Strategy 1: Look for card/row elements ──
    card_selectors = [
        ".item-card",
        ".case-card",
        "[class*='ItemCard']",
        "[class*='item-row']",
        "tr[class*='item']",
        ".MuiCard-root",
        "[class*='Card']",
    ]

    cards = []
    for sel in card_selectors:
        cards = page.query_selector_all(sel)
        if cards:
            print(f"  ✓ Found {len(cards)} cards via selector: {sel}")
            break

    if cards:
        for card in cards:
            try:
                # Optimized line-by-line breakdown to avoid parent-child text contamination
                text_blocks = [t.strip() for t in card.inner_text().split('\n') if t.strip()]
                name = price = profit_chance = invest_roi = ""

                # Name: longest non-numeric string line
                for t in text_blocks:
                    if len(t) > len(name) and not re.match(r"^[\d\$\%\.\,\s\+\-]+$", t):
                        name = t

                # Price: contains $ or looks like a numeric price format
                for t in text_blocks:
                    if "$" in t or re.match(r"^\d+\.\d{2}$", t):
                        price = t
                        break

                # Profit Chance & ROI fields
                pct_fields = [t for t in text_blocks if "%" in t]
                if pct_fields:
                    profit_chance = pct_fields[0]
                if len(pct_fields) > 1:
                    invest_roi = pct_fields[1]

                if name:
                    items.append({
                        "Name": name,
                        "Price": price,
                        "Profit Chance": profit_chance,
                        "Investing ROI": invest_roi,
                    })
            except Exception:
                continue

    # ── Strategy 2: Table Rows fallback ──
    if not items:
        print("  → Falling back to table-row strategy")
        rows = page.query_selector_all("table tbody tr")
        for row in rows:
            cells = [td.inner_text().strip() for td in row.query_selector_all("td")]
            if len(cells) >= 3:
                items.append({
                    "Name": cells[0] if len(cells) > 0 else "",
                    "Price": cells[1] if len(cells) > 1 else "",
                    "Profit Chance": cells[2] if len(cells) > 2 else "",
                    "Investing ROI": cells[3] if len(cells) > 3 else "",
                })

    # ── Strategy 3: Data Attributes fallback ──
    if not items:
        print("  → Falling back to data-attribute strategy")
        els = page.query_selector_all("[data-name], [data-item-name]")
        for el in els:
            name = el.get_attribute("data-name") or el.get_attribute("data-item-name") or ""
            price = el.get_attribute("data-price") or ""
            profit = el.get_attribute("data-profit-chance") or el.get_attribute("data-profit") or ""
            roi = el.get_attribute("data-roi") or el.get_attribute("data-invest-roi") or ""
            if name:
                items.append({"Name": name, "Price": price, "Profit Chance": profit, "Investing ROI": roi})

    print(f"  ✓ Scraped {len(items)} items total")
    return items


# ─── Excel export ─────────────────────────────────────────────────────────────
def save_to_excel(items: list[dict], filename: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "CSROI Data"

    # Header styling
    header_fill = PatternFill("solid", start_color="C00000")   # dark red (CS:GO theme)
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="888888")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["#", "Item Name", "Price (USD)", "Profit Chance", "Investing ROI (1Y)"]
    col_widths = [5, 42, 14, 16, 18]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 22

    # Data rows
    alt_fill = PatternFill("solid", start_color="FFF0F0")   # very light red tint
    data_font = Font(name="Arial", size=10)

    for row_idx, item in enumerate(items, start=2):
        fill = alt_fill if row_idx % 2 == 0 else PatternFill()
        row_data = [
            row_idx - 1,
            item.get("Name", ""),
            item.get("Price", ""),
            item.get("Profit Chance", ""),
            item.get("Investing ROI", ""),
        ]
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = border
            cell.fill = fill
            if col_idx == 1:
                cell.alignment = center
            if col_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Freeze header row
    ws.freeze_panes = "A2"

    # Summary row
    summary_row = len(items) + 2
    ws.cell(row=summary_row, column=1, value="Total Items").font = Font(bold=True, name="Arial")
    ws.cell(row=summary_row, column=2, value=len(items)).font = Font(bold=True, name="Arial")

    wb.save(filename)
    print(f"\n  ✓ Saved {len(items)} items → {filename}")


# ─── Top 3 Picking ────────────────────────────────────────────────────────────
def find_and_display_top_three(items: list[dict]):
    """Sorts all items by their evaluated numeric ROI and prints the top 3."""
    sorted_items = sorted(items, key=lambda x: parse_roi_to_float(x.get("Investing ROI", "")), reverse=True)

    print("\n" + "="*55)
    print("🏆 TOP 3 ITEMS WITH THE HIGHEST INVESTING ROI")
    print("="*55)
    for i, item in enumerate(sorted_items[:3], start=1):
        name = item.get("Name", "Unknown")
        price = item.get("Price", "N/A")
        roi = item.get("Investing ROI", "N/A")
        print(f" {i}. {name}")
        print(f"    ↳ Price: {price} | Investing ROI: {roi}")
    print("="*55 + "\n")


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  CSROI.com Scraper (Firefox Version)")
    print("=" * 55)

    with sync_playwright() as p:
        browser = p.webkit.launch(headless=False)
        context = browser.new_context(
            viewport={"width": 1400, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) "
                "Gecko/20100101 Firefox/125.0"
            ),
        )
        page = context.new_page()

        # ── 1. Load the site ──────────────────────────────────────────────────
        print("\n[1/5] Loading csroi.com …")
        page.goto(URL, wait_until="networkidle", timeout=45000)
        page.wait_for_timeout(2500)

        # ── Cookie Mitigation Layer ───────────────────────────────────────────
        accept_cookies(page)

        # ── 2. Apply filters ─────────────────────────────────────────────────
        print("\n[2/5] Applying filters …")
        for label in ["CASES", "ARMORY", "STICKERS"]:
            click_filter(page, label)

        # ── 3. Sort by Profit Chance ─────────────────────────────────────────
        print("\n[3/5] Setting sort → Profit Chance …")
        select_sort_profit_chance(page)
        page.wait_for_timeout(2000)

        # ── 4. Scroll to load all items ───────────────────────────────────────
        print("\n[4/5] Scrolling to load all items …")
        scroll_to_bottom(page)
        page.wait_for_timeout(1500)

        # ── 5. Scrape ─────────────────────────────────────────────────────────
        print("\n[5/5] Scraping item data …")
        items = scrape_items(page)

        browser.close()

    if not items:
        print("\n⚠  No items scraped. The site structure may have changed.")
        print("   Open the browser (headless=False) and inspect element selectors.")
        return

    # Write all data out to Excel sheet
    save_to_excel(items, OUTPUT_FILE)

    # Pick and display the 3 items with the highest ROI
    find_and_display_top_three(items)

    print("✅ Done! Processes finished successfully.")


if __name__ == "__main__":
    main()