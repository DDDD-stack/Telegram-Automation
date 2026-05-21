import re
import time
import subprocess
import schedule
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

# ─── Config ───────────────────────────────────────────────────────────────────
EXCEL_FILE      = "csroi.xlsx"
RECEIVER_EMAIL  = "receiver@gmail.com"        # <-- change this
RUN_TIME        = "09:00"                     # <-- daily run time (24h format)

# Run this in PowerShell to get your username: $env:USERNAME
# Then replace YourUsername below with it
CHROME_PROFILE  = r"C:\Users\YourUsername\AppData\Local\Google\Chrome\User Data"  # <-- change YourUsername


# ─── Helpers ──────────────────────────────────────────────────────────────────
def parse_roi_to_float(roi_str: str) -> float:
    if not roi_str:
        return -float("inf")
    cleaned = str(roi_str).replace(",", "").strip()
    match = re.search(r"-?[\d]+\.?\d*", cleaned)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return -float("inf")
    return -float("inf")


def load_top3_from_excel(filepath: str) -> list[dict]:
    wb = load_workbook(filepath)
    ws = wb.active

    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    def col(name):
        for i, h in enumerate(headers):
            if h and name.lower() in str(h).lower():
                return i + 1
        return None

    name_col  = col("Name")
    price_col = col("Price")
    roi_col   = col("Investing ROI")

    if not all([name_col, price_col, roi_col]):
        raise ValueError(f"Could not find required columns. Headers found: {headers}")

    items = []
    for row in range(2, ws.max_row + 1):
        name  = ws.cell(row=row, column=name_col).value
        price = ws.cell(row=row, column=price_col).value
        roi   = ws.cell(row=row, column=roi_col).value
        if not name or str(name).strip() in ("", "Total Items"):
            continue
        items.append({
            "Name":          str(name).strip(),
            "Price":         str(price).strip() if price else "N/A",
            "Investing ROI": str(roi).strip()   if roi   else "N/A",
        })

    top3 = sorted(items, key=lambda x: parse_roi_to_float(x["Investing ROI"]), reverse=True)[:3]
    return top3


def build_email_body(top3: list[dict]) -> str:
    medals = ["🥇", "🥈", "🥉"]
    lines = ["🏆 TOP 3 CS2 ITEMS — HIGHEST INVESTING ROI\n"]
    for i, item in enumerate(top3):
        lines.append(f"{medals[i]} {item['Name']}")
        lines.append(f"   Price:         {item['Price']}")
        lines.append(f"   Investing ROI: {item['Investing ROI']}\n")
    lines.append("Source: csroi.com")
    return "\n".join(lines)


def kill_chrome():
    """Force close any open Chrome windows so the profile isn't locked."""
    subprocess.call("taskkill /f /im chrome.exe", shell=True)
    time.sleep(2)
    print("  ✓ Chrome closed.")


# ─── Core job ─────────────────────────────────────────────────────────────────
def run_job():
    print(f"\n{'='*50}")
    print(f"  Running job...")
    print(f"{'='*50}")

    print(f"Reading {EXCEL_FILE}…")
    top3 = load_top3_from_excel(EXCEL_FILE)

    if not top3:
        print("⚠  No items found in Excel file.")
        return

    print("\n🏆 Top 3 items by Investing ROI:")
    for i, item in enumerate(top3, 1):
        print(f"  {i}. {item['Name']} | {item['Price']} | {item['Investing ROI']}")

    subject = "🏆 Top 3 CS2 Items — Highest Investing ROI"
    body    = build_email_body(top3)

    # Kill Chrome so the profile isn't locked
    print("\nClosing any open Chrome windows…")
    kill_chrome()

    with sync_playwright() as p:
        print("Launching Chrome with your profile…")
        context = p.chromium.launch_persistent_context(
            user_data_dir=CHROME_PROFILE,
            channel="chrome",
            headless=False,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
                "--profile-directory=Default",
            ],
        )
        page = context.new_page()

        print("Opening Gmail…")
        page.goto("https://mail.google.com")
        page.wait_for_timeout(4000)

        # If not logged in, wait for manual login
        if "accounts.google.com" in page.url:
            print("⚠  Not logged in — please log in manually in the browser window.")
            page.wait_for_url("**/mail/**", timeout=120000)

        page.wait_for_timeout(3000)

        # ── Compose ───────────────────────────────────────────────────────────
        print("Clicking Compose…")
        page.click("div[gh='cm']")
        page.wait_for_timeout(2000)

        # ── To ────────────────────────────────────────────────────────────────
        to_field = page.query_selector("input[name='to'], [aria-label='To']")
        if to_field:
            to_field.click()
            to_field.fill(RECEIVER_EMAIL)
            page.keyboard.press("Tab")
            page.wait_for_timeout(800)

        # ── Subject ───────────────────────────────────────────────────────────
        subj_field = page.query_selector("input[name='subjectbox'], [aria-label='Subject']")
        if subj_field:
            subj_field.click()
            subj_field.fill(subject)
            page.wait_for_timeout(500)

        # ── Body ──────────────────────────────────────────────────────────────
        body_field = page.query_selector("div[aria-label='Message Body'], div[role='textbox'][aria-multiline='true']")
        if body_field:
            body_field.click()
            body_field.fill(body)
            page.wait_for_timeout(500)

        # ── Send ──────────────────────────────────────────────────────────────
        print("Sending email…")
        send_btn = page.query_selector("div[aria-label^='Send']")
        if send_btn:
            send_btn.click()
            page.wait_for_timeout(3000)
            print("✅ Email sent!")
        else:
            print("⚠  Send button not found — please send manually.")
            time.sleep(10)

        context.close()


# ─── Scheduler ────────────────────────────────────────────────────────────────
def main():
    print(f"⏰ Scheduler started — job will run daily at {RUN_TIME}")
    print("   Press Ctrl+C to stop.\n")

    run_job()  # run immediately on startup to verify

    schedule.every().day.at(RUN_TIME).do(run_job)

    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    main()