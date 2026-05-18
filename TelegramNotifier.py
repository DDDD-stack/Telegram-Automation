"""
CSROI Telegram Notifier
-----------------------
Reads: csroi_data.xlsx
Extracts: Top 3 items by highest Investing ROI
Sends: Format-optimized notification via Telegram Bot API
"""

import os
import re
import requests
from openpyxl import load_workbook

# ─── Config ───────────────────────────────────────────────────────────────────
EXCEL_FILE = "csroi_data.xlsx"

# Replace these placeholders with your actual Telegram bot details
TELEGRAM_BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"
TELEGRAM_CHAT_ID = "YOUR_CHAT_ID_HERE"


# ─── Helpers ──────────────────────────────────────────────────────────────────
def parse_roi_to_float(roi_str: str) -> float:
    """Extract ROI as a float value for accurate sorting, handling negative numbers."""
    if not roi_str:
        return -float('inf')

    # Cast explicitly to string in case openpyxl parsed it as something else
    cleaned = str(roi_str).replace(",", "").strip()

    # Check for negative/positive signs and decimal/integer values
    match = re.search(r"-?[\d]+\.?\d*", cleaned)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return -float('inf')
    return -float('inf')


def read_items_from_excel(filename: str) -> list[dict]:
    """Reads data rows from the generated Excel sheet, ignoring summary labels."""
    if not os.path.exists(filename):
        print(f"✗ Error: The file '{filename}' was not found. Please run the scraper first.")
        return []

    wb = load_workbook(filename, data_only=True)
    ws = wb.active
    items = []

    # Iterate through data rows starting from row 2 (skipping the headers)
    for row in range(2, ws.max_row + 1):
        col1_val = ws.cell(row=row, column=1).value
        item_name = ws.cell(row=row, column=2).value

        # Break or skip if we encounter the 'Total Items' summary row or empty spaces
        if col1_val == "Total Items" or not item_name:
            continue

        price = ws.cell(row=row, column=3).value or "N/A"
        profit_chance = ws.cell(row=row, column=4).value or "N/A"
        invest_roi = ws.cell(row=row, column=5).value or "N/A"

        items.append({
            "Name": str(item_name),
            "Price": str(price),
            "Profit Chance": str(profit_chance),
            "Investing ROI": str(invest_roi),
            "ROI_Float": parse_roi_to_float(invest_roi)
        })

    return items


def send_telegram_message(token: str, chat_id: str, text: str):
    """Sends a markdown-formatted message using the Telegram Bot API."""
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown"
    }

    try:
        response = requests.post(url, json=payload, timeout=10)
        if response.status_code == 200:
            print("✓ Telegram notification sent successfully!")
        else:
            print(f"✗ Failed to send Telegram message. API Response: {response.text}")
    except requests.exceptions.RequestException as e:
        print(f"✗ Network error occurred while connecting to Telegram: {e}")


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    print("=" * 55)
    print("  CSROI Top 3 Item Picker & Telegram Sender")
    print("=" * 55)

    # 1. Read data from Excel
    items = read_items_from_excel(EXCEL_FILE)
    if not items:
        return

    # 2. Sort by the float evaluation of ROI descending
    sorted_items = sorted(items, key=lambda x: x["ROI_Float"], reverse=True)
    top_three = sorted_items[:3]

    # 3. Construct a cleanly formatted Markdown text payload
    message_lines = [
        "🏆 *CSROI.COM TOP 3 INVESTING ROI REPORT*",
        "‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾‾"
    ]

    for index, item in enumerate(top_three, start=1):
        message_lines.append(
            f"*{index}. {item['Name']}*\n"
            f" 💰 *Price:* {item['Price']}\n"
            f" 📈 *Investing ROI:* `{item['Investing ROI']}`\n"
            f" 🎲 *Profit Chance:* {item['Profit Chance']}\n"
        )

    message_lines.append(f"📊 _Total tracked items parsed: {len(items)}_")
    telegram_text = "\n".join(message_lines)

    # 4. Fire the notification
    if TELEGRAM_BOT_TOKEN == "YOUR_BOT_TOKEN_HERE" or TELEGRAM_CHAT_ID == "YOUR_CHAT_ID_HERE":
        print("\n⚠️  Please update the script with your real TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID config.")
        print("\nGenerated message draft:\n")
        print(telegram_text)
    else:
        print(f"→ Processing top 3 items out of {len(items)} collected...")
        send_telegram_message(TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, telegram_text)


if __name__ == "__main__":
    main()