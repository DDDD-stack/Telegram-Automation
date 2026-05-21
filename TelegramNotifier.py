import re
import time
import asyncio
import subprocess
import schedule
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

excel = "csroi.xlsx"
receiver = "receiver@gmail.com"
run = "09:00"

def main():
    with sync_playwright() as p:
        wb = load_workbook(excel)
        ws = wb.active

        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto("https://accounts.google.com/")

        page.fill("input[aria-label='Email or phone']", "shqipg46@gmail.com", timeout=10000)
        page.get_by_role("button", name="Next").click()

        page.fill("input[aria-label='Enter your password']", "Denis1234robi12@#")
        page.get_by_role("button", name="Next").click()

        page.click("div[role='button']: has-text('Compose')")
        page.fill("input[aria-label='To']", "drobi840@gmail.com")
        page.fill("input[name='subjectbox']", "Top 3 best investments")
        page.fill("div[role='textbox]", "This is just a test")

        page.click("div[role='button']: has-text('Send')")

        browser.close()

if __name__ == "__main__":
    main()