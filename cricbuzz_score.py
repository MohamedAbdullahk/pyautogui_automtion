"""
Fetch the latest live cricket score from Cricbuzz, take a screenshot,
and save the score text to a .txt file.

Run:
    pip install playwright
    playwright install chromium
    python cricbuzz_score.py
"""

from playwright.sync_api import sync_playwright
from datetime import datetime
import os

OUTPUT_DIR = "cricbuzz_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

LIVE_SCORES_URL = "https://www.cricbuzz.com/cricket-match/live-scores"


def get_latest_score():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(viewport={"width": 1366, "height": 900})

        # Navigate to the live scores page (lists all ongoing/recent matches)
        page.goto(LIVE_SCORES_URL, wait_until="domcontentloaded", timeout=60000)

        # Dismiss cookie/consent banner if it appears
        try:
            page.click("text=I Accept", timeout=5000)
        except Exception:
            pass

        # Give the match cards a moment to render (Cricbuzz loads scores via JS)
        page.wait_for_timeout(4000)

        # --- Strategy 1: known Cricbuzz match-card classes ---
        match_text = None
        try:
            page.wait_for_selector("a.cb-lv-scrs-well, div.cb-mtch-lst", timeout=8000)
            first_card = page.locator("a.cb-lv-scrs-well, div.cb-mtch-lst").first
            match_text = first_card.inner_text().strip()
        except Exception:
            pass

        # --- Strategy 2: fallback - grab any element with class containing "cb-lv" ---
        if not match_text:
            try:
                cards = page.locator("[class*='cb-lv']")
                if cards.count() > 0:
                    match_text = cards.first.inner_text().strip()
            except Exception:
                pass

        # --- Strategy 3: last resort - just grab visible page text near the top ---
        if not match_text:
            body_text = page.locator("body").inner_text()
            match_text = "\n".join(body_text.splitlines()[:15])

        # Screenshot of the whole page (change full_page=False for just the viewport)
        screenshot_path = os.path.join(OUTPUT_DIR, "cricbuzz_score.png")
        page.screenshot(path=screenshot_path, full_page=True)

        page.close()
        browser.close()

        return match_text, screenshot_path


def save_to_file(text, screenshot_path):
    txt_path = os.path.join(OUTPUT_DIR, "latest_score.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(f"Fetched at: {timestamp}\n")
        f.write(f"Source: {LIVE_SCORES_URL}\n")
        f.write(f"Screenshot: {screenshot_path}\n")
        f.write("-" * 50 + "\n")
        f.write(text + "\n")
    return txt_path


if __name__ == "__main__":
    score_text, screenshot = get_latest_score()
    txt_file = save_to_file(score_text, screenshot)

    print("Latest score:\n")
    print(score_text)
    print(f"\nScreenshot saved to: {screenshot}")
    print(f"Text saved to: {txt_file}")
