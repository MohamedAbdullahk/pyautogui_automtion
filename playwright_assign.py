"""
WhatsApp Message Sender + Smart Data Extractor
------------------------------------------------
Reads contacts from contacts.xlsx (columns: Name, Phone, Message),
sends each contact a personalized WhatsApp Web message, screenshots
the sent message, extracts the last 3 messages from that chat, and
saves a JSON + Excel report.

SETUP:
    pip install playwright pandas openpyxl
    playwright install chromium

FIRST RUN:
    You will see a QR code in the opened browser window. Scan it with
    your phone (WhatsApp > Linked Devices > Link a Device). The session
    is saved in the "wa_session" folder, so you won't need to scan again
    on later runs (unless you log out or delete that folder).

USAGE:
    python playwright_assign.py

contacts.xlsx format:
    | Name      | Phone          | Message                                  |
    |-----------|----------------|-------------------------------------------|
    | John Doe  | +919876543210  | Hi {name}, your order has shipped!         |
    | Jane Roe  | +919812345678  |                                             |   (blank -> uses DEFAULT_MESSAGE)
"""

import json
import os
import random
import time
from datetime import date, datetime

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
CONTACTS_FILE = "contacts.xlsx"
SESSION_DIR = "wa_session"          # persistent Chromium profile (keeps you logged in)
SCREENSHOT_DIR = "screenshots"
DEFAULT_MESSAGE = "Hello {name}, this is an automated message."
MIN_DELAY, MAX_DELAY = 2, 5         # human-like random delay (seconds) between actions
WHATSAPP_URL = "https://web.whatsapp.com"

os.makedirs(SCREENSHOT_DIR, exist_ok=True)


def human_pause(a=MIN_DELAY, b=MAX_DELAY):
    """Sleep a random amount to avoid robotic, ban-prone timing."""
    time.sleep(random.uniform(a, b))


def load_contacts(path):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"'{path}' not found. Create it with columns: Name, Phone, Message"
        )
    df = pd.read_excel(path, dtype=str).fillna("")
    required = {"Name", "Phone"}
    if not required.issubset(df.columns):
        raise ValueError(f"contacts.xlsx must contain columns: {required}")
    return df.to_dict("records")


def wait_for_login(page, timeout_ms=120_000):
    """
    Waits until WhatsApp Web has finished loading a logged-in session
    (i.e. the chat list / search box is visible). If a QR code is shown
    instead, this blocks until the user scans it.

    Uses several fallback selectors since WhatsApp Web's DOM attributes
    (data-tab indices, aria-labels) change across releases. If nothing
    matches within the timeout, saves a debug screenshot + HTML dump so
    you can inspect what state the page was actually in.
    """
    print("Waiting for WhatsApp Web login (scan the QR code if shown)...")

    # Any one of these appearing means "we're past the QR screen and logged in"
    logged_in_selectors = [
        "div[contenteditable='true'][data-tab='3']",     # older WA Web: top search box
        "div[aria-label='Search input textbox']",
        "div[contenteditable='true'][aria-label='Search input textbox']",
        "div[title='Search input textbox']",
        "#pane-side",                                     # the chat list panel wrapper
        "div[aria-label='Chat list']",
        "canvas[aria-label='Scan me!']",                  # (see note below) -- NOT logged in, just so we can detect it separately
    ]

    poll_interval_ms = 2000
    elapsed = 0
    qr_seen = False

    while elapsed < timeout_ms:
        # Report QR presence once, so you know the page loaded correctly
        if not qr_seen:
            qr = page.locator("canvas[aria-label='Scan me!'], div[data-testid='qrcode']")
            if qr.count() > 0 and qr.first.is_visible():
                qr_seen = True
                print("QR code detected -- scan it with your phone now.")

        # Check for any logged-in indicator (skip the QR selector itself)
        for sel in logged_in_selectors[:-1]:
            loc = page.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                print("Logged in.")
                return

        page.wait_for_timeout(poll_interval_ms)
        elapsed += poll_interval_ms

    # Timed out -- save debug artifacts instead of a bare exception
    debug_png = "wa_login_debug.png"
    debug_html = "wa_login_debug.html"
    page.screenshot(path=debug_png, full_page=True)
    with open(debug_html, "w", encoding="utf-8") as f:
        f.write(page.content())
    raise TimeoutError(
        f"Timed out waiting for WhatsApp Web login after {timeout_ms/1000:.0f}s. "
        f"QR code was {'shown' if qr_seen else 'NOT shown'}. "
        f"Saved '{debug_png}' and '{debug_html}' -- open the screenshot to see what "
        f"state the page was actually in (e.g. WhatsApp's "
        f"'use WhatsApp on this browser?' notice, a network error, or a stale selector)."
    )


def open_chat_by_phone(page, phone):
    """
    Opens a chat directly via WhatsApp's send-by-phone deep link,
    which is more reliable than typing into the search box for
    numbers that may not be existing contacts.
    Returns True if the chat loaded, False if the number is invalid/not on WhatsApp.
    """
    clean_phone = phone.replace("+", "").replace(" ", "").replace("-", "")
    page.goto(f"{WHATSAPP_URL}/send?phone={clean_phone}", wait_until="domcontentloaded")

    # WhatsApp shows an explicit invalid-number popup if the phone isn't on WhatsApp
    try:
        page.wait_for_selector("text=Phone number shared via url is invalid", timeout=6000)
        return False
    except PlaywrightTimeoutError:
        pass

    # Otherwise wait for either the message box (success) or a "Use WhatsApp on your phone" wall
    try:
        page.wait_for_selector(
            "div[contenteditable='true'][data-tab='10'], div[contenteditable='true'][data-tab='6']",
            timeout=20000,
        )
        return True
    except PlaywrightTimeoutError:
        return False


def send_message(page, message_text):
    """Types into the message box and sends. Returns True if a send tick was detected."""
    # The composer box's data-tab index has varied across WA Web versions (6, 9, 10...)
    box = page.locator(
        "div[contenteditable='true'][data-tab='10'], div[contenteditable='true'][data-tab='6']"
    ).last
    box.click()
    # Type in small chunks with tiny delays -- looks more human than instant paste
    for line in message_text.split("\n"):
        box.type(line, delay=random.randint(15, 45))
        page.keyboard.down("Shift")
        page.keyboard.press("Enter")
        page.keyboard.up("Shift")
    # Remove trailing newline then send
    page.keyboard.press("Backspace")
    page.keyboard.press("Enter")

    # Confirm the message left the outbox: look for a "sent" (single/double check) icon
    try:
        page.wait_for_selector(
            "span[data-icon='msg-check'], span[data-icon='msg-dblcheck'], "
            "span[data-icon='msg-dblcheck-ack']",
            timeout=15000,
        )
        return True
    except PlaywrightTimeoutError:
        return False


def extract_last_messages(page, count=3):
    """Grabs the text of the last N message bubbles in the currently open chat."""
    try:
        page.wait_for_selector("div.message-in, div.message-out", timeout=8000)
    except PlaywrightTimeoutError:
        return []

    bubbles = page.locator("div.message-in, div.message-out")
    total = bubbles.count()
    texts = []
    start = max(0, total - count)
    for i in range(start, total):
        try:
            # copyable-text span holds the message body in most WA Web versions
            span = bubbles.nth(i).locator("span.selectable-text")
            if span.count() > 0:
                texts.append(span.first.inner_text().strip())
        except Exception:
            continue
    return texts[-count:]


def run():
    contacts = load_contacts(CONTACTS_FILE)
    results = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            SESSION_DIR,
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = context.new_page()
        page.goto(WHATSAPP_URL)
        wait_for_login(page)

        for contact in contacts:
            name = contact.get("Name", "").strip()
            phone = contact.get("Phone", "").strip()
            template = contact.get("Message", "").strip() or DEFAULT_MESSAGE
            message = template.replace("{name}", name)

            record = {
                "name": name,
                "phone": phone,
                "message": message,
                "status": "pending",
                "error": None,
                "screenshot": None,
                "last_messages": [],
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            }

            print(f"\n--- Processing {name} ({phone}) ---")

            if not phone:
                record["status"] = "failed"
                record["error"] = "Missing phone number"
                results.append(record)
                continue

            try:
                found = open_chat_by_phone(page, phone)
                if not found:
                    record["status"] = "not_found"
                    record["error"] = "Contact/number not found on WhatsApp"
                    results.append(record)
                    human_pause()
                    continue

                human_pause()  # let chat fully render before typing

                sent = send_message(page, message)
                record["status"] = "sent" if sent else "send_unconfirmed"

                human_pause(1, 2)

                # Screenshot of the sent message / chat window
                safe_name = "".join(c if c.isalnum() else "_" for c in name or phone)
                shot_path = os.path.join(SCREENSHOT_DIR, f"{safe_name}.png")
                page.screenshot(path=shot_path)
                record["screenshot"] = shot_path

                # Bonus: extract last 3 messages from this chat
                record["last_messages"] = extract_last_messages(page, count=3)

            except Exception as e:
                record["status"] = "failed"
                record["error"] = str(e)

            results.append(record)
            human_pause()  # human-like gap before moving to the next contact

        context.close()

    save_reports(results)


def save_reports(results):
    today = date.today().isoformat()
    json_path = f"whatsapp_report_{today}.json"
    xlsx_path = f"whatsapp_report_{today}.xlsx"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    # Flatten for the Excel summary (last_messages joined into one cell)
    summary_rows = []
    for r in results:
        summary_rows.append(
            {
                "Name": r["name"],
                "Phone": r["phone"],
                "Status": r["status"],
                "Error": r["error"] or "",
                "Screenshot": r["screenshot"] or "",
                "Last Messages": " | ".join(r["last_messages"]),
                "Timestamp": r["timestamp"],
            }
        )
    pd.DataFrame(summary_rows).to_excel(xlsx_path, index=False)

    print(f"\nReports saved:\n  {json_path}\n  {xlsx_path}")


if __name__ == "__main__":
    run()