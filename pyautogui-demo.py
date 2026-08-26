import pyautogui
import time
import pyperclip
from datetime import datetime

# safety delay
pyautogui.PAUSE = 1
time.sleep(3)

# -----------------------------
# STEP 1: Open Chrome
# -----------------------------
pyautogui.press('win')
time.sleep(1)

pyautogui.write('chrome')
pyautogui.press('enter')
time.sleep(5)

# -----------------------------
# STEP 2: Open website
# -----------------------------
pyautogui.write('https://www.google.com')
pyautogui.press('enter')
time.sleep(5)

# -----------------------------
# STEP 3: Copy page title
# -----------------------------
pyautogui.hotkey('ctrl', 'l')  # focus address bar
pyautogui.hotkey('ctrl', 'c')

copied_data = pyperclip.paste()

# -----------------------------
# STEP 4: Open Excel
# -----------------------------
pyautogui.press('win')
time.sleep(1)

pyautogui.write('excel')
pyautogui.press('enter')
time.sleep(8)

# -----------------------------
# STEP 5: Enter data
# -----------------------------
now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

comment = "Automation test - looks good"

pyautogui.write(now)
pyautogui.press('tab')

pyautogui.write(copied_data)
pyautogui.press('tab')

pyautogui.write(comment)

# -----------------------------
# STEP 6: Save file
# -----------------------------
date_str = datetime.now().strftime("%Y-%m-%d")

pyautogui.hotkey('ctrl', 's')
time.sleep(3)

file_name = f"daily_report_{date_str}.xlsx"

pyautogui.write(file_name)
pyautogui.press('enter')
time.sleep(3)

# -----------------------------
# STEP 7: Take screenshot
# -----------------------------
screenshot = pyautogui.screenshot()
screenshot.save(f"daily_report_{date_str}.png")

print("✅ Report created successfully!")