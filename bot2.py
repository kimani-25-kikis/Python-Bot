import xlwings as xw
import pyautogui
import time

print("🤖 Waking up...")app = xw.apps.active
wb = app.books.active
sheet = wb.sheets['Order']

# Find the last row of data
last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
print(f"Found data up to row {last_row}.")

# Let's give you 3 seconds to click on the Excel window so the bot has focus!
print("⚠️ Click on the Excel window NOW! Starting in 3 seconds...")
time.sleep(3)

# How to write a date in Excel? 
# Let's say you want to change it to today's date.2026-11-16    2026-11-20
 
# Excel likes dates in this format: YYYY-MM-DD
new_order_date = "2026-11-16"
new_ship_date = "2026-11-20"

# Start at row 2 (skip the header)
for row in range(2, last_row + 1):
    
    # --- ORDER DATE (Column C) ---
    print(f"Typing Order Date for row {row}...")
    # 1. Select the cell (the bot "clicks" it)
    sheet.range(f'C{row}').select()
    # 2. Physically press the keys to type the date
    pyautogui.typewrite(new_order_date)
    # 3. Press the 'Tab' key to move to the right (or 'Enter' to go down)
    pyautogui.press('tab') 
    time.sleep(0.1) # A tiny pause so it looks human

    # --- SHIP DATE (Column D) ---
    print(f"Typing Ship Date for row {row}...")
    # 1. Select the cell (Wait, column C auto-moved us to D, but let's be safe)
    sheet.range(f'D{row}').select()
    # 2. Type the date
    pyautogui.typewrite(new_ship_date)
    # 3. Press 'Enter'
    pyautogui.press('enter')
    time.sleep(0.1)

    # Let's make the bot scroll down one row so you can see it working
    sheet.range(f'A{row}').select()

# --- SCROLL BACK TO TOP ---
print("✅ Finished all rows! Scrolling back to the top...")
# Select cell A1 to jump back to the top
sheet.range('A1').select()
# Press Ctrl+Home just to be sure we are at the very top
pyautogui.hotkey('ctrl', 'home')

