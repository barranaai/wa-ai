# main.py

import time
import random
import csv
from selenium.webdriver.common.by import By
from login import login_google


# Random sleep
def random_sleep(min_sec=1, max_sec=3):
    delay = random.uniform(min_sec, max_sec)
    print(f"Sleeping for {delay:.2f} seconds...")
    time.sleep(delay)

# Auto-scroll to load full table
def scroll_to_bottom(driver, pause_time=2):
    print("Scrolling to bottom of page to load full table...")
    last_height = driver.execute_script("return document.body.scrollHeight")

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause_time)

        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            print("Reached bottom of page.")
            break
        last_height = new_height

# === Launch driver ===
driver = login_google()

# === Go to staff dashboard ===
print("Navigating to staff dashboard...")
driver.get("https://wceasy.club/staff/index.php")
random_sleep(3, 6)

# === Check login ===
current_url = driver.current_url
if "index.php" in current_url or "table-view.php" in current_url:
    print("=== Already logged in! Proceeding... ===")
else:
    print("ERROR: Not logged in — please check your Chrome profile.")
    driver.quit()
    exit()

# === Go to Table View ===
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(3, 6)

# === Click first View button ===
print("Clicking first View button...")
try:
    first_view_button = driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]")
    first_view_button.click()
    print("Clicked first View button!")
except Exception as e:
    print("Error finding View button:", e)
    driver.quit()
    exit()

random_sleep(5, 8)

# === Scroll to load all data ===
scroll_to_bottom(driver, pause_time=2)

# === Extract all rows ===
print("Extracting rows...")
rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
print(f"Found {len(rows)} rows.")

# === Extract data ===
data = []
header = ["Sl.No", "Name", "Phone", "Whatsapp", "Email", "Place", "Taxable Rupees", "CGST", "SGST", "Total Tax", "Total Rupees", "Payment Mode", "Reason", "Remark", "Whatsapp_Link"]

for i, row in enumerate(rows):
    cells = row.find_elements(By.TAG_NAME, "td")
    row_data = []

    for cell in cells:
        text = cell.text.strip()
        row_data.append(text)

    # Extract Whatsapp link
    whatsapp_link = ""
    try:
        whatsapp_cell = cells[3]  # Column 4 (index 3)
        link_element = whatsapp_cell.find_element(By.TAG_NAME, "a")
        whatsapp_link = link_element.get_attribute("href")
    except Exception as e:
        whatsapp_link = "NO LINK"

    row_data.append(whatsapp_link)

    print(f"Row {i+1}: {row_data}")
    data.append(row_data)
    random_sleep(0.5, 1)

# === Export to CSV ===
csv_filename = "extracted_data.csv"
with open(csv_filename, mode="w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(header)
    writer.writerows(data)

print(f"=== DONE! All data extracted to {csv_filename} ===")

# === FINISH ===
input("Press Enter to close browser...")
driver.quit()

'''
# working extractor with 100 records
from login import login_google
from navigate import get_first_view_link
from extractor import extract_data_from_view
import time

def main():
    print("=== Barrana AI Scraper: Starting ===")
    
    driver = login_google()

    print("Login complete. Navigating to first View link...")
    view_link, table_id = get_first_view_link(driver)

    print(f"Processing Table ID: {table_id} at URL: {view_link}")
    extract_data_from_view(driver, view_link, table_id)

    print("=== DONE! All data extracted to /data/extracted_data.csv ===")

    driver.quit()

if __name__ == "__main__":
    main()

'''

'''
import time
import random
from selenium.webdriver.common.by import By
from login import get_driver

def random_sleep(min_sec=1, max_sec=3):
    delay = random.uniform(min_sec, max_sec)
    print(f"Sleeping for {delay:.2f} seconds...")
    time.sleep(delay)

# === Launch driver ===
driver = get_driver()

# === Go to staff dashboard ===
print("Navigating to staff dashboard...")
driver.get("https://wceasy.club/staff/index.php")
random_sleep(3, 6)

# === Check login ===
current_url = driver.current_url
if "index.php" in current_url or "table-view.php" in current_url:
    print("=== Already logged in! Proceeding... ===")
else:
    print("ERROR: Not logged in — please check your Chrome profile.")
    driver.quit()
    exit()

# === Go to Table View ===
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(3, 6)

# Example: click first View button
print("Clicking first View button...")
try:
    first_view_button = driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]")
    first_view_button.click()
    print("Clicked first View button!")
except Exception as e:
    print("Error finding View button:", e)
    driver.quit()
    exit()

random_sleep(5, 8)

# === Extract first 5 rows as test ===
print("Extracting first 5 rows...")
rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
print(f"Found {len(rows)} rows.")

num_rows_to_extract = min(5, len(rows))
for i in range(num_rows_to_extract):
    row = rows[i]
    cells = row.find_elements(By.TAG_NAME, "td")
    row_data = []
    for cell in cells:
        text = cell.text.strip()
        row_data.append(text)
    print(f"Row {i+1}:", row_data)
    random_sleep(1, 2)

# === DONE ===
print("=== TEST COMPLETE ===")
input("Press Enter to close browser...")
driver.quit()
''' 
