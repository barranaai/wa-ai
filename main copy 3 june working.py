import time
import random
import csv
import os
from selenium.webdriver.common.by import By
from login import login_google
from datetime import datetime

# === Random Sleep ===
def random_sleep(min_sec=1, max_sec=3):
    delay = random.uniform(min_sec, max_sec)
    print(f"Sleeping for {delay:.2f} seconds...")
    time.sleep(delay)

# === Launch browser & login ===
driver = login_google()

# === Go to Table View ===
print("Navigating to Table View...")
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(3, 6)

# === Click First "View" button ===
print("Clicking first View button...")
try:
    first_view_button = driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]")
    first_view_button.click()
    print("Clicked first View button!")
except Exception as e:
    print("ERROR: Cannot find 'View' button:", e)
    driver.quit()
    exit()

random_sleep(5, 8)

# === Detect all Sheet Tabs ===
print("Detecting Sheet Tabs...")
tab_elements = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
print(f"Found {len(tab_elements)} tabs.")

# === Prepare output folder ===
output_folder = "extracted_data"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# === Process each tab ===
for tab_idx, tab in enumerate(tab_elements):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_idx+1}: {tab_name} ===")

    # Click tab
    driver.execute_script("arguments[0].click();", tab)
    random_sleep(4, 7)

    # === Detect column headers ===
    header_elements = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th//p")
    column_names = [h.text.strip() for h in header_elements]
    print("Detected columns:", column_names)

    # === Find column indexes ===
    def get_column_index(col_name):
        try:
            idx = column_names.index(col_name)
            print(f"'{col_name}' column index: {idx}")
            return idx
        except ValueError:
            print(f"WARNING: '{col_name}' column not found!")
            return None

    name_col_index = get_column_index("Name")
    whatsapp_col_index = get_column_index("Whatsapp")

    # === SCROLL TABLE to load all rows ===
    print("Scrolling to load full table...")

    table_wrapper = driver.find_element(By.CLASS_NAME, "tableWraper")
    last_row_count = 0
    max_scroll_tries = 30

    for scroll_try in range(max_scroll_tries):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", table_wrapper)
        random_sleep(1, 2)

        rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
        current_row_count = len(rows)
        print(f"Scroll try {scroll_try+1}: {current_row_count} rows loaded...")

        if current_row_count == last_row_count:
            print("No more rows to load.")
            break
        last_row_count = current_row_count

    # === Extract all rows ===
    print(f"Extracting {current_row_count} rows...")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_folder, f"extracted_{tab_name}_{timestamp}.csv")

    with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(column_names)

        for row_idx, row in enumerate(rows, start=1):
            cells = row.find_elements(By.TAG_NAME, "td")
            row_data = []

            for idx, cell in enumerate(cells):
                text = cell.text.strip()

                if idx == whatsapp_col_index:
                    try:
                        link = cell.find_element(By.TAG_NAME, "a").get_attribute("href")
                        row_data.append(link)
                    except:
                        row_data.append("")
                else:
                    row_data.append(text)

            writer.writerow(row_data)
            print(f"Row {row_idx} extracted.")
            random_sleep(0.2, 0.5)

    print(f"=== DONE Tab {tab_idx+1}: Data saved to {output_file} ===")

# === All Tabs Processed ===
print("\n=== ALL TABS PROCESSED ===")
input("Press Enter to keep browser open...")
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
