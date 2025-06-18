import os
import time
import random
import csv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium import webdriver
from login import login_google

def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

driver = login_google()
driver.get("https://wceasy.club/staff/table-view.php")
input("⚠️ Enable pop-ups and press Enter to continue...")

# === Output Setup ===
output_dir = "sheet_headers_log"
os.makedirs(output_dir, exist_ok=True)
filename = os.path.join(output_dir, f"sheet_tabs_headers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

processed_sheet_ids = set()

with open(filename, "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Sheet Name", "Sheet HREF", "Tab Name", "Header Column"])

    # === Step 1: Gather unique sheet links ===
    all_links = driver.find_elements(By.XPATH, "//tbody[@id='excel_data']//a[contains(@href, 'view-excel.php?id=')]")
    unique_sheets = []

    for link in all_links:
        try:
            href = link.get_attribute("href")
            sheet_id = href.split("id=")[-1]
            if href and link.is_displayed() and sheet_id not in processed_sheet_ids:
                sheet_name = link.text.strip() or f"Sheet_{sheet_id}"
                unique_sheets.append((sheet_name, href))
                processed_sheet_ids.add(sheet_id)
        except Exception:
            continue

    print(f"✅ Found {len(unique_sheets)} unique sheets.")

    # === Step 2: Loop through each sheet ===
    for sheet_index, (sheet_label, sheet_href) in enumerate(unique_sheets):
        sheet_name = f"{sheet_label}_{sheet_index + 1}".replace(" ", "_")
        print(f"\n=== Opening {sheet_name} ===")

        try:
            driver.get(sheet_href)
            random_sleep(4, 6)
        except Exception as e:
            print(f"❌ Failed to open sheet {sheet_name}: {e}")
            continue

        # === Step 3: Wait for sub-tabs ===
        tabs = []
        for attempt in range(10):
            tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
            if tabs:
                break
            print("⏳ Waiting for sub-tabs to load...")
            time.sleep(2)

        if not tabs:
            print("❌ No tabs found. Skipping.")
            driver.back()
            random_sleep(4, 6)
            continue

        print(f"  ➤ Found {len(tabs)} sub-tabs.")

        for tab_index, tab in enumerate(tabs):
            try:
                tab_name = tab.text.strip().replace(" ", "_") or f"Tab_{tab_index+1}"
                print(f"    • Opening tab: {tab_name}")
                driver.execute_script("arguments[0].click();", tab)
                time.sleep(4)

                header_cells = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]")

                if not header_cells:
                    print(f"      ⚠️ No headers found in {tab_name}")
                    continue

                for idx, h in enumerate(header_cells):
                    p = h.find_elements(By.TAG_NAME, "p")
                    col_text = p[0].text.strip() if p else h.text.strip()
                    header = col_text if col_text else f"Column{idx+1}"
                    writer.writerow([sheet_name, sheet_href, tab_name, header])

                print(f"      ✅ Headers logged: {len(header_cells)} columns")

            except Exception as e:
                print(f"      ❌ Failed on tab {tab_index+1}: {e}")

        driver.back()
        random_sleep(5, 7)

print("\n✅ All sheets and tabs processed with headers.")
print(f"📁 CSV saved at: {filename}")
input("Press Enter to exit...")
driver.quit()