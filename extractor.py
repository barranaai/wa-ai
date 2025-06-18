# data extraction module
import time
import pandas as pd
from selenium.webdriver.common.by import By
from config import WAIT_TIME

def extract_data_from_view(driver, view_url, table_id):
    driver.get(view_url)
    time.sleep(WAIT_TIME)

    all_data = []

    # Get list of all tabs
    tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")

    print(f"Found {len(tabs)} tabs (Sheets).")

    for idx, tab in enumerate(tabs):
        try:
            sheet_name = tab.text.strip() or f"Sheet {idx+1}"

            print(f"Switching to Sheet: {sheet_name}...")
            tab.click()
            time.sleep(WAIT_TIME)

            # Extract column headers
            headers = driver.find_elements(By.XPATH, "//table[@id='myTable']//thead//th")
            header_names = [h.text.strip() for h in headers]

            print(f"Headers: {header_names}")

            # Extract rows
            rows = driver.find_elements(By.XPATH, "//table[@id='myTable']//tbody//tr")

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                cell_values = [c.text.strip() for c in cells]

                row_dict = {
                    "Table ID": table_id,
                    "Sheet Name": sheet_name
                }

                for col_name, cell_value in zip(header_names, cell_values):
                    row_dict[col_name] = cell_value

                all_data.append(row_dict)

            print(f"Extracted {len(rows)} rows from {sheet_name}.")

        except Exception as e:
            print(f"Error processing tab {idx+1}: {e}")

    # Save to CSV
    df = pd.DataFrame(all_data)
    df.to_csv("data/extracted_data.csv", index=False)
    print("Data saved to data/extracted_data.csv.")