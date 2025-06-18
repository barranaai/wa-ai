# navigation module
import time
from selenium.webdriver.common.by import By
from config import TABLE_VIEW_URL, WAIT_TIME

def get_first_view_link(driver):
    driver.get(TABLE_VIEW_URL)
    time.sleep(WAIT_TIME)

    # Wait for table to load
    view_buttons = driver.find_elements(By.XPATH, "//a[contains(@href, 'view-excel.php?id=')]")

    if not view_buttons:
        raise Exception("No View links found!")

    first_view_button = view_buttons[0]
    view_url = first_view_button.get_attribute("href")

    # Extract Table ID from URL
    table_id = view_url.split("id=")[-1]

    return view_url, table_id