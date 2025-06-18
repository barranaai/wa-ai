import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By

def login_google():
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = uc.Chrome(options=options)

    driver.get("https://wceasy.club/staff/")

    print("Please complete Google Login manually...")

    # Wait for manual login & 2FA
    while "index.php" not in driver.current_url:
        print("Waiting for successful login...")
        time.sleep(5)

    print("Successfully logged in!")
    return driver