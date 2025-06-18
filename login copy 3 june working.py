import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By

# CONFIGURATION
USER_DATA_DIR = "/Users/faran/Library/Application Support/Google/Chrome"
PROFILE_NAME = "Profile 13"
TARGET_URL = "https://wceasy.club/staff/"

# Delay helper
def random_delay(min_seconds=2, max_seconds=5):
    delay = (max_seconds - min_seconds) * 0.5 + min_seconds
    print(f"[DELAY] Sleeping for {delay:.2f} sec...")
    time.sleep(delay)

def login_google():
    print("=== Launching Chrome with PROFILE ===")
    options = uc.ChromeOptions()
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={PROFILE_NAME}")
    options.add_argument("--start-maximized")
    options.headless = False

    driver = uc.Chrome(options=options)
    print("=== Chrome launched with profile! ===")

    random_delay(2, 4)

    # === Check if TARGET_URL is already open ===
    all_tabs = driver.window_handles
    target_found = False

    print("=== Checking existing tabs for target URL ===")
    for tab in all_tabs:
        driver.switch_to.window(tab)
        try:
            current_url = driver.execute_script("return window.location.href;")
        except Exception:
            current_url = "unknown"

        print("Tab:", current_url)

        if current_url.startswith(TARGET_URL):
            print("=== Target URL already open! Using this tab. ===")
            target_found = True
            break

    # === If not found, open new tab ===
    if not target_found:
        print("=== Target URL NOT open. Opening new tab. ===")
        driver.execute_script(f"window.open('{TARGET_URL}', '_blank');")
        driver.switch_to.window(driver.window_handles[-1])
        random_delay(3, 5)

    # Wait for manual login (Google 2FA)
    print("=== Waiting for manual login & 2FA ===")
    while True:
        current_url = driver.execute_script("return window.location.href;")
        if "index.php" in current_url or "table-view.php" in current_url:
            print("=== LOGIN SUCCESSFUL ===")
            break
        else:
            print("Waiting for login... Current:", current_url)
            time.sleep(5)

    return driver

'''
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
'''