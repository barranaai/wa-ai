# login.py — STANDARD VERSION for servers with ChromeDriver

import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def login_google():
    chrome_path = "/usr/bin/google-chrome"
    driver_path = "/usr/local/bin/chromedriver"

    options = Options()
    options.binary_location = chrome_path
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-notifications")
    options.add_argument("--window-size=1920,1080")
    # options.add_argument("--headless=new")  # Commented for GUI/VNC support

    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)

    # Login page
    driver.get("https://wceasy.club/staff/")
    print("Please complete Google Login manually...")

    # Wait for successful login by URL check
    try:
        while True:
            if "index.php" in driver.current_url:
                break
            print("Waiting for successful login...")
            time.sleep(5)
    except KeyboardInterrupt:
        print("❌ Interrupted by user.")
        driver.quit()
        exit()

    print("✅ Successfully logged in!")
    return driver

''' 
# login.py — SAFE VERSION with crash handling

import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

def login_google():
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = uc.Chrome(options=options)

    driver.get("https://wceasy.club/staff/")
    print("Please complete Google Login manually...")

    # Wait for manual login & 2FA
    try:
        while True:
            try:
                if "index.php" in driver.current_url:
                    break
                print("Waiting for successful login...")
                time.sleep(5)
            except WebDriverException:
                print("❌ Chrome window was closed or crashed during login.")
                driver.quit()
                exit()
    except KeyboardInterrupt:
        print("❌ Interrupted by user.")
        driver.quit()
        exit()

    print("✅ Successfully logged in!")
    return driver
'''

'''
# login.py — SAFE VERSION

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