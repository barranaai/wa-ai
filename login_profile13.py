from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def login_google_profile13():
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9223")

    driver = webdriver.Chrome(options=options)

    print("✅ Attached to existing Chrome Profile 13")
    print("Current URL:", driver.current_url)

    return driver