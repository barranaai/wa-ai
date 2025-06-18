
# fully functional code with AI message generation and WhatsApp integration

import os
import time
import csv
import random
import re
import pyautogui
from datetime import datetime
from dotenv import load_dotenv
from openai import OpenAI
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from login import login_google

# Load API key
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

def get_greeting():
    hour = (datetime.utcnow().hour + 4) % 24
    return "Good Morning" if 5 <= hour < 12 else "Good Afternoon" if 12 <= hour < 17 else "Good Evening"

def generate_ai_message(full_name, last_name, building_name, unit_number, retries=2):
    greeting = get_greeting()
    unit_info = f"your unit {unit_number} in {building_name}" if unit_number and building_name else \
        f"your unit in {building_name}" if building_name else \
        f"your unit {unit_number}" if unit_number else "your property"

    prompt = f"""
    The person's full name is: "{full_name}"
    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g. contains LLC, LIMITED, COMPANY, etc), do NOT use Mr./Ms.
      Instead, use: {greeting}, [Full Name of Company],
    ---
    My name is Omar Bayat, and I"’"m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.
    I’m reaching out regarding {unit_info}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.
    Looking forward to your response. Thank you, and have a lovely day ahead.
    Best regards,  
    Omar Bayat  
    White & Co. Real Estate
    Output ONLY the final message.
    """

    for attempt in range(retries + 1):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=800,
                temperature=0.7
            )
            message = response.choices[0].message.content.strip()

            if len(message.split()) < 20:
                print(f"⚠️ Attempt {attempt+1}: Message too short — retrying...")
                time.sleep(1)
                continue  # Try again

            return message

        except Exception as e:
            print(f"❌ Error generating message (attempt {attempt+1}): {e}")
            time.sleep(1)

    return "{greeting} [Last Name], I'm contacting you regarding your property {unit_info}."

'''
def generate_ai_message(full_name, last_name, building_name, unit_number):
    greeting = get_greeting()
    unit_info = f"your unit {unit_number} in {building_name}" if unit_number and building_name else \
        f"your unit in {building_name}" if building_name else \
        f"your unit {unit_number}" if unit_number else "your property"

    prompt = f"""
    The person's full name is: "{full_name}"
    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g. contains LLC, LIMITED, COMPANY, etc), do NOT use Mr./Ms.
      Instead, use: {greeting}, [Full Name of Company],
    ---
    My name is Omar Bayat, and I’m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.
    I’m reaching out regarding {unit_info}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.
    Looking forward to your response. Thank you, and have a lovely day ahead.
    Best regards,  
    Omar Bayat  
    White & Co. Real Estate
    Output ONLY the final message.
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"ERROR generating message: {e}")
        return ""
    '''

driver = login_google()

# Open WhatsApp Web tab
print("Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com/")
whatsapp_tab = driver.current_window_handle
print("Please scan the QR code...")

# Wait for WhatsApp login
max_wait = 300
elapsed = 0
while elapsed < max_wait:
    try:
        driver.find_element(By.XPATH, "//div[@aria-label='Chat list' or @role='textbox']")
        print("✅ WhatsApp Web login detected!")
        break
    except:
        print("⏳ Waiting for WhatsApp login...")
        time.sleep(5)
        elapsed += 5
else:
    print("❌ WhatsApp login timeout.")
    driver.quit()
    exit()

# Switch to WCEasy
driver.switch_to.new_window('tab')
print("Opening WCEasy Table View...")
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(5, 7)
input("⚠️ Enable pop-ups and press Enter to continue...")

try:
    driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]").click()
except Exception as e:
    print("ERROR: View button not found.", e)
    driver.quit()
    exit()

# === Wait for .tableFooterTabBox to be present ===
tabs = []
for attempt in range(10):  # Retry every 2s up to 20s
    tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
    if tabs:
        break
    print("⏳ Waiting for tabs to load...")
    time.sleep(2)

if not tabs:
    print("❌ No tabs found after retrying. Exiting.")
    driver.quit()
    exit()

print(f"✅ Found {len(tabs)} tabs.")
output_folder = "extracted_data_ai"
os.makedirs(output_folder, exist_ok=True)

for tab_index, tab in enumerate(tabs):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_index+1}: {tab_name} ===")
    driver.execute_script("arguments[0].click();", tab)
    random_sleep(5, 7)

    for _ in range(10):
        try:
            if driver.find_element(By.ID, "table_body").find_elements(By.TAG_NAME, "tr"):
                break
        except:
            time.sleep(1)

    rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
    print(f"Extracting {len(rows)} rows...")

    headers = []
    for h in driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]"):
        p = h.find_elements(By.TAG_NAME, "p")
        headers.append(p[0].text.strip() if p else h.text.strip())
    headers = [h if h else f"Column{idx+1}" for idx, h in enumerate(headers)]

    def match_col(possible): return next((i for i, col in enumerate(headers) if col.lower() in [p.lower() for p in possible]), None)

    idx_name = match_col(['name', 'full name', 'client name', 'owner', 'owner name', 'landlord', 'landlord name', 'contact name', 'property owner'])
    idx_unit = match_col(['unit', 'unit no', 'unit number', 'flat', 'flat no', 'apartment', 'apartment no', 'apartment number', 'unit#', 'flat#', 'apartment#'])
    idx_building = match_col(['building', 'building name', 'tower', 'tower name', 'project', 'project name', 'property', 'property name', 'building/project', 'bldg', 'bldg name', 'development'])

    with open(os.path.join(output_folder, f"{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers + ["AI_Message"])

        for i, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")[1:]
            values = [c.text.strip() for c in cells]
            full_name = values[idx_name] if idx_name is not None else ""
            last_name = full_name.split()[-1] if full_name else ""
            unit_number = values[idx_unit] if idx_unit is not None else ""
            building_name = values[idx_building] if idx_building is not None else ""
            ai_msg = generate_ai_message(full_name, last_name, building_name, unit_number)
            values.append(ai_msg)
            writer.writerow(values)
            print(f"Row {i}: Message generated.")

            try:
                link = row.find_element(By.XPATH, ".//a[contains(@href, 'https://api.whatsapp.com/send?')]")
                href = re.sub(r'phone=\+?\d+', 'phone=923226100103', link.get_attribute("href"))
                href = re.sub(r'text=.*', 'text=Hi', href)

                existing_tabs = driver.window_handles
                driver.execute_script("window.open(arguments[0]);", href)
                time.sleep(2)
                new_tab = list(set(driver.window_handles) - set(existing_tabs))[0]
                driver.switch_to.window(new_tab)

                start = time.time()
                chat_loaded = False
                while time.time() - start < 20:
                    try:
                        for btn in driver.find_elements(By.XPATH, "//a[@id='action-button' or .//span[text()='Continue to Chat']]"):
                            if btn.is_displayed():
                                btn.click()
                                print("✅ Clicked 'Continue to Chat'")
                                time.sleep(2)
                                break
                    except: pass

                    try:
                        for btn in driver.find_elements(By.XPATH, "//a[contains(@href, 'web.whatsapp.com/send') or .//span[text()='use WhatsApp Web']]"):
                            if btn.is_displayed():
                                btn.click()
                                print("✅ Clicked 'Use WhatsApp Web'")
                                time.sleep(2)
                                break
                    except: pass

                    try:
                        driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
                        print("✅ WhatsApp chat loaded.")
                        chat_loaded = True
                        break
                    except: pass
                    time.sleep(1)

                if not chat_loaded:
                    print("❌ Chat did not load — skipping.")
                    driver.close()
                    driver.switch_to.window(driver.window_handles[1])
                    continue

                random_sleep(3, 5)

                # === Robust Message Typing ===
                try:
                    chat_input = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
                    )
                    chat_input.click()

                    if os.name == 'posix':
                        pyautogui.hotkey("command", "a")
                    else:
                        pyautogui.hotkey("ctrl", "a")
                    pyautogui.press("backspace")
                    
                    # ✅ Add a small random wait after clearing input
                    random_sleep(0.5, 1.2)
                    
                    if len(ai_msg.strip().split()) < 5:
                        print(f"⚠️ Warning: Suspiciously short AI message for row {i}: {ai_msg}")

                    print("⌨️ Typing message...")
                    for line in ai_msg.splitlines():
                        pyautogui.typewrite(line)
                        pyautogui.keyDown('shift')
                        pyautogui.press('enter')
                        pyautogui.keyUp('shift')
                    time.sleep(1)                    
                    pyautogui.press('enter')
                    print("✅ Message typed and sent.")
                except Exception as e:
                    print(f"❌ Error typing message: {e}")

                random_sleep(4, 6)
                driver.close()
                driver.switch_to.window(driver.window_handles[1])
            except Exception as e:
                print(f"❌ Failed on row {i}: {e}")
            random_sleep(1, 2)

input("\nAll done. Press Enter to exit...")
driver.quit()


'''
# working with sending messages but only after manually Allowing the pop-up blocked. but Continue to Chat and Use WhatsApp Web are not working. 

import os
import time
import csv
import random
import re
import pyautogui
from datetime import datetime
from dotenv import load_dotenv
from openai import OpenAI
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from login import login_google

# === Load .env ===
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

def get_greeting():
    hour = (datetime.utcnow().hour + 4) % 24
    return "Good Morning" if 5 <= hour < 12 else "Good Afternoon" if 12 <= hour < 17 else "Good Evening"

def generate_ai_message(full_name, last_name, building_name, unit_number):
    greeting = get_greeting()
    unit_info = f"your unit {unit_number} in {building_name}" if unit_number and building_name else \
                f"your unit in {building_name}" if building_name else \
                f"your unit {unit_number}" if unit_number else "your property"

    prompt = f"""
    The person's full name is: "{full_name}"
    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g. contains LLC, LIMITED, COMPANY, etc), do NOT use Mr./Ms.
      Instead, use: {greeting}, [Full Name of Company],
    ---
    My name is Omar Bayat, and I’m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.
    I’m reaching out regarding {unit_info}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.
    Looking forward to your response. Thank you, and have a lovely day ahead.
    Best regards,  
    Omar Bayat  
    White & Co. Real Estate
    Output ONLY the final message.
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"ERROR generating message: {e}")
        return ""

# === LOGIN ===
driver = login_google()

# === Tab 1: Open WhatsApp Web ===
print("Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com/")
whatsapp_tab = driver.current_window_handle
print("Please scan the QR code in WhatsApp Web...")

# === Wait for WhatsApp Login ===
max_wait = 300
elapsed = 0
interval = 5
while elapsed < max_wait:
    try:
        driver.find_element(By.XPATH, "//div[@aria-label='Chat list' or @role='textbox']")
        print("✅ WhatsApp Web login detected!")
        break
    except:
        print("⏳ Waiting for WhatsApp login...")
        time.sleep(interval)
        elapsed += interval
else:
    print("❌ WhatsApp login timeout.")
    driver.quit()
    exit()

# === Tab 2: Open WCEasy Table View ===
driver.switch_to.new_window('tab')
print("Opening WCEasy Table View...")
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(5, 7)

print("⚠️ IMPORTANT: Please allow pop-ups for this site in your Chrome browser.")
for i in range(30, 0, -5):
    print(f"⏳ Waiting {i} seconds for you to allow pop-ups...")
    time.sleep(5)

try:
    driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]").click()
except Exception as e:
    print("ERROR: View button not found.", e)
    driver.quit()
    exit()
random_sleep(6, 8)

tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
print(f"Found {len(tabs)} tabs.")
output_folder = "extracted_data_ai"
os.makedirs(output_folder, exist_ok=True)

for tab_index, tab in enumerate(tabs):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_index+1}: {tab_name} ===")
    driver.execute_script("arguments[0].click();", tab)
    random_sleep(5, 7)

    print("Waiting for table to load...")
    for _ in range(10):
        try:
            if driver.find_element(By.ID, "table_body").find_elements(By.TAG_NAME, "tr"):
                break
        except:
            time.sleep(1)

    rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
    print(f"Extracting {len(rows)} rows...")

    header_cells = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]")
    headers = []
    for h in header_cells:
        p = h.find_elements(By.TAG_NAME, "p")
        headers.append(p[0].text.strip() if p else h.text.strip())
    headers = [h if h else f"Column{idx+1}" for idx, h in enumerate(headers)]

    def match_col(possible):
        for i, col in enumerate(headers):
            if col.lower() in [p.lower() for p in possible]:
                return i
        return None

    idx_name = match_col(['name', 'full name', 'client name'])
    idx_unit = match_col(['unit', 'unit no', 'unit number'])
    idx_building = match_col(['building', 'project', 'tower', 'building name'])

    file_path = os.path.join(output_folder, f"extracted_{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_ai.csv")
    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers + ["AI_Message"])

        for i, row in enumerate(rows, 1):
            tds = row.find_elements(By.TAG_NAME, "td")
            if not tds or len(tds) < 2:
                continue
            cells = tds[1:]
            values = [c.text.strip() for c in cells]
            full_name = values[idx_name] if idx_name is not None and idx_name < len(values) else ""
            last_name = full_name.split()[-1] if full_name else ""
            unit_number = values[idx_unit] if idx_unit is not None and idx_unit < len(values) else ""
            building_name = values[idx_building] if idx_building is not None and idx_building < len(values) else ""
            ai_msg = generate_ai_message(full_name, last_name, building_name, unit_number)
            values.append(ai_msg)
            writer.writerow(values)
            print(f"Row {i}: Message generated.")

            try:
                link = row.find_element(By.XPATH, ".//a[contains(@href, 'https://api.whatsapp.com/send?')]")
                href = link.get_attribute("href")
                href = re.sub(r'phone=\+?\d+', 'phone=923226100103', href)
                href = re.sub(r'text=.*', 'text=Hi', href)

                existing_tabs = driver.window_handles

                # Retry pop-up opening
                for retry in range(3):
                    driver.execute_script("window.open(arguments[0]);", href)
                    time.sleep(3)
                    new_tabs = driver.window_handles
                    if len(new_tabs) > len(existing_tabs):
                        break
                    print(f"⚠️ Attempt {retry+1}: WhatsApp link did not open. Retrying...")

                if len(new_tabs) <= len(existing_tabs):
                    print("❌ WhatsApp pop-up blocked or failed. Skipping this contact.")
                    continue

                new_tab = list(set(new_tabs) - set(existing_tabs))[0]
                driver.switch_to.window(new_tab)

                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
                )
                time.sleep(2)

                # Random delay before sending
                random_sleep(3, 6)

                print("⌨️ Typing message...")
                for line in ai_msg.splitlines():
                    pyautogui.typewrite(line)
                    pyautogui.keyDown('shift')
                    pyautogui.press('enter')
                    pyautogui.keyUp('shift')
                time.sleep(1)
                pyautogui.press('enter')
                print("✅ Message typed and sent.")

                # Random delay after sending
                random_sleep(4, 8)

                driver.close()
                driver.switch_to.window(driver.window_handles[1])  # Back to WCEasy tab

            except Exception as e:
                print(f"❌ Failed to send WhatsApp message on row {i}: {e}")

            random_sleep(0.5, 1)

print("\n✅ All tabs processed. Exiting now.")
driver.quit()

'''

'''
# working with sending messages but only after manually Allowing the pop-up blocked. 
import os
import time
import csv
import random
import re
import pyautogui
import pyperclip
from datetime import datetime
from dotenv import load_dotenv
from openai import OpenAI
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from login import login_google

# === Load .env ===
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

def get_greeting():
    hour = (datetime.utcnow().hour + 4) % 24
    return "Good Morning" if 5 <= hour < 12 else "Good Afternoon" if 12 <= hour < 17 else "Good Evening"

def generate_ai_message(full_name, last_name, building_name, unit_number):
    greeting = get_greeting()
    unit_info = f"your unit {unit_number} in {building_name}" if unit_number and building_name else \
    f"your unit in {building_name}" if building_name else \
    f"your unit {unit_number}" if unit_number else "your property"

    prompt = f"""
    The person's full name is: "{full_name}"
    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g. contains LLC, LIMITED, COMPANY, etc), do NOT use Mr./Ms.
      Instead, use: {greeting}, [Full Name of Company],
    ---
    My name is Omar Bayat, and I’m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.
    I’m reaching out regarding {unit_info}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.
    Looking forward to your response. Thank you, and have a lovely day ahead.
    Best regards,  
    Omar Bayat  
    White & Co. Real Estate
    Output ONLY the final message.
    """
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"ERROR generating message: {e}")
        return ""

# === LOGIN ===
driver = login_google()

# === Tab 1: Open WhatsApp Web ===
print("Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com/")
whatsapp_tab = driver.current_window_handle
print("Please scan the QR code in WhatsApp Web...")

# === Wait for WhatsApp Login ===
max_wait = 300
elapsed = 0
interval = 5
while elapsed < max_wait:
    try:
        driver.find_element(By.XPATH, "//div[@aria-label='Chat list' or @role='textbox']")
        print("✅ WhatsApp Web login detected!")
        break
    except:
        print("⏳ Waiting for WhatsApp login...")
        time.sleep(interval)
        elapsed += interval
else:
    print("❌ WhatsApp login timeout.")
    driver.quit()
    exit()

# === Tab 2: Open WCEasy Table View ===
driver.switch_to.new_window('tab')
print("Opening WCEasy Table View...")
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(5, 7)

try:
    driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]").click()
except Exception as e:
    print("ERROR: View button not found.", e)
    driver.quit()
    exit()
random_sleep(6, 8)

tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
print(f"Found {len(tabs)} tabs.")
output_folder = "extracted_data_ai"
os.makedirs(output_folder, exist_ok=True)

for tab_index, tab in enumerate(tabs):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_index+1}: {tab_name} ===")
    driver.execute_script("arguments[0].click();", tab)
    random_sleep(5, 7)

    print("Waiting for table to load...")
    for _ in range(10):
        try:
            if driver.find_element(By.ID, "table_body").find_elements(By.TAG_NAME, "tr"):
                break
        except:
            time.sleep(1)

    rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
    print(f"Extracting {len(rows)} rows...")

    header_cells = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]")
    headers = []
    for h in header_cells:
        p = h.find_elements(By.TAG_NAME, "p")
        headers.append(p[0].text.strip() if p else h.text.strip())
    headers = [h if h else f"Column{idx+1}" for idx, h in enumerate(headers)]

    def match_col(possible):
        for i, col in enumerate(headers):
            if col.lower() in [p.lower() for p in possible]:
                return i
        return None

    idx_name = match_col(['name', 'full name', 'client name'])
    idx_unit = match_col(['unit', 'unit no', 'unit number'])
    idx_building = match_col(['building', 'project', 'tower', 'building name'])

    file_path = os.path.join(output_folder, f"extracted_{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_ai.csv")
    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers + ["AI_Message"])

        for i, row in enumerate(rows, 1):
            tds = row.find_elements(By.TAG_NAME, "td")
            if not tds or len(tds) < 2:
                continue
            cells = tds[1:]
            values = [c.text.strip() for c in cells]
            full_name = values[idx_name] if idx_name is not None and idx_name < len(values) else ""
            last_name = full_name.split()[-1] if full_name else ""
            unit_number = values[idx_unit] if idx_unit is not None and idx_unit < len(values) else ""
            building_name = values[idx_building] if idx_building is not None and idx_building < len(values) else ""
            ai_msg = generate_ai_message(full_name, last_name, building_name, unit_number)
            values.append(ai_msg)
            writer.writerow(values)
            print(f"Row {i}: Message generated.")

            try:
                link = row.find_element(By.XPATH, ".//a[contains(@href, 'https://api.whatsapp.com/send?')]")
                href = link.get_attribute("href")
                href = re.sub(r'phone=\+?\d+', 'phone=923226100103', href)
                href = re.sub(r'text=.*', 'text=Hi', href)
                
                # Open WhatsApp link in new tab
                existing_tabs = driver.window_handles
                driver.execute_script("window.open(arguments[0]);", href)
                time.sleep(2)
                new_tabs = driver.window_handles

                if len(new_tabs) <= len(existing_tabs):
                    print("❌ WhatsApp link did not open — pop-up blocked or JS failed.")
                    continue

                new_tab = list(set(new_tabs) - set(existing_tabs))[0]
                driver.switch_to.window(new_tab)

                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
                )
                time.sleep(2)

                # Random sleep before typing
                random_sleep(3, 5)

                print("⌨️ Typing message...")
                for line in ai_msg.splitlines():
                    pyautogui.typewrite(line)
                    pyautogui.keyDown('shift')
                    pyautogui.press('enter')
                    pyautogui.keyUp('shift')
                time.sleep(1)
                pyautogui.press('enter')
                print("✅ Message typed and sent.")

                # Random sleep after sending
                random_sleep(4, 6)

                driver.close()
                driver.switch_to.window(driver.window_handles[1])  # Back to WCEasy tab
            except Exception as e:
                print(f"❌ Failed to send WhatsApp message on row {i}: {e}")

            random_sleep(0.5, 1)

print("\nAll tabs processed. Press Enter to exit...")
input()
driver.quit()
'''




'''
# main.py (Safe Version with Intelligent Column Handling)

import time
import random
import csv
import os
from selenium.webdriver.common.by import By
from login import login_google
from datetime import datetime
from openai import OpenAI
from dotenv import load_dotenv

# === Load .env ===
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

# === Random Sleep ===
def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

# === AI Message Generator ===
def generate_ai_message(full_name, last_name, building_name, unit_number):
    greeting = get_greeting()
    unit_info = f"your unit {unit_number} in {building_name}" if unit_number and building_name else \
                f"your unit in {building_name}" if building_name else \
                f"your unit {unit_number}" if unit_number else "your property"

    prompt = f"""
    The person's full name is: "{full_name}"

    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g. contains LLC, LIMITED, COMPANY, etc), do NOT use Mr./Ms.
      Instead, use: {greeting}, [Full Name of Company],

    ---

    My name is Omar Bayat, and I’m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.

    I’m reaching out regarding {unit_info}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.

    Looking forward to your response. Thank you, and have a lovely day ahead.

    Best regards,  
    Omar Bayat  
    White & Co. Real Estate

    Output ONLY the final message.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=400,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"ERROR generating message: {e}")
        return ""

# === Greeting based on UAE time ===
def get_greeting():
    hour = (datetime.utcnow().hour + 4) % 24
    return "Good Morning" if 5 <= hour < 12 else "Good Afternoon" if 12 <= hour < 17 else "Good Evening"

# === Login and Navigate ===
driver = login_google()
print("Opening Table View...")
# === Open WhatsApp Web and wait for login ===
print("Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com/")
print("Please scan the QR code in WhatsApp Web...")

# Wait for login confirmation
max_wait_time = 300  # seconds
wait_interval = 5
elapsed = 0
logged_in = False

while elapsed < max_wait_time:
    try:
        # Detect a UI element that only appears when logged in
        driver.find_element(By.XPATH, "//div[@aria-label='Chat list' or @role='textbox']")
        logged_in = True
        print("✅ WhatsApp Web login detected!")
        break
    except:
        print("⏳ Waiting for WhatsApp login...")
        time.sleep(wait_interval)
        elapsed += wait_interval

if not logged_in:
    print("❌ WhatsApp login timeout. Please try again.")
    driver.quit()
    exit()
driver.get("https://wceasy.club/staff/table-view.php")
random_sleep(5, 7)

# === Click View Button ===
print("Clicking first View button...")
try:
    driver.find_element(By.XPATH, "//a[contains(@href, 'view-excel.php')]").click()
except Exception as e:
    print("ERROR: View button not found.", e)
    driver.quit()
    exit()
random_sleep(6, 8)

# === Detect Sheet Tabs ===
tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
print(f"Found {len(tabs)} tabs.")

output_folder = "extracted_data_ai"
os.makedirs(output_folder, exist_ok=True)

for tab_index, tab in enumerate(tabs):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_index+1}: {tab_name} ===")

    driver.execute_script("arguments[0].click();", tab)
    random_sleep(5, 7)

    print("Waiting for table to load...")
    for _ in range(10):
        try:
            if driver.find_element(By.ID, "table_body").find_elements(By.TAG_NAME, "tr"):
                break
        except: time.sleep(1)

    header_cells = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]")
    headers = []
    for h in header_cells:
        p = h.find_elements(By.TAG_NAME, "p")
        headers.append(p[0].text.strip() if p else h.text.strip())
    headers = [h if h else f"Column{idx+1}" for idx, h in enumerate(headers)]
    print("Detected columns:", headers)

    def match_col(possible):
        for i, col in enumerate(headers):
            if col.lower() in [p.lower() for p in possible]:
                return i
        return None

    idx_name = match_col(['name', 'full name', 'client name'])
    idx_whatsapp = match_col(['whatsapp', 'mobile wn'])
    idx_unit = match_col(['unit', 'unit no', 'unit number'])
    idx_building = match_col(['building', 'project', 'tower', 'building name'])

    table_wrapper = driver.find_element(By.CLASS_NAME, "tableWraper")
    last_row_count = 0
    for _ in range(30):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", table_wrapper)
        random_sleep(1, 2)
        rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")
        if len(rows) == last_row_count:
            break
        last_row_count = len(rows)

    print(f"Extracting {len(rows)} rows...")
    file_path = os.path.join(output_folder, f"extracted_{tab_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_ai.csv")

    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers + ["AI_Message"])

        for i, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")[1:]  # Skip SI.No.
            values = [c.text.strip() for c in cells]

            full_name = values[idx_name] if idx_name is not None and idx_name < len(values) else ""
            last_name = full_name.split()[-1] if full_name else ""
            unit_number = values[idx_unit] if idx_unit is not None and idx_unit < len(values) else ""
            building_name = values[idx_building] if idx_building is not None and idx_building < len(values) else ""

            ai_msg = generate_ai_message(full_name, last_name, building_name, unit_number)
            values.append(ai_msg)
            writer.writerow(values)
            print(f"Row {i}: Message generated.")
            random_sleep(0.3, 0.6)

print("\nAll tabs processed. Press Enter to exit...")
input()
driver.quit()
'''

'''
import time
import random
import csv
import os
from selenium.webdriver.common.by import By
from login import login_google
from datetime import datetime
from openai import OpenAI
from dotenv import load_dotenv

# === Load .env ===
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

# === Random Sleep ===
def random_sleep(min_sec=1, max_sec=3):
    delay = random.uniform(min_sec, max_sec)
    print(f"Sleeping for {delay:.2f} seconds...")
    time.sleep(delay)

# === AI Message Generator ===
def generate_ai_message(full_name, last_name, building_name, unit_number):
    greeting = get_greeting()
    prompt = f"""
    The person's full name is: "{full_name}"
    
    - If this is a person's name, first do your best efforts and research to detect gender and select proper salutation (Mr. or Ms.), and use: {greeting} [Salutation] [Last Name].

    - If this is a company/organization name (for example, if it contains words like: LLC, LIMITED, COMPANY, INC, HOLDINGS, ENTERPRISE, GROUP, DEVELOPMENT, CORPORATION, or similar), then do NOT use Mr./Ms.
    Instead, start with:

    {greeting}, [Full Name of Company],

    ---

    After the greeting, write this full professional WhatsApp message:

    My name is Omar Bayat, and I’m a Property Consultant at White & Co. Real Estate, one of Dubai’s leading British-owned brokerages.

    I’m reaching out regarding your unit {unit_number} in {building_name}. May I ask if it’s currently vacant and available for rent?
    I have a qualified client actively searching in the building and would appreciate the opportunity to present your unit if available.

    Looking forward to your response. Thank you, and have a lovely day ahead.

    Best regards,  
    Omar Bayat  
    White & Co. Real Estate

    ---

    Please output ONLY the final complete message — with placeholders replaced — no explanations.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=300,
            temperature=0.7
        )
        message = response.choices[0].message.content.strip()
        return message
    except Exception as e:
        print(f"ERROR generating message: {e}")
        return ""

# === Greeting based on UAE time ===
def get_greeting():
    uae_hour = (datetime.utcnow().hour + 4) % 24
    if 5 <= uae_hour < 12:
        return "Good Morning"
    elif 12 <= uae_hour < 17:
        return "Good Afternoon"
    else:
        return "Good Evening"

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
output_folder = "extracted_data_ai"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# === Process each tab ===
for tab_idx, tab in enumerate(tab_elements):
    tab_name = tab.text.strip().replace(" ", "_")
    print(f"\n=== Processing Tab {tab_idx+1}: {tab_name} ===")

    # Click tab
    driver.execute_script("arguments[0].click();", tab)
    random_sleep(4, 7)

    # Wait for table to load
    print("Waiting for table to load...")
    table_body = None
    while True:
        try:
            table_body = driver.find_element(By.ID, "table_body")
            if table_body.find_elements(By.TAG_NAME, "tr"):
                print("Table loaded.")
                break
        except:
            pass
        time.sleep(2)

    # === Detect column headers ===
    header_elements = driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th//p")
    original_column_names = [h.text.strip() for h in header_elements]
    column_names_lower = [name.lower() for name in original_column_names]
    print("Detected columns:", original_column_names)

    # === Find column indexes (CASE INSENSITIVE) ===
    def get_column_index(possible_names):
        for name in possible_names:
            name_lower = name.lower()
            if name_lower in column_names_lower:
                idx = column_names_lower.index(name_lower)
                print(f"Matched column '{original_column_names[idx]}' at index {idx}")
                return idx
        print(f"WARNING: Could not find columns matching {possible_names}")
        return None

    # Matching lists (safe lower-case matching)
    name_col_index = get_column_index(['name', 'client name', 'customer name', 'full name'])
    whatsapp_col_index = get_column_index(['whatsapp', 'whatsapp number', 'mobile wn'])
    unit_col_index = get_column_index(['unit number', 'unit', 'unit no', 'apartment'])
    building_col_index = get_column_index(['building name', 'building', 'project', 'tower'])

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
    output_file = os.path.join(output_folder, f"extracted_{tab_name}_{timestamp}_ai.csv")

    with open(output_file, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(original_column_names + ["AI_Message"])

        for row_idx, row in enumerate(rows, start=1):
            cells = row.find_elements(By.TAG_NAME, "td")
            row_data = []

            # Collect raw data
            full_name = ""
            last_name = ""
            unit_number = ""
            building_name = ""
            whatsapp_link = ""

            for idx, cell in enumerate(cells):
                text = cell.text.strip()

                if idx == whatsapp_col_index:
                    try:
                        link = cell.find_element(By.TAG_NAME, "a").get_attribute("href")
                        whatsapp_link = link
                        row_data.append(link)
                    except:
                        row_data.append("")
                else:
                    row_data.append(text)

                # Extract needed fields
                if idx == name_col_index:
                    full_name = text
                    last_name = text.split()[-1] if text else ""
                elif idx == unit_col_index:
                    unit_number = text
                elif idx == building_col_index:
                    building_name = text

            # Generate AI message
            ai_message = generate_ai_message(full_name, last_name, building_name, unit_number)
            row_data.append(ai_message)

            writer.writerow(row_data)
            print(f"Row {row_idx}: AI message generated.")
            random_sleep(0.5, 1)

    print(f"=== DONE Tab {tab_idx+1}: Data saved to {output_file} ===")

# === All Tabs Processed ===
print("\n=== ALL TABS PROCESSED ===")
input("Press Enter to keep browser open...")
driver.quit()
'''