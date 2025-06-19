import os
import time
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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime

# === CONFIGURATION ===
pyautogui.FAILSAFE = False
USE_TEST_NUMBER = True
TEST_NUMBER = "923226100103"  # Replace with your real test number

# === LOAD ENV & OPENAI ===
load_dotenv()
openai_api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=openai_api_key)

# === LOGGING SETUP ===
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot_run_log.txt")

def log(msg, tag=None):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_message = f"[{timestamp}] [{tag if tag else 'INFO'}] {msg}"
    print(log_message)  # Keep printing to console

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(log_message + "\n")

def random_sleep(min_sec=1, max_sec=3):
    time.sleep(random.uniform(min_sec, max_sec))

def get_greeting():
    hour = (datetime.utcnow().hour + 4) % 24
    return "Good Morning" if 5 <= hour < 12 else "Good Afternoon" if 12 <= hour < 17 else "Good Evening"


def generate_ai_message(full_name, last_name, building_name, unit_number, property_type, prompt=None, retries=2):
    greeting = get_greeting()

    property_type_cleaned = property_type.strip().lower()

    if any(pt in property_type_cleaned for pt in ['apartment', 'flat', 'unit', 'office', 'room', 'villa']):
        for pt in ['apartment', 'flat', 'unit', 'office', 'room', 'villa']:
            if pt in property_type_cleaned:
                unit_term = pt
                break
    else:
        unit_term = "property"

    unit_info = f"your {unit_term} {unit_number} in {building_name}" if unit_number and building_name else \
                f"your {unit_term} in {building_name}" if building_name else \
                f"your {unit_term} {unit_number}" if unit_number else f"your {unit_term}"
    if prompt is None:
        prompt = """
        The person's full name is: "{full_name}"
        - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
        - If this is a company name (e.g., contains LLC, LIMITED, COMPANY, BANK etc.), do NOT use Mr./Ms.
          Instead, use: {greeting},

        ---
        My name is Omar Bayat, and I am a Property Consultant at White & Co. Real Estate, one of the leading British-owned brokerages in Dubai.
        I am reaching out regarding {unit_info}. May I ask if it is currently vacant and available for rent?
        I have a qualified client actively searching in the building who would be interested in arranging a viewing at your earliest convenience if the property is available. Kindly confirm its availability and any further details.
        Looking forward to your response. Thank you, and have a lovely day ahead.

        Best regards,  
        Omar Bayat  
        White & Co. Real Estate

        Slightly vary the wording professionally for each message. Output ONLY the final message.
        """
    formatted_prompt = prompt.format(
        greeting=greeting,
        full_name=full_name,
        last_name=last_name,
        unit_info=unit_info
    )

    for attempt in range(retries + 1):
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": formatted_prompt}],
                max_tokens=900,
                temperature=0.7
            )
            message = response.choices[0].message.content.strip()

            log(f"🔎 GPT Response (Attempt {attempt+1}): {message} (Length: {len(message.split())} words)", "info")

            if len(message.split()) < 10:
                log(f"⚠️ Attempt {attempt+1}: Message too short — retrying after 3s ...", "warn")
                time.sleep(3)
                continue

            return message
        except Exception as e:
            log(f"❌ GPT Error (attempt {attempt+1}): {e}", "error")
            time.sleep(3)

    return f"{greeting}, I'm contacting you regarding {unit_info}."

def normalize_name(text):
    text = text.strip().lower()
    text = re.sub(r'[\W]+', '_', text)
    text = re.sub(r'_+', '_', text)
    return text.strip('_')

def scroll_to_load_all_rows(driver, pause_time=2, max_attempts=20):
    attempts = 0
    while attempts < max_attempts:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        random_sleep(pause_time, pause_time + 1)

        try:
            all_loaded = driver.find_element(By.XPATH, "//div[@class='elseDesign1' and contains(.,'All Data Loaded')]")
            if all_loaded.is_displayed():
                log("✅ All Data Loaded element detected. Scrolling complete.", "success")
                break
        except:
            log("⏳ 'All Data Loaded' element not detected yet. Continuing to scroll...", "info")

        attempts += 1
    else:
        log("⚠️ Max scrolling attempts reached. Some data may not have loaded fully.", "warn")

    # Scroll explicitly back to top
    driver.execute_script("window.scrollTo(0, 0);")
    random_sleep(1, 2)
    log("✅ Scrolled back to the top.", "info")

def run_whatsapp_bot(selected_sheet_name: str = None, selected_tabs: list[str] = None, prompt: str = None, log_fn=None):
    def log(msg, tag=None):
        print(msg)
        if log_fn:
            log_fn(msg, tag or "default")
    # Immediately log the received prompt clearly
    if prompt:
        log("🟢 Prompt successfully received from GUI!", "success")
        log(f"📨 Prompt:\n{prompt[:300]}...", "info")
    else:
        log("🔴 No prompt received. Using fallback.", "error")

    fallback_prompt = """
    The person's full name is: "{full_name}"
    - If this is a person's name, detect gender and use: {greeting} [Salutation] [Last Name].
    - If this is a company name (e.g., contains LLC, LIMITED, COMPANY, BANK etc.), do NOT use Mr./Ms.
      Instead, use: {greeting},

    ---
    My name is Omar Bayat, and I am a Property Consultant at White & Co. Real Estate, one of the leading British-owned brokerages in Dubai.
    I am reaching out regarding {unit_info}. May I ask if it is currently vacant and available for rent?
    I have a qualified client actively searching in the building who would be interested in arranging a viewing at your earliest convenience if the property is available. Kindly confirm its availability and any further details.
    Looking forward to your response. Thank you, and have a lovely day ahead.

    Best regards,  
    Omar Bayat  
    White & Co. Real Estate

    Slightly vary the wording professionally for each message. Output ONLY the final message.
    """

    normalized_selected_tabs = [normalize_name(t) for t in selected_tabs] if selected_tabs else []
    normalized_selected_sheet = normalize_name(selected_sheet_name) if selected_sheet_name else ""

    driver = login_google()

    log("Opening WhatsApp Web...")
    driver.get("https://web.whatsapp.com/")
    whatsapp_tab = driver.current_window_handle
    log("Please scan the QR code...")

    for _ in range(60):
        try:
            continue_btn = driver.find_element(By.XPATH, "//button//div[contains(text(),'Continue')]")
            if continue_btn.is_displayed():
                random_sleep(1, 2)
                continue_btn.click()
                log("✅ Clicked 'Continue' button during QR scanning.")
                random_sleep(1, 2)
        except:
            pass
        try:
            driver.find_element(By.XPATH, "//div[@aria-label='Chat list' or @role='textbox']")
            log("✅ WhatsApp Web login detected!")
            break
        except:
            log("⏳ Waiting for WhatsApp login...")
            time.sleep(5)
    else:
        log("❌ WhatsApp login timeout.", "error")
        driver.quit()
        return

    driver.switch_to.new_window('tab')
    log("Opening WCEasy Table View...")
    driver.get("https://wceasy.club/staff/table-view.php")
    random_sleep(5, 7)

    # NEW: Dynamically select the provided sheet (as shown clearly above)
    # Convert internal name to actual displayed name on WCEasy table
    selected_sheet_display_name = re.sub(r'_\d+$', '', selected_sheet_name).replace('_', ' ')

    sheet_xpath = f'//tbody[@id="excel_data"]/tr[td/a[normalize-space(text())="{selected_sheet_display_name}"]]//a[@class="tableViewBtn"]'

    try:
        sheet_element = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, sheet_xpath))
        )
        sheet_element.click()
        log(f"✅ Successfully selected sheet: '{selected_sheet_name}'.", "success")
        random_sleep(4, 6)
    except Exception as e:
        log(f"❌ Failed to select sheet '{selected_sheet_name}': {e}", "error")
        driver.quit()
        return

    log("⚠️ Please ensure pop-ups are enabled in your browser for https://wceasy.club", "info")
    time.sleep(5)

    # === INSERT SCROLLING LOGIC HERE ===
    try:
        tabs_container = driver.find_element(By.CLASS_NAME, "tableFooterTab")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center', inline: 'nearest'});", tabs_container)
        random_sleep(2, 3)

        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", tabs_container)
        random_sleep(1, 2)

        log("✅ Successfully scrolled tabs container into view.", "info")
    except Exception as e:
        log(f"⚠️ Could not explicitly scroll tabs container: {e}", "warn")

    # === START: Sheet and Tabs Selection (Step 1 & 2 combined clearly) ===
    log(f"📑 Selecting Sheet: {selected_sheet_name}", "info")

    # === START FIXED TABS SELECTION LOGIC ===
    try:
        WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".tableFooterTabBox"))
        )
        tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")
        log(f"✅ Tabs loaded successfully.")
    except Exception as e:
        log(f"❌ Tabs did not load: {e}", "error")
        driver.quit()
        return

    frontend_tab_map = {normalize_name(tab.text): tab for tab in tabs}
    normalized_selected_tabs = [normalize_name(tab) for tab in selected_tabs]

    matched_tabs = [frontend_tab_map[norm] for norm in normalized_selected_tabs if norm in frontend_tab_map]

    if not matched_tabs:
        missing_tabs = set(normalized_selected_tabs) - set(frontend_tab_map.keys())
        log(f"❌ None of the selected tabs match tabs from frontend. Missing: {missing_tabs}", "error")
        driver.quit()
        return
    # === END FIXED TABS SELECTION LOGIC ===
    
    '''   # Wait for the table to load  
    for _ in range(10):
        tabs = driver.find_elements(By.CSS_SELECTOR, ".tableFooterTabBox")

        if tabs:
            break
        log("⏳ Waiting for tabs to load...")
        time.sleep(2)

    if not tabs:
        log("❌ No tabs found after retrying. Exiting.", "error")
        driver.quit()
        return

    '''    
    #frontend_tab_map = {normalize_name(tab.text): tab for tab in tabs}
    #matched_tabs = [frontend_tab_map[norm] for norm in normalized_selected_tabs if norm in frontend_tab_map]

    message_count = 0
    for tab_index, tab in enumerate(matched_tabs):
        tab_name = tab.text.strip().replace(" ", "_")
        log(f"\n=== Processing Tab {tab_index + 1}: {tab_name} ===", "info")
        driver.execute_script("arguments[0].click();", tab)
        random_sleep(5, 7)

        # Explicitly load all rows before processing them
        scroll_to_load_all_rows(driver)

        rows = driver.find_elements(By.XPATH, "//table[@id='myTable']/tbody/tr")

        headers = []
        for h in driver.find_elements(By.XPATH, "//table[@id='myTable']/thead/tr/th[position()>1]"):
            p = h.find_elements(By.TAG_NAME, "p")
            headers.append(p[0].text.strip() if p else h.text.strip())
        headers = [h if h else f"Column{idx + 1}" for idx, h in enumerate(headers)]

        def match_col(possible, fuzzy_contains=False):
            for i, col in enumerate(headers):
                col_lower = col.lower()
                if fuzzy_contains:
                    if any(p.lower() in col_lower for p in possible):
                        return i
                else:
                    if col_lower in [p.lower() for p in possible]:
                        return i
            return None

        idx_name = match_col([
            'name', 'full name', 'client name', 'owner', 'owner name', 'owner_name',
            'landlord', 'landlord name', 'contact name', 'property owner'
        ], fuzzy_contains=False)
        idx_unit = match_col([
            'unit', 'unit no', 'unit number', 'flat', 'flat no', 'flat number', 
            'apartment', 'apartment no', 'apartment number', 'unit#', 'flat#', 'apartment#', 
            'rooms', 'rooms description', 'rooms descr', 'unitnumber'
        ], fuzzy_contains=True)
        idx_building = match_col([
            'building', 'building name', 'tower', 'tower name', 'project', 'project name',
            'property', 'property name', 'building/project', 'bldg', 'bldg name', 'development',
            'bldg/project', 'project/building', 'building/project name', 'building_name',
            'tower/building', 'building/tower', 'building_name/project', 'project/building_name',
            'building name/project', 'project name/building', 'building-project', 'project-building',
            'bld', 'bldng', 'bldn'
        ], fuzzy_contains=True)
        idx_property_type = match_col([
            'type', 'property type', 'propertytype', 'propertytypeen'
        ], fuzzy_contains=True)
        log(f"Headers found: {headers}")
        log(f"Property type column identified: '{headers[idx_property_type]}' at index {idx_property_type}" if idx_property_type is not None else "No property type column matched.")
        idx_wn = match_col(['wn'], fuzzy_contains=True)

        for i, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")[1:]
            values = [c.text.strip() for c in cells]
            full_name = values[idx_name] if idx_name is not None else ""
            last_name = full_name.split()[-1] if full_name else ""
            unit_number = values[idx_unit] if idx_unit is not None else ""
            building_name = values[idx_building] if idx_building is not None else ""
            real_number = values[idx_wn] if idx_wn is not None else ""
            unit_column_name = headers[idx_unit].lower() if idx_unit is not None else ""
            property_type = values[idx_property_type] if (idx_property_type is not None and values[idx_property_type].strip()) else ""

            if not property_type:
                if 'apartment' in unit_column_name:
                    property_type = 'apartment'
                elif 'flat' in unit_column_name:
                    property_type = 'flat'
                elif 'office' in unit_column_name:
                    property_type = 'office'
                elif 'room' in unit_column_name:
                    property_type = 'room'
                elif 'villa' in unit_column_name:
                    property_type = 'villa'
                elif 'unit' in unit_column_name:
                    property_type = 'unit'
                else:
                    property_type = 'property'  # default fallback

            log(f"Inferred property type: '{property_type}' from column '{unit_column_name}'")

            number = TEST_NUMBER if USE_TEST_NUMBER else real_number

            if not number:
                log(f"⚠️ Row {i}: No WhatsApp number found. Skipping.", "info")
                continue

            # ✅ New WhatsApp Link Click Logic (updated clearly)
            max_chat_attempts = 2  # number of attempts to retry opening chat
            chat_attempts = 0
            chat_loaded = False

            # Store tab handles clearly
            wceasy_tab = driver.window_handles[1]  # Explicit WCEasy Tab
            whatsapp_web_tab = whatsapp_tab  # Explicit WhatsApp Web Tab

            while chat_attempts < max_chat_attempts:
                chat_attempts += 1
                
                # Open WhatsApp Link explicitly from WCEasy Tab (each retry)
                driver.switch_to.window(wceasy_tab)
                random_sleep(1, 2)
                
                try:
                    # Find all WhatsApp links in the current row
                    whatsapp_links = row.find_elements(By.XPATH, ".//a[contains(@href, 'whatsapp.com/send')]")

                    if not whatsapp_links:
                        log(f"⚠️ Row {i}: No WhatsApp links found. Skipping.", "info")
                        continue

                    whatsapp_link = whatsapp_links[0]  # Pick the first available WhatsApp link

                    #for live server, replace with web.whatsapp.com    
                    original_href = whatsapp_link.get_attribute('href').strip()

                    # Ensure WhatsApp Web format explicitly
                    new_href = original_href.replace("whatsapp://", "https://web.whatsapp.com/")

                    # Replace phone number explicitly with clean TEST_NUMBER
                    new_href = re.sub(r'(phone=\+?[\d]+)', f'phone={TEST_NUMBER}', new_href) if USE_TEST_NUMBER else original_href

                    # Remove unwanted URL encodings and whitespace explicitly
                    new_href = re.sub(r'(%20|\s)+', '', new_href)

                    # Log the final href for debugging
                    log(f"🚀 Opening WhatsApp Link: {new_href}", "info")

                    # OPTIONAL BREAK (first-time popup allowance)
                    if message_count == 0:
                        log("⏸️ Initial pause for manual popup allowance...", "info")
                        random_sleep(30, 40)

                    '''
                    # for local server testing, replace with TEST_NUMBER
                    # Assuming TEST_NUMBER is defined
                    clean_number = re.sub(r'\D', '', TEST_NUMBER)  # removes spaces, pluses, and any non-digit characters
                    original_href = whatsapp_link.get_attribute('href')
                    new_href = re.sub(r'phone=\+?\d+', f'phone={TEST_NUMBER}', original_href) if USE_TEST_NUMBER else original_href
                    '''
                    # Introduce a flag at the start (before loop)
                    # Initialize flag explicitly once outside of retry loop if first chat attempt
                    first_whatsapp_click = (message_count == 0 and chat_attempts == 1)

                    # SINGLE OPENING OF WHATSAPP LINK
                    driver.execute_script(f"window.open('{new_href}', '_blank');")
                    driver.switch_to.window(driver.window_handles[-1])

                    if first_whatsapp_click:
                        log("⏳ First WhatsApp link clicked, pausing initially for popup allowance...", "info")
                        time.sleep(20)  # Adjust as needed
                    else:
                        random_sleep(3, 5)

                except Exception as e:
                    log(f"❌ Row {i}: Error opening WhatsApp link: {e}", "error")
                    continue
                
                # Start checking if chat loads
                start = time.time()
                chat_loaded = False

                chat_loaded = False
                max_wait_time = 20
                end_time = time.time() + max_wait_time

                while time.time() < end_time:
                    try:
                        continue_chat_btn = WebDriverWait(driver, 2).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[@id='action-button' or .//span[text()='Continue to Chat']]"))
                        )
                        continue_chat_btn.click()
                        log("✅ Clicked 'Continue to Chat' button quickly.")
                        random_sleep(0.5, 1)
                    except:
                        pass  # button not found yet

                    try:
                        use_whatsapp_web_btn = WebDriverWait(driver, 2).until(
                            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'web.whatsapp.com/send') or .//span[text()='use WhatsApp Web']]"))
                        )
                        use_whatsapp_web_btn.click()
                        log("✅ Clicked 'Use WhatsApp Web' button quickly.")
                        random_sleep(0.5, 1)
                    except:
                        pass  # button not found yet

                    try:
                        continue_popup_btn = WebDriverWait(driver, 1).until(
                            EC.element_to_be_clickable((By.XPATH, "//button//div[contains(text(),'Continue')]"))
                        )
                        continue_popup_btn.click()
                        log("✅ Clicked 'Continue' popup button quickly.")
                        random_sleep(0.5, 1)
                    except:
                        pass  # popup button not found yet

                    # Quickly check if chat loaded successfully
                    try:
                        driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
                        chat_loaded = True
                        log("✅ WhatsApp chat loaded quickly.")
                        break
                    except:
                        pass  # chat box not loaded yet

                    random_sleep(0.5, 1)  # Short sleep for next retry

                    time.sleep(1)

                if chat_loaded:
                    break  # Exit retry loop if successful

                # Chat did not load, close WhatsApp tab explicitly and handle popup explicitly
                log(f"❌ Chat did not load on attempt {chat_attempts} — closing and retrying.", "warn")
                driver.close()
                
                # Switch explicitly back to WhatsApp Web main tab to handle popup
                driver.switch_to.window(whatsapp_web_tab)
                random_sleep(1, 2)

                # Handle "Use here" popup explicitly
                try:
                    use_here_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[text()='Use here']"))
                    )
                    use_here_button.click()
                    log("✅ Clicked 'Use here' button on WhatsApp Web popup (retry logic).")
                    random_sleep(0.5, 1.5)
                except Exception as e:
                    log(f"⚠️ No 'Use here' popup appeared or could not be clicked during retry: {e}", "info")

            if not chat_loaded:
                log(f"❌ Chat did not load on attempt {chat_attempts} — explicitly closing the tab and retrying.", "warn")
                driver.close()
                driver.switch_to.window(whatsapp_web_tab)
                random_sleep(1, 2)

            # ✅ PLACE YOUR MESSAGE GENERATION HERE, AFTER CHAT LOAD SUCCESS
            log(f"🛠️ Generating message with parameters:\n"
                f"   - Full Name: '{full_name}'\n"
                f"   - Last Name: '{last_name}'\n"
                f"   - Building Name: '{building_name}'\n"
                f"   - Unit Number: '{unit_number}'\n"
                f"   - Property Type: '{property_type}'", "info")

            unit_column_name = headers[idx_unit].lower() if idx_unit is not None else ""
            message = generate_ai_message(full_name, last_name, building_name, unit_number, property_type, prompt=prompt)

            log(f"📩 Generated AI message:\n{message}\n{'-'*50}", "info")

            # Now proceed with your existing typing and sending logic.

            random_sleep(3, 5)

            chat_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
            )
            chat_input.click()
            random_sleep(0.3, 0.5)
            chat_input.click()
            random_sleep(0.3, 0.5)

            try:
                chat_input.send_keys(Keys.CONTROL + 'a')  # Select all text
                random_sleep(0.2, 0.4)
                chat_input.send_keys(Keys.DELETE)  # Delete selected text
                random_sleep(0.2, 0.4)
            except Exception:
                chat_input.send_keys(Keys.COMMAND + 'a')  # For Mac compatibility
                random_sleep(0.2, 0.4)
                chat_input.send_keys(Keys.DELETE)  # Delete selected text
                random_sleep(0.2, 0.4)

            # Additional explicit clearing and focusing (KEY STEP)
            chat_input.click()
            random_sleep(0.3, 0.6)
            pyautogui.press("backspace", presses=3, interval=0.2)  # explicitly clear residual invisible characters
            random_sleep(0.3, 0.5)

            '''
            # Explicitly press backspace 3 times
            for _ in range(5):
                pyautogui.press("delete")
                random_sleep(0.1, 0.2)

            '''

            # Safety check for short message
            if len(message.strip().split()) < 5:
                log(f"⚠️ Warning: Suspiciously short AI message for row {i}: {message}")

            log("⌨️ Typing message manually with pyautogui...")
            for line in message.splitlines():
                print(f"Typing line: {line}")
                random_sleep(0.1, 0.3)
                pyautogui.typewrite(line, interval=0.02)  # typing speed controlled here
                pyautogui.keyDown('shift')
                pyautogui.press('enter')  # Shift+Enter creates a newline without sending
                pyautogui.keyUp('shift')
                random_sleep(0.1, 0.5)

            # Explicitly click the chat box again to ensure it's focused
            chat_input.click()
            random_sleep(0.1, 0.3)

            # Send Enter explicitly
            pyautogui.keyDown('enter')
            random_sleep(0.2, 0.4)
            pyautogui.keyUp('enter')

            log(f"✅ Row {i}: Message typed and sent to {number}.", "success")
            random_sleep(1, 3)
            # Ensure correct tab is closed explicitly and log clearly
            try:
                current_tab = driver.current_window_handle
                driver.close()
                log(f"✅ Successfully closed WhatsApp chat tab: {current_tab}.", "info")
            except Exception as e:
                log(f"❌ Error closing WhatsApp chat tab: {e}", "error")

            # Switch back explicitly to original WhatsApp Web tab
            try:
                driver.switch_to.window(whatsapp_tab)
                log("✅ Switched back to WhatsApp Web main tab successfully.", "info")
            except Exception as e:
                log(f"❌ Error switching back to WhatsApp main tab: {e}", "error")
            random_sleep(2, 7)

            # Handle "WhatsApp open in another window" popup
            
            # ✅ New logic: Click "Continue" button popup
            try:
                continue_btn = driver.find_element(By.XPATH, "//button//div[contains(text(),'Continue')]")
                if continue_btn.is_displayed():
                    continue_btn.click()
                    log("✅ Clicked 'Continue' button popup")
                    random_sleep(3, 8)
            except:
                pass
            
            try:
                use_here_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[text()='Use here']"))
                )
                random_sleep(2, 3)
                use_here_button.click()
                log("✅ Clicked 'Use here' button on WhatsApp Web popup.")
                random_sleep(1, 3)
            except Exception as e:
                log(f"⚠️ No 'Use here' popup appeared or could not be clicked: {e}", "info")

            # IMPORTANT: Switch explicitly back to WCEasy tab
            wceasy_tab = driver.window_handles[1]  # assuming your WCEasy tab is second tab
            random_sleep(1, 3)
            driver.switch_to.window(wceasy_tab)
            random_sleep(1, 5)
            log(f"✅ Row {i}: Successfully processed and sent message.", "success")
            message_count += 1
            if message_count % 10 == 0:
                log("⏸️ Pausing for 30 seconds after 10 messages...", "info")
                random_sleep(20, 60)

    log("🏁 All done!", "success")
    driver.quit()
    log(f"Total messages sent: {message_count}", "info")



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