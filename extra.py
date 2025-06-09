import time
import pandas as pd
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, UnexpectedAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
import easyocr
from test import preprocess_image
from captcha import save_captcha_from_driver

# Constants
URL = "https://results.vtu.ac.in/DJcbcs25/index.php"
USN_LIST = ["1AY23IS048", "1AY23IS013", "1AY23IS049"]

KNOWN_SUBJECTS = [
    "MATHEMATICS FOR COMPUTER SCIENCE",
    "DIGITAL DESIGN & COMPUTER ORGANIZATION",
    "OPERATING SYSTEMS",
    "DATA STRUCTURES AND APPLICATIONS",
    "DATA STRUCTURES LAB",
    "SOCIAL CONNECT AND RESPONSIBILITY",
    "YOGA",
    "OBJECT ORIENTED PROGRAMMING WITH JAVA",
    "DATA VISUALIZATION WITH PYTHON"
]

# Setup Chrome options
options = webdriver.ChromeOptions()
# options.add_argument("--headless=new")  # Uncomment for headless mode
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

caps = webdriver.DesiredCapabilities.CHROME.copy()
caps['goog:loggingPrefs'] = {'browser': 'ALL'}

# Initialize driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Initialize OCR reader
reader = easyocr.Reader(['en'])

def get_captcha_text():
    """Extract CAPTCHA text using OCR."""
    saved = save_captcha_from_driver(driver)
    if not saved:
        print("❌ Failed to save CAPTCHA image.")
        return None

    processed_image_path = preprocess_image("captcha.png")
    if not processed_image_path:
        print("❌ Failed to preprocess CAPTCHA image.")
        return None

    result = reader.readtext(processed_image_path)
    if result:
        text = result[0][1].strip().replace(" ", "")
        print(f"🔍 OCR Detected CAPTCHA: {text}")
        return text
    else:
        print("❌ CAPTCHA OCR failed.")
        return None

def handle_possible_alert():
    """Handle alert if present."""
    try:
        alert = driver.switch_to.alert
        alert_text = alert.text
        alert.accept()
        print(f"⚠️ Alert Dismissed: {alert_text}")
        return alert_text
    except NoAlertPresentException:
        return None

def print_browser_logs():
    """Print JavaScript errors/warnings in browser console."""
    logs = driver.get_log('browser')
    if logs:
        print("🛠️ Browser console logs:")
        for log in logs:
            print(log)

def main():
    results_data = []
    saved_usns = []

    # STEP 1: Save all screenshots
    for usn in USN_LIST:
        print(f"\n🎯 Processing USN (screenshot only): {usn}")
        attempts = 0
        screenshot_path = f"screenshots/{usn}_result.png"

        while attempts < 10:
            attempts += 1
            print(f"🔁 Attempt {attempts} for {usn}")

            try:
                driver.delete_all_cookies()
                driver.get(URL)

                wait = WebDriverWait(driver, 10)
                wait.until(EC.presence_of_element_located((By.NAME, "lns")))

                captcha_text = get_captcha_text()
                if not captcha_text or len(captcha_text) != 6:
                    print("⚠️ Invalid CAPTCHA text length. Retrying...")
                    continue

                # Fill in form
                usn_input = driver.find_element(By.NAME, "lns")
                captcha_input = driver.find_element(By.NAME, "captchacode")
                submit_btn = driver.find_element(By.ID, "submit")

                usn_input.clear()
                captcha_input.clear()
                usn_input.send_keys(usn)
                captcha_input.send_keys(captcha_text)
                submit_btn.click()

                # Check for CAPTCHA length warning on page
                try:
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Please lengthen this text')]"))
                    )
                    print("⚠️ CAPTCHA length error on page. Retrying...")
                    continue
                except TimeoutException:
                    pass

                # Check for alert
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert_msg = handle_possible_alert()
                    print(f"❌ CAPTCHA/USN Error: {alert_msg}")
                    continue  # Retry
                except TimeoutException:
                    pass  # CAPTCHA accepted

                print("✅ CAPTCHA accepted. Waiting for result page to load...")
                time.sleep(3)  # wait for page to render

                # Take screenshot
                print("🖼️ Running result.py to capture screenshot...")
                subprocess.run(["python", "result.py", usn], check=True)

                saved_usns.append(usn)
                break

            except Exception as e:
                print(f"❌ Unexpected error occurred: {e}")
                continue

        if attempts >= 10:
            print(f"⚠️ Max attempts reached for {usn}. Moving on.")
            results_data.append({"USN": usn, "Result": "❌ Screenshot not saved"})
        else:
            results_data.append({"USN": usn, "Result": "✅ Screenshot saved"})

    # STEP 2: Extract marks from screenshots
    print("\n📊 Starting marks extraction for all saved screenshots...\n")
    for usn in saved_usns:
        screenshot_path = f"screenshots/{usn}_result.png"
        try:
            subprocess.run(["python", "marks.py", screenshot_path], check=True)
            print(f"✅ Marks extracted for {usn}")
        except Exception as e:
            print(f"❌ Failed to extract marks for {usn}: {e}")
            results_data.append({"USN": usn, "Result": "❌ Failed to extract marks"})

    # Save summary
    df = pd.DataFrame(results_data)
    df.to_csv("vtu_results.csv", index=False)
    print("\n📁 All results saved to 'vtu_results.csv'")

    driver.quit()


if __name__ == "__main__":
    main()