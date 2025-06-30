# This script takes the screenshot of result page
import time
import pyautogui
import os
import sys
from PIL import Image

# Settings
WAIT_TIME = 1  # Time to switch to the browser window
SCROLL_PIXELS = 400
IMAGE_WIDTH, IMAGE_HEIGHT = pyautogui.size()

# Get USN from command line
if len(sys.argv) < 2:
    print("❌ USN not provided.")
    sys.exit(1)

usn = sys.argv[1]

# Ensure 'screenshots/' folder exists
output_dir = os.path.join(os.getcwd(), "screenshots")
os.makedirs(output_dir, exist_ok=True)

print(f"⏳ You have {WAIT_TIME} seconds to open the result page for {usn}...")
time.sleep(WAIT_TIME)

print("📸 Capturing first screenshot...")
screenshot1 = pyautogui.screenshot()

# Scroll and capture second screenshot
pyautogui.scroll(-SCROLL_PIXELS)
time.sleep(1)
print("📸 Capturing second screenshot after scroll...")
screenshot2 = pyautogui.screenshot()

# Combine
combined_image = Image.new('RGB', (IMAGE_WIDTH, IMAGE_HEIGHT * 2))
combined_image.paste(screenshot1, (0, 0))
combined_image.paste(screenshot2, (0, IMAGE_HEIGHT))

# Save inside screenshots/ with USN as filename
output_path = os.path.join(output_dir, f"{usn}_result.png")
combined_image.save(output_path)
print(f"✅ Screenshot saved at: {output_path}")
