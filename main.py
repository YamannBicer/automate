from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import time
import pyautogui
from wrt import wrt
import dotenv
import os

dotenv.load_dotenv(override=True)

username = os.getenv("USERNAME")
password = os.getenv("PASSWORD")

url = "https://10.34.20.126/sense/app/86c3becd-4334-4d7c-b2eb-59873a09b204/overview"

# Set Chrome options to skip the certificate warning page
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--allow-insecure-localhost")
chrome_options.add_argument("--allow-running-insecure-content")
chrome_options.add_argument("--disable-features=BlockCredentialedSubresources")

# Initialize the driver
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

driver.get(url)
driver.maximize_window()

time.sleep(1)
wrt(username)
pyautogui.press("tab")

time.sleep(1)
wrt(password)
pyautogui.press("tab")

time.sleep(1)
pyautogui.press("enter")










