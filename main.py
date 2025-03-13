from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains

import time
import pyautogui
from wrt import wrt
import dotenv
import os
import pandas as pd

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

# xpath - MAIN TABLE REAL-TIME
# //*[@id="approved-sheet-section"]/li[13]/div/div/div/div[1]/div

##approved-sheet-section > li.qv-content-li.last > div > div > div > div.qv-thumb-wrap.ng-isolate-scope > div

# time.sleep(10)
#
# links = driver.find_elements()
#
# for link in links:
#     print(link.get_attribute("innerHTML"))





wait = WebDriverWait(driver, 10)
element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="approved-sheet-section"]/li[13]/div/div/div/div[1]/div')))
time.sleep(1)
element.click()





time.sleep(1)
# Find the element by XPath
element = driver.find_element(By.XPATH, "//*[@id='grid']/div[3]/div[1]/div/div[1]/article/div[1]")

# Create an instance of ActionChains
actions = ActionChains(driver)

# Right-click (context click) on the element
actions.context_click(element).perform()

time.sleep(1)

# Click on the three-dots
# threedots = driver.find_element(By.XPATH, "/html/body/div[8]/div/div/div[2]/div[1]/div")

threedots = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".qv-objectmenu-item-container")))

time.sleep(1)

threedots.click()

# overlay = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".lui-list")))
# (Adjust the CSS selector to match the container that holds your menu items)


download = wait.until(EC.presence_of_element_located((By.ID, "export-group")))
time.sleep(1)
download.click()

download = wait.until(EC.presence_of_element_located((By.ID, "export")))
time.sleep(1)
download.click()

pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")

time.sleep(1)

pyautogui.press("enter")

time.sleep(1)
pyautogui.press("tab")
time.sleep(1)
pyautogui.press("enter")

time.sleep(5)

# Define the Downloads folder path
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# List all files in the Downloads folder
files = [os.path.join(downloads_folder, f) for f in os.listdir(downloads_folder)
         if os.path.isfile(os.path.join(downloads_folder, f))]

# Get the most recently created file
latest_file = max(files, key=os.path.getctime)

print("The latest file is:", latest_file)

# Check if the latest file was modified in the last 30 seconds
if (time.time() - os.path.getctime(latest_file)) > 30:
    raise ValueError(f"The latest file was not modified in the last 30 seconds: {latest_file}")

# Check if the latest file is an Excel file with the .xlsx extension
if not (latest_file.lower().endswith('.xlsx')):
    raise ValueError(f"The latest file is not an Excel file: {latest_file}")

# Read the Excel file using pandas
df = pd.read_excel(latest_file)

print("Successfully read Excel file:", latest_file)
print(df.head())







