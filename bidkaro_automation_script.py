import time  # Importing time module for adding delays
import random  # Importing random module for generating random numbers
from selenium import webdriver  # Importing webdriver from selenium for browser automation
import sys  # Importing sys module for system-specific parameters and functions
import os  # Importing os module for interacting with the operating system
import glob  # Importing glob module for file pattern matching
import time, datetime  # Importing time and datetime modules for time-related functions
from datetime import datetime  # Importing datetime class from datetime module

from selenium.webdriver.common.by import By  # Importing By for locating elements
from selenium.webdriver.chrome.options import Options  # Importing Options for setting Chrome options
from selenium.webdriver.support.ui import WebDriverWait  # Importing WebDriverWait for waiting conditions
from selenium.webdriver.support import expected_conditions as EC  # Importing expected_conditions for waiting conditions
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException  # Importing exceptions
import time  # Importing time module for adding delays
from selenium import webdriver  # Importing webdriver from selenium for browser automation
import os  # Importing os module for interacting with the operating system
from webdriver_manager.chrome import ChromeDriverManager  # Importing ChromeDriverManager for managing Chrome driver
from selenium.webdriver.common.by import By  # Importing By for locating elements
from selenium.webdriver.chrome.options import Options  # Importing Options for setting Chrome options
import glob  # Importing glob module for file pattern matching
import time, datetime  # Importing time and datetime modules for time-related functions
from datetime import datetime  # Importing datetime class from datetime module

# Get current time and format it for folder name
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M')

opts = Options()  # Initialize options for Chrome

# Define the folder name with current time
folder_name = f"Bidkaro\\Bidkaro{current_time}"
current_folder = folder_name
# Define the path where the output will be saved
folder_path = os.path.join(r"D:\Automation\OutputW", folder_name)
# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Define the path to the Chrome driver
Chrome_driver_path = r"D:/Automation/Scripts/chromedriver.exe"

# Set preferences for the Chrome browser (download directory)
prefs = {
    "download.default_directory": folder_path  # Use folder_path here
}
opts.add_experimental_option('prefs', prefs)  # Add experimental options to the Chrome instance
opts.add_argument(f"chromedriver-executable-dir = {Chrome_driver_path}")  # Set executable directory

# Initialize the Chrome browser with specified options
browser = webdriver.Chrome(options=opts)
browser.maximize_window()  # Maximize the browser window

try:
    # 1. Open the website and fill in login credentials
    browser.get("https://www.bidkaro.net/")
    time.sleep(5)  # Wait for 5 seconds to load the page
    browser.find_element(By.ID, 'username').send_keys("Pstopcars")  # Enter username
    browser.find_element(By.XPATH, "//input[@placeholder='password']").send_keys("Rahul@1234")  # Enter password
    time.sleep(3)  # Wait for 3 seconds
    browser.find_element(By.XPATH, "//button[normalize-space()='sign in']").click()  # Click sign in button
    browser.get("https://www.bidkaro.net/events/live-auctions")  # Navigate to live auctions page
    time.sleep(4)  # Wait for 4 seconds
    count = browser.find_element(By.XPATH, "/html/body/bidkaro-root/div/bidkaro-live-auctions/div/tabset/ul/li[1]/a/span")
    num = count.text.strip().split(" ")[1]  # Extract the number of live auctions
    num1 = int(num.replace("(", "").replace(")", ""))  # Convert the extracted number to integer
    number_of_pages = (num1 // 100) + 1  # Calculate the number of pages
    page_number = 0  # Initialize page number

    while True:
        # 2. Download all the files on the current page
        download_buttons = browser.find_elements(By.XPATH, "//button[@class='btn btn-sm btn-outline-secondary']")
        for i, button in enumerate(download_buttons):
            try:
                button.click()  # Click the download button
                time.sleep(4)  # Wait for 4 seconds for the file to download

                # Change directory to the download folder
                os.chdir(f"D:\\Automation\\OutputW\\{current_folder}\\")
                # Get the list of downloaded files
                list_of_files = glob.glob(f"D:\\Automation\\OutputW\\{current_folder}\\*")
                latest_file = max(list_of_files, key=os.path.getctime)  # Get the latest downloaded file
                oldname = latest_file.split("\\")[-1]  # Extract the old file name
                addname = oldname.split(".")[0]  # Extract the file name without extension

                # Extract auction details for renaming the file
                file1 = browser.find_element(By.XPATH, f"/html/body/bidkaro-root/div/bidkaro-live-auctions/div/tabset/div/tab[1]/bidkaro-live-auction-list/bidkaro-auction-list/div/p-table/div/div[2]/table/tbody/tr[{i+1}]/td[1]")
                file2 = browser.find_element(By.XPATH, f"/html/body/bidkaro-root/div/bidkaro-live-auctions/div/tabset/div/tab[1]/bidkaro-live-auction-list/bidkaro-auction-list/div/p-table/div/div[2]/table/tbody/tr[{i+1}]/td[3]")
                day = file2.text.split("-")[0]  # Extract day from date
                month = file2.text.split("-")[1]  # Extract month from date
                word = day + month  # Concatenate day and month
                file = word + '-' + file1.text  # Create new file name
                newname2 = addname + "-" + file  # Combine old name with new file details

                # Check and replace '/' with '-' in the new file name
                if '/' not in newname2:
                    newname = f'{newname2}.xlsx'
                    os.rename(oldname, newname)  # Rename the file
                else:
                    newname2 = newname2.replace("/", "-")
                    newname = f'{newname2}.xlsx'
                    os.rename(oldname, newname)  # Rename the file

            except Exception as e:
                print(f"Error downloading file: {e}")

        # 3. Check if next page is available and click on it
        try:
            next_button = browser.find_element(By.XPATH, "/html/body/bidkaro-root/div/bidkaro-live-auctions/div/tabset/div/tab[1]/bidkaro-live-auction-list/bidkaro-auction-list/div/p-table/div/p-paginator/div/a[3]")
            next_button.click()  # Click the next page button
            time.sleep(3)  # Wait for 3 seconds

            page_number += 1  # Increment the page number
            if(page_number >= number_of_pages):  # Check if all pages are processed
                break

        except NoSuchElementException:
            # 4. If the "Next" button is not available, exit the loop
            print("No more pages available. Exiting...")
            break

except Exception as e:
    print(f"Contact to Vipin: {e}")

finally:
    # Close the browser
    print('done')
    browser.quit()


