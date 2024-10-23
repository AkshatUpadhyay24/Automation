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
folder_name = f"Ediig\\Ediig{current_time}"
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

# Open the website
browser.get("https://ediig.com/")
time.sleep(3)  # Wait for 3 seconds to load the page

# Enter login credentials and sign in
browser.find_element(By.ID, 'buyerLoginName').send_keys("Pstopcars")  # Enter username
browser.find_element(By.XPATH, "//input[@id='buyerPassword']").send_keys("eDiig@123")  # Enter password
browser.find_element(By.XPATH, "//button[contains(text(),'Sign In')]").click()  # Click sign in button
time.sleep(2)  # Wait for 2 seconds

# Navigate to Sale Calendar
browser.find_element(By.XPATH, "//a[contains(text(),'Sale Calendar')]").click()
time.sleep(4)  # Wait for 4 seconds to load the page

# Initialize a condition for the while loop
condition = True

# Loop to download files and handle pagination
while condition:
    time.sleep(3)  # Wait for 3 seconds
    # Find all the elements for downloading files
    download_files_ediig = browser.find_elements(By.XPATH, "//td[@class='sorting_1']")
    for i, button in enumerate(download_files_ediig):
        try:
            button.click()  # Click the download button
            time.sleep(3)  # Wait for 3 seconds for the file to download

            # Change directory to the download folder
            os.chdir(f"D:\\Automation\\OutputW\\{current_folder}\\")
            # Get the list of downloaded files
            list_of_files = glob.glob(f"D:\\Automation\\OutputW\\{current_folder}\\*")
            latest_file = max(list_of_files, key=os.path.getctime)  # Get the latest downloaded file
            oldname = latest_file.split("\\")[-1]  # Extract the old file name
            addname = oldname.split(".")[0]  # Extract the file name without extension

            # Extract auction details for renaming the file
            file1 = browser.find_element(By.XPATH, f"//*[@id='saleCalendarTabel']/tbody/tr[{i+1}]/td[1]")
            seller_name = browser.find_element(By.XPATH, f"//*[@id='saleCalendarTabel']/tbody/tr[{i+1}]/td[3]")
            location = browser.find_element(By.XPATH, f"//*[@id='saleCalendarTabel']/tbody/tr[{i+1}]/td[4]")            
            newname2 = addname + "-" + file1.text + "-" + seller_name.text + "-" + location.text  # Create new file name

            # Check and replace '/' with '-' in the new file name
            if '/' not in newname2:
                newname = f'{newname2}.xls'
                os.rename(oldname, newname)  # Rename the file
            else:
                newname2 = newname2.replace("/", "-")
                newname = f'{newname2}.xls'
                os.rename(oldname, newname)  # Rename the file

        except:
            print('button is not clickable')

    try:
        # Find the next page button and click it
        pages = browser.find_element(By.XPATH, "//*[@id='saleCalendarTabel_next']/a")
        pages.click()
        time.sleep(2)  # Wait for 2 seconds
    except:
        print('no click element available')
        condition = False  # Exit the loop if no more pages are available

print('Done')  # Print done when all files are downloaded
browser.quit()  # Close the browser

