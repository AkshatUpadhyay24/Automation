import time  # Importing the time module to add delays in the script
from selenium import webdriver  # Importing webdriver from selenium for browser automation
from selenium.webdriver.common.by import By  # Importing By for locating elements
import pandas as pd  # Importing pandas for data manipulation and analysis
import numpy as np  # Importing numpy for numerical operations
import datetime  # Importing datetime module for date and time operations
from datetime import datetime  # Importing datetime class from datetime module
import os  # Importing os module for interacting with the operating system
from selenium.webdriver.chrome.options import Options  # Importing Options for setting Chrome options

# Get current time and format it for folder name
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M')
folder_name = f"Cardekho/Cardekho{current_time}"  # Define the folder name with current time
opts = Options()  # Initialize options for Chrome

# Define the path where the output will be saved
folder_path = os.path.join(r"D:/Automation/OutputW", folder_name)
# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Define the path to the Chrome driver
Chrome_driver_path = f"D:/Automation/Scripts/chromedriver.exe"

# Set preferences for the Chrome browser (download directory)
prefs = {
    "download.default_directory": folder_path  # Use folder_path here
}
opts.add_experimental_option('prefs', prefs)  # Add experimental options to the Chrome instance
opts.add_argument(f"chromedriver-executable-dir = {Chrome_driver_path}")  # Set executable directory

# Initialize the Chrome browser with specified options
browser = webdriver.Chrome(options=opts)
browser.maximize_window()  # Maximize the browser window

# Open the target URL
browser.get("https://auctions.cardekho.com/#/")
time.sleep(3)  # Wait for 3 seconds to load the page

try:
    # Fill in the login credentials and submit the form
    browser.find_element(By.XPATH, '//header/div[1]/div[1]/form[1]/div[1]/input[1]').send_keys("MH202144416")
    browser.find_element(By.XPATH, "//header/div[1]/div[1]/form[1]/div[2]/input[1]").send_keys("Rahul@123")
    browser.find_element(By.XPATH, "//header/div[1]/div[1]/form[1]/button[1]").click()
    time.sleep(4)  # Wait for 4 seconds to ensure the login completes

    # Navigate to a specific section of the website after logging in
    browser.find_element(By.XPATH, "html/body/div[1]/div[1]/div[4]/div/ul/li[4]/a").click()
    time.sleep(4)  # Wait for 4 seconds

    # Navigate to the auction listing page
    browser.get("https://auctions.cardekho.com/#/auctionListing/bank_live")
    time.sleep(2)  # Wait for 2 seconds to load the page

    # Find elements containing auction names and vehicle counts
    auction_name = browser.find_elements(By.XPATH, "/html/body/div[1]/div[3]/div/table[1]/tbody[2]/tr/td[1]")
    vehicles = browser.find_elements(By.XPATH, '//table[@class="live_auctions"]//tbody[2]//td[6]')

    l = []  # Initialize list for auction names
    m = []  # Initialize list for vehicle counts

    # Iterate over auction names and vehicle counts
    for i in range(len(auction_name)):
        l.append(auction_name[i].text)  # Append auction name to the list
        vehicle_data = vehicles[i].text  # Get vehicle data text
        vehicle_list = vehicle_data.split(" ")  # Split vehicle data into a list
        total_vehicle_no = sum([int(v.strip()) for v in vehicle_list])  # Calculate total number of vehicles
        m.append(total_vehicle_no)  # Append total vehicle count to the list

    # Create a DataFrame from the lists
    df_final = pd.DataFrame(list(zip(l, m)), columns=['Client', 'Count'])

except Exception as e:
    # Print error message if an exception occurs
    print("Contact to Vipin1", e)

# Define the output file name
file = f"cardekho{current_time}"

try:
    # Save the DataFrame to an Excel file
    df_final.to_excel(f"D:/Automation/OutputW/{folder_name}/{file}.xlsx")
    print('Done')
except Exception as e:
    # Print error message if an exception occurs
    print("contact to vipin2", e)

# Close the browser
browser.quit()