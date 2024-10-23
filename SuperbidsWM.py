
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
import datetime
from datetime import datetime
import os
from selenium.webdriver.chrome.options import Options

# Set current time to a formatted string for folder name
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M')

# Set up Selenium Chrome options
opts = Options()

# Define folder path to save the output
folder_name = f"Superbids/Superbids{current_time}"
folder_path = os.path.join(r"D:/Automation/OutputW", folder_name)

# Create the folder if it doesn't exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# Set Chrome driver path
Chrome_driver_path = f"D:/Automation/Scripts/chromedriver.exe"

# Set preferences for Chrome browser (download directory)
prefs = {
    "download.default_directory": folder_path  # Use folder_path here
}
opts.add_experimental_option('prefs', prefs)
opts.add_argument(f"chromedriver-executable-dir = {Chrome_driver_path}")  # Set executable directory

# Initialize Chrome browser with specified options
browser = webdriver.Chrome(options=opts)
browser.maximize_window()

# Open the target URL
browser.get("https://sheerdrivepro.com/home/")
time.sleep(2)  # Wait for 2 seconds to load the page

# Fill in the login credentials
browser.find_element(By.ID, "mat-input-0").send_keys("SB106124C")
browser.find_element(By.ID, "mat-input-1").send_keys("Test@123")
time.sleep(5)  # Wait for 5 seconds

# Click the login button
browser.find_element(By.XPATH, "/html/body/app-root/app-home/div/app-header/mat-sidenav-container/mat-sidenav-content/div/section[2]/div/mat-card/div/mat-card-content/div/form/button[1]").click()
time.sleep(3)  # Wait for 3 seconds to ensure the login completes

# Find all event card elements on the page
elements = browser.find_elements(By.XPATH, "/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card")
print(len(elements))  # Print the number of event cards found

count = 0  # Initialize a counter
headings = ['Client', 'Start Date', 'End Date', 'Count', 'Category', 'Zone']  # Define column names for the DataFrame
temp_data = pd.DataFrame(columns=headings)  # Create an empty DataFrame with specified columns

# Iterate over the indices of elements
for i in range(len(elements)):
    # Extract details of each event card using XPath
    title = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-header/div[2]/a/strong")
    start_time = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-content/div[1]/div[1]/p[1]/span")
    end_time = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-content/div[1]/div[1]/p[1]/span")
    count_vehicle = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-content/div[2]/p[1]")
    category_vehicle = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-content/div[1]/div[2]/p")
    zone_vehicle = browser.find_element(By.XPATH, f"/html/body/app-root/app-events/app-side-nav/mat-sidenav-container/mat-sidenav-content/div/mat-drawer-container/mat-drawer-content/div/div[2]/div/mat-card[{i+1}]/mat-card-header/div[2]/p")

    # Increment counter
    count += 1

    # Append the extracted data to the DataFrame
    temp_data.loc[len(temp_data)] = [title.text, start_time.text, end_time.text, count_vehicle.text, category_vehicle.text, zone_vehicle.text]

# Replace empty strings with NaN
temp_data.replace({col: {'' : np.nan} for col in headings}, inplace=True)

# Drop rows with NaN values
temp_data.dropna(inplace=True)

# Convert start and end date to specified format
temp_data['Start Date'] = temp_data['Start Date'].apply(lambda x: pd.to_datetime(x).strftime('%d-%b-%Y'))
temp_data['End Date'] = temp_data['End Date'].apply(lambda x: pd.to_datetime(x).strftime('%d-%b-%Y'))

# Extract numeric part from the 'Count' column
temp_data['Count'] = temp_data['Count'].apply(lambda x: int(re.search(r'\d+', str(x)).group()))

# Attempt to split 'Zone' column and handle exception if it fails
try:
    temp_data['Zone'] = temp_data['Zone'].apply(lambda x: x.split('\n')[1])
except:
    temp_data['Zone'] = temp_data['Zone']

# Replace commas with spaces in 'Category' column
temp_data['Category'] = temp_data['Category'].str.replace(',', ' ')

# Create columns 'Cat_first' and 'Cat_last' by splitting 'Category'
temp_data['Cat_first'] = temp_data['Category'].apply(lambda x: x.split(',')[0]).str.strip()
temp_data['Cat_last'] = temp_data['Category'].apply(lambda x: x.split(',')[-1]).str.strip()

# Split 'Client' column based on 'Cat_first' and 'Cat_last'
temp_data['Split_Client_first'] = temp_data.apply(lambda row: row['Client'].split(row['Cat_first'])[0].strip(), axis=1)
temp_data['Split_Client_last'] = temp_data.apply(lambda row: row['Client'].split(row['Cat_last'])[-1].strip(), axis=1)

# Update 'Client' and 'Location' columns based on certain conditions
for index, row in temp_data.iterrows():
    if 'Insurance Salvage' in row['Split_Client_first']:
        temp_data.at[index, 'Client'] = 'Insurance Salvage'
        temp_data.at[index, 'Location'] = row['Zone']
    else:
        temp_data.at[index, 'Client'] = temp_data.at[index, 'Split_Client_first'].split('-')[1].strip()
        temp_data.at[index, 'Location'] = temp_data.at[index, 'Split_Client_last'].split('Online')[0]

# Select and reorder the columns for the final DataFrame
final_temp = temp_data[['Client', 'Start Date', 'End Date', 'Count', 'Category', 'Location', 'Zone']]

# Save the final DataFrame to an Excel file
current_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
file = f"SuperBids{current_time}"
try:
    final_temp.to_excel(f"D:/Automation/OutputW/{folder_name}/{file}.xlsx")
    print('Done')
except Exception as e:
    print('Contact to vipin', e)

# Close the browser
browser.quit()
