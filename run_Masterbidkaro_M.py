import math  # Importing math module for mathematical functions
import pandas as pd  # Importing pandas for data manipulation and analysis
pd.set_option('display.float_format', lambda x: '%0.3f' % x)  # Setting display option for float format in pandas
pd.set_option('display.max_columns', 100)  # Setting maximum number of columns to display
pd.set_option('display.max_rows', 1000)  # Setting maximum number of rows to display
import warnings  # Importing warnings module to manage warning messages
from datetime import datetime  # Importing datetime for handling date and time
import numpy as np  # Importing numpy for numerical operations
import os  # Importing os module for interacting with the operating system
import re  # Importing re module for regular expressions

# Ignore DeprecationWarning
warnings.simplefilter("ignore")
warnings.filterwarnings('ignore')

import matplotlib.pyplot as plt  # Importing matplotlib for plotting
from datetime import datetime, timedelta  # Importing datetime and timedelta from datetime module

# Read the Excel file with Bidkaro vertical data
sheet1 = pd.read_excel("D://Automation//Scripts//final_verticle_file.xlsb", sheet_name='Bidkaro_Verticle')
sheet1['Bidkaro'] = sheet1['Bidkaro'].str.strip()  # Strip whitespace from 'Bidkaro' column
sheet1['FinalVertical'] = sheet1['FinalVertical'].str.strip()  # Strip whitespace from 'FinalVertical' column

# Define state abbreviations and regions
state_df = {
    'State_': [
        'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh',
        'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir',
        'Jharkhand', 'Karnataka', 'Kerala', 'Madhya Pradesh', 'Maharashtra',
        'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Orissa',
        'Punjab', 'Rajasthan', 'Sikkim', 'Telangana', 'Tripura', 'Uttarakhand',
        'Uttar Pradesh', 'West Bengal', 'Tamil Nadu', 'Tripura', 'Andaman and Nicobar Islands',
        'Chandigarh', 'Dadra and Nagar Haveli', 'Daman and Diu', 'Delhi',
        'Lakshadweep', 'Pondicherry', 'East', 'West', 'North', 'South', 'Delhi'
    ],
    'Abbreviation': [
        'AP', 'AR', 'AS', 'BR', 'CG',
        'GA', 'GJ', 'HR', 'HP', 'JK',
        'JH', 'KA', 'KL', 'MP', 'MH',
        'MN', 'ML', 'MZ', 'NL', 'OR',
        'PB', 'RJ', 'SK', 'TS', 'TR', 'UK',
        'UP', 'WB', 'TN', 'TR', 'AN', 'CH',
        'DH', 'DD', 'DL', 'LD', 'PY', 'East', 'West', 'North', 'South', 'North'
    ],
    'Region_': [
        'South', 'East', 'East', 'East', 'West',
        'West', 'West', 'North', 'North', 'North',
        'East', 'South', 'South', 'West', 'West',
        'East', 'East', 'East', 'East', 'East',
        'North', 'North', 'East', 'South', 'East',
        'North', 'North', 'West', 'South', 'East', 'South',
        'North', 'West', 'West', 'North', 'South', 'South', 'East', 'West', 'North', 'South', 'North'
    ]
}

# Create a DataFrame from the state dictionary
df_state = pd.DataFrame(state_df)

# Set the working directory to the specified path
os.chdir(f"D:\\Automation\\OutputW\\Bidkaro\\")
folder_path = f"D:\\Automation\\OutputW\\Bidkaro\\"

# Get list of all Excel files in the directory and subdirectories
excel_files_list = [
    os.path.join(root, file)
    for root, _, files in os.walk(folder_path)
    for file in files
    if (file.endswith(".xlsx") or file.endswith(".xls"))
]

data = []  # Initialize an empty list to store data
current_year = datetime.now().year  # Get the current year

# Loop through each Excel file in the list
for file in excel_files_list:
    try:
        # Split the file path to get file name components
        file_name = file.split('\\')[-1].split('-')
    except:
        print('File name issue')

    try:
        # Read the Excel file and drop rows with missing values in 'Loan A/C No' column
        excel = pd.read_excel(file, usecols=[4]).dropna()
        for i in excel['Loan A/C No']:
            loan = i  # Loan account number
            auction_date = file_name[5] + '-' + str(current_year)  # Auction date
            client = file_name[6]  # Client name
            state = file_name[7]  # State
            make = file_name[8].split('.')[0]  # Make of the vehicle
            current_date = datetime.now()  # Current date and time
            formatted_date = current_date.strftime("%d-%b-%Y")  # Formatted date

            # Append data to the list
            data.append({
                'Make': make,
                'Product': make,
                'Lan No': loan,
                'Client': client,
                'Comp Name': 'Bidkaro',
                'Location': state,
                'State': state,
                'Region': 'nan',
                'Auction Date': auction_date,
                'Auction Type': 'Online',
                'Verticle': 'nan',
                'Month': formatted_date
            })
    except:
        print('file issue', file)

# Create a DataFrame from the data list
df = pd.DataFrame(data)
df['Client'] = df['Client'].str.strip()  # Strip whitespace from 'Client' column

# Merge DataFrame with state DataFrame to get the region
try:
    merged_df = pd.merge(df, df_state, how="left", left_on="State", right_on="State_")
    df['Region'] = merged_df['Region_'].combine_first(df['Region'])

    # Merge DataFrame with vertical data to get the vertical
    merged_df1 = pd.merge(df, sheet1, how="left", left_on="Client", right_on="Bidkaro")
    df['Verticle'] = merged_df1['FinalVertical'].combine_first(df['Verticle'])
    df.drop_duplicates(subset=['Lan No'], keep='first', inplace=True)  # Remove duplicate loan numbers

except:
    print('check region and verticle')

# Define folder path to save the output
folder_path = 'D://Automation//OutputM//'  # Replace with the desired parent directory path
folder_name = 'bidkaro'  # Replace with the desired folder name

# Create the folder if it doesn't exist
full_folder_path = os.path.join(folder_path, folder_name)
if not os.path.exists(full_folder_path):
    os.makedirs(full_folder_path)

# Save the DataFrame to a CSV file
df.to_csv(os.path.join(full_folder_path, 'masterfile_bidkaro.csv'), index=False)
print('Done')