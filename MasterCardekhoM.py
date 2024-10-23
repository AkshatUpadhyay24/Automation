#pip install pyxlsb
import openpyxl
import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# Create an empty dataframe
headings = ['Make','Product','Lan No', 'Client','Comp Name','Location','State','Region','Auction Date','Auction Type',
            'Verticle','Month','Count']
data = []
sheet1 = pd.read_excel("D://Automation//Scripts//final_verticle_file.xlsb",sheet_name = 'Cardekho_Verticle')

# Define the data as a dictionary
state_data = {
    'State_': [
        'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh',
        'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir',
        'Jharkhand', 'Karnataka', 'Kerala', 'Madhya Pradesh', 'Maharashtra',
        'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Orissa',
        'Punjab', 'Rajasthan', 'Sikkim', 'Telangana', 'Tripura', 'Uttarakhand',
        'Uttar Pradesh', 'West Bengal', 'Tamil Nadu', 'Tripura', 'Andaman and Nicobar Islands','Chandigarh',
        'Dadra and Nagar Haveli', 'Daman and Diu', 'Delhi','Lakshadweep', 'Pondicherry','East','West','North','South','Delhi','North'
    ],
    'Abbreviation': [
        'AP', 'AR', 'AS', 'BR', 'CG',
        'GA', 'GJ', 'HR', 'HP', 'JK',
        'JH', 'KA', 'KL', 'MP', 'MH',
        'MN', 'ML', 'MZ', 'NL', 'OR',
        'PB', 'RJ', 'SK', 'TS', 'TR', 'UK',
        'UP', 'WB', 'TN', 'TR', 'AN', 'CH',
        'DH', 'DD', 'DL', 'LD', 'PY','East','West','North','South','Delhi','NE'
    ],
    'Region_': [
        'South', 'East', 'East', 'East', 'West',
        'West', 'West', 'North', 'North', 'North',
        'East', 'South', 'South', 'West', 'West',
        'East', 'East', 'East', 'East', 'East',
        'North', 'North', 'East', 'South', 'East',
        'North', 'North', 'West', 'South', 'East', 'South',
        'North', 'West', 'West', 'North', 'South', 'South','East','West','North','South','North','North'
    ]
}

# Create the DataFrame
df_state = pd.DataFrame(state_data)
folder_path = f"D://Automation//OutputW//Cardekho//"

excel_files_list = [
    os.path.join(root, file)
    for root, _, files in os.walk(folder_path)
    for file in files
    if (file.endswith(".xlsx") or file.endswith(".xls"))
]
#files = [file for file in os.listdir(folder_path) if (file.endswith(".xlsx") or file.endswith(".xls"))]
data = []
data1 = []
#appended_df = pd.DataFrame(columns=df.columns)
for file in excel_files_list:
    df = pd.DataFrame(columns = headings)
    appended_df = pd.DataFrame(columns=df.columns)
    #file_path = os.path.join(folder_path, file)
    #print(file_path)
    cardekho_combine = pd.read_excel(file,index_col = 0)
    data1.append(cardekho_combine)
cardekho = pd.concat(data1, ignore_index=True)
cardekho.drop_duplicates(inplace=True)
    
cardekho = cardekho[cardekho['Client'].str.len() >= 8].reset_index(drop = True)

print(cardekho['Count'].sum())
keywords = ['4W', 'FE', 'CV', '2W','CE','3W']
pattern = '|'.join(rf'\b{k}\b' for k in keywords)
#print(pattern)
df['Client']=cardekho['Client'].str.split(pattern).str[0]
df['Make'] = cardekho['Client'].str.split('Online Auction').str[0].str.split().str[-1]
keywords = ['4W', 'FE', 'CV', '2W','CE','3W']
mask = df['Make'].isin(keywords)
df=df[mask]
df['Product'] = df['Make']
df['Comp Name'] = 'Cardekho'


df['Auction Date'] = cardekho['Client'].str.split('Online Auction').str[1].str.split().str[-1]
df['Auction Date'] = pd.to_datetime(df['Auction Date'])
df['Auction Date'] = df['Auction Date'].dt.strftime('%d-%b-%y')

df['Location'] = cardekho['Client'].str.split('Online Auction').str[1].str.split().str[0]
df['Location'] = df['Location'].str.replace('\d+', '', regex=True) # remove the number from up1 and so on

merged_df = pd.merge(df,df_state,how = "left", left_on = "Location",right_on = "Abbreviation")
df['State'] = merged_df['State_'].combine_first(df['State'])
df['Region'] = merged_df['Region_'].combine_first(df['Region'])

df['Auction Type'] = 'Online'
#df['Auction Date'] = superbids['End Date'] 
current_date = datetime.now()
formatted_date = current_date.strftime("%d-%b-%Y")

df['Month'] = formatted_date
df['Count']= cardekho['Count']  


df['Client']= df.Client.str.strip().values
sheet1['CarDekho'] = sheet1['CarDekho'].str.strip().values
sheet1['FinalVertical'] = sheet1['FinalVertical'].str.strip().values


merged_df = pd.merge(df,sheet1, how="left", left_on="Client", right_on="CarDekho")
df['Verticle'] = merged_df['FinalVertical'].combine_first(df['Verticle'])

for index, row in df.iterrows():
    category = row['Make']
    count = int(row['Count'])
    cat_split = category.split()

    # single category logic
    if len(category.split())==1:
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*count,ignore_index = True)
        #appended_df['Lan No'] = 'Dummy ' + (appended_df.index + 1).astype(str)



    #double category with even number logic
    elif len(category.split()) ==2 and (count % 2 == 0):
        #print('##', category, count)
        #cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)

    # double category with odd number logic    
    elif len(category.split()) ==2 and (count % 2 == 1):
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//2)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)




     # triple category with even number logic   
    elif len(category.split()) ==3 and (count % 3 == 0):
        #print('###', category, count)
        #cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)

    elif len(category.split()) ==3 and (count % 3 == 1):
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True) 
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*((count//3)+1),ignore_index = True)

    elif len(category.split()) ==3 and (count % 3 == 2):
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*((count//3)+1),ignore_index = True) 
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*((count//3)+1),ignore_index = True)





     # quad category with even number logic
    elif len(category.split()) ==4 and (count % 4 == 0):
        #print('###', category, count)
        #cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)


    elif len(category.split()) ==4 and (count % 4 == 1):
        #print('###', category, count)
        cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)       


    elif len(category.split()) ==4 and (count % 4 == 2):
        print('###', category, count)
        cat_split = category.split()
        print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)       


    elif len(category.split()) ==4 and (count % 4 == 3):
        print('###', category, count)
        cat_split = category.split()
        print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)    
data.append(appended_df)
final_df_cardekho = pd.concat(data,ignore_index = True) 
dummy_values = ["Dummy " + str(i) for i in range(1, len(final_df_cardekho) + 1)]
final_df_cardekho["Lan No"].fillna(pd.Series(dummy_values), inplace=True)    
columns_to_delete = ["Count"]
final_df_cardekho = final_df_cardekho.drop(columns=columns_to_delete, axis=1)
folder_path = 'D://Automation//OutputM//'  # Replace with the desired parent directory path
folder_name = 'cardekho'  # Replace with the desired folder name

# Create the folder
full_folder_path = os.path.join(folder_path, folder_name)
if not os.path.exists(full_folder_path):
    os.makedirs(full_folder_path)
final_df_cardekho.to_csv(os.path.join(full_folder_path, 'masterfile_cardekho.csv'), index=False)
print('Done')    