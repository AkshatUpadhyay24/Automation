#!pip install pyxlsb
import pandas as pd
import numpy as np
import os
import re
from datetime import datetime
import warnings

# Ignore DeprecationWarning
warnings.simplefilter("ignore")


# Create an empty dataframe
headings = ['Make','Product','Lan No', 'Client','Comp Name','Location','State','Region','Auction Date','Auction Type',
            'Verticle','Month','count']
#df = pd.DataFrame(columns = headings)

sheet1 = pd.read_excel("D://Automation//Scripts//final_verticle_file.xlsb",sheet_name = 'Superbids_Verticle')
#C:\Users\Piyush\Desktop\Automation\SuperBids
folder_path = "D://Automation//OutputW//Superbids//"
excel_files_list = [
    os.path.join(root, file)
    for root, _, files in os.walk(folder_path)
    for file in files
    if (file.endswith(".xlsx") or file.endswith(".xls"))
]


data = []
data1 = []


for file in excel_files_list:
    df = pd.DataFrame(columns = headings)
    appended_df = pd.DataFrame(columns=df.columns)

    #file_path = os.path.join(folder_path, file)
    #print(file_path)
    superbid = pd.read_excel(file,index_col = 0)
    data1.append(superbid)
superbids = pd.concat(data1, ignore_index=True)
superbids.drop_duplicates(inplace=True)
df['Client'] = superbids['Client']



#for word in words_to_remove:
    #df["State"] = df["State"].str.replace(word, "")
df['Location']=superbids['Location']
df['Region']=superbids['Zone']
df['State']=superbids['Location']

#df['Location'] = superbids['State']
df['Comp Name'] = 'Superbids'
#df['Lan No'] = 'Dummy ' + (df.index + 1).astype(str)
#df['Lan No'] = 'Dummy 1'
df['Auction Type'] = 'Online'
df['Auction Date'] = superbids['End Date'] 
current_date = datetime.now()
formatted_date = current_date.strftime("%d-%b-%Y")

df['Month'] = formatted_date
df['Make'] = superbids['Category']
df['Product'] = superbids['Category']
df['count']= superbids['Count'] 



# logic of Verticle
df['Client'] = df['Client'].str.strip()
merged_df = pd.merge(df,sheet1, how="left", left_on="Client", right_on="SuperbidsClient")
#print(df)
df['Verticle'] = merged_df['FinalVertical'].combine_first(df['Verticle'])
#print('df',df)    
for index, row in df.iterrows():
    #print('###',index,row)
    category = row['Make']
    try:
        count = int(float(row['count']))
    except:
        print('count is not readble')
    cat_split = category.split()
    #print('#####', category, count, cat_split, index)
   # single category logic
    if (len(category.split())==1) :
        #print('#',category, count)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*count,ignore_index = True)
        #print('##########################################################',appended_df)


    # double category with even number logic
    elif len(category.split()) ==2 and (count % 2 == 0):
        #cat_split = category.split()
        #print('##', category, count)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)

        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)
        #print('##########################################################',appended_df)

    # double category with odd number logic    
    elif len(category.split()) ==2 and (count % 2 == 1):
        #print('##@', category, count)

        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//2)+1),ignore_index = True)

        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//2),ignore_index = True)




     # triple category with even number logic   
    elif len(category.split()) ==3 and (count % 3 == 0):
        #print('###', category, count)

        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)

    elif len(category.split()) ==3 and (count % 3 == 1):
        #print('###@', category, count)

        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*(count//3),ignore_index = True) 
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*((count//3)+1),ignore_index = True)

    elif len(category.split()) ==3 and (count % 3 == 2):
        #print('###@@', category, count)

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
        #cat_split = category.split()
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
        #print('###', category, count)
        #cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)       


    elif len(category.split()) ==4 and (count % 4 == 3):
        #print('###', category, count)
        #cat_split = category.split()
        #print(cat_split)
        row['Make'] = cat_split[0]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[1]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[2]
        appended_df = appended_df._append([row]*((count//4)+1),ignore_index = True)
        row['Make'] = cat_split[3]
        appended_df = appended_df._append([row]*(count//4),ignore_index = True)    
data.append(appended_df)
final_df_superbids = pd.concat(data,ignore_index = True)
dummy_values = ["Dummy " + str(i) for i in range(1, len(final_df_superbids) + 1)]
final_df_superbids["Lan No"].fillna(pd.Series(dummy_values), inplace=True)    
columns_to_delete = ["count"]
final_df_superbids = final_df_superbids.drop(columns=columns_to_delete, axis=1)
folder_path = 'D://Automation//OutputM//'  # Replace with the desired parent directory path
folder_name = 'superbids'  # Replace with the desired folder name
#D:\Automation\OutputW
# Create the folder
full_folder_path = os.path.join(folder_path, folder_name)
if not os.path.exists(full_folder_path):
    os.makedirs(full_folder_path)
final_df_superbids.to_csv(os.path.join(full_folder_path, 'masterfile_superbids.csv'), index=False)
print('Done')