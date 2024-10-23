import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import re
from datetime import datetime
import datetime
from datetime import datetime, timedelta

#from sklearn.model_selection import RandomizedSearchCVa
pd.set_option('display.float_format', lambda x: '%0.3f' % x)
pd.set_option('display.max_columns',100)
pd.set_option('display.max_rows',1000)
import warnings
warnings.filterwarnings('ignore')


sheet1 = pd.read_excel("D://Automation//Scripts//final_verticle_file.xlsb",sheet_name = 'Ediig_Verticle')
sheet1['Ediig']=sheet1['Ediig'].str.strip()
sheet1['FinalVertical']=sheet1['FinalVertical'].str.strip()


# Define the data as a dictionary
state_df = {
    'State_': [
        'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh',
        'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir',
        'Jharkhand', 'Karnataka', 'Kerala', 'Madhya Pradesh', 'Maharashtra',
        'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Orissa',
        'Punjab', 'Rajasthan', 'Sikkim', 'Telangana', 'Tripura', 'Uttarakhand',
        'Uttar Pradesh', 'West Bengal', 'Tamil Nadu', 'Tripura', 'Andaman and Nicobar Islands',
        'Chandigarh', 'Dadra and Nagar Haveli', 'Daman and Diu', 'Delhi',
        'Lakshadweep', 'Pondicherry','East','West','North','South','Delhi'
    ],
    'Abbreviation': [
        'AP', 'AR', 'AS', 'BR', 'CG',
        'GA', 'GJ', 'HR', 'HP', 'JK',
        'JH', 'KA', 'KL', 'MP', 'MH',
        'MN', 'ML', 'MZ', 'NL', 'OR',
        'PB', 'RJ', 'SK', 'TS', 'TR', 'UK',
        'UP', 'WB', 'TN', 'TR', 'AN', 'CH',
        'DH', 'DD', 'DL', 'LD', 'PY','East','West','North','South','Delhi'
    ],
    'Region_': [
        'South', 'East', 'East', 'East', 'West',
        'West', 'West', 'North', 'North', 'North',
        'East', 'South', 'South', 'West', 'West',
        'East', 'East', 'East', 'East', 'East',
        'North', 'North', 'East', 'South', 'East',
        'North', 'North', 'West', 'South', 'East', 'South',
        'North', 'West', 'West', 'North', 'South', 'South','East','West','North','South','North'
    ]
}

# Create the DataFrame
df_state = pd.DataFrame(state_df)

def is_valid_alphanumeric(value):
    return bool(re.match(r'^[A-Za-z0-9]+$', str(value)))
os.chdir(f"D:\\Automation\\OutputW\\Ediig\\")
folder_path = f"D:\\Automation\\OutputW\\Ediig\\"
excel_files_list = [
    os.path.join(root, file)
    for root, _, files in os.walk(folder_path)
    for file in files
    if (file.endswith(".xlsx") or file.endswith(".xls"))
]
data = []
try:
    for file in excel_files_list:
        file_path = os.path.join(folder_path, file)

        try:
            df = pd.read_excel(file_path,header=None)
            #print('#####',file_path)
        except:
            print('Issue in file please check',file_path) #AGRNO  Contract No


        keywords_to_find = ['Contract Number','Agreement Number','AgreementNo','LOS','HP NO.','Loan No','Loan No.', 'Contract No.',
                            'Contract No / Loan No','Agreement No.','Agreement Number*','Deal No','AGRNO','Loan Account #',      
                            'KOTAK Agmt. No.','Contract No','Loan Agreement No','LAN','Loan Number','CONTRACTNUMBER',
                            'Agreement No.','AGREEMENTID']#AGRNO


        found_keywords = []
        try:
            for column in df.columns:
                for keyword in keywords_to_find:
                    if keyword in df[column].values:
                        found_keywords.append(keyword)
                        final_found_word = set(found_keywords)
                        final_found_words = list(final_found_word)
                        #print('keywords found',final_found_words,file_path)
                        print("Found Keywords:", '###',final_found_words, found_keywords)
        except:
            print('keywords not found ')
        try:


            row_index, col_index = [(i, j) for i, row in df.iterrows() for j, cell in enumerate(row) if cell in final_found_words], [(j, i) for i, col in df.items() for j, cell in enumerate(col) if cell in final_found_words]
            
        except:
            print('###row index not found',file)
            

        try:
            if (len(found_keywords)== 1):
                #print(count3)
                #print(f"Row Index: {row_index[0]}, Column Index: {col_index[0]}")
                df1= df.iloc[row_index[0][0]:,[row_index[0][1]]].values

                df2 = pd.DataFrame(df1)
                df2.columns = df2.iloc[0]
                df2 = df2[1:]
                df2 = df2.reset_index(drop=True)
                #df2.columns = df2.columns.str.replace(r'^\d+\s+', '', regex=True)
                #print(df2)df.rename(columns=lambda x: x.lstrip('0 '), inplace=True)
                df2.dropna(inplace = True)

                df_final1 = df2[df2[df2.columns[0]].apply(lambda x: is_valid_alphanumeric(x) and (not (str(x).isdigit() and len(str(x)) < 5)))]
                #print(df_final1.columns)
                for i in df_final1.values:
                    lan_no = i[0]
                    #print('lan no1',lan_no,file)
                    #file_path = os.path.join(folder_path, file)
                    file_name = file_path.split('/')[-1].split('-')


                    # logic of location and state
                    location = file_name[-1].strip().split('.')[0] # location state
                    final_location = location.title()

                    Client_name = file_path.split('/')[-1].split('-')[-2] # client
                    final_client_name = Client_name.title()
                    try:
                        demo = Client_name + '-' + location + '.xls'
                    except:
                        demo = Client_name + '-' + location + '.xlsx'

                    company_name = 'Ediig'# company name

                    # Logic of month
                    current_date = datetime.now()
                    formatted_date = current_date.strftime("%d-%b-%Y") # month


                    # logic of make and product
                    try:
                        category = file_path.split('/')[-1].split("ET")[1].strip()
                    except:
                        pass

                    result_2 = category.replace(demo, "").strip()
                    #print(category,'#####', demo,'@@@@',result_2)
                    make = result_2[1:-1] if result_2.startswith("-") and result_2.endswith("-") else result_2.lstrip("-").rstrip("-") # make and product

                    auction_type = "Online"  #auction type

                    # Logic of auction date
                    try:
                        auction_date = file_name[1] +'-' + file_name[2] + '-' + str((datetime.now().year)% 100)
                    except:
                        auction_date = file_name[1] + '-' + str((datetime.now().year)% 100)
                    data.append({'Make':make,'Product':make,'Lan No':lan_no,'Client':final_client_name,'Comp Name':'Ediig','Location':final_location,'State':final_location,'Region':'nan','Auction Date': auction_date,'Auction Type':'Online','Verticle':'nan','Month':formatted_date})

            elif (len(found_keywords)== 2):
                # for first occurence 
                df11= df.iloc[row_index[0][0]:,[row_index[0][1]]]
                df11=df11.reset_index(drop = True)
                df11.columns = df11.iloc[0]
                df11 = df11[1:]
                #df1.dropna(inplace = True)
                df11=df11.reset_index(drop = True)
                try:
                    nan_index = df11.index[df11[df11.columns[0]].isna()].tolist()[0]
                    df11_cleaned = df11.iloc[:nan_index]

                except:
                    print('file issue2')

                # for second occurence
                try:
                    df22= df.iloc[row_index[1][0]:,[row_index[1][1]]]
                    df22=df22.reset_index(drop = True)
                    df22.columns = df22.iloc[0]
                    df22 = df22[1:]
                    df22=df22.reset_index(drop = True)
                except:
                    print('file issue2')
                # combine dataframe
                try:
                    combined_df = pd.concat([df11_cleaned,df22], ignore_index = True)
                    df_final1 = combined_df[combined_df[combined_df.columns[0]].apply(lambda x: is_valid_alphanumeric(x) and (not (str(x).isdigit() and len(str(x)) < 5)))]
                #print(df_final1.columns)
                except:
                    print('file issue3')
                for i in df_final1.values:
                    lan_no = i[0]
                    #print('lan no2',lan_no,file)

                    #file_path = os.path.join(folder_path, file)
                    file_name = file_path.split('/')[-1].split('-')


                    # logic of location and state
                    location = file_name[-1].strip().split('.')[0] # location state
                    final_location = location.title()
                    #final_location = location.title()
                    #print(final_location)
                    Client_name = file_path.split('/')[-1].split('-')[-2] # client
                    final_client_name = Client_name.title()
                    try:
                        demo = Client_name + '-' + location + '.xls'
                    except:
                        demo = Client_name + '-' + location + '.xlsx'

                    company_name = 'Ediig'# company name

                    # Logic of month
                    current_date = datetime.now()
                    formatted_date = current_date.strftime("%d-%b-%Y") # month


                    # logic of make and product
                    try:
                        category = file_path.split('/')[-1].split("ET")[1].strip()
                    except:
                        pass
                    result_2 = category.replace(demo, "").strip()
                    make = result_2[1:-1] if result_2.startswith("-") and result_2.endswith("-") else result_2.lstrip("-").rstrip("-") # make and product

                    auction_type = "Online"  #auction type

                    # Logic of auction date
                    try:
                        auction_date = file_name[1] +'-' + file_name[2] + '-' + str((datetime.now().year)% 100)
                    except:
                        auction_date = file_name[1] + '-' + str((datetime.now().year)% 100)
                    data.append({'Make':make,'Product':make,'Lan No':lan_no,'Client':final_client_name,'Comp Name':'Ediig','Location':final_location,'State':final_location,'Region':'nan','Auction Date': auction_date,'Auction Type':'Online','Verticle':'nan','Month':formatted_date})
        except:
            print('issue in keyword length code ')
        final_df = pd.DataFrame(data)
except:
    print('file issue last except')
    
    
#final_df = final_df.dropna(subset=['Lan No'])
try:
    final_df['Client']=final_df['Client'].str.strip()
    final_df['Make'] = final_df['Make'].str.upper()
    final_df = final_df[final_df['Make'] != "SLV"].reset_index(drop = True)



    #Region
    final_df["Make"] = final_df["Make"].apply(lambda x: x.split('-')[0] if '-' in x else x)

    merged_df = pd.merge(final_df,df_state,how = "left", left_on = "State",right_on = "State_")
    final_df['Region'] = merged_df['Region_'].combine_first(final_df['Region'])



    # Vertical code
    merged_df1 = pd.merge(final_df,sheet1, how="left", left_on="Client", right_on="Ediig")
    final_df['Verticle'] = merged_df1['FinalVertical'].combine_first(final_df['Verticle'])
except:
    print("issue in verticle and region")


folder_path = 'D://Automation//OutputM//'  # Replace with the desired parent directory path
folder_name = 'ediig'  # Replace with the desired folder name

# Create the folder
full_folder_path = os.path.join(folder_path, folder_name)
if not os.path.exists(full_folder_path):
    os.makedirs(full_folder_path)
final_df.to_csv(os.path.join(full_folder_path, 'masterfile_ediig.csv'), index=False)
print('Done')
