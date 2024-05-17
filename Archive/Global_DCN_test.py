# import datetime
# import glob
# import pandas as pd
# import os
# import csv
# import re
# import time
# import urllib.request
# from collections import Counter
# from tkinter import messagebox, simpledialog,scrolledtext
# import tkinter as tk
# import sys
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.service import Service
# import openpyxl
# import tkinter.font as tkFont
# from openpyxl import Workbook
# from PyPDF2 import PdfReader
# from pandas.core.common import flatten
# from openpyxl.styles import PatternFill
# from openpyxl.styles import Font
# from selenium import webdriver
# from PyQt5.QtWidgets import QApplication, QMessageBox
# import win32com.client as win32
# from selenium.common.exceptions import (NoSuchElementException,
#                                         NoSuchFrameException,
#                                         WebDriverException,TimeoutException)
# from selenium.webdriver.chrome.service import Service as ChromeService
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.select import Select
# from selenium.webdriver.support.ui import WebDriverWait
# from webdriver_manager.chrome import ChromeDriverManager

# ###Creating format for date_time
# TS = datetime.datetime.now().strftime("%d, %m, %Y, %H, %M, %S") 
# TS = TS.split(', ')
# TS_date = TS[:3]
# TS_date = '-'.join(TS_date)
# TS_time = TS[3:]
# TS_time = '-'.join(TS_time)
# date_time = '{} {}'.format(TS_date,TS_time)

# user = os.getlogin() #System login user id
# download_path = r'C:\Python_SPI\Global_DCN_Test\downloads'
# files_to_delete = [f for f in os.listdir(download_path)]#deleting all existing  files in the downloads folder
# for file in files_to_delete:
#     os.remove(os.path.join(download_path, file))
# folder_path = r'C:\Python_SPI\Global_DCN_Test'

# def DCN_Sign_OFF_TA():
#     # Read Excel file into a DataFrame
#     DCNfrmkola_path = r'C:\Python_SPI\Global_DCN_Test\DCNfrmKola.xlsx'
#     DCNinJobQ_path = r'C:\Python_SPI\Global_DCN_Test\DCNinJobQ.xlsx'
    
#     df_kola = pd.read_excel(DCNfrmkola_path, sheet_name='DCNfrmKola')
#     df_jobq = pd.read_excel(DCNinJobQ_path, sheet_name='DCNinJobQ')

#     # Perform Inner Join
#     merged_df = pd.merge(df_kola, df_jobq, left_on='DCNG', right_on='DCN Design Change Notice')

#     # Apply the SQL-like conditions
#     query_result = merged_df[
#         (merged_df['Type'] == 'Technical Authorisation') &
#         (merged_df['DCNinJobQ'] == 'Yes') &
#         (merged_df['PPLNotSigned'].isnull())
#     ]

#     # Select specific columns
#     result_columns = [
#         'Product Class', 'Object', 'DCNG', 'Heading', 'Type', 'DIS Time',
#         'DCNinJobQ', 'PPLNotSigned', 'DCN PSU ID', 'DCN Job Role'
#     ]

#     final_result = query_result[result_columns]
#     # current_date = datetime.now().strftime('%Y%m%d')
#     DCNSignOFFTA_path = f'C:\Python_SPI\Global_DCN_Test\DCNSignOFFTA_{date_time}.xlsx'
#     final_result.to_excel(DCNSignOFFTA_path, index=False)
#     print("DCNSignOFFTA Completed")
#     # Remove filtered rows from original DCNfrmKola DataFrame
#     df_kola = df_kola[~df_kola['DCNG'].isin(query_result['DCNG'])]

#     # Save the modified DataFrame to the original Excel file
#     with pd.ExcelWriter(DCNfrmkola_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#         df_kola.to_excel(writer, sheet_name='DCNfrmKola', index=False)
#     print("DCNfrmKola updated")
    
# ############################################################################################################


# def DCN_LinkId():
#     # Read Excel files into DataFrames
#     DCNfrmkola_path = r'C:\Python_SPI\Global_DCN_Test\DCNfrmKola.xlsx'
#     AppPC_path = r'C:\Python_SPI\Global_DCN_Test\AppPC.xlsx'
    
#     df_kola = pd.read_excel(DCNfrmkola_path, sheet_name='DCNfrmKola')
#     df_app_pc = pd.read_excel(AppPC_path, sheet_name='AppPC')

#     # Perform Inner Join
#     merged_df = pd.merge(df_kola, df_app_pc, left_on='Product Class', right_on='PC')

#     # Apply the SQL-like conditions
#     query_result = merged_df[
#         (merged_df['DIS Time'].notnull()) &
#         (merged_df['Type'] == 'Product Structure') &
#         (merged_df['DCNOrdered'].isnull()) &
#         (merged_df['DCNinJobQ'] == 'Yes') &
#         (merged_df['DCNBySCP'].isnull()) &
#         (merged_df['PPLNotSigned'].isnull()) &
#         (merged_df['QJ'].isin(['QJ', 'Q-']) == False) &
#         (merged_df['AMObjects'].isnull())
#     ]

#     # Select specific columns
#     result_columns = [
#         'Product Class', 'Object', 'DCNG', 'Heading', 'DIS Time', 'Type'
#     ]

#      # Select specific columns and keep only distinct 'Product Class'
#     final_result = query_result[result_columns].sort_values(result_columns).groupby('Product Class').apply(lambda group: group).drop_duplicates(subset=result_columns)
    
#     # Save the result to Excel with current date
#     # current_date = pd.Timestamp.now().strftime('%Y%m%d')
#     DCNListLIndID_path = f'C:\Python_SPI\Global_DCN_Test\DCNListLIndID {date_time} .xlsx'
#     final_result.to_excel(DCNListLIndID_path , index=False)

#     print(f'DCNListLIndID Completed. Results saved to: {DCNListLIndID_path }')

# ##############################################################################################################



# def login_EDB(user_id,password):
#     #########webdriver setup################
#     options = webdriver.ChromeOptions()
#     prefs = {"download.default_directory" : download_path}
#     options.add_experimental_option("prefs",prefs)
#     options.add_argument("--headless")
#     service = Service(executable_path="C:\Python_SPI\chromedriver.exe")
#     driver = webdriver.Chrome(service=service, options=options)
#     #############EDB Login ###########################
#     driver.get("http://edb.volvo.net/edb2/index.htm")
#     driver.switch_to.frame("banner")
#     driver.find_element(By.ID,"action").click()
#     driver.forward()
#     cred1 = driver.find_element(By.NAME, "username")
#     cred1.clear()
#     cred1.send_keys(user_id)
#     cred2 = driver.find_element(By.NAME, "password")
#     cred2.clear()
#     cred2.send_keys(password)
#     driver.find_element(By.CLASS_NAME, "button").click()
#     driver.forward()
#     driver.maximize_window()
#     driver.switch_to.frame("menu")
#     # updater.update_status("Login","Logged in successfully")
#     return(driver)

# def generate_LinkIdList(driver):
#     driver.find_element(By.PARTIAL_LINK_TEXT, "KOLA+").click()
#     driver.find_element(By.PARTIAL_LINK_TEXT,'Partno Data').click()
#     driver.find_element(By.PARTIAL_LINK_TEXT,'KOLA Partno Info').click()
#     driver.switch_to.parent_frame()
#     driver.switch_to.frame("edb_main")
#     temp=Select(driver.find_element(By.NAME,'func'))
#     temp.select_by_value("135")
#     textarea=WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//textarea[@name='art']")))
#     textarea.click()
#     folder_path = r'C:\Python_SPI\Global_DCN_Test'
#     latest_file = max([f for f in os.listdir(folder_path) if f.startswith('DCNListLIndID')], key=os.path.getctime)
#     DCNListLIndID_path = os.path.join(folder_path, latest_file)
#     df = pd.read_excel(DCNListLIndID_path)
#     #Get DCNG column values
#     dcng_values = df['DCNG']
#     textarea.click()
#     for i in dcng_values:
#         textarea.send_keys(str(i) + '\n')
        
#     driver.find_element(By.XPATH,"//input[@name='Fetch']").click()
#     try:
#         WebDriverWait(driver,2500).until(EC.presence_of_element_located((By.LINK_TEXT,"DOWNLOAD EXCEL FILE"))).click()
#     except TimeoutException:
#         print("Timeout: The 'DOWNLOAD EXCEL FILE' link did not appear within the specified time.")
#         # driver.find_element(By.LINK_TEXT,'DOWNLOAD EXCEL FILE').click()
#         time.sleep(10)
#     latest_csv = max((f for f in os.listdir(folder_path) if f.endswith('.csv')), key=os.path.getctime)
#     latest_csv_path = os.path.join(folder_path, latest_csv)





# # ################################################################################################################
        

# def DCN_SignOffLinkID():
#     ##First Create DCNtoOrder excel
#     DCNfrmkola_path = r'C:\Python_SPI\Global_DCN_Test\DCNfrmKola.xlsx'
#     DCNLinkID_path=r'C:\Users\A478237\Downloads\DCNLinkID.xlsx'
#     df_kola = pd.read_excel(DCNfrmkola_path, sheet_name='DCNfrmKola')
#     df_DCNLinkID=pd.read_excel(DCNLinkID_path)
#     # Merge DataFrames based on the DCN column
#     merged_df = pd.merge(df_DCNLinkID, df_kola, left_on='DCN', right_on='DCN')
#     # Filter rows based on the Parttype column
#     merged_df.columns = merged_df.columns.map(lambda x: x.rstrip('_x'))
#     filtered_df = merged_df[merged_df['Parttype'].isin(['K', 'D', 'S', 'P'])]
#     filtered_df.to_excel((f'C:\Python_SPI\Global_DCN_Test\oilter_df{date_time}.xlsx'),index=False)

#     # Select specific columns
#     result_columns = ['Product Class', 'DCNG', 'Object', 'Heading', 'KOLA Time', 'DIS Time']
#     # final_result = filtered_df[result_columns].sort_values(by=result_columns).drop_duplicates()
#     final_result = filtered_df[result_columns].sort_values(result_columns).groupby('Product Class').apply(lambda group: group).drop_duplicates(subset=result_columns)
#     DCNtoOrder_path= f'C:\Python_SPI\Global_DCN_Test\DCNtoOrder {date_time}.xlsx'
#     final_result.to_excel(DCNtoOrder_path, index=False)
#     print("DCNtoOrder created")
#         ###########################STARTING SIGNOFFLINKID########################
#     ##Assuming you have the dataframes loaded from your Excel files
#     ##Replace the file paths and sheet names accordingly
#     folder_path = r'C:\Python_SPI\Global_DCN_Test'
#     latest_file = max([f for f in os.listdir(folder_path) if f.startswith('DCNtoOrder')], key=os.path.getctime)
#     DCNtoOrder_path = os.path.join(folder_path, latest_file)
#     DCNtoOrder_df = pd.read_excel(DCNtoOrder_path)
#     DCNtoOrder_df = DCNtoOrder_df.rename(columns={'DCNG': 'DCNG_y'})
#     DCNtoOrder_df = DCNtoOrder_df.rename(columns={'Heading': 'Heading_y'})


#     latest_file = max([f for f in os.listdir(folder_path) if f.startswith('DCNListLIndID')], key=os.path.getctime)
#     DCNListLIndID_path = os.path.join(folder_path, latest_file)
#     DCNListLIndId_df = pd.read_excel(DCNListLIndID_path)
#     # DCNListLIndId_df = DCNListLIndId_df.rename(columns={'Heading': 'Heading_x'})
#     DCNinJobQ_path = r'C:\Python_SPI\Global_DCN_Test\DCNinJobQ.xlsx'
#     DCNinJobQ_df=pd.read_excel(DCNinJobQ_path,sheet_name='DCNinJobQ')
    
#     # Perform the SQL-like query
#     temp_leftjoin_df = DCNListLIndId_df.merge(DCNtoOrder_df, left_on='DCNG',right_on='DCNG_y',how='left')
#     # print("temp_leftjoin_df:", temp_leftjoin_df.columns)
#     # print(temp_leftjoin_df)
#     merged_df = temp_leftjoin_df.merge(DCNinJobQ_df, left_on='DCNG', right_on='DCN Design Change Notice', how='inner') 
#     # print("merged_df:", merged_df.columns)
#     # temp_leftjoin_df.to_excel((f'C:\Python_SPI\Global_DCN_Test\emp_leftjoin_df{date_time}.xlsx'),index=False)
#     # print(merged_df)
#     filtered_df = merged_df[merged_df['DCNG_y'].isna()]
#     # merged_df.columns = merged_df.columns.map(lambda x: x.rstrip('_x'))
#     # merged_df.columns = merged_df.columns.map(lambda x: x.rstrip('_y'))
#     # print(merged_df)
#     # print("filtered_df:", filtered_df.columns)
#     # filtered_df = filtered_df.rename(columns={'DCNG_y': 'DCNG_y'})
#     # print(filtered_df)
#     # print("print(filtered_df['DCNG_y']",filtered_df['DCNG_y'])
#     # Select specific columns
#     result_columns = ['DCNG', 'Heading', 'DCN PSU ID', 'DCN Job Role', 'Product Class_x', 'Object_x', 'DIS Time_x']
#     # Get unique values based on the specified columns
#     unique_result_df = filtered_df[result_columns].sort_values(result_columns).groupby('DCNG').apply(lambda group: group).drop_duplicates(subset=result_columns)
#     print("unique_result_df",unique_result_df)
#     # Save the result to Excel with current date
#     # current_date = pd.Timestamp.now().strftime('%Y%m%d')
#     DCN_SignOffLinkID_path = f'C:\Python_SPI\Global_DCN_Test\DCN_SignOffLinkID {date_time} .xlsx'
#     unique_result_df.to_excel(DCN_SignOffLinkID_path , index=False)

#     print(f'DCN_SignOffLinkID_path Completed. Results saved to: {DCN_SignOffLinkID_path }')

# ##########################################################################################################


# def DDCNList():
#     # Assuming you have DataFrames for DDCNListingS1, ObjPrfToExcDDCN, and AMObjects
#     ##first create DCNListingS1 using DCNinJobQ and Applicable PSU table
#     DCNinJobQ_path = r'C:\Python_SPI\Global_DCN_Test\DCNinJobQ.xlsx'
#     Applicable_PSU_path= r'C:\Python_SPI\Global_DCN_Test\ApplicablePSU.xlsx'
#     Applicable_PSU_df=pd.read_excel(Applicable_PSU_path,sheet_name='ApplicablePSU')
#     DCNinJobQ_df=pd.read_excel(DCNinJobQ_path,sheet_name='DCNinJobQ')
#     # print(DCNinJobQ_df)
#     DCNinJobQ_df['DCN PSU ID'] = pd.to_numeric(DCNinJobQ_df['DCN PSU ID'], errors='coerce')
#     Applicable_PSU_df['PSU'] = pd.to_numeric(Applicable_PSU_df['PSU'], errors='coerce')
#     merged_df = DCNinJobQ_df.merge(Applicable_PSU_df, left_on='DCN PSU ID', right_on='PSU', how='inner')
#     # print(merged_df)
#      # Step 3: Apply WHERE conditions
#     filtered_df = merged_df[
#         (merged_df['DCN Archive Date'] != "1901-01-01") &
#         (merged_df['DCN Design Change Notice 1-1'] == "D") &
#         ((merged_df['DCN Object Number'].str[:2] != "QJ") & (merged_df['DCN Object Number'].str[:2] != "Q-"))
#     ]
#     # print(filtered_df)
#     result_df = pd.DataFrame({
#     'DCN Design Change Notice': filtered_df['DCN Design Change Notice'],
#     'DCN Object Number': filtered_df['DCN Object Number'],
#     'DCN PSU ID': filtered_df['DCN PSU ID'],
#     'DCN Archive Date': filtered_df['DCN Archive Date'],
#     'Site': filtered_df['Site'],
#     'Prefix': [x[:3] for x in filtered_df['DCN Object Number']],  # Prefix
#     'Prefix1': [x[:2] for x in filtered_df['DCN Object Number']],  # Prefix1
#     })

#     # print(result_df)

#     result_columns =['DCN Design Change Notice','DCN Object Number', 'DCN PSU ID', 'DCN Archive Date', 'Site', 'Prefix', 'Prefix1']
#     final_result = result_df[result_columns].sort_values(result_columns).groupby('DCN Design Change Notice').apply(lambda group: group).drop_duplicates(subset=result_columns)
#     final_result['DCN PSU ID'] = final_result['DCN PSU ID'].astype('Int64')  # Convert to nullable integer
#     final_result['DCN PSU ID'] = final_result['DCN PSU ID'].astype(str)  # Convert to string
#     # print(final_result)
#     DCNListingS1_path = f'C:\Python_SPI\Global_DCN_Test\DCNListingS1 {date_time} .xlsx'
#     final_result.to_excel(DCNListingS1_path , index=False)

#     print(f'DCNListingS1 Completed. Results saved to: {DCNListingS1_path}')

#     #########################STARTING DDCNLISTINGS2#####################################
#     ##find lastest file DDCNlistings1
#     latest_file = max([f for f in os.listdir(folder_path) if f.startswith('DCNListingS1')], key=os.path.getctime)
#     DCNListingS1_path = os.path.join(folder_path, latest_file)
#     DCNListingS1_df = pd.read_excel(DCNListingS1_path)
#     AM_Objects_path=r'C:\Python_SPI\Global_DCN_Test\AMObjects.xlsx'
#     AM_Objects_df=pd.read_excel(AM_Objects_path)
#     ObjPrfToExcDDCN_path=r'C:\Python_SPI\Global_DCN_Test\ObjPrfToExcDDCN.xlsx'
#     ObjPrfToExcDDCN_df=pd.read_excel(ObjPrfToExcDDCN_path)
#     ObjPrfToExcDDCN_df = ObjPrfToExcDDCN_df.rename(columns={'Prefix': 'Prefix_y'})

#     # Perform the SQL-like query
#     merged_df = DCNListingS1_df.merge(ObjPrfToExcDDCN_df, left_on='Prefix', right_on='Prefix_y', how='left',suffixes=('', '_right'))
#     print(merged_df.columns)
#     merged_df = pd.merge(merged_df, AM_Objects_df, left_on='DCN Object Number', right_on='AMObject', how='left')
#     # print(merged_df.columns)
#     # Filter rows based on conditions
#     filtered_df = merged_df[(merged_df['Prefix_y'].isna()) & (merged_df['AMObject'].isna())]
#     # print(filtered_df)
#     filtered_df['ArchYear']=filtered_df['DCN Archive Date'].astype(str).str[:4]
#     # Select specific columns
#     result_columns = ['DCN Design Change Notice', 'DCN Object Number', 'DCN PSU ID', 'Site', 'ArchYear']
#     result_df = filtered_df[result_columns]
#     # print(result_df)

#     # Save the result to Excel with current date and time
#     result_path = f'C:\Python_SPI\Global_DCN_Test\DDCNList {date_time}.xlsx'
#     result_df.to_excel(result_path, index=False)

#     print(f'DDCNList created. Results saved to: {result_path}')





    













#     # Wait for some time to let the download link appear


# ##Call the function to execute the code
# # DCN_Sign_OFF_TA()
# # ##Call the function to execute the code
# # DCN_LinkId()
# # driver=login_EDB("A478237","Hopethebest123$")
# # generate_LinkIdList(driver)
# DCN_SignOffLinkID()
# # DDCNList()

from datetime import datetime,timedelta

Input_week = "202349"
year, weeknum = Input_week[:4], int(Input_week[4:])
input_date = datetime.strptime(f"{year}-{weeknum}-1", "%Y-%W-%w")
before_date = input_date - timedelta(weeks=15)
after_date = input_date + timedelta(weeks=15)
current = before_date
print(current)

while current <= after_date:
    current_year, current_week = current.strftime("%Y-%W").split('-')
    
    # Check if the week is in the next year
    if int(current_week) < int(weeknum):
        current_year = str(int(current_year) + 1)

    # Special handling for the transition from the last week to the first week of the year
    if current_week == '00':
        current_week = '01'
        current_year = str(int(current_year) + 1)

    # Special handling for the transition from the first week to the last week of the year
    elif current_week == '51' and int(weeknum) == 1:
        current_week = '52'
        current_year = str(int(current_year) - 1)

    s=f"{current_year}{current_week}"
    print("printing s",s )
    current +=timedelta(weeks=1)



    
