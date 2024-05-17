import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox, simpledialog,scrolledtext,font
from datetime import datetime,timedelta
from Global_DCNOrdering_function import *
##Initialize Tkinter
root = tk.Tk()
root.title("Function Status")
root.geometry("600x300")  # Adjusted the initial size of the window
root.configure(bg="#000000")
# Create an instance of StatusUpdater
status_updater = StatusUpdater(root)
updater = status_updater
#-----------------------------------------------------------------------------------------------

user_id = simpledialog.askstring(title='credentials', prompt='Please enter user_id')
Password_EDB = simpledialog.askstring(title='credentials',show='*', prompt='Please enter EDB password')
Input_week = simpledialog.askstring(title='Week', prompt='Please enter Intro Week(YYYYWW-eg-202430)')
print("Input_week",Input_week)
Refresh="Please Refresh the Master_Data(DCNinJobQ) and then Click OK to proceed further "
messagebox.showinfo("Custom Message Box", Refresh)
#---------------------------------------------------------------
 
# Input_week = simpledialog.askstring(title='Week', prompt='Please enter Week')
year, weeknum = Input_week[:4], int(Input_week[4:])
print(f"year:{year}, weeknum:{weeknum}")
input_date = datetime.strptime(f"{year}-{weeknum}-1", "%Y-%W-%w")
print("input date",input_date)
before_date = input_date - timedelta(weeks=0)
print("before_date",before_date)
after_date = input_date + timedelta(weeks=15)
print("after_date",after_date)
current = before_date
print("current",current)
while current <= after_date:
    current_year, current_week = current.strftime("%Y-%W").split('-')
    print(f"current year in while loop:{current_year}, current week in while loop:{current_week}")  
    # Check if the week is in the next year
    if int(current_week) < int(weeknum):
        # current_year = str(int(current_year) + 1)
        print("if cond current year",current_year)
        pass
    # Special handling for the transition from the last week to the first week of the year
    if current_week == '53':
        current_week = '01'
        current_year = str(int(current_year) + 1)
        print("sec if cond current year",current_year)
    # Special handling for the transition from the first week to the last week of the year
    elif current_week == '51' and int(weeknum) == 1:
       
        # current_week = '52'
        # current_year = str(int(current_year) - 1)  
        print("elif cond current year",current_year)
        pass
 
    cw=f"{current_year}{current_week}"
    print("cw",cw)
    updater.update_status("Current_week", cw)
    driver=login_EDB(user_id, Password_EDB)#--1
    DCNDisTime(driver, cw, updater,EDB123_path)#--2
    current += timedelta(weeks=1)

# # --------------------------------------------------------------------------------------------------
EDB123_path = (r"C:\Users\a323151\Desktop\Alten_Automation\Global_DCN_Ordering\Query_Files\EDB123.xlsx")
if os.path.exists(EDB123_path):
    workbook = openpyxl.load_workbook(EDB123_path)
    worksheet = workbook.active
    for cell in worksheet['I'][1:]:
        if cell.value is not None:
            formula_values = cell.value[4:-2]
            cell.value = formula_values
        else:
            print("novalue in EDB sheet to convert the formula format")
    workbook.save(EDB123_path)
# ---------------------------------------------------------------------------------------------------
DCNfrmkola_path=DCNListDis(updater) #--1
Import(updater) ##--2
DCNfrmkola_path=UpdateK_DCN(updater)##--4
DCN_Sign_OFF_TA(updater) ##--5
DCN_LinkID(updater) ##--6
files = glob.glob("C:/Users/a323151/Desktop/Alten_Automation/Global_DCN_Ordering/Temp_downloads/*.csv")
for file in files:
    os.remove(file)
driver=login_EDB(user_id, Password_EDB) #--7
Generate_LinkIdList(driver,updater) #--8
DCN_SignOffLinkID(updater)#--8
KDCN_toOrder(updater)##--9
KDCNListQJAM_DDCNListQJAM(updater)#--10
DDCNList(updater)#--11
DCNNotSignedByPPL(updater)#--13
