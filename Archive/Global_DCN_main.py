import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox, simpledialog,scrolledtext,font
from Global_DCN_main_test import *

##Initialize Tkinter
root = tk.Tk()
root.title("Function Status")
root.geometry("600x300")  # Adjusted the initial size of the window
root.configure(bg="#000000")
# Create an instance of StatusUpdater
status_updater = StatusUpdater(root)
updater = status_updater

#-------------------------------------------------------------- 
###taking user_id and password from the user to login into websites
user_id = simpledialog.askstring(title='credentials', prompt='Please enter user_id')
Password_EDB = simpledialog.askstring(title='credentials',show='*', prompt='Please enter EDB password')
Input_week = simpledialog.askstring(title='Week', prompt='Please enter Week')
before_week =int(Input_week)-3
after_week=int(Input_week)+3
Refresh="Please Refresh the Master_Data(DCNinJobQ) and Click OK to proceed further "
messagebox.showinfo("Custom Message Box", Refresh)

#--------------------------------------------------

year, weeknum = Input_week[:4], int(Input_week[4:])
input_date = datetime.strptime(f"{year}-{weeknum}-1", "%Y-%W-%w")
before_date = input_date - datetime.timedelta(weeks=15)
after_date = input_date + datetime.timedelta(weeks=15)
current = before_date

while current <= after_date:
    current_year, current_week, current_weekday = current.strftime("%Y-%W-%w")   
    # Check if the week is in the next year
    if int(current_week) < int(weeknum):
        current_year = str(int(current_year) + 1)
    updater.update_status("Current_week", current)
    driver = login_EDB(user_id, Password_EDB)
    DCNDisTime(driver, current, updater)
    current += datetime.timedelta(weeks=1)
# current_week=202340
# for i in range(before_week,after_week):
#     updater.update_status("Current_week", i)
#     driver = login_EDB(user_id, Password_EDB) # Repeat the process for 5 iterations as an example
#     DCNDisTime(driver, i, updater)
#     # current_week += 1   
EDB123_path = (r"C:\Python_SPI\Global_DCN_Test\Query_files\EDB123.xlsx")
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
DCNListDis(updater) #--1
Import(updater) ##--2
UpdateK_DCN(updater)##--4
DCN_Sign_OFF_TA(updater) ##--5
DCN_LinkID(updater) ##-6
driver=login_EDB(user_id, Password_EDB) #--7
Generate_LinkIdList(driver,updater) #--8
DCN_SignOffLinkID(updater)#-8
KDCN_toOrder(updater)##--9
KDCNListQJAM_DDCNListQJAM(updater)#----10
DDCNList(updater)#--11
DCNNotSignedByPPL(updater)#--13







# Close the driver when done


root.mainloop()
