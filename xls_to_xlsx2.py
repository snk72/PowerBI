import win32com.client as win32
import os

import shutil

directory_path = r"C:\Users\sindre\AppData\Local\Temp\gen_py\3.10"

def delete_directory_contents(directory):
    for root, dirs, files in os.walk(directory, topdown=False):
        for file in files:
            file_path = os.path.join(root, file)
            os.remove(file_path)
        for dir in dirs:
            dir_path = os.path.join(root, dir)
            shutil.rmtree(dir_path)

# Call the function to delete the contents of the directory
delete_directory_contents(directory_path)


xlsx_file_path = r"C:\Users\sindre\OneDrive - BRAbank ASA\Documents - BRAbank Likviditet\General\Collateral\XLSX\LEA_BANKFX.xlsx"

# check if the xlsx file already exists and delete it if it does
if os.path.exists(xlsx_file_path):
    os.remove(xlsx_file_path)


fname = r"C:\Users\sindre\OneDrive - BRAbank ASA\Documents - BRAbank Likviditet\General\Collateral\LEA_BANKFX.xls"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(xlsx_file_path, FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

# AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap' -->  C:\Users\sindre\AppData\Local\Temp\gen_py\3.10