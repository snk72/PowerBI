import win32com.client
xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
wb = xl.Workbooks.Open(r'C:\Users\sindre\OneDrive - BRAbank ASA\Documents - BRAbank Likviditet\macro.xlsm')
xl.Application.Run('macro.xlsm!Module1.macro2')
wb.Save()
xl.Application.Quit()