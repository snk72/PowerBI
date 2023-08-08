import win32com.client
xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app
wb = xl.Workbooks.Open(r'C:\Users\sindre\OneDrive - BRAbank ASA\Documents - BRAbank Likviditet\Likviditet MV5.xlsm')
xl.Application.Run('Likviditet MV5.xlsm!mail.email')
wb.Save()
xl.Application.Quit()