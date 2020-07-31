import win32com.client
import os

file_dir = "\\CorruptedFiles"
for filename in os.listdir(file_dir):
    print(filename)
    file= os.path.splitext(filename)[0]
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    wb = o.Workbooks.Open
        (file_dir  + filename) 
    wb.ActiveSheet.SaveAs
        (file_dir  + file + ".xlsx", 51) 
    o.Application.Quit()