import os
import pythoncom
from tkinter import filedialog
import win32com.client as win32


def xls_to_xlsx(filename):
    pythoncom.CoInitialize()
    if filename.endswith('.xls'):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(filename)
        wb.SaveAs(filename + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()


default_dir = os.getcwd()
file_selection = filedialog.askopenfilenames(initialdir=default_dir)

if len(file_selection) == 1:
    file_name = file_selection[0].replace('/', '\\')
    xls_to_xlsx(file_name)

else:
    for i in range(len(file_selection)):
        xls_to_xlsx(file_selection[i].replace('/', '\\'))

print('Done')
