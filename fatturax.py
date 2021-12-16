
import glob,time
import os
import subprocess
import re
from shutil import copy2
from re import sub
from decimal import Decimal
import win32com.client
from datetime import datetime



excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
dirpath = os.path.abspath(os.path.dirname(__file__))
wb2 = excel.Workbooks.Open(dirpath+"\\fattura.xlsx")

data_di_oggi = datetime.today().strftime('%d/%m/%Y')

#ws2 = wb.Worksheets('Sheet')
ws2 = wb2.ActiveSheet

ws2.Range('G5').Value = 23
ws2.Range('H5').Value = data_di_oggi

#wb2.Save()
#excel.Application.Quit()

