#Takes the snapshot from the excel file and save it as image

import sys
from pathlib import Path
import win32com.client as win32
from PIL import ImageGrab

excel_path = r'C:\Users\argautam\General\budget.xlsx'
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
workbook = excel.Workbooks.Open(excel_path)
worksheets = workbook.Worksheets(1)

win32c = win32.constants
worksheets.Range("B4:C7").CopyPicture(Format=win32c.xlBitmap)
img = ImageGrab.grabclipboard()
image_path = r'C:\Users\argautam\General\budget.png'
img.save(image_path)
