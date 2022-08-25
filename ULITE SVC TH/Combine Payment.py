from ast import If
import datetime
import glob
import os
import shutil
from tkinter import HIDDEN
import uuid
from doctest import DocFileTest
from email.utils import format_datetime
from math import fabs
from operator import index
from pickle import NONE
import pandas as pd
import pyodbc
import xlsxwriter
from matplotlib.pyplot import axis

data_file_folder = "Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH\\Payment\\"
df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='payment_reconcile_report'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
df_combine = df_combine[11:]
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
#df_combine = df_combine[['SO No.','วันที่ชำระ','ผู้ซื้อ','ยอดรับชำระรวม(บาท)']]


todaysdate_filename = str(
    datetime.datetime.now().strftime("Combine Ulite %H%M %Y%m%d")) + '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)
print("\n",df_combine, f"{todaysdate_filename }""\n")

df_combine.to_excel(writer, index=False, sheet_name= 'Combine_Ulite')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Combine_Ulite']


# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0.00'})
format3 = workbook.add_format({'num_format': 'mm/dd/yyyy'})

# Set the column width and format.
worksheet.set_column('A:A', 12)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 15)
worksheet.set_column('D:D', 15)
worksheet.set_column('E:E', 15)
worksheet.set_column('F:F', 10)
worksheet.set_column('G:G', 18)
worksheet.set_column('H:H', 18)
worksheet.set_column('I:I', 14)
worksheet.set_column('J:J', 12)
worksheet.set_column('K:K', 26)
worksheet.set_column('L:L', 28)
worksheet.set_column('M:M', 26)

#Formula 
#worksheet.write_dynamic_array_formula('E2:E500', '=B2:B500&""')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Open file or folder on OS
path_url = "Z:\\MIS\\Fone Wasin\\Python\\ULITE SVC TH\\"
path_file = path_url + "\*.xls*"
for filex in glob.iglob(path_file, recursive=True):
    os.path.realpath(path_url)
    # FBI OPEN UP!!!!
    #os.startfile(path_url)
    os.startfile(filex)
    print('Opened File&Folder:', filex)
