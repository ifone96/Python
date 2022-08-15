from doctest import DocFileTest
from email.utils import format_datetime
from operator import index
import os
from pickle import NONE
from matplotlib.pyplot import axis
import pandas as pd
import datetime
import xlsxwriter
import uuid
import pyodbc 

from doctest import DocFileTest
from email.utils import format_datetime
from operator import index
import os
import pandas as pd
import datetime
import xlsxwriter

data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\GRAB SVC TH Payment\\DAX\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='TH_ECA_DAX_Daily_Repayment'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
df_combine2 = df_combine[['report_date', 'debt_id', 'debtor_id', 'last_payment']]
today = datetime.datetime.today()
print("\n",df_combine2, f"{today:%d-%m-%Y\n}")

writer = pd.ExcelWriter('DAX Combine.xlsx', engine = 'xlsxwriter')
df_combine2.to_excel(writer, index=False, sheet_name= 'Combine DAX')
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Combine DAX']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})

# Set the column width and format.
worksheet.set_column('A:A', 12, format1)
worksheet.set_column('B:B', 12, format1)
worksheet.set_column('C:C', 12, format1)
worksheet.set_column('D:D', 15, format1)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
