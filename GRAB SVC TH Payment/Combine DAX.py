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

data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\GRAB SVC TH Payment\\DAX\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='TH_ECA_DAX_Daily_Repayment'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
df_combine2 = df_combine[['report_date', 'debt_id', 'last_payment', 'xm_debtor_id']]
df_combine3 = df_combine2.assign(debt_id_text = "=B2",
                                DAX_ID = "=LEFT(D2,7)",
                                Account_Number = "=IFNA(XLOOKUP(F2,invoice!A:A,invoice!A:A),XLOOKUP(E2,invoice!A:A,invoice!A:A))",
                                Card_Number= "=G2",
                                Description = "",
                                Amount = "=C2*1",
                                Amount_Amount_in_LCY = "=C2*1",
                                Effective_transaction_date = "=DATE(RIGHT(A2,4),MID(A2,4,2),LEFT(A2,2))",
                                Transaction_Date_Posting = "=TODAY()"
                                ) 


# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectiusdwhph.database.windows.net' 
database = 'dwh_th_2022' 
username = 'atiwat' 
password = '2a#$dfERat^%' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
sql_cmd = """
    select
            a.alternis_invoicenumber
    from stage.alternis_account a
    where   alternis_portfolioidname = 'GRAB SVC TH'
    """
df_sql = pd.read_sql(sql_cmd, cnxn)

today = datetime.datetime.today()
print("\n",df_combine2, f"{today:%d-%m-%Y\n}")

writer = pd.ExcelWriter('DAX Combine.xlsx', engine = 'xlsxwriter')
df_combine3.to_excel(writer, index=False, sheet_name= 'Combine DAX')
df_sql.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name= 'invoice')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Combine DAX']
worksheet2 = writer.sheets['invoice']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})

# Set the column width and format.
worksheet.set_column('A:A', 12, format1)
worksheet.set_column('B:B', 12, format1)
worksheet.set_column('C:C', 12, format1)
worksheet.set_column('D:D', 15, format1)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25)
worksheet.set_column('G:G', 25)
worksheet.set_column('H:H', 25)
worksheet.set_column('J:J', 25)
worksheet.set_column('K:K', 25)
worksheet.set_column('L:L', 30)
worksheet.set_column('M:M', 30)

worksheet2.set_column('A:A', 30)

# Close the Pandas Excel writer and output the Excel file.
writer.save()