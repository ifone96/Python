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

data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\"

df = []
for file in os.listdir(data_file_folder):
    if file.endswith('.xlsb'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), engine='pyxlsb',sheet_name='Sheet1'))
    if file.endswith('.xlsx'):
        print('Loading file {0}...'.format(file))
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='Sheet1'))
        
# Len(df)
df_combine = pd.concat(df, axis=0)
#df_combine2 = df_combine.iloc[:,[0,1,7,11]]
#Pick up column with headernamer

df_combine2 = df_combine[['bill_id']].astype('string')
df_combine3 = df_combine2.assign(debt_id_text = "=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!B:B)",
                                phone = "=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!C:C)",
                                type = "=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!D:D)",
                                name = "=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!E:E)",
                                idnumber = "=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!F:F)") 

# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectiusdwhph.database.windows.net' 
database = 'dwh_th_2022' 
username = 'atiwat' 
password = '2a#$dfERat^%' 
cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
sql_cmd = """
    SELECT DISTINCT
    b.alternis_invoicenumber,
    b.alternis_accountid as uuid,
    REPLACE(a.alternis_number,'*','') AS phone,
    a.alternis_phonetypename as type,
    a.alternis_contactidname as name,
    b.alternis_idnumber as idnumber
    FROM stage.alternis_phone a
    JOIN stage.alternis_account b
    ON a.alternis_contactid = b.alternis_contactid
    WHERE alternis_portfolioidname IN ('SEAMONEY SPL SVC TH','SEAMONEY BCL SVC TH')
    ORDER BY a.alternis_contactidname
    """
df_sql = pd.read_sql(sql_cmd, cnxn)

#Set name file with date/times
todays_date_name = str(datetime.datetime.now().strftime("SMN %H%M") )+ '.xlsx'
writer = pd.ExcelWriter(todays_date_name)

df_combine3.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name= 'SMN')
df_sql.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name= 'SQL MAP')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['SMN']
worksheet2 = writer.sheets['SQL MAP']

# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0000000000'})


# Set the column width and format.
worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:C', 16 ,format2)
worksheet.set_column('D:D', 10)
worksheet.set_column('E:E', 16)
worksheet.set_column('F:F', 22)
worksheet2.set_column('A:A', 25, format1)
worksheet2.set_column('B:B', 40, format1)
worksheet2.set_column('C:C', 13, format2)
worksheet2.set_column('D:D', 22, format1)
worksheet2.set_column('E:E', 20, format1)
worksheet2.set_column('F:F', 22, format1)


# Close the Pandas Excel writer and output the Excel file.
writer.save()

