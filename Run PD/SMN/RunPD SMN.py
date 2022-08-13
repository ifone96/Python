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

data_file_folder = "C:\\Users\\wasin.k\\Desktop\\Python\\Run PD\\SMN\\From TL"

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
df_combine = df_combine[['bill_id']].astype('string') #=XLOOKUP($A2,'SQL MAP'!$A:$A,'SQL MAP'!B:B)
df_combine = df_combine.assign(uuid = "", 
                    phone = "",
                    type = "",
                    name = "",
                    idnumber = "") 

# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = 'collectiusdwhph.database.windows.net' 
database = 'dwh_th_2022' 
username = 'atiwat' 
password = '2a#$dfERat^%' 
connect_database = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
sql_cmd = """
    SELECT DISTINCT
    b.alternis_invoicenumber as bill_id, 
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
df_sql = pd.read_sql(sql_cmd, connect_database)

#f(x) xlookup python pandas
def xlookup(lookup_value, lookup_array, return_array, if_not_found:str = ''):
    match_value = return_array.loc[lookup_array == lookup_value]
    if match_value.empty:
        #return f'"{lookup_value}" is NULL' if if_not_found == '' else if_not_found
        return f'NULL' if if_not_found == '' else if_not_found

    else:
        return match_value.tolist()[0]

#xlookup_data = xlookup('1625976231738529792', df_sql['bill_id'],df_sql['uuid'])
df_combine['uuid'] = df_combine['bill_id'].apply(xlookup, args = (df_sql['bill_id'], df_sql['uuid']))
df_combine['phone'] = df_combine['bill_id'].apply(xlookup, args = (df_sql['bill_id'], df_sql['phone']))
df_combine['type'] = df_combine['bill_id'].apply(xlookup, args = (df_sql['bill_id'], df_sql['type']))
df_combine['name'] = df_combine['bill_id'].apply(xlookup, args = (df_sql['bill_id'], df_sql['name']))
df_combine['idnumber'] = df_combine['bill_id'].apply(xlookup, args = (df_sql['bill_id'], df_sql['idnumber']))

#delete some u don't need
del df_combine['bill_id']

#Join table
#join_data = pd.merge(df_combine, df_sql, on ='bill_id', how ='outer')
#join_data.drop('bill_id', inplace=True, axis=1)

#Set name file with date/times
todaysdate_filename = str(datetime.datetime.now().strftime("SMN %H%M") )+ '.xlsx'
writer = pd.ExcelWriter(todaysdate_filename)

df_combine.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name='Output')
#join_data.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name='test')
df_sql.to_excel(writer, index=False, engine='xlsxwriter' ,sheet_name='SQL MAP')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Output']
worksheet2 = writer.sheets['SQL MAP']


# Add some cell formats.
format1 = workbook.add_format({'num_format': '@'})
format2 = workbook.add_format({'num_format': '0000000000'})
header_format = workbook.add_format({'bold': True})

# Set the column width and format.
worksheet.set_row(0, None, header_format)
worksheet.set_column('A:A', 40)
worksheet.set_column('B:B', 16)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 30)
worksheet.set_column('E:E', 25)
worksheet.set_column('F:F', 25)

worksheet2.set_row(0, None, header_format)
worksheet2.set_column('A:A', 25)
worksheet2.set_column('B:B', 40)
worksheet2.set_column('C:C', 16)
worksheet2.set_column('D:D', 20)
worksheet2.set_column('E:E', 30)
worksheet2.set_column('F:F', 25)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

