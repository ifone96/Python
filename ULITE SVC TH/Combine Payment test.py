from ast import If
import datetime
from email import header
import glob
import os
import shutil
from tkinter import HIDDEN
from unittest import skip
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
        df.append(pd.read_excel(os.path.join(data_file_folder, file), sheet_name='payment_reconcile_report', header=11))
        
# Len(df)
df_combine = pd.concat(df, axis=0)

reCol = {
    'ผู้ซื้อ': 'Account Number',
    'SO No.': 'Invoice/Card Number',
    'ยอดรับชำระรวม(บาท)' : 'Amount',
    'วันที่ชำระ': 'Effective Date' 
}
# call rename () method
df_combine.rename(columns=reCol, inplace=True)
df_combine = df_combine[['Account Number','Invoice/Card Number','Amount','Effective Date']]
df_combine = df_combine.assign(**{ 
                                'DD': '',
                                'MM': '',
                                'Account Number+': '',	
                                'Card Number+':'',	
                                'Description+':'',	 
                                'Amount+':'', 	 
                                'Amount Amount in LCY+':'', 
                                'Effective Transaction Date+': '',
                                'Transaction Date Posting+': '=TODAY()',
                                'Payment Channel+': '',
                                'Product Type+': '',
                                'Statement Reference+': ''
                                })

df_combine.dropna(inplace=True)
# add column
#df_combine.insert(2, "Description", "",True)
#df_combine.insert(4, "Amount In IYC", "",True)
df_combine