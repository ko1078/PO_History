from operator import index
import numpy as np3
import warnings
warnings.simplefilter("ignore", UserWarning)
import pandas as pd
import dataclasses
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
from pandas.tseries.offsets import DateOffset
import pyodbc 
import pypyodbc
import win32com.client as win32
import os
from os.path import join
import os.path
import concurrent.futures
from multiprocessing import freeze_support
from pathlib import Path
import time
import shutil
from datetime import date

import glob

PO = []     #blank list of work Orders

n = int(input('Enter Number of PO#\'s: '))  # request count of certs from user.

for i in range(0, n):
    WO = str(input('PO Number - '))
    PO.append(WO) # these blocks collect the certs from the user based on number of certs told

PO_Requested = pd.DataFrame(PO,columns= ['PO'])

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=STAUBCAD\SIGMANEST;'
                      'Database=SNDBase22;'
                      'Trusted_Connection=yes;')

SERVER_NAME = 'STAUBCAD\SIGMANEST'
DATABASE_NAME = 'SNDBase22'

sql_query = "SELECT WoNumber, SheetName, PartFileName  FROM [dbo].[PartArchive] "

part = pd.read_sql(sql_query, conn) # request SQL  Table from STAUBCAD

sql_query1 = "SELECT SheetName,  PrimeCode, Material, Thickness  FROM [dbo].[StockArchive] "

stock = pd.read_sql(sql_query1, conn) # request SQL  Table from STAUBCAD



stock_shortened = stock[stock['PrimeCode'].isin(PO_Requested['PO'])]          # Removes all un requested Work Orders from the parts list.


part_shortened = part[part['SheetName'].isin(stock['SheetName'])]          # removes all Sheets from the stock list that aernt required for the WO Numbers Requested.



merged_inner = pd.merge(left=stock_shortened, right=part_shortened,how='left', left_on='SheetName', right_on='SheetName') # merges the two data frames of the database and the PO Recietps spreadsheet to matching PO_MTL fields.
merged_inner = merged_inner.drop_duplicates(subset= ['PrimeCode','SheetName','PartFileName'])
merged_inner['CustomerName'] = merged_inner['PartFileName'].apply( lambda x : r'\\' + x.split('\\')[-2]  )
merged_inner['Customer'] = merged_inner['CustomerName'].str[2:]
merged_inner['PartFileName'] = merged_inner['PartFileName'].apply( lambda x :  x.split('\\')[-1]  )
merged_inner['Part'] = merged_inner['PartFileName'].apply( lambda x :  x.split('.PRS')[0]  )

merged_inner.to_excel(r'C:\Users\GGehring\Documents\Jobs_on_PO_List.xlsx', columns=['PrimeCode','Material','Thickness','WoNumber','Customer','Part','SheetName'],index = False)
