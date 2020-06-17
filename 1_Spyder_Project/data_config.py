import os
import sqlite3
import time
import datetime

THIS_MONTH = '2019/10/1'
LAST_YEAR = '2018/10/1' 
LAST_MONTH = '2019/9/1'

CONN = sqlite3.Connection(r'./db.sqlite')
CURSOR = CONN.cursor()

ENTRY = '.'
DB_FOLDER = os.path.join(ENTRY, 'SRC')
DB_FILES = os.listdir(DB_FOLDER)
DB_SRC = [os.path.join(DB_FOLDER, file) for file in DB_FILES]

TBNAME = 'Data'

MAINTAIN_TIME = datetime.datetime.today().strftime('%Y-%m-%d')

STAFFID = 'staff_id'

EXCEL_SRC = os.path.join(ENTRY,'performance.xlsm')

# LINE
TOKEN = 'li4Dy05aUYcrZbzoDbE3Zb1LFElmMnCEQBlhV1q7QZB'
MSG1 = '\n Finish Creating DB'
MSG2_1 = '\n Excel & DB Mnthly Data: Correct!'
MSG2_2 = '\n Excel & DB Mnthly Data: Error!'
MSG3_1 = '\n Excel & DB Cumulative Data: Correct!'
MSG3_2 = '\n Excel & DB Cumulative Data: Error!'
MSG4_1 = '\n Excel & Excel Checking/Monthly Table: Correct!'
MSG4_2 = '\n Excel & Excel Checking/Monthly Table: Error!'
MSG5_1 = '\n Excel & Excel Checking/Monthly Chart: Correct!'
MSG5_2 = '\n Excel & Excel Checking/Monthly Chart: Error!'