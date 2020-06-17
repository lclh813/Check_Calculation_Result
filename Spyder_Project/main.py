from data_config import *
from query_excel import Query_Excel
from query_db import File_Loader, Create_DB, Query_DB
from check_excel_db import Check_Excel_DB
from check_excels import Check_Excels
from line_msg import Line_MSG

try:
    Create_DB().empty_table() 
except Exception:
    print('Table already exist')
finally:
    Create_DB().delete_table()
    Create_DB().insert_table()
    Line_MSG(TOKEN, MSG1)
    
    Check_Excel_DB().excel_db_monthly
    Check_Excel_DB().excel_db_cumsum()
    Check_Excels().excel_check_table()
    Check_Excels().excel_check_chart()
    







    




