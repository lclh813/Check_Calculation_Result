from data_config import *
from query_excel import Query_Excel
from query_db import Query_DB
from line_msg import Line_MSG

import pandas as pd
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tqdm import tqdm

class Check_Excel_DB:
    def __init__(self):
        p0, p1, p2, p3 = Query_Excel().get_all_sheet()
        self.db_data = Query_DB().clean_data()
        self.excel_ref = p0 
        self.excel_check = p1 
        self.excel_table = p2 
        self.excel_chart = p3 
        self.dist_list = Query_Excel().dist_list 
        self.item_list = p0.range('B4:D99').value 
        self.excel_time = [THIS_MONTH, LAST_YEAR, LAST_MONTH]       
        self.db_time = [datetime.strptime(t, '%Y/%m/%d') for t in self.excel_time]   

    def excel_db_monthly(self): 
        df1 = self.excel_check
        db = self.db_data        

        df1.range('E5:E7').value = [[t] for t in self.excel_time]
        
        check_results = []
        
        for item in tqdm(self.item_list):            
            df1.range('H5:H7').value = [[elem] for elem in item]           
            
            for key in self.dist_list:
                for i in range(len(self.dist_list[key])): 
                    df1.range('K6:K7').value = [[key], [self.dist_list[key][i]]]                    
                    
                    excel_sales_this_month, excel_sales_last_year, excel_sales_last_month = \
                    [round(num) for num in df1.range('E10:E12').value]
                    
                    sqlstr_item = item[2]
                    sqlstr_dist = self.dist_list[key][i]                    
                    db_time_str = [str(db_t)[:10] for db_t in self.db_time]                    

                    this_month_query = \
                    'Time=@db_time_str[0] & Category==@sqlstr_item & Channel==@sqlstr_dist'
                    last_year_query = \
                    'Time==@db_time_str[1] & Category==@sqlstr_item & Channel==@sqlstr_dist'
                    last_month_query = \
                    'Time==@db_time_str[2] & Category==@sqlstr_item & Channel==@sqlstr_dist'
                
                    db_sales_this_month = db.query(this_month_query)['Sales']
                    db_sales_last_year = db.query(last_year_query)['Sales']
                    db_sales_last_month = db.query(last_month_query)['Sales']
                    
                    if db_sales_this_month.empty or \
                       db_sales_last_year.empty or \
                       db_sales_last_month.empty:
                           continue
                    
                    else:               
                        db_sales_this_month = round(float(db_sales_this_month))
                        db_sales_last_year = round(float(db_sales_last_year))
                        db_sales_last_month = round(float(db_sales_last_month))    
                    
                    if (excel_sales_this_month != db_sales_this_month) or \
                       (excel_sales_last_year != db_sales_last_year) or \
                       (excel_sales_last_month != db_sales_last_month):
                           check_result = f'Error!: Category:{item[2]}, Channel:{key}, Sub-Channel:{self.dist_list[key][i]};'
                           check_result += f'Excel/Database: Sales of This Month:{[excel_sales_this_month, db_sales_this_month]};'
                           check_result += f'Excel/Database: Sales of Last Year:{[excel_sales_last_year, db_sales_last_year]};'
                           check_result += f'Excel/Database: Sales of Last Month:{[excel_sales_last_month, db_sales_last_month]}'
                           check_results.append(check_result)
        
        if len(check_results)==0:
            Line_MSG(TOKEN, MSG2_1)
        else:
            Line_MSG(TOKEN, MSG2_2)
            with open('excel_db_monthly.txt', mode='wt', encoding='utf-8') as file:
                for check_result in check_results:
                    file.write('\n'.join(str(line) for line in check_result.split(';')))
                    file.write('\n ------------------------ \n')
                    
    def excel_db_cumsum(self):
        df1 = self.excel_check
        db = self.db_data
        
        df1.range('E5:E7').value = [[t] for t in self.excel_time]
        
        check_results = []
        
        for item in tqdm(self.item_list):
            df1.range('H5:H7').value = [[elem] for elem in item]           
            for key in self.dist_list:
                for i in range(len(self.dist_list[key])): 
                    df1.range('K6:K7').value = [[key], [self.dist_list[key][i]]] 
                       
                    excel_cumsum_this_month = df1.range('E35').value
                    excel_cumsum_last_year = df1.range('E51').value                    
               
                    sqlstr_item = item[2]
                    sqlstr_dist = self.dist_list[key][i]
                    
                    db_cumsum_this_month = []
                    for j in range(self.db_time[0].month):
                        date = str(self.db_time[0] - relativedelta(months=j))[:10]
                        
                        tmp_this_month_query = \
                        'Time==@date & Category==@sqlstr_item & Channel==@sqlstr_dist'
                        
                        tmp_cumsum_this_month = db.query(tmp_this_month_query)['Sales']         
                        
                        if tmp_cumsum_this_month.empty:
                            db_cumsum_this_month.append(tmp_cumsum_this_month)
                        else:
                            db_cumsum_this_month.append(float(tmp_cumsum_this_month))
                          
                    db_cumsum_last_year = []
                    for j in range(self.db_time[1].month):
                        date = str(self.db_time[1] - relativedelta(months=j))[:10]
                        
                        tmp_last_year_query = \
                        'Time==@date & Catefory==@sqlstr_item & Channel==@sqlstr_dist'
                                                
                        tmp_cumsum_last_year = db.query(tmp_last_year_query)['Sales']
                        if tmp_cumsum_last_year.empty:
                            db_cumsum_last_year.append(tmp_cumsum_last_year)
                        else:
                            db_cumsum_last_year.append(float(tmp_cumsum_last_year))
                    
                    if (type(sum(db_cumsum_this_month)) == type(pd.Series([]))) or \
                       (type(sum(db_cumsum_last_year)) == type(pd.Series([]))):
                        continue
                    
                    if (round(excel_cumsum_this_month) != round(sum(db_cumsum_this_month))) or \
                       (round(excel_cumsum_last_year) != round(sum(db_cumsum_last_year))): 
                        check_result = f'Error!: Categoty:{item[2]}, Channel:{key}, Sub-Channel:{self.dist_list[key][i]};'
                        check_result += f'Excel/Database: Cumulative Sales of This Year:{[excel_cumsum_this_month, db_cumsum_this_month]};'
                        check_result += f'Excel/Database: Cumulative Sales of Last Year:{[excel_cumsum_last_year, db_cumsum_last_year]}'
                        check_results.append(check_result)
        
        if len(check_results)==0:
            Line_MSG(TOKEN, MSG3_1)
        else:
            Line_MSG(TOKEN, MSG3_2)
            with open('excel_db_cumsum.txt', mode='wt', encoding='utf-8') as file:
                for check_result in check_results:
                    file.write('\n'.join(str(line) for line in check_result.split(';')))
                    file.write('\n ------------------------ \n')