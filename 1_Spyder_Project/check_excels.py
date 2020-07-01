from data_config import *
from query_excel import Query_Excel
from query_db import Query_DB
from line_msg import Line_MSG

import pandas as pd
import xlwings as xw
from datetime import datetime
from dateutil.relativedelta import relativedelta
from itertools import groupby
from tqdm import tqdm

class Check_Excels:
    def __init__(self):
        p0, p1, p2, p3 = Query_Excel().get_all_sheet()
        self.excel_check = p1 
        self.excel_table = p2 
        self.excel_chart = p3 
        self.dist_list = Query_Excel().dist_list
        self.item_list = p0.range('B4:D99').value 
        self.excel_time = [THIS_MONTH, LAST_YEAR, LAST_MONTH] 
        
    def dist_trans(self):
        dist_list = self.dist_list
        dist_trans = []
        for key in dist_list:
            for i in range(len(dist_list[key])):
                dist_trans.append([[key], [dist_list[key][i]]])  
        return dist_trans
    
    def get_small_item(self, groupby_index:int, target_index:int) -> list: 
        things = self.item_list
        small_item_list = []  
        for key, group in groupby(things, lambda x: x[groupby_index]):
            small_item = []     
            for thing in group:
                small_item.append(thing[target_index])
            small_item_list.append(small_item)            
        return small_item_list
    
    def excel_check_table(self):
        item_list = self.item_list
        df1 = self.excel_check
        df2 = self.excel_table
        df1.range('E5:E7').value = [[t] for t in self.excel_time] # 指定時間
        df2.range('B5').value = [self.excel_time[0]] # 指定時間
        col1 = self.dist_trans()
        col2 = ['F', 'J', 'N', 'T']
        df2_key = ['Offline_Total', 'Online_Total', 'Online_C', 'Online_D']
        
        df1_dict, df2_dict = {}, {}
        for item in tqdm(item_list):
        
            df1_dict[item[2]], df1_key, df1_val = {}, [], []
            for c in col1[2:]:
                df1.range('H5:H7').value = [[elem] for elem in item]
                df1.range('K6:K7').value = c       
                df1_key.append(df1.range('K7').value)
                df1_val.append(df1.range('E13').value)
            df1_dict[item[2]]= {key: value for key, value in zip(df1_key, df1_val)}              

            df2_dict[item[2]], df2_val = {}, []
            if item[2] in ['Headset', 'Toothbrush']:
                for index, value in enumerate(df2.range('B1:B120').value):
                    if value == item[2]:
                        cells = [c + str(index+1) for c in col2]
            else:
                r = df2.api.UsedRange.Find(item[2])    
                if r == None:
                    continue
                else:
                    s = r.address
                    cells = [c + s.split('$B$')[1] for c in col2]
                
            for cell in cells:
                df2_val.append(df2.range(cell).value)
            df2_dict[item[2]]= {key: value for key, value in zip(df2_key, df2_val)}
            
        check_results = []
        for key in df1_dict.keys():
            check_result_list = []
            for subkey in (df1_dict[key].keys()):  
                if len(df2_dict[key]) == 0:
                    continue
                else:
                    df1_tmp, df2_tmp = df1_dict[key][subkey], df2_dict[key][subkey]
                    df1_res = round(0 if df1_tmp is None or type(df1_tmp) is str else df1_tmp, 4)
                    df2_res = round(0 if df2_tmp is None or type(df2_tmp) is str else df2_tmp, 4)
                    if round(df1_res, 4) != round(df2_res, 4):
                        check_result = f'Error!: Category:{key}'
                        check_result += f'Checking/Monthly Table({subkey}):{[df1_res, df2_res]}'
                        check_result_list.append(check_result)
                check_results.append(check_result_list)
                
        if sum([len(check_result) for check_result in check_results])==0:
            Line_MSG(TOKEN, MSG4_1)
        else:
            Line_MSG(TOKEN, MSG4_2)
            with open('excel_check_table.txt', mode='wt', encoding='utf-8') as file:
                for check_result_list in check_results:
                    file.write('\n'.join(str(check_result) for check_result in check_result_list))
                    file.write('\n------------------------ \n')
    
    def excel_check_chart(self):
        df1 = self.excel_check
        df3 = self.excel_chart
        df1.range('E5:E7').value = [[t] for t in self.excel_time]
        df1.range('K6').value = 'Online'
        col1 = {'Share':['H','K'], 'Sales':'E'}
        row1 = {'Share':[12,19,20], 'Sales':[10,13,14]}
        col3 = {'Category':'T', 
                'Share':'V', 'Share MoM':'X', 'Share YoY':'Z',
                'Sales':'AD', 'Sales MoM':'AF', 'Sales YoY':'AH'}               
        row3 = {'Share':[21,22,23,24], 'Sales':[21,23]}
        df1_share_key = ['online_share', 'panel_share', 
                         'online_share_yoy', 'panel_share_yoy',
                         'online_share_mom', 'panel_share_mom']
        df1_rev_key = ['rev', 'rev_yoy', 'rev_mom']
        df3_share_key = ['online_c_online_share', 'online_d_online_share',
                         'online_c_panel_share', 'online_d_panel_share',
                         'online_c_online_share_mom', 'online_d_online_share_mom',
                         'online_c_panel_share_mom', 'online_d_panel_share_mom',
                         'online_c_online_share_yoy', 'online_d_online_share_yoy',
                         'online_c_panel_share_yoy', 'online_d_panel_share_yoy']        
        df3_rev_key = ['online_c_rev', 'online_d_rev',
                       'online_c_rev_mom', 'online_d_rev_mom',
                       'online_c_rev_yoy', 'online_d_rev_yoy']
        item_list = self.item_list
        small_item_list = self.get_small_item(1, 2) 

        df1_dict = {}
        for item in tqdm(item_list):    
            df1_dict[item[2]] = {}    
            df1.range('H5:H7').value = [[elem] for elem in item]           
            
            df1.range('K7').value = 'online_c'
            df1_online_c_share = [df1.range(key+str(i)).value for i in row1['Share'] for key in col1['Share']]
            df1_online_c_rev = [df1.range(key+str(i)).value for i in row1['Sales'] for key in col1['Sales']]        
            df1_dict_online_c_share = {key: value for key, value in zip(['online_c_' + elem for elem in df1_share_key], df1_online_c_share)}
            df1_dict_online_c_rev = {key: value for key, value in zip(['online_c_' + elem for elem in df1_rev_key], df1_online_c_rev)}
            
            df1.range('K7').value = 'Online_D'
            df1_online_d_share = [df1.range(key+str(i)).value for i in row1['Share'] for key in col1['Share']]
            df1_online_d_rev = [df1.range(key+str(i)).value for i in row1['Sales'] for key in col1['Sales']]
            df1_dict_online_d_share = {key: value for key, value in zip(['online_d_' + elem for elem in df1_share_key], df1_online_d_share)}
            df1_dict_online_d_rev = {key: value for key, value in zip(['online_d_' + elem for elem in df1_rev_key], df1_online_d_rev)}
            
            df1_dict[item[2]] = {**df1_dict_online_c_share, **df1_dict_online_c_rev,
                                 **df1_dict_online_d_share, **df1_dict_online_d_rev}
            
        df3_dict = {}
        index = 37       
        for small_items in tqdm(small_item_list):
            for small_item in small_items:
                df3_dict[small_item] = {}  
                df3.range(col3['Category'] + str(index)).value = small_item
                df3_share = [df3.range(col3['Share'] + str(index+i)).value for i in row3['Share']]
                df3_share_mom = [df3.range(col3['Share MoM'] + str(index+i)).value for i in row3['Share']]
                df3_share_yoy = [df3.range(col3['Share YoY'] + str(index+i)).value for i in row3['Share']]
                df3_rev = [df3.range(col3['Sales'] + str(index+i)).value for i in row3['Sales']]
                df3_rev_yoy = [df3.range(col3['Sales YoY'] + str(index+i)).value for i in row3['Sales']]
                df3_rev_mom = [df3.range(col3['Sales MoM'] + str(index+i)).value for i in row3['Sales']]
                df3_dict_share = {key: value for key, value in zip(df3_share_key, df3_share + df3_share_mom + df3_share_yoy)}
                df3_dict_rev = {key: value for key, value in zip(df3_rev_key, df3_rev + df3_rev_mom + df3_rev_yoy)}
            
                df3_dict[small_item] = {**df3_dict_share, **df3_dict_rev} 
                
            index = index + 28
        
        check_results = []
        for key in df1_dict.keys():
            check_result_list = []
            for subkey in (df1_dict[key].keys()):
                df1_tmp, df3_tmp = df1_dict[key][subkey], df3_dict[key][subkey]
                df1_res = round(0 if df1_tmp is None or type(df1_tmp) is str else df1_tmp, 4)
                df3_res = round(0 if df3_tmp is None or type(df3_tmp) is str else df3_tmp, 4)
                if round(df1_res, 4) != round(df3_res, 4):
                    check_result = f'Error!: Category:{key}, '
                    check_result += f'Checking/Monthly Table({subkey}):{[df1_res, df3_res]}'
                    check_result_list.append(check_result)
                check_results.append(check_result_list)
            
        if sum([len(check_result) for check_result in check_results])==0:
            Line_MSG(TOKEN, MSG5_1)
        else:
            Line_MSG(TOKEN, MSG5_2)
            with open('excel_check_chart.txt', mode='wt', encoding='utf-8') as file:
                for check_result_list in check_results:
                    file.write('\n'.join(str(check_result) for check_result in check_result_list))
                    file.write('\n------------------------ \n')
