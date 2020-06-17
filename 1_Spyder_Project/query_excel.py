from adj_data_config import *
import pandas as pd
import xlwings as xw

class Query_Excel:
    def __init__(self):
        self.path = EXCEL_SRC
        self.sheet_name = ['Sheet1','Sheet2','Sheet3','Sheet4'] 
        self.dist_list = {'Offline': ['Offline_A', 'Offline_B', 'Offline_Total'], 
                          'Online': ['Online_C', 'Online_D', 'Online_Total']}
        
    def get_workbook(self):
        xw.App.visible = False
        wb = xw.Book(self.path)
        return wb
    
    def get_all_sheet(self):
        sheet_list = []
        for i in range(len(self.sheet_name)):
            sheet_list.append(self.get_workbook().sheets[self.sheet_name[i]])
        return sheet_list
        
    def get_cell(self, sheet_index:int, cell:str):
        cell_value = self.get_sheet(sheet_index).range(cell).value
        return cell_value    
