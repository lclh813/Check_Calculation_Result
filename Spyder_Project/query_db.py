from data_config import *
import pandas as pd
from tqdm import tqdm

class File_Loader:
    def __init__(self):
        self.path = DB_SRC
        
    def import_data(self) -> pd.DataFrame:
        data = pd.DataFrame([])
        for file in self.path:
            data = data.append(pd.read_excel(file, encoding='ANSI'))
        return data       
        
class Create_DB:
    def __init__(self):
        self.table_name = TBNAME       
        self.maintain_time = MAINTAIN_TIME
        self.staff_id = STAFFID 
        self.conn = CONN
        self.col_name = list(File_Loader().import_data().columns) \
                        + ['Maintain Time', 'Staff ID']
        self.col_type = ['datetime2(7)'] + ['nvarchar(50)']*2 + \
                        ['float', 'datetime2(7)','nvarchar(50)']
        self.df = File_Loader().import_data()
                 
    def empty_table(self) -> pd.DataFrame:
        conn = self.conn
        cursor = conn.cursor()
        pairs = zip(self.col_name, self.col_type)
        sql_str = f'CREATE TABLE "{self.table_name}" ('
        sql_str += ',\n'.join([f'"{i}" {j}' for i, j in pairs]) + ')'
        cursor.execute(sql_str)
        conn.commit()
        cursor.close()
        
    def delete_table(self):
        conn = self.conn
        cursor = conn.cursor()
        sql_del = f'DELETE FROM "{self.table_name}"'
        cursor.execute(sql_del)
        conn.commit()
        cursor.close()
        
    def insert_table(self) -> pd.DataFrame:
        conn = self.conn
        cursor = conn.cursor()
        df = self.df
        maintain_time = self.maintain_time
        for index, row in tqdm(df.iterrows()):
            sql_add = f'INSERT INTO "{self.table_name}" VALUES ('
            sql_add += str(list(row.astype(str)))[1:-1].replace('"NULL"','NULL') 
            sql_add += f', "{self.maintain_time}", "{self.staff_id}")'
            cursor.execute(sql_add)
            conn.commit()
        cursor.close()

class Query_DB:
    def __init__(self):
        self.table_name = TBNAME
        self.conn = CONN           

    def get_sqlstr(self) -> str:
        sqlstr = f'SELECT * FROM "{self.table_name}"'
        return sqlstr
    
    def get_data(self) -> pd.DataFrame:
        conn = self.conn
        sqlstr = self.get_sqlstr()
        df = pd.read_sql(sqlstr, conn)
        return df
    
    def clean_data(self) -> pd.DataFrame:
        df = Query_DB().get_data()  
        df['Time'] = pd.to_datetime(df['Time'])
        return df
        
        
        
        

