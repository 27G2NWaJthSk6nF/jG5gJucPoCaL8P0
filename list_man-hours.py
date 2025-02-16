from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import os

directory_path = 'data/'
path_result = 'result_man-hours.xlsx'

"""
実工数読み込み
"""
sheet_name = '実工数'
all_data = []
excel_files = [f for f in os.listdir(directory_path) if f.endswith('.xlsx')]
df = pd.DataFrame(columns=[])
for file in excel_files:
    file_path = os.path.join(directory_path, file)
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    #df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', index_col=0)
    all_data.append(df)
merged_df = pd.concat(all_data, ignore_index=True)
print(merged_df)
grouped_df = merged_df.groupby('担当').sum().reset_index()
#grouped_df = grouped_df.dropna()
print(grouped_df)

"""
出力
"""
#df.to_excel('out.xlsx', index=False, engine='openpyxl')
wb = Workbook()
ws = wb.active
#for r in dataframe_to_rows(grouped_df, index=True, header=True):
#import pdb; pdb.set_trace()
for r in dataframe_to_rows(grouped_df):
    ws.append(r)
ws.auto_filter.ref = ws.dimensions
wb.save(path_result)
