import os
import sys
import pandas as pd
import xlsxwriter
import datetime
from lxml import *

def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

print('----------------------------------------------------')
print('작업시작')
# 경로찾기
if getattr(sys, 'frozen', False):
    path = os.path.dirname(os.path.abspath(sys.executable))
else:
    path = os.path.dirname(os.path.abspath(__file__))
print('현재경로 : ', path)



# xls 파일 찾기
file_list = os.listdir(path)
file_list_xls = [file for file in file_list if file.endswith(".xls")]

for file_xls in file_list_xls:
    print('파일명 : ', file_xls)
    try:
        # xls 파일 읽기
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]

        df = xls.loc[:, ['소재지 지번주소', '상태']]
        df['소재지 지번주소'] = df['소재지 지번주소'].str.split(' ').str[2]

        df_group = df.groupby(['소재지 지번주소', '상태']).size().reset_index(name='count')

        targetStatus = ['조사완료', '검증완료', '입력완료']
        df_count = pd.DataFrame(df_group.groupby(['소재지 지번주소']).sum())
        df_target_count = df_group[df_group['상태'].isin(targetStatus)].groupby(['소재지 지번주소'])['count'].sum().reset_index(name='target_count')
        df_result = pd.merge(df_count, df_target_count, on='소재지 지번주소', how='left')
        df_result.fillna(0, inplace=True)
        df_result.rename(columns={'count': '전체'}, inplace=True)
        
        # 파일 내보내기
        export_name = file_xls.split('.')[0]
        c_date = datetime.datetime.now().date()
        # 1. 엑셀 파일 열기 w/ExcelWriter
        createDirectory(f'{path}/result')
        writer = pd.ExcelWriter(f'{path}/result/{export_name}_result_{c_date}.xlsx' , engine='xlsxwriter')

        # 2. 시트별 데이터 추가하기(1개 시트로 저장할때와 유사)
        df_group.to_excel(writer, sheet_name= 'sheet1', index=False)
        df_result.to_excel(writer, sheet_name= 'sheet2', index=False)

        for column in df_group:
            maxLen = max([len(str(x)) for x in df_group[column]])
            col_idx = df_group.columns.get_loc(column)
            writer.sheets['sheet1'].set_column(col_idx, col_idx, max(maxLen, len(column)) + 10)
        for column in df_result:
            maxLen = max([len(str(x)) for x in df_result[column]])
            col_idx = df_result.columns.get_loc(column)
            writer.sheets['sheet2'].set_column(col_idx, col_idx, max(maxLen, len(column)) + 10)
            

        # 3. 엑셀 파일 저장하기
        writer.save()
        print(f'{export_name}_result_{c_date}.xlsx 저장 완료')
    except Exception as e:
        print(e)
        print(f'{file_xls} 에러 발생')

print('----------------------------------------------------')

os.system('Pause')