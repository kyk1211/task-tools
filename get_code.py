import os
import sys
import pandas as pd
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
print("xls파일 리스트 : ", file_list_xls)
for file_xls in file_list_xls:
    print('타겟 파일명 : ', file_xls)
    try:
        # xls 파일 읽기
        df = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8', converters={c:lambda x: str(x) for c in df.columns})[0]
        df = xls.iloc[:, [1,2,3,4]]
        df = df[df['상태'] == '검증완료']
        arr = []
        for i in df.itertuples():
            if (i[2] == 'CW'):
                s = '-'.join([i[2], i[1], i[3]])
            else:
                s = '-'.join([i[1], i[2], i[3]])
            arr.append(s)
        df['관리번호'] = arr
        df = df[['관리번호', '상태']]
        # 파일 내보내기
        export_name = file_xls.split('.')[0]
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{export_name}_result.xlsx')
        df.to_excel(save_dir, index=False)
        print(f'{export_name}_result.xlsx 저장 완료')
    except Exception as e:
        print(e)
        print(f'{file_xls} 에러 발생')

print('----------------------------------------------------')

os.system('Pause')