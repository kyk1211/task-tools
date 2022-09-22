import os
import sys
import pandas as pd
import numpy as np

def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

def mngNumber(list: list):
    new_list = []
    prev_val = ""
    for index, value in enumerate(list):
        val = ""
        if (value != value):
            new_list.append(np.nan)
        elif (index == 0):
            new_list.append("00001")
        elif (prev_val == value):
            val = new_list[-1]
            new_list.append(val)
        elif (prev_val != value):
            val = str(int(new_list[-1]) + 1).zfill(5)
            new_list.append(val)
        prev_val = value
    return new_list

def orderNumber(list: list):
    new_list = []
    dic = dict()
    for i in list:
        dic[i] = dic.get(i, 0) + 1
    for key, value in dic.items():
        temp = [str(x).zfill(2) for x in range(1, value + 1)]
        new_list += temp
    return new_list

print('----------------------------------------------------')
print('작업시작')

if getattr(sys, 'frozen', False):
    path = os.path.dirname(os.path.abspath(sys.executable))
else:
    path = os.path.dirname(os.path.abspath(__file__))

print('현재경로: ', path)

file_list = os.listdir(path)
file_list_xls = [file for file in file_list if file.endswith(".xls")]
print("xls파일 리스트: ", file_list_xls)

for file_xls in file_list_xls:
    print("타겟: ", file_xls)
    try:
        df = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8', converters={c:lambda x: str(x) for c in df.columns})[0]
        columns = xls.columns
        xls = xls.sort_values(by=['지주번호', columns[2]], axis=0)
        xls = pd.DataFrame(xls)
        xls = xls[['지주번호', columns[2]]]
        code_list = xls['지주번호'].values.tolist()
        code_list_new = mngNumber(code_list)
        xls["지주번호정리"] = code_list_new
        order_list_new = orderNumber(code_list_new)
        xls['순번정리'] = order_list_new

        export_name = file_xls.split('.')[0]
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{export_name}_result.xlsx')
        xls.to_excel(save_dir, index=False)
        print(f'{export_name}_result.xlsx 저장 완료')
    except Exception as e:
        print(e)


print('----------------------------------------------------')

os.system('Pause')