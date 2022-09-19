import os
import sys
import pandas as pd

if getattr(sys, 'frozen', False):
    path = os.path.dirname(os.path.abspath(sys.executable))
else:
    path = os.path.dirname(os.path.abspath(__file__))
print('현재경로 : ', path)
print("-------------------")
print("작업시작")
dir_list = os.listdir(path)

for entry in dir_list:
    target_dir = os.path.join(path, entry)
    if (os.path.isdir(target_dir)):
        try:  
            file_list = os.listdir(target_dir) 
            file_list = list(filter(lambda x: x.endswith(".jpg"), file_list))
            file_list = list(map(lambda x: x.split(".jpg")[0], file_list))
            if (len(file_list) == 0):
                print(f'{entry} 이미지없음')
                continue
            save_dir = os.path.join(path, f'{entry}.xlsx')
            df = pd.DataFrame(file_list, columns = ['파일 이름'])
            df.to_excel(save_dir)
            print(f"{entry} 작업완료")   
        except Exception as e:
            print(e)
            print(f"{entry} 작업실패")

print("-------------------")
os.system('Pause')