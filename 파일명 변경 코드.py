import os
import pandas as pd

#### 작업폴더(targetDir) 엑셀파일(df) 현재이름목록(curlist) 바꿀이름목록(wantlist) 
#### 변경필요
targetDir = r'C:\Users\USER\Desktop\rename1'
wantDir = r'C:\Users\USER\Desktop\work'
df = pd.read_excel(r'C:\Users\USER\Desktop\리네임.xlsx')
curlist = df['기존이미지'] 
wantlist = df['리네임이미지']
list = []
#### 파일명 수정 코드
num = 0
fail_list = []
for idx, name in enumerate(curlist):
    try:
        src = os.path.join(targetDir, name)
        dst = wantlist[idx]
        dst1 = os.path.join(wantDir, dst)
        os.rename(src, dst1)
    except:
        fail_list.append((name, dst))
        print(f'{name} => {dst} failed')
print(fail_list)