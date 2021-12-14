import os
import pandas as pd

checkDir = r'C:\Users\USER\Desktop\work'
df = pd.read_excel(r'C:\Users\USER\Desktop\리네임.xlsx')
checklist = df['리네임이미지']

num = 0
for i in checklist:
    if not os.path.isfile(f'{checkDir}\{i}'):
        num += 1
print(num)