import pandas as pd
from pyproj import Proj, transform
epsg5174= Proj(init="epsg:5174")
wgs84=Proj(init='epsg:4326')
df = pd.read_excel(r'C:\Users\USER\Desktop\xy좌표변환.xlsx')
lats = df['위도']
lngs = df['경도']

df['X'], df['Y'] = transform(wgs84, epsg5174, lngs.tolist(), lats.tolist())

df.to_excel(r'C:\Users\USER\Desktop\xy좌표변환완료.xlsx')