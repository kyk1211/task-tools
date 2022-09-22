from unittest import result
import pandas as pd
import sys
import os

def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")

print('----------------------------------------------------')
print('작업시작')

if getattr(sys, 'frozen', False):
    path = os.path.dirname(os.path.abspath(sys.executable))
else:
    path = os.path.dirname(os.path.abspath(__file__))

print('현재경로 : ', path)

file_list = os.listdir(path)
file_list_xls = [file for file in file_list if file.endswith(".xls")]
print("xls파일 리스트 : ", file_list_xls)

for file_xls in file_list_xls:
    print('타겟 파일명 : ', file_xls)
    try:
        df = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8', converters={c:lambda x: str(x) for c in df.columns})[0]        
        xls = xls.sort_values(by=['지주번호', '과속방지턱 관리번호'], axis=0)
        area = xls[['시군구명']].values.tolist()[0]
        if (area == '구리시'):
          area_code = '31120'
        elif (area == '양주시'):
          area_code = '31260'
        elif (area == '포천시'):
          area_code = '31270'
        elif (area == '하남시'):
          area_code = '31180'
        elif (area == '김해시'):
          area_code = '38070'
        else:
          raise Exception("지역이 잘못되었습니다.")
        result_df = pd.DataFrame()
        mng_code = []
        
        for i in xls.itertuples():
            s = '-'.join([area_code, str(i[3]), str(i[2]), str(i[4])])
            mng_code.append(s)
        result_df["과속방지턱관리번호"] = mng_code
        result_df["시도명"] = xls["시도명"]
        result_df["시군구명"] = xls["시군구명"]
        result_df["도로명"] = xls["도로노선명"]
        result_df["소재지도로명주소"] = xls["소재지 도로명주소"]
        result_df["소재지지번번호"] = xls["소재지 지번주소"]
        result_df["설치장소"] = xls["설치장소"]
        result_df["과속방지턱재료"] = xls["과속방지턱 재료"]
        result_df["과속방지턱형태구분"] = xls["과속방지턱 형태구분"]
        result_df["과속방지턱높이"] = xls["과속방지턱 높이"]
        result_df["과속방지턱폭"] = xls["과속방지턱 폭"]
        result_df["과속방지턱연장"] = xls["과속방지턱 연장"]
        result_df["도로유형구분"] = xls["도로유형 구분"]
        result_df["규격여부"] = xls["규격 여부"]
        result_df["위도"] = xls["위도"]
        result_df["경도"] = xls["경도"]
        result_df["보차분리여부"] = xls["보차분리 여부"]
        result_df["연속형여부"] = xls["연속형 여부"]
        result_df["과속방지턱설치연도"] = xls["과속방지턱 설치일자"]
        result_df["관리기관명"] = xls["관리기관명"]
        result_df["관리기관전화번호"] = xls["관리기관 전화번호"]
        result_df["데이터기준일자"] = '2022-12-15'
        result_df['기타사항']=xls["기타사항"]
        export_name = file_xls.split('.')[0]
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{export_name}_result.xlsx')
        result_df.to_excel(save_dir, index=False)
        print(f'{export_name}_result.xlsx 저장 완료')
    except Exception as e:
        print(e)

print('----------------------------------------------------')

os.system('Pause')