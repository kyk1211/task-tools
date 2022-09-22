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
        xls = xls.sort_values(by=['지주번호', '신호등번호'], axis=0)
        area = xls[['시군구명']].value_counts().index.tolist()[0][0]
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
            s = '-'.join([area_code, str(i[2]), str(i[3]), str(i[4])])
            mng_code.append(s)
        result_df["시도명"] = xls["시도명"]
        result_df["시군구명"] = xls["시군구명"]
        result_df["도로종류"] = xls["도로종류"]
        result_df["도로노선번호"] = xls["도로노선번호"]
        result_df["도로노선명"] = xls["도로노선명"]
        result_df["도로노선방향"] = xls["도로노선방향"]
        result_df["소재지도로명주소"] = xls["소재지 도로명주소"]
        result_df["소재지지번번호"] = xls["소재지 지번주소"]
        result_df["위도"] = xls["위도"]
        result_df["경도"] = xls["경도"]
        result_df["신호기설치방식"] = xls["신호기 설치방식"]
        result_df["도로형태"] = xls["도로형태"]
        result_df["주도로여부"] = xls["주도로 여부"]
        result_df["신호등관리번호"] = mng_code
        result_df["신호등구분"] = xls["신호등 구분"]
        result_df["신호등색종류"] = xls["신호등색 종류"]
        result_df["신호등화방식"] = xls["신호등화 방식"]
        result_df["신호등화순서"] = xls["신호등화 방식 순서"]
        result_df["신호등화시간"] = xls["신호등화 방식 시간"]
        result_df["광원종류"] = xls["광원 종류"]
        result_df["신호제어방식"] = xls["신호제어방식"]
        result_df["신호시간결정방식"] = xls["신호시간결정방식"]
        result_df["점멸등운영여부"] = xls["점멸등 운영 방식"]
        result_df["점멸등운영시작시각"] = xls["점멸등 운영 시작 시간"]
        result_df["점멸등운영종료시각"] = xls["점멸등 운영 종료 시간"]
        result_df["보행자작동신호기유무"] = xls["보행자 작동 신호기 유무"]
        result_df["잔여시간표시기유무"] = xls["잔여시간 표시기 유무"]
        result_df["시각장애인용음향신호기유무"] = xls["시각장애인용 음향신호기 유무"]
        result_df["도로안내표지일련번호"] = xls["도로안내표지 일련번호"]
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