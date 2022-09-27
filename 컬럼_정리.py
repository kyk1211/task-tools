import datetime
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

filter_dic = {
    '구리시': ['갈매동', '사노동', '수택동', '아천동'],
    '양주시': ['남방동', '마전동', '산북동', '어둔동', '유양동', '고읍동', '광사동', '만송동', '삼숭동', '덕정동', '봉양동', '덕계동', '희정동', '고암동','옥정동', '율정동', '회암동'],
    '포천시': ['가산면', '관인면', '군내면', '내촌면', '동교동', '선단동', '설운동', '어룡동', '영북면', '영중면', '이동면', '일동면', '자작동', '창수면', '화현면'],
    '하남시': ['감복동', '감이동', '감일동', '광암동', '교산동', '망월동', '미사동', '배알미동', '상사창동', '상산곡동', '선동','초이동', '초일동', '춘궁동', '풍산동', '하사창동', '하산곡동', '학암동', '항동']
}

for file_xls in file_list_xls:
    print('타겟 파일명 : ', file_xls)
    try:
        df = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8', converters={c:lambda x: str(x) for c in df.columns})[0]
        xls = xls[xls['상태'].isin(['검증완료'])]
        area = xls[['시군구명']].value_counts().index.tolist()[0][0]
        dic = {
          'SS': '도로안전표지 번호',
          'CW': '횡단보도 번호',
          "TL": '신호등번호',
          'SB': '과속방지턱 관리번호'
        }
        if ('분류번호' in xls.columns):
            code = xls.loc[:, ['분류번호']].value_counts().index.tolist()[0][0]
        else:
            code = 'TL'

        filter_target = filter_dic[area]
        if (area == '양주시' and (code == 'CW' or code == 'SS')):
            for i in filter_target:
                xls = xls[~xls['소재지 지번주소'].str.contains(i)]

        if (area in ['하남시', '구리시', '포천시'] and (code == 'SS')):
            for i in filter_target:
                xls = xls[~xls['소재지 지번주소'].str.contains(i)]

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
        xls = xls.sort_values(by=['지주번호', dic[code]], axis=0)
        result_df = pd.DataFrame()
        mng_code = []
        for i in xls.itertuples():
            if (code == 'CW' or code == 'SB'):
                s = '-'.join([area_code, str(i[3]), str(i[2]), str(i[4])])
            else:
                s = '-'.join([area_code, str(i[2]), str(i[3]), str(i[4])])
            mng_code.append(s)
            
        if (code == 'SS'):
            result_df['고유번호'] = xls['고유번호']
            result_df['안전표지일련번호'] = mng_code
            result_df['도로종류'] = xls['도로종류']
            result_df['도로노선번호'] = xls['도로노선번호']
            result_df['도로노선명'] = xls['도로노선명']
            result_df['도로노선방향'] = xls['도로노선방향']
            result_df['도로형태'] = xls['도로형태']
            result_df['차로수'] = xls['차로수']
            result_df['도로폭'] = xls['도로폭']
            result_df['소재지도로명주소'] = xls['소재지 도로명주소']
            result_df['소재지지번주소'] = xls['소재지 지번주소']
            result_df['위도'] = xls['위도']
            result_df['경도'] = xls['경도']
            result_df['안전표지구분'] = xls['도로안전표지 구분']
            result_df['안전표지종별일련번호'] = xls['도로안전표지 종별 일련번호']
            result_df['주행제한속도'] = xls['주행제한속도']
            result_df['안전표지설명'] = xls['도로안전표지 설명']
            result_df['지주형식'] = xls['지주형식']
            result_df['제2외국어표기여부'] = xls['제2외국어 표기여부']
            result_df['설치일자'] = xls['설치일자']
            result_df['관리기관명'] = xls['관리기관명']
            result_df['관리기관전화번호'] = xls['관리기관 전화번호']
            result_df['데이터기준일자'] = '2022-12-15'
            result_df['기타사항']=xls["기타사항"]
        if (code == 'CW'):
            result_df['고유번호'] = xls['고유번호']
            result_df['시도명'] = xls['시도명']
            result_df['시군구명'] = xls['시군구명']
            result_df['도로명']= xls['도로명']
            result_df['소재지도로명주소'] = xls['소재지 도로명주소']
            result_df['소재지지번주소'] = xls['소재지 지번주소']
            result_df['횡단보도관리번호'] = mng_code
            result_df['횡단보도종류'] = xls['횡단보도 종류']
            result_df['자전거횡단도겸용여부'] = xls['자전거횡단도 겸용여부']
            result_df['고원식적용여부'] = xls['고원식 적용여부']
            result_df['위도'] = xls['위도']
            result_df['경도'] = xls['경도']
            result_df['차로수'] = xls['차로수']
            result_df['횡단보도폭'] = xls['횡단보도 폭']
            result_df['횡단보도연장'] = xls['횡단보도 길이(연장)']
            result_df['보행자신호등유무'] = xls['보행자 신호등 유무']
            result_df['보행자작동신호기유무'] = xls['보행자 작동 신호기 유무']
            result_df['음향신호기설치여부'] = xls['음향신호기 유무']
            result_df['녹색신호시간'] = xls['녹색신호시간']
            result_df['적색신호시간'] = xls['적색신호시간']
            result_df['교통섬유무'] = xls['교통섬 유무']
            result_df['보도턱낮춤여부'] = xls['보도턱 낮춤여부']
            result_df['점자블록유무'] = xls['점자블록 유무']
            result_df['집중조명시설유무'] = xls['집중조명시설 유무']
            result_df['관리기관명'] = xls['관리기관명']
            result_df['관리기관전화번호'] = xls['관리기관 전화번호']
            result_df['데이터기준일자'] = '2022-12-15'
            result_df['기타사항']=xls["기타사항"]
        if (code == 'SB'):
            result_df['고유번호'] = xls['고유번호']
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
        if (code == 'TL'):
            result_df['고유번호'] = xls['고유번호']
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
            result_df['신호등화방식특이사항'] = xls['신호등화 방식 특이사항']
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
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{area}-{code}-{datetime.datetime.now().date()}_result.xlsx')
        result_df.to_excel(save_dir, index=False)
        print(f'{area}-{code}-{datetime.datetime.now().date()}_result.xlsx 저장완료')
    except Exception as e:
        print(e)

print('----------------------------------------------------')

os.system('Pause')