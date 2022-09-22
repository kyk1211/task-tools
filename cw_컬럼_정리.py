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
file_list_xls = [file for file in file_list if file.endswith(".xls") or file.endswith(".xlsx")]
print("xls파일 리스트 : ", file_list_xls)

for file_xls in file_list_xls:
    print('타겟 파일명 : ', file_xls)
    try:
        df = pd.read_html(f'{path}/{file_xls}', encoding='utf-8')[0]
        xls = pd.read_html(f'{path}/{file_xls}', encoding='utf-8', converters={c:lambda x: str(x) for c in df.columns})[0]        
        xls = xls.sort_values(by=['지주번호', '횡단보도 번호'], axis=0)
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
            s = '-'.join([area_code, str(i[3]), str(i[2]), str(i[4])])
            mng_code.append(s)
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
        export_name = file_xls.split('.')[0]
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{export_name}_result.xlsx')
        result_df.to_excel(save_dir, index=False)
        print(f'{export_name}_result.xlsx 저장 완료')
    except Exception as e:
        print(e)

print('----------------------------------------------------')

os.system('Pause')