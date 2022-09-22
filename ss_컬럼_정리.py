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
        xls = xls.sort_values(by=['지주번호', '도로안전표지 번호'], axis=0)
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
        export_name = file_xls.split('.')[0]
        createDirectory(f'{path}/result')
        save_dir = os.path.join(f'{path}/result', f'{export_name}_result.xlsx')
        result_df.to_excel(save_dir, index=False)
        print(f'{export_name}_result.xlsx 저장 완료')
    except Exception as e:
        print(e)

print('----------------------------------------------------')

os.system('Pause')