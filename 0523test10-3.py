#import library

import requests, xmltodict, json
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import time
from urllib import parse
import datetime as dt
import urllib3
import traceback
import os

def mkfilename(): # datetime과 지역구로 파일명 생성
    x = dt.datetime.now()
    str(x)
    print(x)
    x2 = x.strftime("%a") + x.strftime("%b") + x.strftime("%d") + x.strftime("%f") + x.strftime("%y")
    print(x2)
    filename = location + ee + x2
    print(filename)
    time.sleep(3)
    return filename

def mkdataframe(url): #url로 호출한 api pandas dataframe로 받기
    print (url)
    print ("3초뒤에 계속됩니다")
    time.sleep(3)

    res = requests.get(url)

    #파싱 및 변수명 재설정
    dict_res = xmltodict.parse(res.content)
    json_string = json.dumps(dict_res['response']['body']['items'], ensure_ascii=False)

    jsonObj = json.loads(json_string)
    df = pd.DataFrame(jsonObj['item'])

    return df

def printarea():
    area = float(input("면적을 숫자로만 입력하세요 (m2) : "))
    areas = area - 3.3
    areab = area + 3.3
    df['전용면적'] = df['전용면적'].astype(float)
    df1 = df.loc[(df['전용면적'] >= (areas)) & (df['전용면적'] <= (areab))]
    print(df1)
    df1.to_excel('C:/Users/skflc/Desktop/0523result/' + filename + ' filtered.xlsx')

def printarea1():
    area = float(input("면적을 숫자로만 입력하세요 (m2) : "))
    areas = area - 66
    areab = area + 66
    df['연면적'] = df['연면적'].astype(float)
    df1 = df.loc[(df['연면적'] >= (areas)) & (df['연면적'] <= (areab))]
    print(df1)
    df1.to_excel('C:/Users/skflc/Desktop/0523result/' + filename + ' filtered.xlsx')

def printarea2():
    area = float(input("면적을 숫자로만 입력하세요 (m2) : "))
    areas = area - 3.3
    areab = area + 3.3
    df['계약면적'] = df['계약면적'].astype(float)
    df1 = df.loc[(df['계약면적'] >= (areas)) & (df['계약면적'] <= (areab))]
    print(df1)
    df1.to_excel('C:/Users/skflc/Desktop/0523result/' + filename + ' filtered.xlsx')

def main():
    while True:
        try:
            global location #지역
            global filename #파일네임
            # urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            # 경고가 거슬리면 위의 urllib앞 주석(#) 제거
            # 법정동코드 api로써 조회
            location = input("건물 주소를 도-시-구-동 까지 입력해주세요 : ")
            location_encoded = parse.quote(location)

            print("입력한 건물주소는 : " + location + " 입니다. ")
            print("encoding한 주소는 [" + location_encoded + "] 입니다")
            print("3초뒤에 계속됩니다.")
            time.sleep(3)

            url0 = 'https://apis.data.go.kr/1741000/StanReginCd/getStanReginCdList?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&pageNo=1&numOfRows=3&type=xml&locatadd_nm=' + location_encoded
            res0 = requests.get(url0, verify=False)
            dict_res0 = xmltodict.parse(res0.content)
            json_string0 = json.dumps(dict_res0['StanReginCd'], ensure_ascii=False)

            jsonObj0 = json.loads(json_string0)
            df0 = pd.DataFrame(jsonObj0['row'], index=[0])
            df0.columns = ['법정동코드전체', '시도코드', '시군구코드', '읍면동코드', '리코드', '주민지역코드', '지적지역커드', '지역주소명', '서열', '비고',
                           '상위지역코드', '최하위지역명', '생성일']

            print(df0)
            search = df0[['법정동코드전체', '시도코드', '시군구코드']].head()
            search2 = df0.loc[0, ['법정동코드전체']]
            print(search)
            print(search2.head())

            # api 요청변수 입력부분 법정동코드와 연월

            house = input("건물 종류를 숫자로 입력해주세요 1.아파트 2.연립/다세대 3.오피스텔 4.단독/다가구 : ")
            sale = input("거래종류를 숫자로 입력해 주세요 1.매매 2.전/월세 : ")
            LAWD_CD = input("법정동코드 앞 5자리를 입력해 주세요 : ")
            YMD = input("연월을 입력해 주세요 예) 202204 : ")
            global ee
            global url
            global df
            global rent  # 월세
            global area  # 필터링용면적
            global deposit  # 보증금
            ee = 123

            print("요청 자료는" + house + "와" + sale + "입니다")

            if house == '1' and sale == '1':
                print("아파트 매매입니다")
                ee = '아파트 매매'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['거래금액', '거래유형', '건축년도', '년', '법정동', '아파트', '월', '일', '전용면적', '중개사소재지', '지번', '지역코드', '층',
                              '해제사유발생일', '해제사유']

            elif house == '1' and sale == '2':
                print("아파트 전월세입니다.")
                ee = '아파트 전월세'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptRent?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '법정동', '보증금액', '아파트', '월', '월세금액', '일', '전용면적',
                              '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']

            elif house == '2' and sale == '1':
                print("연립/다세대 매매입니다.")
                ee = '연립 다세대 매매'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcRHTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['거래금액', '거래유형', '건축년도', '년', '대지권면적', '법정동', '연립다세대', '월', '일', '전용면적', '중개사소재지', '지번',
                              '지역코드', '층', '해제사유발생일', '해제여부']

            elif house == '2' and sale == '2':
                print("연립/다세대 전월세입니다.")
                ee = '연립 다세대 전월세'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcRHRent?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '법정동', '보증금액', '연립다세대', '월', '월세금액', '일', '전용면적',
                              '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']

            elif house == '3' and sale == '1':
                print("오피스텔 매매입니다.")
                ee = '오피스텔 매매'
                print(ee)
                # url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=11110&DEAL_YMD=201512'
                url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['거래금액', '거래유형', '건축년도', '년', '단지', '법정동', '시군구', '월', '일', '전용면적', '중개사소재지', '지번', '지역코드',
                              '층', '해제사유발생일', '해제사유']

            elif house == '3' and sale == '2':
                print("오피스텔 전월세입니다.")
                ee = '오피스텔 전월세'
                print(ee)
                url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiRent?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '단지', '법정동', '보증금', '시군구', '월', '월세', '일', '전용면적',
                              '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']

            elif house == '4' and sale == '1':
                print("단독/다가구 매매입니다.")
                ee = '단독 다가구 매매'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcSHTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['거래금액', '거래유형', '건축년도', '년', '대지면적', '법정동', '연면적', '월', '일', '주택유형', '중개사소재지', '지역코드',
                              '해제사유발생일', '해제여부']

            elif house == '4' and sale == '2':
                print("단독/다가구 전월세입니다.")
                ee = '단독 다가구 전월세'
                print(ee)
                url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcSHRent?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=' + LAWD_CD + '&DEAL_YMD=' + YMD
                df = mkdataframe(url)
                df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '계약면적', '년', '법정동', '보증금액', '월', '월세금액', '일',
                              '종전계약보증금', '종전계약월세', '지역코드']

            else:
                print("그외의 경우 입니다 3초뒤 처음으로 돌아갑니다.")
                time.sleep(3)
                continue

            # elif로 했지만 더 좋은 방법이 있을것... 본 코드도 def가 전혀 없다.
            c = input("그만하려면 x를 누르세요 : ")
            if c == 'x':
                break

            print(ee)
            time.sleep(1)

            # url 작성란 요청변수는 입력을 받아야 할 수 있음 moit.go.kr이 원래 url 문제는 파일 명을 건물 분류마다 다르게 할 필요가 있음
            # url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=11710&DEAL_YMD=201512'

            # url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD='+LAWD_CD+'&DEAL_YMD='+YMD

            # df=mkdataframe(url)
            # df.columns = ['거래금액','거래유형','건축년도','년','법정동','아파트','월','일','전용면적','중개사소재지','지번','지역코드','층','해제사유발생일','해제사유']

            print(df)

            filename = mkfilename()
            df.to_excel(filename + ".xlsx")
            df.to_excel('C:/Users/skflc/Desktop/0523result/' + filename + '.xlsx')  # 폴더경로+ 파일명 + .xlsx
            time.sleep(3)

            if house == '1' and sale == '1':
                print("아파트 매매입니다")
                printarea()
            elif house == '1' and sale == '2':
                print("아파트 전월세입니다.")
                printarea()

            elif house == '2' and sale == '1':
                print("연립/다세대 매매입니다.")
                printarea()
            elif house == '2' and sale == '2':
                print("연립/다세대 전월세입니다.")
                printarea()
            elif house == '3' and sale == '1':
                print("오피스텔 매매입니다.")
                printarea()
            elif house == '3' and sale == '2':
                print("오피스텔 전월세입니다.")
                printarea()
            elif house == '4' and sale == '1':
                print("단독/다가구 매매입니다.")
                printarea1()
            elif house == '4' and sale == '2':
                print("단독/다가구 전월세입니다.")
                printarea2()
            else:
                print("그외의 경우 입니다 3초뒤 처음으로 돌아갑니다.")
                time.sleep(3)
                continue

            # 액셀을 pandas로 열고 쓰고 닫는방식

            # 단순 끝지점 체크용
            print("excel 파일 생성이 완료되었습니다.")
            print(" 본 액셀시트는 구 단위로 검색됩니다.")

            b = input(" 종료하려면 X를 누르세요 : ")
            if b == 'X':
                break

        except KeyError as e:
            print(e)
            print(traceback.format_exc())
            print('잘못된 주소입력입니다.')
            time.sleep(1)
            print('주소를 한글로 정확히 입력했는지 확인바랍니다.')
            time.sleep(1)
            print('광역시는 광역시 이름을 붙여야 합니다 예)서울시 >> X 서울특별시 >> O')
            time.sleep(1)
            print('3초뒤 다시 시작합니다')
            time.sleep(3)
        except TypeError as e:
            print(e)
            print(traceback.format_exc())
            print('잘못된 법정동코드/연월 입력입니다.')
            time.sleep(1)
            print(' 법정동코드/연월을 다시 확인해주세요')
            time.sleep(1)
            print('3초뒤 다시 시작합니다')
            time.sleep(3)
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            print('예기치 못한 오류가 발생했습니다.')
            print('3초뒤 다시 시작합니다')
            time.sleep(3)

main()

print ("X를 누르셨습니다. 곧 종료됩니다.")


# 개선점들
# 1. 요청변수를 여기선 고정했는데 요청변수를 사용자 입력을 받아서 그걸 url로 지정해야함 =>해결
# 2. csv 저장하는데 퍼미션 디나이드가 있음 번호도 일단 랜덤하게 정하던지 해서 피해야함 =>해결
# 3. api를 총 7개를 씀 단독 다가구까지 쓰면 10개까지도 가능한 부분 특히 법정동 코드는 그부분만 읽어야 하는데.... 이건 별개 =>사용자 입력받음
# 4. 가장 마지막 작업일 듯 한데 결국 이걸 하나의 파이썬 exe 파일로 만들어야 할듯 함 아님 액셀 파이썬 연동을 하던가.
# 5. 건물 분류와 전세,월세에 따라 다른 api를 적용해야하는 문제.. 현재도 가능은 하나 뭔가 무식하게 하는듯한 기분이 ->해결
# 6. 파일명을 날짜+시간+아파트분류+주소로 적어야할 필요성이 생김 따라서 toexel을 다시써야될듯? => 해결
# 7. 다른게 맞다 하나는 16열이고 하나는 15열 이거 하나때문에 따로 만들기에는 너무 비효율적이다. 근데 하나만 이런게 아니라 다 지금 api별로 틀리다. =>해결
# 8 filename지정만 def로 뺐는데 나머지도 좀 가능하면 좋은데... 데이터 프레임 동적할당이 안되서 url따로 지정하는게 안되는게 슬프다. 30줄 X6이니까 40줄로 해결될게 180줄 잡아먹는거 좀 비효율 적인데... =>해결
# 9. 본 버젼은 함수 지정 버젼 그래봐야 url이 제각기지만 url만 지정이 되면 되는거니가. =>해결 6번까지
# 10. 함수 각 지정 => 해결
# 11 예외처리, 오류처리가 전혀 없다. #8에서 try로 박아두는거 생각할것 =>해결
# 12. 궁극적인 목표 GUI 구현은 어떻게 할까.. pyqt5 기억안나는데 아마 따로 진행해야 할것 같다 기능은 다 나왔거든
# 13. pyinstaller도 결국 puqt5로 만든걸 실행파일 하는거고 이게 있어야 파이썬이 없는 환경에서도 이게 되는거라서

# 3,12,13 만 해결하면 사실상 완성 이건 기록으로 남겨두고 싶다. 파이썬과 액셀은 정말 명기다 이거 액셀에서 안되서 파이참으로 간건데
# test9이 이전버전 test10,11은 test9의 필터링과 pyqt5로  구현하는 것이 목적이 됨
# 필터링 기능은 부가적인 기능으로 보류 test9로 pyqt진행 여기서는 면적만 받는것으로 함

# 본 코드는 원래 전체의 코드가 def화 됨