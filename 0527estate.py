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
    filename = ee +' '+ x2
    print(filename)
    time.sleep(3)
    return filename

def mkdataframe(url): #url로 호출한 api pandas dataframe로 받기
    print (url)
    print ("3초뒤에 계속됩니다")
    time.sleep(3)

    res = requests.get(url,verify=False)

    #파싱 및 변수명 재설정
    dict_res = xmltodict.parse(res.content)
    json_string = json.dumps(dict_res['response']['fields'], ensure_ascii=False)

    jsonObj = json.loads(json_string)
    df = pd.DataFrame(jsonObj['field'])

    return df


while True :
    try:
        #urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        # 경고가 거슬리면 위의 urllib앞 주석(#) 제거

        c = input("그만하려면 x를 누르세요 : ")
        if c == 'x':
            break

        ee = ' 부동산개발업 목록  '
        print('부동산 개발업 사무소 액셀 생성 py입니다.')
        print(ee)
        url = 'https://apis.data.go.kr/1611000/nsdi/EstateDevlopService/attr/getEDOfficeInfo?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&format=xml&numOfRows=2836&pageNo=1'
        # 본 코드는 이 url만 옮겨가며 작성됨
        df = mkdataframe(url)
        df.columns = ['부동산개발업등록번호', '부동산개발업상호', '대표자', '지번', '도로명주소코드', '도로명주소본번', '도로명주소부번', '도로명주소기타주소', '도로명주소읍면동일련번호', '전화번호', '사무소일련번호', '사무소구분코드', '사무소구분', '사무소소유구분코드', '사무소소유구분', '상태코드', '상태명', '#사무실면적', '데이터기분일자', '도로명주소지하코드', '고유번호', '법정동코드', '법정동명', '대장구분코드', '대장구분명', '기타주소']

        print(ee)
        time.sleep(1)

        #url 작성란 요청변수는 입력을 받아야 할 수 있음 moit.go.kr이 원래 url 문제는 파일 명을 건물 분류마다 다르게 할 필요가 있음
        #url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=11710&DEAL_YMD=201512'


        #url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD='+LAWD_CD+'&DEAL_YMD='+YMD

        #df=mkdataframe(url)
        #df.columns = ['거래금액','거래유형','건축년도','년','법정동','아파트','월','일','전용면적','중개사소재지','지번','지역코드','층','해제사유발생일','해제사유']

        print(df)

        filename = mkfilename()
        df.to_excel(filename+".xlsx")
        df.to_excel('C:/Users/skflc/Desktop/0527result/'+filename +'.xlsx') # 폴더경로+ 파일명 + .xlsx
        time.sleep(3)

        ef = df[df['상태명'] =='정상']
        ef = ef.drop(['부동산개발업등록번호', '고유번호', '법정동코드', '대장구분코드', '대장구분명', '도로명주소코드', '도로명주소지하코드', '도로명주소본번', '도로명주소부번', '도로명주소읍면동일련번호', '사무소일련번호', '사무소구분코드', '사무소소유구분코드', '상태코드', '데이터기준일자'])
        ef.to_excel('C:/Users/skflc/Desktop/0527result/'+filename +' filtered.xlsx')

        #else :
        #    print("그외의 경우 입니다 3초뒤 처음으로 돌아갑니다.")
        #    time.sleep(3)
        #    continue


        # 액셀을 pandas로 열고 쓰고 닫는방식

        # 단순 끝지점 체크용
        print("excel 파일 생성이 완료되었습니다.")
        print(" 본 액셀시트는 부동산 시행사 목록입니다..")

        b = input(" 종료하려면 X를 누르세요 : ")
        if b == 'X' :
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
print ("X를 누르셨습니다. 곧 종료됩니다.")


# 개선점들
# 1. 필터링이 좀 2번 거쳐야 할듯함 등록일자와 등록상태 처리상태로
# 필터링 기능은 부가적인 기능으로 보류 test9로 pyqt진행 여기서는 면적만 받는것으로 함