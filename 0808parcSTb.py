# streamlit 으로 선택을 연습하는데 중점을 둠
# 안타깢디만 주소검색은 현재 보이지 않음

import streamlit as st
import pandas as pd
import numpy as np
import requests, xmltodict, json
import time
from urllib import parse
import datetime as dt
import traceback
from dateutil.relativedelta import relativedelta as rd
from io import BytesIO
from io import StringIO

def mkfilename(): # datetime과 지역구로 파일명 생성
    x = dt.datetime.now()
    str(x)
    print(x)
    x2 = x.strftime("%a") + x.strftime("%b") + x.strftime("%d") + x.strftime("%f") + x.strftime("%y")
    print(x2)
    filename = location + x2
    print(filename)
    time.sleep(3)
    return filename

def mkdataframe(url): #url로 호출한 api pandas dataframe로 받기
    #st.write(url)
    st.write("1초뒤에 계속됩니다")
    time.sleep(1)

    res = requests.get(url, verify=False)

    #파싱 및 변수명 재설정
    dict_res = xmltodict.parse(res.content)
    json_string = json.dumps(dict_res['response']['body']['items'], ensure_ascii=False)

    jsonObj = json.loads(json_string)
    df = pd.DataFrame(jsonObj['item'])

    return df

def convert_df(df):
    return df.to_csv().encode('cp949')

#   return df.to_excel().encode('cp949')

def main():
    try:
        global location  # 지역
        global filename  # 파일네임
        global ee
        global df
        global url
        global YMD
        global Qdate
        global df_list
        global df_all
        KEY = 'kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D'
        st.title('streamlit 연습')
        st.text("Hello Streamlit! 이 글은 튜토리얼 입니다.")

        html = """
            <div style='
                background-color:red;
                color:white;
            '>
                안녕하세요
                </div>
            """

        st.markdown(html ,unsafe_allow_html=True) # 이자리가 원래 주소찾기 자리지만
        #code = '''def hello():
        #            print("test, streamlit")'''

        #st.code(code, language='python')

        ## Checkbox
        location = '서울특별시 종로구'
        location = st.text_input("건물 주소를 구 까지 입력해주세요 서울시는 서울특별시로 입력해 주세요 ex) 서울특별시 송파구, 경기도 성남시 중원구,")
        location_encoded = parse.quote(location)

        url0 = 'https://apis.data.go.kr/1741000/StanReginCd/getStanReginCdList?serviceKey='+ KEY +'&pageNo=1&numOfRows=3&type=xml&locatadd_nm=' + location_encoded
        res0 = requests.get(url0, verify=False)
        dict_res0 = xmltodict.parse(res0.content)
        json_string0 = json.dumps(dict_res0['StanReginCd'], ensure_ascii=False)

        jsonObj0 = json.loads(json_string0)
        df0 = pd.DataFrame(jsonObj0['row'], index=[0])
        df0.columns = ['법정동코드전체', '시도코드', '시군구코드', '읍면동코드', '리코드', '주민지역코드', '지적지역커드', '지역주소명', '서열', '비고',
                       '상위지역코드', '최하위지역명', '생성일']

        st.write(df0)
        search = df0[['법정동코드전체', '시도코드', '시군구코드']].head()
        search2 = df0.loc[0, ['법정동코드전체']]
        sd_code = df0.iloc[0]['시도코드']
        sgg_code = df0.iloc[0]['시군구코드']
        bdleft = sd_code+sgg_code
        st.write(search)
        st.write(search2.head())
        st.write('['+location + '] 의 법정동코드는 [' + bdleft + '] 입니다.' )

        if st.checkbox("체크박스위젯"):
            st.write("체크박스가 선택되었습니다.")

        ## Radio button
        status = st.radio("매매, 전/월세를 선택하세요.", ("매매", "전/월세"))
        if status == "매매":
            st.success("매매입니다.")
            deposit = st.number_input("보증금(매매금액)을 만원단위로 입력하세요")
        elif status == "전/월세":
            st.success("전/월세입니다")
            rent = st.number_input("월세를 만원 단위로 입력하세요")
            deposit = st.number_input("보증금을 만원단위로 입력하세요")
        else:
            st.write("오류입니다.")

        ## Select Box
        occupation = st.selectbox("건물종류를 선택하세요.",
                                  ["아파트",
                                   "연립,다세대",
                                   "단독,다가구",
                                   "오피스텔"])
        if occupation == "단독,다가구":
            aarea = st.number_input("전용면적을 입력해주세요 (m2)")

        YMD = st.date_input("연월을 입력해주세요")
        intbd = int(bdleft)
        st.write('['+location + '] 의 법정동 주소 앞 5자리는 ['+ str(intbd) +']입니다.')
        st.write(YMD)
        YMD = YMD - rd(months=1)
        Qdate = YMD.strftime('%Y%m') #연월만 추출 이역시 문자열
        #Rdate = YMD.isoformat() # 날짜를 문자형으로 변환
        #Sdate = YMD.year # 날짜를 문자형으로 바꿔줌
        #Sdate2 = YMD.month # dt.strftime은 일부만 추출용
        st.write('입력된 연월은 [' + Qdate + '] 입니다.')
        st.write(bdleft)
        #st.write(Rdate)
        #st.write(Sdate)
        #st.write(Sdate2)
        #st.write(type(YMD)) # type()은 단순 자료형 검사기
        #st.write(type(Sdate2))

        st.write("당신은 [", occupation,"] ,[",status, "] 를 선택하셨습니다.")
        st.write("보증금은 ", deposit,"만원입니다.")
        if status == "월세":
            st.write("전/월세는 ", rent, "만원입니다.")

        df_list = []
        df_all = pd.DataFrame(df_list)

        if occupation == '아파트' and status == '매매': #house => occupation 전/월세 ==> status
            ee = '아파트 매매'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTrade?serviceKey='+ KEY +'&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['거래금액', '거래유형', '건축년도', '년', '법정동', '아파트', '월', '일', '전용면적', '중개사소재지', '지번', '지역코드', '층',
                          '해제사유발생일', '해제사유']
            st.write(df)

        elif occupation == '아파트' and status == '전/월세':  # house => occupation 전/월세 ==> status
            ee = '아파트 전월세'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptRent?serviceKey=' + KEY + '&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '법정동', '보증금액', '아파트', '월', '월세금액', '일', '전용면적',
                          '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']
            st.write(df)

        elif occupation == '연립,다세대' and status == '매매':  # house => occupation 전/월세 ==> status
            ee = '연립 다세대 매매'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcRHTrade?serviceKey=' + KEY + '&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['거래금액', '거래유형', '건축년도', '년', '대지권면적', '법정동', '연립다세대', '월', '일', '전용면적', '중개사소재지', '지번',
                          '지역코드', '층', '해제사유발생일', '해제여부']
            st.write(df)

        elif occupation == '연립,다세대' and status == '전/월세':
            ee = '연립 다세대 전월세'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcRHRent?serviceKey=' + KEY + '&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '법정동', '보증금액', '연립다세대', '월', '월세금액', '일', '전용면적',
                          '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']
            st.write(df)

        elif occupation == '오피스텔' and status == '매매':
            ee = '오피스텔 매매'
            st.write(ee)
            for i in range(0, 6):
                ZMD = YMD - rd(months=i) # 등차수열됨
                Qdate = ZMD.strftime('%Y%m')
                # url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiTrade?serviceKey=kGN3LmvLVDIB26IbtOB9522yukFv74%2FVrqMvR2k6fqQftVvcqVj%2FyhIlztF%2BndXlgRCasv4LQd0rklNaCAdMWw%3D%3D&LAWD_CD=11110&DEAL_YMD=201512'
                url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiTrade?serviceKey=' + KEY + '&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
                df = mkdataframe(url)
                df.columns = ['거래금액', '거래유형', '건축년도', '년', '단지', '법정동', '시군구', '월', '일', '전용면적', '중개사소재지', '지번', '지역코드',
                              '층', '해제사유발생일', '해제사유']
                st.write(df)
                st.write(str(ZMD) + '       '+Qdate)
                df_all = pd.concat([df_all,df])
                print(df_all)



        elif occupation == '오피스텔' and status == '전/월세':
            ee = '오피스텔 전월세'
            st.write(ee)
            url = 'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcOffiRent?serviceKey='+ KEY +'&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '년', '단지', '법정동', '보증금', '시군구', '월', '월세', '일', '전용면적',
                          '종전계약보증금', '종전계약월세', '지번', '지역코드', '층']
            st.write(df)

        elif occupation == '단독,다가구' and status == '매매':
            ee = '단독 다가구 매매'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcSHTrade?serviceKey='+ KEY +'&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['거래금액', '거래유형', '건축년도', '년', '대지면적', '법정동', '연면적', '월', '일', '주택유형', '중개사소재지', '지역코드',
                          '해제사유발생일', '해제여부']
            st.write(df)

        elif occupation == '단독,다가구' and status == '전/월세':
            ee = '단독 다가구 전월세'
            st.write(ee)
            url = 'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcSHRent?serviceKey='+ KEY +'&LAWD_CD=' + bdleft + '&DEAL_YMD=' + Qdate
            df = mkdataframe(url)
            df.columns = ['갱신요구권사용', '건축년도', '계약구분', '계약기간', '계약면적', '년', '법정동', '보증금액', '월', '월세금액', '일',
                          '종전계약보증금', '종전계약월세', '지역코드']
            st.write(df)

        else:
            st.warning('아직 건물./ 거래종류가 입력되지 않았습니다.')

        df = df.fillna('NaN')
        filename = mkfilename()
        Xdata = convert_df(df_all)
        st.download_button(
            label= ee + ' 실거래가를 xlsx로 출력',
            data = Xdata,
            file_name=filename + '.csv',
            mime='text/csv'
        )

        if st.button('계속하기'):
            st.write('곧 다시 검색합니다')
        else:
            st.write('중단점 테스트')




    except TypeError as e:
        print(e)
        print(traceback.format_exc())
        st.warning('잘못된 법정동코드/연월 입력입니다.')
        time.sleep(1)
        st.warning(' 법정동코드/연월을 다시 확인해주세요')
        time.sleep(1)
        print('3초뒤 다시 시작합니다')
        time.sleep(3)
    except KeyError as e:
        print(e)
        print(traceback.format_exc())
        print('잘못된 주소입력입니다.')
        st.warning('잘못된 주소입력입니다.')
        time.sleep(1)
        st.warning('주소를 한글로 정확히 입력했는지 확인바랍니다.')
        time.sleep(1)
        st.warning('광역시는 광역시 이름을 붙여야 합니다 예)서울시 >> X 서울특별시 >> O')
        time.sleep(1)
        print('3초뒤 다시 시작합니다')
        st.warning('다시 입력해 주세요')
        time.sleep(3)
    except Exception as e:
        print(e)
        print(traceback.format_exc())
        print('예기치 못한 오류가 발생했습니다.')
        print('3초뒤 다시 시작합니다')
        st.write('예기치 못한 오류가 발생했습니다.')
        st.write('3초뒤 다시 시작합니다')
        time.sleep(3)

main()

# 아마 안될것 같은데 그래서 예제도 전혀 없지 싶은데
# st.markdown으로 <script> 태그를 이용해서 적는걸 한번은 해봐야된다.
# to_excel로 할시 계속 excel_writer 지정한게 없다 난리난리를 친다. 결국 csv로
# excel_writer는 저장 경로인데 이건 컴퓨터마다 다를 것이므로..
# 액셀 다운법이 아예 없진 않은데 https://discuss.streamlit.io/t/download-button-with-excel-file/20794
# 저 방법은 반복문 노가다 매우매우 끔찍
# https://docs.streamlit.io/library/components/components-api components는 내가 생각한 방법이 아닌거같다?
# 원래는 이중 거래금액과 면적에 맞는 데이터들을 또 필터링 해서 보여줘야 한다. 이 과정이 생략되었다.
# 그리고 원래도 최근 6개월 간의 자료를 보여줘야 한다 이는 아마 반복문을 쓰게 될텐데.. 개월이 음수가 되면 연도를 깍아야 하는 if elif가 골치아프다. 날짜 계산이 좀 그래
# csv 다운하면 한번 새로고침이 되서 api 호출을 한번 더해버린다 어떡하지?
# 항시 새로고침을 하는 것도 문제
