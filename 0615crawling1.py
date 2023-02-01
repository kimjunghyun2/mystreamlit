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
import openpyxl
from bs4 import BeautifulSoup

#공지사항에 쓰임 목적에따라 달라질 수 있음

def void(): #main 격의 자리이며 그외 함수는 이 위에 def로 지정
    while True:
        try:
            print("hello")
            c = input("종료하려면 x를 누르세요 : ")
            if c == 'x':
                break
            time.sleep(1)

            excel_file = openpyxl.Workbook()
            excel_sheet = excel_file.active
            excel_sheet.title = '테스트'
            excel_sheet.append(['번호', '제목', '내용(selectone)', '내용(select)', '작성자', '작성시각'])
            dir = 'C:\\Users\\skflc\\Desktop\\test0615\\testnotice.xlsx'
            j = 1  # 글번호
            # 여기서 부터 반복문
            for i in range(44, 111):

                jj = str(j)
                ii = str(i)
                url = 'https://www.costac.co.kr/bbs/board.php?bo_table=notice&wr_id=' + ii
                print(url)
                source = requests.get(url, verify=False)

                #분기점 source.status_code

                html = source.text
                soap = BeautifulSoup(html, 'html.parser')
                article = soap.select_one('#bo_v_con')
                article2 = soap.select('#bo_v_atc')
                title = soap.select('#bo_v_title') #제목
                auth = soap.select('#bo_v_info > ul > li:nth-child(1) > strong > span') #작성자
                wtime = soap.select('#bo_v_info > ul > li:nth-child(2) > strong') #작성시각

                # time.sleep(1)

                ## print(soap)
                #print('-----------------------------')
                #print(article)
                #print('-----------------------------')
                ## print(article.get_text())
                brticle = str(article)
                brticle2 = str(article2)
                btitle = str(title)
                buth = str(auth)
                btime = str(wtime)
                excel_sheet.append([jj, btitle, brticle, brticle2, buth, btime])
                excel_sheet.append(['공백'])
                print(jj+'번째 실행')
                j = j+1

                # time.sleep(1)

                ## e = input("종료하려면 x를 누르세요 : ")

                ## if e == 'x':
                ##    break

            ## source = requests.get(url, verify=False)

            ## article = soap2.select_one('#bo_v_con')
            ## article2 = soap2.select_one('#bo_v_atc')

            time.sleep(1)
            # 글내용 bo_v_con

            excel_file.save(dir)

            time.sleep(1)
            print("액셀파일 생성이완료됨")
            d = input("종료하려면 x를 누르세요 : ")
            if d == 'x':
                break

        except Exception as e:
            print(e)
            print(traceback.format_exc())
            print('예기치 못한 오류가 발생했습니다.')
            print('3초뒤 다시 시작합니다')
            time.sleep(3)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) #경고무시창
requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += 'HIGH:!DH:!aNULL'
try:
    requests.packages.urllib3.contrib.pyopenssl.DEFAULT_SSL_CIPHER_LIST += 'HIGH:!DH:!aNULL'
except AttributeError:
    # no pyopenssl support used / needed / available
    pass

# 일단 복붙은 했는데 도저히 무슨원린지 모르겠다... 이건 verify 문제도 아닌가본데? 참고 : https://stackoverflow.com/questions/38015537/python-requests-exceptions-sslerror-dh-key-too-small

void()


# 거의 모든경우 이렇게 진행
# 페이지번호도 따로 읽어야 한다. 페이지의 모든경우를 다읽었을 경우에도 다음페이지로 넘어가는 코드가 있어야한다. 단 번호만으로도 가능한경우엔 상관없다.
# 개별 번호 역시 순차적으로 되지 않기 때문에 200인 경우에만 시도하도록 짜야한다.
# if 200 = append 글내용 else = append error 404