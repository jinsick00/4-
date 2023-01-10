import sys
import pandas as pd
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog
# from PyQt5.QtGui import *
# from PyQt5.QtCore import *

from pathlib import Path
class myapp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()
        

    def initUI(self):
        
        self.button()
        self.setGeometry(300, 300, 600, 500) # 창의 위치, 크기
        self.setWindowTitle("4대보험계산프로그램") # 프로그램 이름
        self.show()


    def button(self):
        self.btn0 = QPushButton("프로그램 사용 방법(미구현)", self)
        self.btn0.setFixedSize(240, 40)
        self.btn0.move(30, 40)
        self.btn0.clicked.connect(self.btn0_clicked)

        self.btn1 = QPushButton("1. 건강보험 파일 선택", self)
        self.btn1.setFixedSize(240, 40)
        self.btn1.move(30, 90)
        self.btn1.clicked.connect(self.btn1_clicked)

        self.btn2 = QPushButton("2. 국민연금 파일 선택", self)
        self.btn2.setFixedSize(240, 40)
        self.btn2.move(30, 140)
        self.btn2.clicked.connect(self.btn2_clicked)

        self.btn3 = QPushButton("3. 고용보험 파일 선택", self)
        self.btn3.setFixedSize(240, 40)
        self.btn3.move(30, 190)
        self.btn3.clicked.connect(self.btn3_clicked)

        self.btn4 = QPushButton("4. 산재보험 파일 선택", self)
        self.btn4.setFixedSize(240, 40)
        self.btn4.move(30, 240)
        self.btn4.clicked.connect(self.btn4_clicked)

        self.btn5 = QPushButton("4대보험 합치기(1~4 작업 후 실행)", self)
        self.btn5.setFixedSize(240, 40)
        self.btn5.move(30, 290)
        self.btn5.clicked.connect(self.btn5_clicked)

        self.btn6 = QPushButton("결과 저장(합치기 작업 후 실행)", self)
        self.btn6.setFixedSize(240, 40)
        self.btn6.move(30, 340)
        self.btn6.clicked.connect(self.btn6_clicked)

    def btn0_clicked(self):
        pass

    def btn1_clicked(self): 
        files1 = QFileDialog.getOpenFileName(None, "건강보험 파일 선택", '.', "엑셀 파일(*.xlsx *.xls);; CSV 파일(*.csv)")[0]
        if files1: 
            #self.lineEdit_openfile.setText(files) # 이게 파일경로 열어주는 기능
            global 건강2
            global 건강기본
            건강2 = pd.read_excel(files1)
            건강2.rename(columns= {"성명" : "근로자명"} , inplace=True) # 성함 열 근로자명으로 열 이름 바꾸기
            건강2.rename(columns= {"주민등록번호" : "생년월일"} , inplace=True) # 주민번호 열 생년월일으로 열 이름 바꾸기
            건강2["생년월일"] = 건강2["생년월일"].str[: -8] # 생년월일 양식 통일(주민번호 뒷자리 삭제)
            건강2 = 건강2.set_index(["근로자명", "생년월일"]) # 인덱스 처리
            건강기본 = 건강2.iloc[: , -1 : ].copy()
            건강기본["개인부담금"] = 건강기본.가입자총납부할보험료 # 개인부담금 열 생성
            건강기본["사업주부담금"] = 건강기본.가입자총납부할보험료 # 사업주부담금 열 생성
            건강기본.drop(columns = ["가입자총납부할보험료"], inplace=True) # 불필요한 열 삭제
            
        # openfile1 = QFileDialog.getOpenFileName(self, '파일 열기', './', ("*.xlsx"))
        # filename = openfile1[0]
        # df1 = pd.read_excel(filename)
        # f = open(openfile1, pd.read_excel(openfile1,  header = 1))   # self.pd.read_excel("a",header = 1)
        # print(f)
    
    def btn2_clicked(self):
        files2 = QFileDialog.getOpenFileName(None, "국민연금 파일 선택", '.', "엑셀 파일(*.xlsx *.xls);; CSV 파일(*.csv)")[0]
        if files2:
            global 국민
            global 국민기본
            국민2 = pd.read_csv(files2, encoding = "cp949") # 국민연금 세팅
            국민2.rename(columns= {"가입자명" : "근로자명"} , inplace=True) # 성함 열 근로자명으로 열 이름 바꾸기
            국민2.rename(columns= {"주민번호" : "생년월일"} , inplace=True) # 주민번호 열 생년월일으로 열 이름 바꾸기
            국민2["생년월일"] = 국민2["생년월일"].str[: -8] # 생년월일 양식 통일(주민번호 뒷자리 삭제)
            국민기본 = 국민2.set_index(["근로자명", "생년월일"]) # 인덱스 처리
            국민기본 = 국민기본.iloc[: , -1 : ].copy()
            국민기본["개인부담금"] = 국민기본.결정보험료/2 # 개인부담금 열 생성
            국민기본["사업주부담금"] = 국민기본.결정보험료/2 # 사업주부담금 열 생성
            국민기본.drop(columns = ["결정보험료"], inplace=True) # 불필요한 열 삭제
            
        
    def btn3_clicked(self):
        files3 = QFileDialog.getOpenFileName(None, "고용보험 파일 선택", '.', "엑셀 파일(*.xlsx *.xls);; CSV 파일(*.csv)")[0]
        if files3:
            global 고용기본
            global 고용기본2
            고용기본 = pd.read_excel(files3,header = 1) # 고용보험 기본 세팅
            생년월일통일 = 고용기본["생년월일"].replace("-", "", inplace = True, regex = True) # 생년월일 양식 통일
            고용기본 = 고용기본.set_index(["근로자명", "생년월일"]) # 근로자명, 생년월일 인덱스 처리
            # 생년월일통일 = 고용기본.columns = ["근로자명", "생년월일", "근로자실업급여보험료.3", "사업주실업급여보험료.3", "사업주고안직능보험료.3"] 22개 22개로 맞춰야 대서 안댄다
            고용기본2 = 고용기본.iloc[: , -3 : ].copy()
            고용기본2.columns = ["개인부담금_고용", "사업주실업급여보험료", "사업주고안직능보험료"] # 고용보험 열 이름 변경
            고용기본2["사업주부담금"] = 고용기본2.사업주실업급여보험료 + 고용기본2.사업주고안직능보험료 # 고용보험 사업주부담금 열 생성
            고용기본2.drop(columns = ["사업주실업급여보험료", "사업주고안직능보험료"], inplace=True) # 고용보험 불필요열 삭제
            


    def btn4_clicked(self):
        files4 = QFileDialog.getOpenFileName(None, "산재보험 파일 선택", '.', "엑셀 파일(*.xlsx *.xls);; CSV 파일(*.csv)")[0]
        if files4:
            global 산업
            global 산업기본2
            산업 = pd.read_excel(files4) # 산재 기본 세팅
            주민번호통일 = 산업["생년월일"].replace("-", "", inplace = True, regex = True) # 생년월일 양식 통일
            산업기본 = 산업.set_index(["근로자명", "생년월일"]) # 근로자명, 생년월일 인덱스처리
            산업기본2 = 산업기본.iloc[: , -1 : ].copy() # 열 하나만 가져온다
            산업기본2.columns = ["사업주부담금"] # 가져온 열 이름 변경
            


    def btn5_clicked(self):
       # files5 = QFileDialog.getOpenFileName(None, "Open Excel File", '.', "(*.xlsx)")[0]
       # if files5:
        global 건강국민
        global 고용산재2
        global 건강국민고용산재
        건강국민 = 건강기본.merge(국민기본, how = "outer", on = ["근로자명", "생년월일"], suffixes= ("_건강", "_국민"), indicator= True ) # 건강보험, 국민연금 파일 합치기
        건강국민.drop(columns = ["_merge"], inplace=True) # 불필요한 열 삭제
        건강국민 = 건강국민.fillna(0) # Nan 값 0으로 교체

        고용산재2= 고용기본2.merge(산업기본2, how = "outer", on = ["근로자명", "생년월일"], suffixes= ("_고용", "_산재"), indicator= True ) # 고용보험, 산재보험 합치기
        고용산재2.drop(columns = ["_merge"], inplace=True) # 불필요한 열(merge) 삭제
        고용산재2 = 고용산재2.fillna(0)

        건강국민고용산재 = 건강국민.merge(고용산재2, how = "outer", on = ["근로자명", "생년월일"], suffixes= ("_건강", "_국민"), indicator= True ) # 4대보험 파일 합치기
        건강국민고용산재.drop(columns = ["_merge"], inplace=True) # 불필요한 열 삭제
        건강국민고용산재 = 건강국민고용산재.fillna(0) # Nan값 0으로 교체
        건강국민고용산재["사업주부담금합"] = 건강국민고용산재.사업주부담금_건강 + 건강국민고용산재.사업주부담금_국민 + 건강국민고용산재.사업주부담금_고용 + 건강국민고용산재.사업주부담금_산재 # 개인별 4대보험 사업주부담금 열 만들기
        건강국민고용산재 = 건강국민고용산재.sort_index() # 각 행을 근로자명 내림차순 배열
        건강국민고용산재.reset_index(inplace = True) # 인덱스 행 초기화
        건강국민고용산재["직종"] = "" # 직종열 생성(빈 칸임)
        건강국민고용산재 = 건강국민고용산재.reindex(columns = ['직종', '근로자명', '생년월일', '개인부담금_건강', '사업주부담금_건강', '개인부담금_국민', '사업주부담금_국민',
        '개인부담금_고용', '사업주부담금_고용', '사업주부담금_산재', '사업주부담금합']) # 열 순서 변경(직종을 맨 앞으로)
        
        
        
    def btn6_clicked(self):
        

        files6 = QFileDialog.getSaveFileName(None, "결과 저장하기", '.', "(*.xlsx)")[0]
        if files6:
            건강국민고용산재.to_excel(files6)

app = QApplication(sys.argv)
exc = myapp()
app.exec_()


# line edit: 선택한 파일 경로 표시