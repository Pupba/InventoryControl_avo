# DB 읽어와서 엑셀로
#import
import wx
import os
import pandas as pd
import numpy as np
import pymysql
pymysql.install_as_MySQLdb() # pymysql를 이용해 mysql연동객체 설치
import MySQLdb # 임포트
from sqlalchemy import create_engine
from matplotlib import pyplot


path = os.getcwd() # 작업위치


class Pick(wx.Dialog):
    def __init__(self,parent):
        wx.Dialog.__init__(self, parent=None, title='Pick', size=(300, 300))

        login = Getinfo(self) # 로그인
        login.ShowModal()
        self.name,self.pw,self.dbn = login.getvalue()

        # 위젯
        self.btn1 = wx.Button(self,label = '주문서')
        self.btn2 = wx.Button(self,label = '상품개수')
        self.btn3 = wx.Button(self, label = '거래처별 정리')
        self.btn4 = wx.Button(self, label = '병 개수')
        self.btn5 = wx.Button(self, label= '세트 정리')

        # 사이저
        self.gbox = wx.BoxSizer(wx.VERTICAL)
        self.gbox.Add(self.btn1, 1, wx.EXPAND, 10)
        self.gbox.Add(self.btn2, 1, wx.EXPAND, 10)
        self.gbox.Add(self.btn3, 1, wx.EXPAND, 10)
        self.gbox.Add(self.btn4, 1, wx.EXPAND, 10)
        self.gbox.Add(self.btn5, 1, wx.EXPAND, 10)
        self.SetSizer(self.gbox)  # 셋 사이저

        # 바인딩
        self.btn1.Bind(wx.EVT_BUTTON, self.onclick1)
        self.btn2.Bind(wx.EVT_BUTTON, self.onclick2)
        self.btn3.Bind(wx.EVT_BUTTON, self.onclick3)
        self.btn4.Bind(wx.EVT_BUTTON, self.onclick4)
        self.btn5.Bind(wx.EVT_BUTTON, self.onclick5)
        
        
    # 이벤트 핸들러 셋
    def onclick1(self,event):
        try :
            start = ConnectMySQL(self,1,self.name,self.pw,self.dbn)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)

    def onclick2(self,event):
        try :
            start = ConnectMySQL(self,2,self.name,self.pw,self.dbn)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)

    def onclick3(self,event):
        try :
            start = ConnectMySQL(self,3,self.name,self.pw,self.dbn)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)

    def onclick4(self,event):
        try :
            start = ConnectMySQL(self,4,self.name,self.pw,self.dbn)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)

    def onclick5(self,event):
        try :
            start = ConnectMySQL(self,5,self.name,self.pw,self.dbn)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)

# db접속 정보 가져오기
class Getinfo(wx.Dialog):
    def __init__(self,parent):
        wx.Dialog.__init__(self,parent = None,title = 'Info',size = (300,300))

        # 위젯
        self.text1 = wx.StaticText(self,label = 'Username')
        self.text2 = wx.StaticText(self,label = 'Password')
        self.text3 = wx.StaticText(self,label = 'DBname')
        self.name = wx.TextCtrl(parent = self,name = 'Username')
        self.pw = wx.TextCtrl(parent = self,name = 'Password')
        self.dbn = wx.TextCtrl(parent = self,name = 'DBname')

        # OK, CANCEL buttons
        _btns_sizer = wx.BoxSizer(wx.HORIZONTAL)
        okbtnSizer = self.CreateStdDialogButtonSizer(wx.OK)
        canbtnSizer = self.CreateStdDialogButtonSizer(wx.CANCEL)
        _btns_sizer.Add(okbtnSizer, 1, wx.ALIGN_CENTER_VERTICAL, 0)
        _btns_sizer.Add(canbtnSizer, 1, wx.ALIGN_CENTER_VERTICAL, 0)

        #사이저 셋팅
        hsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        hsizer1.Add(self.text1,0,wx.ALIGN_CENTER_VERTICAL | wx.LEFT,10)
        hsizer1.Add(self.name, 0, wx.ALIGN_CENTER_VERTICAL |wx.RIGHT| wx.LEFT, 10)
        hsizer2 = wx.BoxSizer(wx.HORIZONTAL)
        hsizer2.Add(self.text2, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        hsizer2.Add(self.pw, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 10)
        hsizer3 = wx.BoxSizer(wx.HORIZONTAL)
        hsizer3.Add(self.text3, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        hsizer3.Add(self.dbn, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 10)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(wx.StaticText(self,label = 'DB LOGIN'),0,wx.ALIGN_CENTER | wx.TOP,10)
        sizer.Add(hsizer1, 0, wx.ALIGN_CENTER | wx.TOP,10)
        sizer.Add(hsizer2, 0, wx.ALIGN_CENTER | wx.TOP,10)
        sizer.Add(hsizer3, 0, wx.ALIGN_CENTER | wx.TOP,10)
        sizer.Add(wx.StaticLine(self, size=(250, 2)), 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)
        sizer.Add(_btns_sizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)
        self.SetSizer(sizer)

        #값 읽어오기

    def getvalue(self):
        return self.name.GetValue(),self.pw.GetValue(),self.dbn.GetValue()

class ConnectMySQL(wx.Dialog):
    def __init__(self,parent,num,name,pw,dbn):
        wx.Dialog.__init__(self, parent=None, title='Connecting', size=(300, 300))
        # 접속 정보
        HOSTNAME = 'localhost'
        PROT = 3306
        USERNAME = name
        PASSWORD = pw
        DATABASE = dbn
        CHARSET1 = 'utf8'  # mysql 에서 사용할 셋
        CHARSET2 = 'utf-8' # 파이썬에서 사용할 셋


        # 로그인 정보 전달
        USERNAME = name # 이름
        PASSWORD = pw # 비밀번호
        DATABASE = dbn # DB명


        # 문자열 변환
        USERNAME = str(USERNAME)
        PASSWORD = str(PASSWORD)
        DATABASE = str(DATABASE)

        #정보 셋팅
        con_str_fmt = 'mysql+mysqldb://{0}:{1}@{2}:{3}/{4}?charset={5}' # 포멧팅
        con_str = con_str_fmt.format(USERNAME, PASSWORD, HOSTNAME, PROT, DATABASE, CHARSET1)

        #db접속
        try :
            global path
            # db 연결
            engine = create_engine(con_str, encoding=CHARSET2)
            conn = engine.connect()
            if num == 1 : #주문서
                p_name = wx.TextEntryDialog(self, '날짜 입력!')  # 다이얼로그 생성
                p_name.ShowModal()
                pname = p_name.GetValue()  # 상품명 추출
                qury = 'select * from '+pname+'orders'
                df = pd.read_sql(qury,con=conn)
                fdf = df.astype({'수량':'int','단가':'int','택배비':'int'}) # 형변환
                fdf.to_excel(path+"/"+pname+'주문서.xlsx',index=False)
                wx.MessageBox("완료!!", 'connect', wx.OK)
                self.Destroy()
            if num == 2 : #상품개수
                p_name = wx.TextEntryDialog(self, '날짜 입력!')  # 다이얼로그 생성
                p_name.ShowModal()
                pname = p_name.GetValue()  # 상품명 추출
                qury = 'select * from '+pname+'product'
                df = pd.read_sql(qury,con=conn)
                fdf = df.astype({'수량':'int'}) # 형변환
                fdf.to_excel(path+"/"+pname+'상품개수.xlsx',index=False)
                wx.MessageBox("완료!!", 'connect', wx.OK)
                self.Destroy()
            if num == 3 : #거래처별 정리
                p_name = wx.TextEntryDialog(self, '날짜 입력!')  # 다이얼로그 생성
                p_name.ShowModal()
                pname = p_name.GetValue()  # 상품명 추출
                qury = 'select * from '+pname+'client'
                df = pd.read_sql(qury,con=conn)
                fdf = df.astype({'수량':'int'}) # 형변환
                fdf.to_excel(path+"/"+pname+'거래처별정리.xlsx',index=False)
                wx.MessageBox("완료!!", 'connect', wx.OK)
                self.Destroy()
            if num == 4 : #병개수
                qury = 'select * from bottles'
                df = pd.read_sql(qury,con=conn)
                fdf = df.astype({'수량':'int'}) # 형변환
                fdf.to_excel(path+'/병개수.xlsx',index=False)
                wx.MessageBox("완료!!", 'connect', wx.OK)
                self.Destroy()
            if num == 5 : #세트 정리
                qury = 'select * from setproducts'
                df = pd.read_sql(qury,con=conn)
                fdf = df.astype({'수량':'int'}) # 형변환
                fdf.to_excel(path+'/세트상품.xlsx',index=False)
                wx.MessageBox("완료!!", 'connect', wx.OK)
                self.Destroy()


        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)






class Window(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, parent=None, title = 'Ver.4.0')
        self.SetSize(300,300) # 사이즈 설정

        #위젯 생성
        self.btn1 = wx.Button(self,label = 'DB To Execl')
        self.btn2 = wx.Button(self,label = '끝내기')

        #바인딩
        self.btn1.Bind(wx.EVT_BUTTON, self.onClickbtn1)
        self.btn2.Bind(wx.EVT_BUTTON, self.onClickbtn2)

        # 사이저
        self.gbox = wx.BoxSizer(wx.VERTICAL)
        self.gbox.Add(self.btn1, 2, wx.EXPAND, 10)
        self.gbox.Add(self.btn2, 1, wx.EXPAND, 10)
        self.SetSizer(self.gbox) # 셋 사이저

        # 버튼 이벤트 선언 및 바인딩
    def onClickbtn1(self,event):
        try:
            start = Pick(self)
            start.ShowModal()
            start.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅜ"+t1, "DB TO EXCEL", wx.OK)
    def onClickbtn2(self,event):
        try:
            wx.MessageBox("종료합니다.", "종료", wx.OK)
            wx.Exit()  # 종료
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅜ"+t1, "종료", wx.OK)



if __name__ == "__main__":
    app = wx.App()
    frame = Window()
    frame.Show()

    app.MainLoop()