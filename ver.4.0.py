# ver.4.0
# mysql 연동



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

# 전역 변수
orders = pd.DataFrame({}) # orders
product = pd.DataFrame({}) # product
bottles = pd.DataFrame({'병 종류':['퓨어250ml','퓨어500ml','퓨어1L','버진250ml','버진500ml','미녀플랜아보카도오일250ml',
                                '올리브250ml','톡톡300ml(1box)','톡톡100ml','톡톡10ml','세븐화이바','키토썸MCT250ml',
                                '키토썸아보카도오일250ml','즐거운유아보카도오일500ml'],'수량':[0,0,0,0,0,0,0,0,0,0,0,0,0,0]})
set1 = None # 세트상품명
# bottles
numb = None
product1,product2,product3,product4,product5 = None, None, None, None, None
num1,num2,num3,num4,num5 = 0,0,0,0,0
setnn = None


class backend :
    # 주문서 정리
    def orders(self,path,name,uname,pw,dbn):
        global orders
        df = pd.read_excel(path)
        df['택배비'] = np.where(df["택배비"] == "무료", "0", df['택배비'])  # 택배비 무료 0처리
        df_client = df["거래처"].dropna(axis=0, how='any')
        df_product = df["상품명"].dropna(axis=0, how='any')
        df_num = df["수량"].dropna(axis=0, how='any')
        df_countprice = df["단가"].dropna(axis=0, how='any')
        df_postprice = df["택배비"].dropna(axis=0, how='any')
        temp = pd.concat([df_client, df_product, df_num, df_countprice, df_postprice], axis=1)
        temp_ = temp.astype({"수량":'int','단가':'int','택배비':'int'})
        orders = temp_

        # MySQL Connect
        start = ConnectMySQL(self,1,name,uname,pw,dbn)
        start.ShowModal()
    # 상품 총 개수 정리
    def product(self,name,uname,pw,dbn):
        # MySQL Connect
        start = ConnectMySQL(self,2,name,uname,pw,dbn)
        start.ShowModal()
    # 병개수 정리
    def bottles(self,name,uname,pw,dbn):
        # MySQL Connect
        start = ConnectMySQL(self,3,name,uname,pw,dbn)
        start.ShowModal()
    # 업체별 정리
    def client(self,name,uname,pw,dbn):
        # MySQL Connect
        start = ConnectMySQL(self, 4,name,uname,pw,dbn)
        start.ShowModal()
    # 세트 상품 분류
    def setproduct(self,name,uname,pw,dbn):
        # MySQL Connect
        start = ConnectMySQL(self, 5, name,uname,pw,dbn)
        start.ShowModal()


    def bottlecount(self,df_):
        global bottles
        temp = df_
        p_temp = temp['상품명']
        plist = p_temp.tolist()
        setp = set(plist)
        rlist = list(setp) # 상품명 추출
        blist = ['퓨어250ml','퓨어500ml','퓨어1L','버진250ml','버진500ml','미녀플랜아보카도오일250ml',
                '올리브250ml','톡톡300ml(1box)','톡톡100ml','톡톡10ml','세븐화이바','키토썸MCT250ml',
                '키토썸아보카도오일250ml','즐거운유아보카도오일500ml']
        for Np in rlist :
            start = PickDialog(self,Np)
            start.ShowModal()
            global numb
            global product1, product2, product3, product4, product5
            global num1, num2, num3, num4, num5
            global setnn # 세트 개수
            if numb == 1:
                count = 0
                for p in blist :
                    if p == product1:
                        bottles['수량'][count]+= int(num1)
                    else :
                        count+=1

            elif numb ==2:
                count = 0
                for p in blist:
                    if p == product1:
                        bottles['수량'][count]+=int(num1)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product2:
                        bottles['수량'][count]+=int(num2)
                    else:
                        count += 1

            elif numb == 3:
                count = 0
                for p in blist:
                    if p == product1:
                        bottles['수량'][count]+=int(num1)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product2:
                        bottles['수량'][count]+=int(num2)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product3:
                        bottles['수량'][count]+=int(num3)
                    else:
                        count += 1

            elif numb == 4:
                count = 0
                for p in blist:
                    if p == product1:
                        bottles['수량'][count]+=int(num1)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product2:
                        bottles['수량'][count]+=int(num2)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product3:
                        bottles['수량'][count]+=int(num3)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product4:
                        bottles['수량'][count]+=int(num4)
                    else:
                        count += 1
            elif numb == 5:
                count = 0
                for p in blist:
                    if p == product1:
                        bottles['수량'][count]+=int(num1)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product2:
                        bottles['수량'][count]+=int(num2)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product3:
                        bottles['수량'][count]+=int(num3)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product4:
                        bottles['수량'][count]+=int(num4)
                    else:
                        count += 1
                count = 0
                for p in blist:
                    if p == product5:
                        bottles['수량'][count]+=int(num5)
                    else:
                        count += 1




class PickDialog(wx.Dialog):
    def __init__(self,parent,str1):
        wx.Dialog.__init__(self, parent = None, title='상품종류개수')
        self.SetSize(300, 300)  # 사이즈 설정
        sizer = wx.BoxSizer(wx.VERTICAL) # 큰 틀 사이저
        self.str1 = str1
        nlist = ['1개','2개','3개','4개','5개']

        # 위젯 생성
        message = wx.StaticText(self, label = self.str1 + ' <--상품의 종류의 개수')
        self.combo1 = wx.ComboBox(self, choices = nlist) # 종류 선택
        self.okbtn = wx.Button(self, label="ok")
        self.stbtn = wx.Button(self, label="세트일때")

        # 바인딩
        self.okbtn.Bind(wx.EVT_BUTTON, self.okBtn)
        self.stbtn.Bind(wx.EVT_BUTTON, self.setBtn)

        # 수평 사이저 셋팅
        wsizer = wx.BoxSizer(wx.HORIZONTAL)
        wsizer.Add(self.combo1,0,wx.ALIGN_CENTER_VERTICAL | wx.LEFT,10)
        ocsizer = wx.BoxSizer(wx.HORIZONTAL)
        ocsizer.Add(self.okbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ocsizer.Add(self.stbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 10)

        # 수직 사이저 셋팅
        sizer.Add(message, 0, wx.ALIGN_CENTER | wx.TOP , 10)
        sizer.Add(wsizer,0,wx.ALIGN_CENTER | wx.TOP,10)
        sizer.Add(ocsizer,0,wx.ALIGN_CENTER | wx.TOP|wx.BOTTOM,10)
        self.SetSizer(sizer)


    def okBtn(self,event):
        global numb
        global product1, product2, product3, product4, product5
        global num1, num2, num3, num4, num5
        selc = None
        selc = int(self.combo1.GetValue().replace("개",""))
        if selc == 1:
            numb = selc
            start = GetDialog(self, selc)
            ddd = start.ShowModal()
            product1, num1 = start.getvalue(selc)
            start.Destroy()
            self.Destroy()

        elif selc == 2:
            numb = selc
            start = GetDialog(self, selc)
            ddd = start.ShowModal()
            product1, num1, product2, num2 = start.getvalue(selc)
            start.Destroy()
            self.Destroy()

        elif selc == 3:
            numb = selc
            start = GetDialog(self, selc)
            ddd = start.ShowModal()
            product1, num1, product2, num2, product3, num3 = start.getvalue(selc)
            start.Destroy()
            self.Destroy()

        elif selc == 4:
            numb = selc
            start = GetDialog(self, selc)
            ddd = start.ShowModal()
            product1, num1, product2, num2, product3, num3, product4, num4 = start.getvalue(selc)
            start.Destroy()
            self.Destroy()

        elif selc == 5:
            numb = selc
            start = GetDialog(self, selc)
            ddd = start.ShowModal()
            product1, num1, product2, num2, product3, num3, product4, num4, product5, num5 = start.getvalue(selc)
            start.Destroy()
            self.Destroy()

    def setBtn(self,event):
        self.Destroy()


class GetDialog(wx.Dialog):
    def __init__(self,parent,num):
        wx.Dialog.__init__(self, parent = None ,title = "입력",size = (300,300))

        # 상품 종류 리스트
        plist = ['퓨어250ml', '퓨어500ml', '퓨어1L', '버진250ml', '버진500ml', '미녀플랜아보카도오일250ml',
                 '올리브250ml', '톡톡300ml(1box)', '톡톡100ml', '톡톡10ml', '세븐화이바',
                 '키토썸MCT250ml', '키토썸아보카도오일250ml', '즐거운유아보카도오일500ml']
        # OK, CANCEL buttons
        _btns_sizer = wx.BoxSizer(wx.HORIZONTAL)
        okbtnSizer = self.CreateStdDialogButtonSizer(wx.OK)
        canbtnSizer = self.CreateStdDialogButtonSizer(wx.CANCEL)
        _btns_sizer.Add(okbtnSizer, 1, wx.ALIGN_CENTER_VERTICAL, 0)
        _btns_sizer.Add(canbtnSizer, 1, wx.ALIGN_CENTER_VERTICAL, 0)

        # 사이저 생성
        sizer = wx.BoxSizer(wx.VERTICAL)
        if num == 1 :
        # 입력 개수에 따라 위젯이름 셋팅
            self.combo1 = wx.ComboBox(self, choices = plist) # 종류 선택
            self.num1 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box1 = wx.BoxSizer(wx.HORIZONTAL)
            box1.Add(self.combo1, 0, flag = wx.ALIGN_LEFT)
            box1.Add(self.num1, 0, flag=wx.ALIGN_LEFT)
            sizer.Add(box1,0,border = 10,flag = wx.TOP)
        elif num == 2 :
            self.combo1 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num1 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box1 = wx.BoxSizer(wx.HORIZONTAL)
            box1.Add(self.combo1, 0, flag=wx.ALIGN_LEFT)
            box1.Add(self.num1, 0, flag=wx.ALIGN_LEFT)
            self.combo2 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num2 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box2 = wx.BoxSizer(wx.HORIZONTAL)
            box2.Add(self.combo2, 0, flag=wx.ALIGN_LEFT)
            box2.Add(self.num2, 0, flag=wx.ALIGN_LEFT)
            sizer.Add(box1, 0,border = 10, flag=wx.TOP)
            sizer.Add(box2, 0,border = 10, flag=wx.TOP)
        elif num == 3 :
            self.combo1 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num1 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box1 = wx.BoxSizer(wx.HORIZONTAL)
            box1.Add(self.combo1, 0, flag=wx.ALIGN_LEFT)
            box1.Add(self.num1, 0, flag=wx.ALIGN_LEFT)
            self.combo2 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num2 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box2 = wx.BoxSizer(wx.HORIZONTAL)
            box2.Add(self.combo2, 0, flag=wx.ALIGN_LEFT)
            box2.Add(self.num2, 0, flag=wx.ALIGN_LEFT)
            self.combo3 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num3 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box3 = wx.BoxSizer(wx.HORIZONTAL)
            box3.Add(self.combo3, 0, flag=wx.ALIGN_LEFT)
            box3.Add(self.num3, 0, flag=wx.ALIGN_LEFT)
            sizer.Add(box1, 0,border = 10, flag=wx.TOP)
            sizer.Add(box2, 0,border = 10, flag=wx.TOP)
            sizer.Add(box3, 0, border=10, flag=wx.TOP)
        elif num == 4 :
            self.combo1 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num1 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box1 = wx.BoxSizer(wx.HORIZONTAL)
            box1.Add(self.combo1, 0, flag=wx.ALIGN_LEFT)
            box1.Add(self.num1, 0, flag=wx.ALIGN_LEFT)
            self.combo2 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num2 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box2 = wx.BoxSizer(wx.HORIZONTAL)
            box2.Add(self.combo2, 0, flag=wx.ALIGN_LEFT)
            box2.Add(self.num2, 0, flag=wx.ALIGN_LEFT)
            self.combo3 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num3 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box3 = wx.BoxSizer(wx.HORIZONTAL)
            box3.Add(self.combo3, 0, flag=wx.ALIGN_LEFT)
            box3.Add(self.num3, 0, flag=wx.ALIGN_LEFT)
            self.combo4 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num4 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box4 = wx.BoxSizer(wx.HORIZONTAL)
            box4.Add(self.combo4, 0, flag=wx.ALIGN_LEFT)
            box4.Add(self.num4, 0, flag=wx.ALIGN_LEFT)
            sizer.Add(box1, 0,border = 10, flag=wx.TOP)
            sizer.Add(box2, 0,border = 10, flag=wx.TOP)
            sizer.Add(box3, 0, border=10, flag=wx.TOP)
            sizer.Add(box4, 0, border=10, flag=wx.TOP)
        elif num == 5 :
            self.combo1 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num1 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box1 = wx.BoxSizer(wx.HORIZONTAL)
            box1.Add(self.combo1, 0, flag=wx.ALIGN_LEFT)
            box1.Add(self.num1, 0, flag=wx.ALIGN_LEFT)
            self.combo2 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num2 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box2 = wx.BoxSizer(wx.HORIZONTAL)
            box2.Add(self.combo2, 0, flag=wx.ALIGN_LEFT)
            box2.Add(self.num2, 0, flag=wx.ALIGN_LEFT)
            self.combo3 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num3 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box3 = wx.BoxSizer(wx.HORIZONTAL)
            box3.Add(self.combo3, 0, flag=wx.ALIGN_LEFT)
            box3.Add(self.num3, 0, flag=wx.ALIGN_LEFT)
            self.combo4 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num4 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box4 = wx.BoxSizer(wx.HORIZONTAL)
            box4.Add(self.combo4, 0, flag=wx.ALIGN_LEFT)
            box4.Add(self.num4, 0, flag=wx.ALIGN_LEFT)
            self.combo5 = wx.ComboBox(self, choices=plist)  # 종류 선택
            self.num5 = wx.TextCtrl(parent=self, name='수량')  # 수량
            box5 = wx.BoxSizer(wx.HORIZONTAL)
            box5.Add(self.combo5, 0, flag=wx.ALIGN_LEFT)
            box5.Add(self.num5, 0, flag=wx.ALIGN_LEFT)
            sizer.Add(box1, 0,border = 10, flag=wx.TOP)
            sizer.Add(box2, 0,border = 10, flag=wx.TOP)
            sizer.Add(box3, 0, border=10, flag=wx.TOP)
            sizer.Add(box4, 0, border=10, flag=wx.TOP)
            sizer.Add(box5, 0, border=10, flag=wx.TOP)

        sizer.Add(wx.StaticLine(self, size=(250, 2)), 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)
        sizer.Add(_btns_sizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)
        self.SetSizer(sizer)

    def getvalue(self,num):
        if num == 1 :
            return self.combo1.GetValue() ,self.num1.GetValue()
        elif num == 2 :
            return self.combo1.GetValue(), self.num1.GetValue(),self.combo2.GetValue(), self.num2.GetValue()
        elif num == 3 :
            return self.combo1.GetValue(), self.num1.GetValue(),self.combo2.GetValue(), self.num2.GetValue(),\
                   self.combo3.GetValue(), self.num3.GetValue()
        elif num == 4 :
            return self.combo1.GetValue(), self.num1.GetValue(),self.combo2.GetValue(), self.num2.GetValue(),\
                   self.combo3.GetValue(), self.num3.GetValue(),self.combo4.GetValue(), self.num4.GetValue()
        elif num == 5 :
            return self.combo1.GetValue(), self.num1.GetValue(),self.combo2.GetValue(), self.num2.GetValue(),\
                   self.combo3.GetValue(), self.num3.GetValue(),self.combo4.GetValue(), self.num4.GetValue(),\
                   self.combo5.GetValue(), self.num5.GetValue()



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
    def __init__(self,parent,control,name,uname,pw,dbn):
        wx.Dialog.__init__(self, parent=None, title='Connecting', size=(300, 300))
        self.text = wx.StaticText(self,label = '연결중....')
        # 접속 정보
        HOSTNAME = 'localhost'
        PROT = 3306
        USERNAME = None
        PASSWORD = None
        DATABASE = None
        CHARSET1 = 'utf8'  # mysql 에서 사용할 셋
        CHARSET2 = 'utf-8'  # 파이썬에서 사용할 셋

        # 정보 전달
        USERNAME = uname
        PASSWORD = pw
        DATABASE = dbn

        # 형변환
        USERNAME = str(USERNAME)
        PASSWORD = str(PASSWORD)
        DATABASE = str(DATABASE)

        #정보 셋팅
        con_str_fmt = 'mysql+mysqldb://{0}:{1}@{2}:{3}/{4}?charset={5}' # 포멧팅
        con_str = con_str_fmt.format(USERNAME, PASSWORD, HOSTNAME, PROT, DATABASE, CHARSET1)

        #db접속
        try :
            # db 연결
            engine = create_engine(con_str, encoding=CHARSET2)
            conn = engine.connect()
            if control == 1 :
                global orders
                try:
                    orders.to_sql(name=name+'orders', con=conn, if_exists='append', index=None)
                    wx.MessageBox("완료!!", "db접속", wx.OK)
                except Exception as ex:
                    t1 = str(ex)
                    wx.MessageBox("실패했습니다ㅠㅜ" + t1, "db접속", wx.OK)
                conn.close()
                self.Destroy()
            elif control == 2 :
                global product
                try:
                    product = pd.read_sql('select 상품명,sum(수량) as 수량 from '+name+'orders group by 상품명', con=conn)
                    product_ = product.astype({'수량': 'int'}) # 형변환
                    product.to_sql(name=name+'product', con=conn, if_exists='append', index=None)
                    wx.MessageBox("완료!!", "db접속", wx.OK)
                except Exception as ex:
                    t1 = str(ex)
                    wx.MessageBox("실패했습니다ㅠㅜ" + t1, "db접속", wx.OK)
                conn.close()
                self.Destroy()
            elif control == 3 :
                global bottles
                try:
                    start = backend()
                    temp = pd.read_sql('select 상품명,sum(수량) as 수량 from '+name+'orders group by 상품명',con = conn)
                    temp_ = temp.astype({'수량': 'int'})  # 형변환

                    start.bottlecount(temp_)
                    # 날짜
                    l = len(bottles)  # 세로 길이
                    today = []  # 리스트
                    for i in range(l):
                        today.append(name)
                    bottles.insert(0, 'date', today)  # 날짜칸 생성
                    bottles.to_sql(name = 'bottles',con = conn,if_exists='replace', index=None)
                    wx.MessageBox("완료!!", "db접속", wx.OK)
                except Exception as ex:
                    t1 = str(ex)
                    wx.MessageBox("실패했습니다ㅠㅜ" + t1, "db접속", wx.OK)
                conn.close()
                self.Destroy()
            elif control == 4:
                try:
                    client = pd.read_sql('select 거래처,상품명,sum(수량) as 수량 from '+name+'orders group by 상품명 order by 거래처',con = conn)
                    client_ = client.astype({'수량': 'int'})  # 형변환
                    client_.to_sql(name = name+'client', con=conn, if_exists='append', index=None)
                    wx.MessageBox("완료!!", "db접속", wx.OK)
                except Exception as ex:
                    t1 = str(ex)
                    wx.MessageBox("실패했습니다ㅠㅜ" + t1, "db접속", wx.OK)
                conn.close()
                self.Destroy()
            elif control == 5:
                global set1
                try:
                    set2 = '\''+set1+'\''
                    sett = pd.read_sql('select 거래처,상품명,sum(수량) as 수량 from ' + name + 'client where 상품명 = '+ set2,con = conn)
                    sett_ = sett.astype({'수량': 'int'})  # 형변환

                    # 날짜
                    l = len(sett_)  # 세로 길이
                    today = []  # 리스트
                    for i in range(l):
                        today.append(name)
                    sett_.insert(0, 'date', today)  # 날짜칸 생성

                    sett_.to_sql(name = 'setproducts',con =conn,if_exists='append',index=None)
                    wx.MessageBox("완료!!", "db접속", wx.OK)
                except Exception as ex:
                    t1 = str(ex)
                    wx.MessageBox("실패했습니다ㅠㅜ" + t1, "db접속", wx.OK)
                conn.close()
                self.Destroy()


        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("오류!! : "+t1,'connect',wx.OK)


class Pickcontrol(wx.Dialog):
    def __init__(self,parent,filepath,filename):
        wx.Dialog.__init__(self,parent = None,title = 'Pick',size = (300,300))
        # 멤버 초기화
        self.filepath = filepath
        self.filename = filename

        # 위젯
        self.btn1 = wx.Button(self,label = '주문서')
        self.btn2 = wx.Button(self,label = '상품개수')
        self.btn3 = wx.Button(self,label = '병 개수')
        self.btn4 = wx.Button(self, label= '거래처별 정리')
        self.btn5 = wx.Button(self, label= '세트상품')
        self.exbten = wx.Button(self,label ='나가기')

        # 사이저
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.btn1, 0,  wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(self.btn2, 0,  wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(self.btn3, 0,  wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(self.btn4, 0,  wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(self.btn5, 0,  wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(wx.StaticLine(self, size=(250, 2)), 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 30)
        sizer.Add(self.exbten,0,wx.ALIGN_CENTER | wx.BOTTOM,10)
        self.SetSizer(sizer)
        
        # 바인딩
        self.btn1.Bind(wx.EVT_BUTTON, self.orders)
        self.btn2.Bind(wx.EVT_BUTTON, self.product)
        self.btn3.Bind(wx.EVT_BUTTON, self.bottles)
        self.btn4.Bind(wx.EVT_BUTTON, self.client)
        self.btn5.Bind(wx.EVT_BUTTON, self.setproduct)
        self.exbten.Bind(wx.EVT_BUTTON, self.exit)

        # 정보가져오기
        startd = Getinfo(self)
        startd.ShowModal()
        self.uname,self.pw,self.dbn = startd.getvalue()  # 접속정보 가져오기



    # 이벤트 셋
    def orders(self,event):
        start = backend()
        start.orders(self.filepath,self.filename,self.uname,self.pw,self.dbn)
    def product(self,event):
        start = backend()
        start.product(self.filename,self.uname,self.pw,self.dbn)
    def bottles(self,event):
        start = backend()
        start.bottles(self.filename,self.uname,self.pw,self.dbn)
    def client(self,event):
        start = backend()
        start.client(self.filename,self.uname,self.pw,self.dbn)
    def setproduct(self,event):
        global set1
        text = wx.TextEntryDialog(self,'세트상품명 입력')
        text.ShowModal()
        set1 = text.GetValue() # 세트상품명 추출
        start = backend()
        start.setproduct(self.filename,self.uname,self.pw,self.dbn)
    def exit(self,event):
        self.Destroy()

        



# 도움말 다이얼로그
class Helpdialog(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, parent=None, title='도움말',size = (600,600))
        sizer = wx.BoxSizer(wx.VERTICAL)
        message = wx.StaticText(self, label='★도움말★')
        help = wx.StaticText(self, label = """1. 모든 입력은 파일명, 엑셀에 있는 데이터값 그대로 적어주세요.\n
2. 상품이 없다고 뜨는 상품들은 이야기해주시면 추가해서 프로그램 업데이트 해드리겠습니다.\n
3. 에로사항이나 문의사항 있으시면 010-2094-7805 정광원 으로 카톡아니면 문자 주세요!!\n
4. 프로그램과 저장되는 폴더들은 같은 폴더안에 둬주세요(그래야 오류안뜸!!)\n
5. 주문서 엑셀 형식은 test.xlsx 파일과 같은 포멧으로 해주세요 그래야 오류 안뜸\n\n\n\n\n\n\n
made by Pupba.J""")
        sizer.Add(message, 10, wx.ALIGN_CENTER | wx.TOP, 5)
        sizer.Add(help, 1, wx.ALIGN_CENTER | wx.TOP, 5)

class Window(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, parent=None, title = 'Ver.4.0')
        self.SetSize(300,300) # 사이즈 설정

        #위젯 생성
        self.btn1 = wx.Button(self)
        self.btn1.SetLabel("주문서처리")
        self.btn2 = wx.Button(self)
        self.btn2.SetLabel("도움말")
        self.btn3 = wx.Button(self)
        self.btn3.SetLabel("끝내기")

        #바인딩
        self.btn1.Bind(wx.EVT_BUTTON, self.onClickbtn1)
        self.btn2.Bind(wx.EVT_BUTTON, self.onClickbtn2)
        self.btn3.Bind(wx.EVT_BUTTON, self.onClickbtn3)

        # 사이저
        self.gbox = wx.BoxSizer(wx.VERTICAL)
        self.gbox.Add(self.btn1, 2, wx.EXPAND, 10)
        self.gbox.Add(self.btn2, 1, wx.EXPAND, 10)
        self.gbox.Add(self.btn3, 1, wx.EXPAND, 10)
        self.SetSizer(self.gbox) # 셋 사이저

        # 버튼 이벤트 선언 및 바인딩
    def onClickbtn1(self,event):
        try:
            p = os.getcwd()
            fileD = wx.FileDialog(self,'파일선택',p,'','*.*') #파일다이얼로그 셋 : 현재 작업 위치에서 모든종류 파일
            fileD.ShowModal()
            filpath = fileD.GetPath() # 경로반환 str
            filename = fileD.GetFilename().replace('.xlsx','') # 파일이름

            # 다이얼로그 실행
            start = Pickcontrol(self,filpath,filename)
            start.ShowModal()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅜ"+t1, "주문서처리", wx.OK)
    def onClickbtn2(self,event):
        try:
            helpd = Helpdialog()
            helpd.ShowModal()
            helpd.Destroy()
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅜ"+t1, "도움말", wx.OK)
    def onClickbtn3(self,event):
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