# ver.3.0
# 파일 선택 다이얼로그 적용
# 체크박스 적용하여 제품 종류 선택 더 쉽게
# 상품 종류 선택시 예외 적용
# 세트 상품일 때 따로 저장 요청

import os
import pandas as pd
import wx

# 전역변수 선언
order = pd.DataFrame({}) # 빈 데이터프레임(주문서)
bottle = pd.DataFrame({'퓨어250ml':[0],'퓨어500ml':[0],'퓨어1L':[0],'버진250ml':[0],'버진500ml':[0],'미녀플랜아보카도오일250ml':[0],
                       '올리브250ml':[0],'톡톡300ml(1box)':[0],'톡톡100ml':[0],'톡톡10ml':[0],'세븐화이바':[0],
                       '키토썸MCT250ml':[0],'키토썸아보카도오일250ml':[0],'즐거운유아보카도오일500ml':[0]}) # 병 종합
path = str(os.getcwd()) # 경로 설정 작업 폴더
rbottle = pd.read_excel(path+'/정리/주문병 수.xlsx') # 기본 주문병 수 불러오기
numb = None
product1,product2,product3,product4,product5 = None, None, None, None, None
num1,num2,num3,num4,num5 = None, None, None, None, None
setnn = None


# 상품 종류, 개수 가져오기
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

# 세트 상품일 때 실행 될 다이얼로그
class SetDialog(wx.Dialog):
    def __init__(self,parent):
        wx.Dialog.__init__(self, parent = None ,title = "세트입력",size = (300,300))
        sizer = wx.BoxSizer(wx.VERTICAL)  # 큰 틀 사이저
        message = wx.StaticText(self, label = '개수')
        self.num = wx.TextCtrl(parent=self, name='수량')  # 수량
        self.okbtn = wx.Button(self, label="ok")
        self.exbtn = wx.Button(self, label="exit")

        # 바인딩
        self.okbtn.Bind(wx.EVT_BUTTON, self.okBtn)
        self.exbtn.Bind(wx.EVT_BUTTON, self.exBtn)

        wsizer = wx.BoxSizer(wx.HORIZONTAL)
        wsizer.Add(self.num, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ocsizer = wx.BoxSizer(wx.HORIZONTAL)
        ocsizer.Add(self.okbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ocsizer.Add(self.exbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 10)

        sizer.Add(message, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(wsizer, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.Add(ocsizer, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)

        self.SetSizer(sizer)

    def okBtn(self,event):
        global numb
        global setnn
        self.Close(True)
        selc = None
        selc = int(self.num.GetValue())
        numb = 6
        setnn = selc # 세트 개수


    def exBtn(self,event):
        wx.CANCEL
# 상품 종류 개수 or 세트 상품 선택
class PickDialog(wx.Dialog):
    def __init__(self,parent,str1):
        wx.Dialog.__init__(self, parent = None, title='상품종류개수')
        self.SetSize(300, 300)  # 사이즈 설정
        sizer = wx.BoxSizer(wx.VERTICAL) # 큰 틀 사이저
        self.str1 = str1
        nlist = ['1개','2개','3개','4개','5개']

        # 위젯 생성
        message = wx.StaticText(self, label = str1 + '의 종류입력, 세트상품일 때 버튼을 눌러주세요')
        self.combo1 = wx.ComboBox(self, choices = nlist) # 종류 선택
        self.sbtn1 = wx.Button(self,label = "세트일때")
        self.okbtn = wx.Button(self, label="ok")
        self.exbtn = wx.Button(self, label="exit")

        # 바인딩
        self.okbtn.Bind(wx.EVT_BUTTON, self.okBtn)
        self.exbtn.Bind(wx.EVT_BUTTON, self.exBtn)
        self.sbtn1.Bind(wx.EVT_BUTTON, self.setBtn)

        # 수평 사이저 셋팅
        wsizer = wx.BoxSizer(wx.HORIZONTAL)
        wsizer.Add(self.combo1,0,wx.ALIGN_CENTER_VERTICAL | wx.LEFT,10)
        wsizer.Add(self.sbtn1,0,wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT,10)
        ocsizer = wx.BoxSizer(wx.HORIZONTAL)
        ocsizer.Add(self.okbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        ocsizer.Add(self.exbtn, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 10)

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

    def exBtn(self,event):
        self.Destroy()

    def setBtn(self,event):
        start = SetDialog(self)
        ddd = start.ShowModal()
        wx.CLOSE
        start.Destroy()
        self.Destroy()

class Backend :
    # 1. 주문서 파일 읽어오기
    def getOrder(self,str1,filename):
        global order # df 사용하겠다 선언
        global path
        temp = pd.read_excel(str1) # 주문서 읽어오기
        temp.sort_values(by = '거래처',ascending=False) # 거래처 기준으로 내림차순 정렬
        # 필요한 데이터만 뽑아서 데이터프레임 만들기
        temp_pr = temp['상품명']
        temp_nu = temp['수량']
        temp_cli = temp['거래처']
        temp_ = pd.concat([temp_cli,temp_pr,temp_nu],axis=1)
        ftemp1 = temp_.dropna(axis=0) # 결측치 제거
        ftemp2 = ftemp1.astype({'수량':'int'}) # 정수형변환
        # 거래처 별 상품 개수 파악 및 단가 계산
        na = ftemp2['거래처']
        nlist = na.tolist()
        setn = set(nlist)
        name = list(setn) # 거래처 리스트 작성

        for Nl in name :
            cn = ftemp2[ftemp2['거래처']==Nl] # 거래처 행만 추출
            pn = cn['상품명'] # 상품 데이터 추출
            plist = pn.tolist()
            setp = set(plist)
            pnm = list(setp) # 상품명 리스트 작성
            index = list(range(len(pnm))) # 인덱스 길이 생성
            nwd = pd.DataFrame([{'거래처':'','상품명':'','수량':[0]}],index=index)
            for i in range(len(pnm)) :
                nwd['거래처'][i] = Nl
                t = str(pnm[i]) # 문자열
                nwd['상품명'][i] = t
                nun = cn[cn['상품명'] == pnm[i]] # 이 상품이 들어간 행만 추출
                nun_ = nun['수량']
                sum_ = nun_.sum() #개수 총합
                nwd['수량'][i] = sum_ # 총수량 데이터프레임에 입력
                i+=1 # 반복문 제어
            order = pd.concat([order,nwd],axis=0)
            order = order.dropna() # 결측치 제거


        #2. 병 수량
        order_pname = order['상품명']
        order_num = order['수량']
        tempbottle = pd.concat([order_pname,order_num],axis=1)
        #상품명별로 병 종류 및 총 병 개수 종합
        global bottle
        pplist = order_pname.tolist()
        setpp = set(pplist)
        pnmm = list(setpp)  # 상품명 리스트 작성
        ibottle = pd.DataFrame({})  # 임시 저장 데이타 프레임
        lpname = []  # 입력받는 병 종류 리스트
        for Nd in pnmm :
            temp__ = tempbottle[tempbottle['상품명'] == Nd] # 상품별 추출 데이터프레임
            # 중복 오류 제거
            ftemp__ = pd.DataFrame([{}])
            tname = pd.DataFrame([{'상품명':Nd}])
            tnu = temp__['수량'].sum()
            ttnu = pd.DataFrame([{'수량':[0]}])
            ttnu['수량'][0] = tnu
            ftemp__ = pd.concat([tname,ttnu],axis=1)
            pname = Nd
            start1 = PickDialog(self,pname) # 클래스 선언
            start1.ShowModal()
            global numb
            global product1, product2, product3, product4, product5
            global num1, num2, num3, num4, num5
            global setnn # 세트 개수
            setdp = pd.DataFrame({},index = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])
            retnum = numb  # 개수 반환
            # 개수에 따라 리턴값이 달라짐
            if retnum == 1: # 개수 한개일 때
                num = int(num1)  # 정수형변환
                tt = pd.Series(num)
                product = str(product1)
                lpname.append(product)
                temp__ = ftemp__.astype({'수량': 'int64'})
                num_ = ftemp__['수량'][0] * tt[0]  # 같은 변수형으로 해야 오류 안남
                bottle_ = pd.DataFrame([{product: [0]}])  # 저장할 데이터프레임
                inum_ = num_.astype('int64')
                bottle_[product] = num_
                ibottle = pd.concat([bottle_, ibottle], axis=0, ignore_index=True)
                ibottle = ibottle.fillna(int(0))  # 결측치 0으로
            elif retnum == 2 :
                knum = [num1,num2]
                kproduct = [product1,product2]
                for II in range(0,retnum):
                    num = int(knum[II])  # 정수형변환
                    tt = pd.Series(knum)
                    product = str(kproduct[II])
                    lpname.append(kproduct[II])
                    temp__ = ftemp__.astype({'수량': 'int64'})
                    num_ = ftemp__['수량'][0] * tt[0]  # 같은 변수형으로 해야 오류 안남
                    bottle_ = pd.DataFrame([{product: [0]}])  # 저장할 데이터프레임
                    num_ = int(num_)
                    bottle_[product] = num_
                    ibottle = pd.concat([bottle_, ibottle], axis=0, ignore_index=True)
                    ibottle = ibottle.fillna(int(0))  # 결측치 0으로
            elif retnum == 3 :
                knum = [num1, num2,num3]
                kproduct = [product1, product2,product3]
                for II in range(0, retnum):
                    num = int(knum[II])  # 정수형변환
                    tt = pd.Series(knum)
                    product = str(kproduct[II])
                    lpname.append(kproduct[II])
                    temp__ = ftemp__.astype({'수량': 'int64'})
                    num_ = ftemp__['수량'][0] * tt[0]  # 같은 변수형으로 해야 오류 안남
                    bottle_ = pd.DataFrame([{product: [0]}])  # 저장할 데이터프레임
                    num_ = int(num_)
                    bottle_[product] = num_
                    ibottle = pd.concat([bottle_, ibottle], axis=0, ignore_index=True)
                    ibottle = ibottle.fillna(int(0))  # 결측치 0으로
            elif retnum == 4 :
                knum = [num1, num2, num3,num4]
                kproduct = [product1, product2, product3,product4]
                for II in range(0, retnum):
                    num = int(knum[II])  # 정수형변환
                    tt = pd.Series(knum)
                    product = str(kproduct[II])
                    lpname.append(kproduct[II])
                    temp__ = ftemp__.astype({'수량': 'int64'})
                    num_ = ftemp__['수량'][0] * tt[0]  # 같은 변수형으로 해야 오류 안남
                    bottle_ = pd.DataFrame([{product: [0]}])  # 저장할 데이터프레임
                    num_ = int(num_)
                    bottle_[product] = num_
                    ibottle = pd.concat([bottle_, ibottle], axis=0, ignore_index=True)
                    ibottle = ibottle.fillna(int(0))  # 결측치 0으로
            elif retnum == 5 :
                knum = [num1, num2, num3, num4,num5]
                kproduct = [product1, product2, product3, product4,product5]
                for II in range(0, retnum):
                    num = int(knum[II])  # 정수형변환
                    tt = pd.Series(knum)
                    product = str(kproduct[II])
                    lpname.append(kproduct[II])
                    temp__ = ftemp__.astype({'수량': 'int64'})
                    num_ = ftemp__['수량'][0] * tt[0]  # 같은 변수형으로 해야 오류 안남
                    bottle_ = pd.DataFrame([{product: [0]}])  # 저장할 데이터프레임
                    num_ = int(num_)
                    bottle_[product] = num_
                    ibottle = pd.concat([bottle_, ibottle], axis=0, ignore_index=True)
                    ibottle = ibottle.fillna(int(0))  # 결측치 0으로
            elif retnum == 6 :
                # 세트일 때
                num = int(setnn)  # 개수
                product = pname.replace('/','') # 상품이름 / 제거 경로 오류
                temps = pd.DataFrame({product:num},index = [0]) # 임시 데이터 프레임
                setdp = pd.concat([temps,setdp],axis=0,ignore_index=True)
                setdp = setdp.fillna("") # 결측치 제거
                setdp.to_excel(path + '/세트/'+filename+product+'세트개수.xlsx', index=False)







        # 상품추출
        sll = set(lpname)
        llpname = list(sll) # 종류 추출
        for kk in llpname :
            if kk == '퓨어250ml':
                pure250 = ibottle['퓨어250ml'].sum()
                bottle['퓨어250ml']+=pure250
            elif kk == '퓨어500ml':
                pure500 = ibottle['퓨어500ml'].sum()
                bottle['퓨어500ml'] += pure500
            elif kk == '퓨어1L':
                pure1L = ibottle['퓨어1L'].sum()
                bottle['퓨어1L'] += pure1L
            elif kk == '버진250ml':
                ver250 = ibottle['버진250ml'].sum()
                bottle['버진250ml'] += ver250
            elif kk == '버진500ml':
                ver500 = ibottle['버진500ml'].sum()
                bottle['버진500ml'] += ver500
            elif kk == '올리브250ml':
                oliv250 = ibottle['올리브250ml'].sum()
                bottle['올리브250ml'] += oliv250
            elif kk == '키토썸MCT250ml':
                kitosmct = ibottle['키토썸MCT250ml'].sum()
                bottle['키토썸MCT250ml'] += kitosmct
            elif kk == '미녀플랜아보카도오일250ml':
                mi = ibottle['미녀플랜아보카도오일250ml'].sum()
                bottle['미녀플랜아보카도오일250ml'] +=mi
            elif kk == '톡톡300ml(1box)':
                tt1 = ibottle['톡톡300ml(1box)'].sum()
                bottle['톡톡300ml(1box)'] +=tt1
            elif kk == '톡톡100ml':
                tt2 = ibottle['톡톡100ml'].sum()
                bottle['톡톡100ml'] +=tt2
            elif kk == '톡톡10ml':
                tt3 = ibottle['톡톡10ml'].sum()
                bottle['톡톡10ml'] +=tt3
            elif kk == '세븐화이바':
                sev = ibottle['세븐화이바'].sum()
                bottle['세븐화이바'] +=sev
            elif kk == '키토썸아보카도오일250ml':
                kitoav = ibottle['키토썸아보카도오일250ml'].sum()
                bottle['키토썸아보카도오일250ml'] +=kitoav
            elif kk == '즐거운유아보카도오일500ml':
                happ = ibottle['즐거운유아보카도오일500ml'].sum()
                bottle['즐거운유아보카도오일500ml'] +=happ

        # 기존 엑셀에 더하기
        bottle1 = bottle.astype({'퓨어250ml':'int','퓨어500ml':'int','퓨어1L':'int','버진250ml':'int','버진500ml':'int','미녀플랜아보카도오일250ml':'int',
                       '올리브250ml':'int','톡톡300ml(1box)':'int','톡톡100ml':'int','톡톡10ml':'int','세븐화이바':'int',
                       '키토썸MCT250ml':'int','키토썸아보카도오일250ml':'int','즐거운유아보카도오일500ml':'int'})
        global rbottle
        rbottle += bottle1
        # 엑셀로 만들기
        bottle.to_excel(path+'/정리/'+filename+'주문병수.xlsx',index=False)
        os.remove(path+'/정리/주문병 수.xlsx')
        rbottle.to_excel(path+'/정리/주문병 수.xlsx',index=False,sheet_name = "병 수")
        bottle = pd.DataFrame({'퓨어250ml':[0],'퓨어500ml':[0],'퓨어1L':[0],'버진250ml':[0],'버진500ml':[0],'미녀플랜아보카도오일250ml':[0],
                       '올리브250ml':[0],'톡톡300ml(1box)':[0],'톡톡100ml':[0],'톡톡10ml':[0],'세븐화이바':[0],
                       '키토썸MCT250ml':[0],'키토썸아보카도오일250ml':[0],'즐거운유아보카도오일500ml':[0]})  # 초기화

class Helpdialog(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, parent=None, title='도움말',size = (600,600))
        sizer = wx.BoxSizer(wx.VERTICAL)
        message = wx.StaticText(self, label='★도움말★')
        help = wx.StaticText(self, label = """1. 모든 입력은 파일명, 엑셀에 있는 데이터값 그대로 적어주세요.\n
2. 상품이 없다고 뜨는 상품들은 이야기해주시면 추가해서 프로그램 업데이트 해드리겠습니다.\n
3. 에로사항이나 문의사항 있으시면 010-2094-7805 정광원 으로 문자 주세요!!\n
4. 프로그램과 저장되는 폴더들은 같은 폴더안에 둬주세요(그래야 오류안뜸!!)\n
5. 주문서 엑셀 형식은 test.xlsx 파일과 같은 포멧으로 해주세요 그래야 오류 안뜸\n\n\n\n\n\n\n
made by Pupba.J""")
        sizer.Add(message, 10, wx.ALIGN_CENTER | wx.TOP, 5)
        sizer.Add(help, 1, wx.ALIGN_CENTER | wx.TOP, 5)

class Window(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, parent=None, title='test')
        
        self.SetSize(300,300) # 사이즈 설정

        self.btn1 = wx.Button(self, label = "주문서 처리")
        self.btn2 = wx.Button(self, label = "도움말" )
        self.btn3 = wx.Button(self, label="끝내기")

        # 사이저 세팅
        self.gbox = wx.BoxSizer(wx.VERTICAL) # 박스 사이저
        self.gbox.Add(self.btn1, 2, wx.EXPAND,10)
        self.gbox.Add(self.btn2, 1, wx.EXPAND,10)
        self.gbox.Add(self.btn3, 1, wx.EXPAND,10)
        self.SetSizer(self.gbox)  # 셋 사이저

        self.btn1.Bind(wx.EVT_BUTTON, self.onClickbtn1)
        self.btn2.Bind(wx.EVT_BUTTON, self.onClickbtn2)
        self.btn3.Bind(wx.EVT_BUTTON, self.onClickbtn3)


    def onClickbtn1(self, event):
        try :
            p = os.getcwd()
            fileD = wx.FileDialog(self,'파일선택',p,'','*.*') # 파일 다이얼로그
            if fileD.ShowModal() == wx.ID_OK:
                filep = fileD.GetPath() # 파일 경로 반환 str
                filen = fileD.GetFilename().replace('.xlsx','')
                # 백엔드 실행
                start1 = Backend()
                start1.getOrder(filep,filen)
            fileD.Destroy()
            wx.MessageBox("Click", "Warning", wx.OK, )
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "주문서 불러오기", wx.OK)

    def onClickbtn2(self, event):
        try:
            start = Helpdialog()
            try:
                start.ShowModal() == wx.ID_OK
            finally:
                start.Destroy()  # 다이얼로그 끄기
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "도움말", wx.OK)

    def onClickbtn3(self,event):
        try:
            wx.MessageBox("종료합니다.", "종료", wx.OK)
            wx.Exit()  # 종료
        except Exception as ex:
            t1 = str(ex)
            wx.MessageBox("실패했습니다ㅠㅜ", "종료", wx.OK)



if __name__ == "__main__":
    app = wx.App()
    frame = Window()
    frame.Show()

    app.MainLoop()