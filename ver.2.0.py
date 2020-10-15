# ver.2.1
# 단가계산기 삭제
# 각종 오류 수정

# 모듈 import
import openpyxl
import pandas as pd
import os
import numpy as np
import wx


# 전역변수 선언
order = pd.DataFrame({}) # 빈 데이터프레임(주문서)
bottle = pd.DataFrame({'퓨어250ml':[0],'퓨어500ml':[0],'버진250ml':[0],'버진500ml':[0],'올리브':[0],'키토썸MCT오일':[0],'톡톡아보오일세트':[0]}) # 병 종합
path = str(os.getcwd()) # 경로 설정 작업 폴더
rbottle = pd.read_excel(path+'/정리/주문병 수.xlsx') # 기본 주문병 수 불러오기

app = wx.App()
frame = wx.Frame(None)
# 사이즈 설정
fsize = wx.Size(300, 300)  # 사이즈 설정
frame.SetSize(fsize)
fpos = wx.Point(300, 100)  # 위치 설정
frame.SetPosition(fpos)
frame.SetTitle("주문서 처리")  # 윈도우바 타이틀 설정
frame.SetWindowStyle(wx.DEFAULT_FRAME_STYLE & ~wx.RESIZE_BORDER & ~wx.MAXIMIZE_BOX)  # 크기 변경 불가

#버튼생성
btn1 = wx.Button(frame, label = '주문서 처리')
btn2 = wx.Button(frame, label = '도움말')
btn3 = wx.Button(frame, label = '끝내기')



#사이저 셋팅
gbox = wx.GridSizer(2,2,15,15) # 그리드사이저 설정 2행 2열 15픽셀 간격
frame.SetSizer(gbox) # 셋 사이저
gbox.Add(btn1, 0, wx.EXPAND)
gbox.Add(btn2, 0, wx.EXPAND)
gbox.Add(btn3, 0, wx.EXPAND)

#다이얼로그 생성
def btn1textdialog():
    dIg = wx.TextEntryDialog(message='파일이름을 입력해주세요!',parent=None) # 다이얼로그 생성
    try:
        if dIg.ShowModal() == wx.ID_OK:
           vdIg = dIg.GetValue() # 값 추출
    finally:
        dIg.Destroy() # 다이얼로그 파괴
    return vdIg  # 값을 반환

def inputPrice(product):
    dIg = wx.TextEntryDialog(message= product+'의 단가를 입력해주세요',parent=None) # 다이얼로그 생성
    try:
        if dIg.ShowModal() == wx.ID_OK:
            vdIg = dIg.GetValue() # 값 추출
            price = int(vdIg)
    finally:
        dIg.Destroy() # 다이얼로그 파괴
    return price  # 값을 반환


class Numandkind(wx.Dialog):
    # 다이얼로그 입력값 여러개
    def __init__(self,te):
        wx.Dialog.__init__(self, parent=None, title='상품 입력', size=(500, 300))
        self.te = te
        sizer = wx.BoxSizer(wx.VERTICAL)

        message = wx.StaticText(self, label=te+'입력\n띄어쓰기 주의!!\n종류 : 버진250ml,500ml,퓨어250ml,500ml,키토썸MCT오일,올리브,톡톡아보오일세트')

        # 종류
        pname_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.pname_edit = wx.TextCtrl(parent=self, name='종류')
        pname_sizer.Add(wx.StaticText(self, label='종류: '), 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        pname_sizer.Add(self.pname_edit, 4, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 10)

        # 수량
        num_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.num_edit = wx.TextCtrl(parent=self, name='수량')
        num_sizer.Add(wx.StaticText(self, label='수량: '), 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)
        num_sizer.Add(self.num_edit, 4, wx.ALIGN_CENTER_VERTICAL | wx.LEFT | wx.RIGHT, 10)

        # OK, CANCEL buttons
        _btns_sizer = wx.BoxSizer(wx.HORIZONTAL)
        okbtnSizer = self.CreateStdDialogButtonSizer(wx.OK)
        _btns_sizer.Add(okbtnSizer, 1, wx.ALIGN_CENTER_VERTICAL, 0)

        #
        sizer.Add(message, 1, wx.ALIGN_CENTER | wx.TOP, 5)
        sizer.Add(pname_sizer, 1, wx.ALIGN_CENTER | wx.TOP, 5)
        sizer.Add(num_sizer, 1, wx.ALIGN_CENTER | wx.TOP, 5)
        sizer.Add(wx.StaticLine(self, size=(250, 2)), 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 5)
        sizer.Add(_btns_sizer, 1, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 5)
        self.SetSizer(sizer)

    def GetValue(self):
        return self.pname_edit.GetLineText(0), self.num_edit.GetLineText(0)  # 값 반환


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








# back end 작성
class Backend :
    # 1. 주문서 파일 읽어오기
    def getOrder(self,str1):
        global path
        global order # df 사용하겠다 선언
        temp = pd.read_excel(path+'/주문서/'+str1+'.xlsx') # 주문서 읽어오기
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
        lpname = [] # 입력받는 병 종류 리스트
        #상품명별로 병 종류 및 총 병 개수 종합
        global bottle
        pplist = order_pname.tolist()
        setpp = set(pplist)
        pnmm = list(setpp)  # 상품명 리스트 작성
        ibottle = pd.DataFrame({})  # 임시 저장 데이타 프레임
        for Nd in pnmm :
            temp__ = tempbottle[tempbottle['상품명'] == Nd] # 상품별 추출 데이터프레임
            # 중복 오류 제거
            ftemp__ = pd.DataFrame([{}])
            tname = pd.DataFrame([{'상품명':Nd}])
            tnu = temp__['수량'].sum()
            ttnu = pd.DataFrame([{'수량':[0]}])
            ttnu['수량'][0] = tnu
            ftemp__ = pd.concat([tname,ttnu],axis=1)

            start2 = Numandkind(Nd) # 클래스 선언
            product = None
            num = None
            num_ = None  # 초기화
            try:
                if start2.ShowModal() == wx.ID_OK:
                    product_, num_ = start2.GetValue()  # 값가져오기
                    num = int(num_)# 정수형변환
                    tt = pd.Series(num)
                    product = str(product_)
                    lpname.append(product)

            finally:
                start2.Destroy()  # 다이얼로그 끄기

            temp__ = ftemp__.astype({'수량':'int64'})
            num_ =  ftemp__['수량'][0] * tt[0] # 같은 변수형으로 해야 오류 안남
            bottle_ = pd.DataFrame([{product:[0]}])  # 저장할 데이터프레임
            inum_ = num_.astype('int64')
            bottle_[product] = num_
            ibottle = pd.concat([bottle_,ibottle],axis=0,ignore_index=True)
            ibottle = ibottle.fillna(int(0)) # 결측치 0으로


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
            elif kk == '버진250ml':
                ver250 = ibottle['버진250ml'].sum()
                bottle['버진250ml'] += ver250
            elif kk == '버진500ml':
                ver500 = ibottle['버진500ml'].sum()
                bottle['버진500ml'] += ver500
            elif kk == '올리브':
                oliv = ibottle['올리브'].sum()
                bottle['올리브'] += oliv
            elif kk == '키토썸MCT오일':
                kitos = ibottle['키토썸MCT오일'].sum()
                bottle['키토썸MCT오일'] += kitos
            elif kk == '톡톡아보오일세트':
                tt = ibottle['톡톡아보오일세트'].sum()
                bottle['톡톡아보오일세트'] +=tt

        # 기존 엑셀에 더하기
        bottle1 = bottle.astype({'퓨어250ml':'int','퓨어500ml':'int','버진250ml':'int','버진500ml':'int','올리브':'int','키토썸MCT오일':'int','톡톡아보오일세트':'int'})
        global rbottle
        rbottle += bottle1
        # 엑셀로 만들기
        bottle.to_excel(path+'/정리/'+str1+'주문병수.xlsx')
        rbottle.to_excel(path+'/정리/주문병 수.xlsx',index=False)
        bottle = pd.DataFrame({'퓨어250ml': [0], '퓨어500ml': [0], '버진250ml': [0], '버진500ml': [0], '올리브': [0], '키토썸MCT오일': [0],'톡톡아보오일세트':[0]})  # 초기화


# event setting
def onClickbtn1(event) :
    try :
        start = Backend()
        orde=btn1textdialog() # 주문서 파일 이름 입력 받기
        sorde = str(orde)
        start.getOrder(sorde) # 주문서 처리 시작
        wx.MessageBox("완료!", "주문서처리", wx.OK)
    except Exception as ex:
        t1 = str(ex)
        wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "주문서처리", wx.OK)
btn1.Bind(wx.EVT_BUTTON, onClickbtn1) # 버튼 1
def onClickbtn2(event) :
    try :
        start = Helpdialog()
        try:
            start.ShowModal() == wx.ID_OK
        finally:
            start.Destroy()  # 다이얼로그 끄기
    except Exception as ex:
        t1 = str(ex)
        wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "도움말", wx.OK)
btn2.Bind(wx.EVT_BUTTON, onClickbtn2) # 버튼 2

def onClickbtn3(event) :
    try :
        wx.MessageBox("종료합니다.", "종료", wx.OK)
        wx.Exit() #종료
    except Exception as ex:
        t1 = str(ex)
        wx.MessageBox("실패했습니다ㅠㅜ", "종료", wx.OK)
btn3.Bind(wx.EVT_BUTTON, onClickbtn3) # 버튼 3


frame.Show(True)
app.MainLoop()