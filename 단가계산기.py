import openpyxl
import pandas as pd
import wx
import numpy as np

# 업체별 정리된 파일 읽어와서 단가 입력 받고 하기

df_total = pd.DataFrame([{}]) # 전역변수 데이터프레임

app = wx.App()
frame = wx.Frame(None)
# 사이즈 설정
fsize = wx.Size(200, 200)  # 사이즈 설정
frame.SetSize(fsize)
fpos = wx.Point(300, 100)  # 위치 설정
frame.SetPosition(fpos)
frame.SetTitle("단가계산")  # 윈도우바 타이틀 설정
frame.SetWindowStyle(wx.DEFAULT_FRAME_STYLE & ~wx.RESIZE_BORDER & ~wx.MAXIMIZE_BOX)  # 크기 변경 불가


#버튼생성
btn1 = wx.Button(frame, label = '셀불러오기')
btn2 = wx.Button(frame, label = '끝내기')

gbox = wx.GridSizer(2,1,15,15) # 그리드사이저 설정 2행 1열 15픽셀 간격
frame.SetSizer(gbox) # 셋 사이저
gbox.Add(btn1, 0, wx.EXPAND)
gbox.Add(btn2, 0, wx.EXPAND)

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

def totalprice(str1):
    df = pd.read_excel('업체별 정리/'+str1+'.xlsx') #업체별 정리 데이터 프레임 가져오기
    df_N = df['거래처']
    llist = df_N.tolist()  # 거래처 컬럼 데이터 리스트로 변환
    list_set = set(llist)  # 집합으로 변환해 중복 제거
    fdf_N = list(list_set)  # 다시 리스트로 변환하여 주문서에있는 상품명 리스트로 작성
    global df_total

    # 업체별로 총액 계산 해야함
    for Nl in fdf_N :
        we = df[df['거래처'] == Nl]  # 위메프만 추출
        we_ = we.astype({'수량': 'int64','택배비':'int64'})  # 컬럼 데이터 타입 변경
        wep = we['상품명'] # 상품 데이터 추출
        llist1 = wep.tolist()  # 상품명 컬럼 데이터 리스트로 변환
        list_set1 = set(llist1)  # 집합으로 변환해 중복 제거
        wep_n = list(list_set1)  # 다시 리스트로 변환하여 상품명 리스트로 작성
        index = list(range(len(wep_n)))
        we_df = pd.DataFrame([{'거래처': '', '상품명': '', '수량': [0],'택배비':[0],'총금액':[0]}],index=index)
        for i in range(len(wep_n)) :
            we_df['거래처'][i] = Nl
            t = str(wep_n[i])
            we_df['상품명'][i] = t
            nwe = we_[we_['상품명'] == wep_n[i]] # 이 상품이 들어간 열만 추출
            snwe = nwe['수량']
            isnwe = snwe.sum() # 개수 총합
            isnwe = isnwe.astype("int")
            we_df['수량'][i] = isnwe # 총 수량 데이터프레임에 입력

            # 단가 입력 받기
            price = inputPrice(t) # 다이얼로그 생성
            postprices = we_['택배비']
            postprice = postprices.sum() # 택배비 총합
            ipost = postprice.astype('int')
            we_df['택배비'][i] = ipost # 택배비
            totalprice = (price * isnwe) + ipost


            i = i + 1
        we_df['총금액'] = totalprice  # 총 금액 입력
        df_total = pd.concat([df_total, we_df], axis=0)
        df_total = df_total.dropna()  # 결측치 제거

    #엑셀로 만들기
    df_total.to_excel('총금액/'+str1+'총금액.xlsx',index = False)

def onClickbtn1(event) :
    try :
        a = btn1textdialog()
        totalprice(a) # 실행
        wx.MessageBox("완료되었습니다" ,"계산", wx.OK)
    except Exception as ex:
        t1 = str(ex)
        wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "계산", wx.OK)
btn1.Bind(wx.EVT_BUTTON, onClickbtn1)  # 버튼 1
def onClickbtn2(event) :
    try :
        wx.MessageBox("종료합니다.","계산", wx.OK)
        wx.Exit()
    except Exception as ex :
        t1 = str(ex)
        wx.MessageBox("실패했습니다ㅠㅠ : " + t1, "계산", wx.OK)
btn2.Bind(wx.EVT_BUTTON, onClickbtn2) #버튼2

frame.Show(True)
app.MainLoop()
