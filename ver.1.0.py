# ver.1
# 누적파일 생성
# 제품명이 업체별로 다 틀려서 상품명 별로 제품 수량 따로 파악 필요함!

import pandas as pd
import os
import numpy as np

class Getorder :
    def getOrder(self,str):
        # 주문서 입력
        # 주문서 파일 이름 입력
        putorder = pd.read_excel("ver.1/주문서/"+ str +'.xlsx')
        putorder.sort_values(by = '상품명',ascending=False) #데이터 정렬
        putorder['택배비'] = np.where(putorder['택배비'].str.len()== 2, '0' ,putorder['택배비']) #택배비 무료 0원으로 변환
        name = putorder['상품명'] # 상품명
        num = putorder['수량'] # 수량
        price = putorder['단가'] # 단가
        postprice = putorder['택배비'] # 택배비
        account = putorder['거래처'] # 거래처

        # 상품명 재정리
        # 아보카도오일 퓨어 1+1 = 퓨어 2병 / 아보카도오일 퓨어 1+1

        l = pd.concat([account,name,num,price,postprice],axis =1) # 거래처,상품명,수량,단가,택배비 데이터 베이스
        l.to_excel('ver.1/정리/'+str+'정리.xlsx',index = False) # 정리파일 저장


        #종합 파일 불러오기
        #all = pd.read_excel('ver.1/종합/종합.xlsx')


        return l # 정리 파일 반환


    def countproduct(self,dataframe): # 총 생산 병 개수 파악

        # 입력받은 데이터프레임 반환

        l = dataframe # 데이타프레임 저장
        fl_name=l['상품명'].dropna() # 결측치 제거 상품명 default or 0 행기준 / axis = 1 열기준
        fl_num=l['수량'].dropna() # 결측지 제거 수량
        f = pd.concat([fl_name,fl_num],axis=1) # 상품명 , 수량

        # 데이터프레임에서 상품명 추출
        llist =fl_name.tolist() # 상품명 컬럼 데이터 리스트로 변환
        list_set = set(llist) # 집합으로 변환해 중복 제거
        listN = list(list_set) # 다시 리스트로 변환하여 주문서에있는 상품명 리스트로 작성

        #주문 총 병 개수 파악 
        allex = pd.DataFrame({'퓨어':[0],'버진':[0],'올리브':[0]}) #상품별 목록 데이터프레임 생성
        for i in listN:  # for 문으로 목록 돌리기

            a=0 # 제품명 있는지 없는지 체크

            # 1. 아보카도오일 핑크골드
            if i == '아보카도오일 핑크골드1호':
                f1 = f['상품명'] == i  # 맞으면 True 아니면 False 반환
                f1_ = f[f1]
                f1_s = f1_['수량']
                sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                sf1 = f1_s.sum()  # 더하기
                sf1 = sf1.astype('int')  # 정수 변환
                sf1 = sf1 * 2  # 퓨어 2병
                allex['퓨어'] += sf1  # 데이터 프레임 누적


            # 2. 아보카도오일 퓨어 1+1 총 2병 / 선물용 박스제품(요청사항에 입력)
            elif i == '아보카도오일 퓨어 1+1 총 2병 / 선물용 박스제품(요청사항에 입력)':
                f1 = f['상품명'] == i  # 맞으면 True 아니면 False 반환
                f1_ = f[f1]
                f1_s = f1_['수량']
                sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                sf1 = f1_s.sum()  # 더하기
                sf1 = sf1.astype('int')  # 정수 변환
                sf1 = sf1 * 2  # 퓨어 2병
                allex['퓨어'] += sf1  # 데이터 프레임 누적

            #3. 아보카도오일 엑스트라버진 250ml 1병 (낱병)
            elif i == '아보카도오일 엑스트라버진 250ml 1병 (낱병)':
                f1 = f['상품명'] == i  # 맞으면 True 아니면 False 반환
                f1_ = f[f1]
                f1_s = f1_['수량']
                sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                sf1 = f1_s.sum()  # 더하기
                sf1 = sf1.astype('int')  # 정수 변환
                allex['버진'] += sf1  # 데이터 프레임 누적

            #4. 퓨어 250ml 2병+10ml2포증정
            elif i == '퓨어 250ml 2병+10ml2포증정':
                f1 = f['상품명'] == i  # 맞으면 True 아니면 False 반환
                f1_ = f[f1]
                f1_s = f1_['수량']
                sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                sf1 = f1_s.sum()  # 더하기
                sf1 = sf1.astype('int')  # 정수 변환
                sf1 = sf1 * 2 #퓨어 2병
                allex['퓨어'] += sf1  # 데이터 프레임 누적

            #5. 핑크골드
            elif i == '핑크골드':
                f1 = f['상품명'] == i  # 맞으면 True 아니면 False 반환
                f1_ = f[f1]
                f1_s = f1_['수량']
                sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                sf1 = f1_s.sum()  # 더하기
                sf1 = sf1.astype('int')  # 정수 변환
                sf1 = sf1 * 2  # 퓨어 2병
                allex['퓨어'] += sf1  # 데이터 프레임 누적

            else :
                a += 1 # 상품이 없을 떄 1 추가
                if a!= 0 :
                    print(i + ' <--이 상품은 저장된 목록에 없습니다.')
                    print(allex)
                    #예외처리
                    name = input('상품 이름을 입력해주세요 ★주의!! 엑셀에 있는 값 그대로 써주세요★ : ')  # 새로운 상품명 입력 받기
                    pname = input('퓨어,버진,올리브 선택 : ')
                    num = int(input('개수 입력 : '))
                    # 상품명 리스트에서 입력한 상품이있는지 없는지 검사
                    for k in listN:
                        if k == name:  # 리스트에 입력한 상품이 있을 때
                            f1 = f['상품명'] == name  # 맞으면 True 아니면 False 반환
                            f1_ = f[f1]  # T 인것만 f1_에 저장
                            f1_s = f1_['수량']  # 수량 컬럼에 값을 f1_s에 저장
                            sf1 = pd.DataFrame({'총 개수': [0]})  # 데이터프레임 초기화
                            sf1 = f1_s.sum()  # 수량 컬럼 데이터 더하기
                            sf1 = sf1.astype('int')  # 데이터 정수형 변환
                            sf1 = sf1 * num  # 퓨어 2병
                            allex[pname] += sf1  # 데이터 프레임 누적
                        else:
                            print('셀에 상품이 없습니다.')
        print(allex)
        allex.to_excel("ver.1/정리/주문병 수.xlsx",index=False) # 상품 목록 정리 저장


    def priceCalculator(self,dataframe):
        l = dataframe #데이터 프레임
        fl_account = l['거래처'].dropna() # 결측치 제거 거래처
        fl_name = l['상품명'].dropna()  # 결측치 제거 상품명 default
        # or 0 행기준 / axis = 1 열기준
        fl_num = l['수량'].dropna()  # 결측지 제거 수량
        fl_price = l['단가'].dropna() # 결측치 제거 단가
        fl_post = l['택배비'].dropna() # 결측치 제거 택배비
        total = pd.DataFrame({}) # 초기화
        s_total = pd.DataFrame({}) # 초기화
        total['총액'] = (fl_num * fl_price) # 개수 * 단가
        total['총액']=total['총액'].astype('int64') # 형변환
        fl_post=fl_post.astype('int64') # 형변환
        total['총액'] = total['총액'] + fl_post # 총액 = 개수 * 단가 + 택배비
        s_total = total['총액'].sum()


        df = pd.DataFrame({}) # 데이터 프레임 생성
        df = pd.concat([fl_account,fl_name,fl_num,fl_price,fl_post,total],axis=1)
        df.to_excel("ver.1/총금액/총금액.xlsx")
        return s_total #총금액 반환











f = input('주문서 파일이름 입력 : ')
start = Getorder() # 클래스 활성화
a=start.getOrder(f) # 주문서 불러오기
start.countproduct(a) # 수량 파악
totalP=start.priceCalculator(a) # 총액계산
print('총금액은 %d원입니다.'%totalP)
