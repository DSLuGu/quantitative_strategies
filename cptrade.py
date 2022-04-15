import sys
import win32com.client


class Cp6033:
    '''계좌별 잔고 및 주문체결 평가 현황 데이터를 요청 및 수신'''
    
    def __init__(self):
        
        self.objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        # 주문을 하기 위한 예비 과정을 수행
        # -1: 오류, 0: 정상, 1: 업무 키 입력 잘못 됨, 2: 계좌 비밀번호 입력 잘못 됨, 3: 취소
        initCheck = self.objTrade.TradeInit(0)
        if initCheck != 0:
            print("Fail to initialize order...")
            return None
        
        account = self.objTrade.AccountNumber[0] # 계좌번호
        # -1: 전체, 1: 주식, 2: 선물/옵션, 16: EUREX, 32: FX 마진, 64: 해외선물
        # 3: 주식(1) + 선물/옵션(2), 96: FX 마진(32) + 해외선물(64)
        accountFlag = self.objTrade.GoodsList(acc, 1) # 주식상품 구분
        print("[Account] {}\n[Account flag] {}".format(account, accountFlag))
        
        self.objRq = win32com.client.Dispatch('CpTrade.CpTd6033')
        self.objRq.SetInputValue(0, account) # 계죄번호
        self.objRq.SetInputValue(1, accountFlag[0]) # 상품구분 - 주식상품 중 첫 번째
        self.objRq.SetInputValue(2, 50) # 요청건수(최대 50)
    
    def _check_rq_status(self):
        '''BlockRequest() 요청 후 통신상태 검사'''
        
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        
        if rqStatus == 0:
            print("Request Status is normal[{}]{}".format(rqStatus, rqRet))
        else:
            print("Request Status is abnormal[{}]{}".format(rqStatus, rqRet))
            return False
        
        return True
    
    def rq6033(self, retCode):
        '''실제적인 6033 통신 처리'''
        
        self.objRq.BlockRequest()
        if not self._check_rq_status(): return False
        
        cnt = self.objRq.GetHeaderValue(7) # 7: 수신개수
        for i in range(cnt):
            # 0: 종목명, 
            # 1: 신용구분 {'Y': 신용융자/유통융자, 'D': 신용대주/유통대주, 'B': 담보대출, 
            #           'M': 매입담보대출, 'P': 플러스론 대출, 'I': 자기융자/유통융자}
            # 2: 대출일, 3: 결제 잔고수량, 4: 결제 장부단가, 5: 전일체결수량, 6: 금일체결수량
            # 7: 체결잔고수량, 9: 평가금액(단위:원) - 천원 미만은 내림, 
            # 10: 평가손익(단위:원) - 천원 미만은 내림, 
            # 11: 수익률, 12: 종목코드, 13: 주문구분, 15: 매도가능수량, 16: 만기일, 
            # 17: 체결장부단가, 18: 손익단가
            stockCode = self.objRq.GetDataValue(12, i) # 종목코드
            stockName = self.objRq.GetDataValue(0, i) # 종목명
            retCode.append(stockCode)
            if len(retCode) >= 200: break # 최대 200 종목만?
            # cashFlag = self.objRq.GetDataValue(1, i)
            # date = self.objRq.GetDataValue(2, i)
            tradeAmount = self.objRq.GetDataValue(3, i)
            tradeBookUnitPrice = self.objRq.GetDataValue(4, i)
            prevConclusionAmount = self.objRq.GetDataValue(5, i)
            conclusionAmount = self.objRq.GetDataValue(6, i)
            amount = self.objRq.GetDataValue(7, i)
            evalValue = self.objRq.GetDataValue(9, i)
            evalPerc = self.objRq.GetDataValue(10, i)
            earningsRatio = self.objRq.GetDataValue(11, i)
            orderFlag = self.objRq.GetDataValue(13, i)
            sellableAmount = self.objRq.GetDataValue(15, i)
            dueDate = self.objRq.GetDataValue(16, i)
            buyPrice = self.objRq.GetDataValue(17, i)
            percUnitPirce = self.objRq.GetDataValue(18, i)
        
        return None
    
    def request(self, retCode):
        
        self.rq6033(retCode)
        
        while self.objRq.Continue:
            self.rq6033(retCode)
            if len(retCode) >= 200: break
        
        return True