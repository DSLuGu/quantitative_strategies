import sys
import win32com.client


g_objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')


def check_plus_status(originalFunc):
    '''originalFunc 콜하기 전에 PLUS 연결 상태 체크하는 데코레이터'''
    
    def wrapper(*args, **kwargs):
        
        # 프로세스 관리자 권한 실행 여부
        isAdmin = ctypes.windll.shell32.IsUserAnAdmin()
        if not isAdmin:
            print("It isn't a process running in user-admin mode...")
            return False
        
        # 연결 여부 체크
        bConnect = g_objCpCybos.IsConnect
        if bConnect == 0:
            print("PLUS is not propertly connected...")
            return False
        
        return originalFunc(*args, **kwargs)
    
    return wrapper


def avoid_rq_limit_warning():
    '''Creon API의 요청제한 횟수를 지키기 위한 메소드'''
    
    remainTime = self.g_objCpCybos.LimitRequestRemainTime
    remainCount = self.g_objCpCybos.GetLimitRemainCount(1)
    if remainCount <= 3: time.sleep(remainTime / 1000)
    
    return None


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
        
        self.account = self.objTrade.AccountNumber[0] # 계좌번호
        # -1: 전체, 1: 주식, 2: 선물/옵션, 16: EUREX, 32: FX 마진, 64: 해외선물
        # 3: 주식(1) + 선물/옵션(2), 96: FX 마진(32) + 해외선물(64)
        self.accountFlag = self.objTrade.GoodsList(self.account, 1) # 주식상품 구분
        print("[Account] {}\n[Account flag] {}".format(account, accountFlag))
        
        self.objRq = win32com.client.Dispatch('CpTrade.CpTd6033')
        self.objRq.SetInputValue(0, self.account) # 계죄번호
        self.objRq.SetInputValue(1, self.accountFlag[0]) # 상품구분 - 주식상품 중 첫 번째
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


class Cp3011:
    
    @check_plus_status
    def __init__(self):
        
        self.objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        # 주문을 하기 위한 예비 과정을 수행
        # -1: 오류, 0: 정상, 1: 업무 키 입력 잘못 됨, 2: 계좌 비밀번호 입력 잘못 됨, 3: 취소
        initCheck = self.objTrade.TradeInit(0)
        if initCheck != 0:
            print("Fail to initialize order...")
            return None
        
        self.account = self.objTrade.AccountNumber[0]
        self.accountFlag = self.objTrade.GoodsList(self.account, 1)
        
        self.objRq = win32com.client.Dispatch('CpTrade.CpTd3011') # 매도/매수 object
    
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
    
    def rq3011(self, buyOrSell, stockCode, orderCnt, orderPrice, orderFlag="01"):
        '''주식 주문 메소드
        
        :param buyOrSell: 주문 구분, {1:매도, 2:매수}
        :param stockCode: 종목코드
        :param orderCnt: 주문수량
        :param orderPrice: 주문금액
        :param orderFlag: 주문호가구분코드, {
            01:보통, 02:임의, 03:시장가, 05:조건부지정가, 06:희망대량, 
            09:자사주, 10:스톡옵션자사주, 11:금전신탁자사주, 
            12:최유리지정가, 13:최우선지정가, 23:임의시장가, 
            25:임의조건부지정가, 51:장중대량, 52:장중바스켓, 
            61:개시전종가, 62:개시전종가대량, 63:개시전시간외바스켓, 
            67:개시전금전신탁자사주, 69:개시전대량자기, 71:신고대량(전장시가), 
            72:시간외대량, 73:신고대량(종가), 77:금전신탁종가대량, 
            79:시간외대량자기, 80:시간외바스켓, 
        }
        
        :return:
        '''

        self.objRq.SetInputValue(0, buyOrSell)
        self.objRq.SetInputValue(1, self.account)
        self.objRq.SetInputValue(2, self.accountFlag[0]) # 상품구분 - 주식상품 중 첫 번째
        self.objRq.SetInputValue(3, stockCode)
        self.objRq.SetInputValue(4, orderCnt)
        self.objRq.SetInputValue(5, orderPrice)
        self.objRq.SetInputValue(7, "0") # 주문조건구분코드 매매조건, {0:없음, 1:IOC, 2:FOK}
        self.objRq.SetInputValue(8, orderFlag)
        
        self.objRq.BlockRequest()
        if not self._check_rq_status(): return False
        
        return True


class NotConcludedData:
    
    def __init__(self):
        
        self.stockCode = ""     # 종목코드
        self.stockName = ""     # 종목명
        self.orderNum = 0       # 주문번호
        self.orderPrevNum = 0   # 원주문번호
        self.orderDesc = ""     # 주문구분내용
        self.amount = 0         # 주문수량
        self.price = 0          # 주문단가
        self.concAmount = 0     # 체결수량
        self.credit = ""        # 신용구분 (현금, 유통융자, 자기융자, 유통대주, 자기대주)
        self.modAvaliAmount = 0 # 정정/취소 가능 수량
        self.buyOrSell = ""     # 매매구분코드 {1:매도, 2:매수}
        self.creditDate = ""    # 대출일
        self.orderFlag = ""     # 주문호가 구분코드
        self.orderFlagDesc = "" # 주문호가 구분코드 내용
        
        self.concDic = {'1': '체결', '2': '확인', '3': '거부', '4': '접수'}
        self.buyOrSellDic = {'1': '매도', '2': '매수'}


class Cp5339:
    '''계좌별 미체결 잔량 데이터를 요청하고 수신'''
    
    @check_plus_status
    def __init__(self):
        
        self.objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        # 주문을 하기 위한 예비 과정을 수행
        # -1: 오류, 0: 정상, 1: 업무 키 입력 잘못 됨, 2: 계좌 비밀번호 입력 잘못 됨, 3: 취소
        initCheck = self.objTrade.TradeInit(0)
        if initCheck != 0:
            print("Fail to initialize order...")
            return None
        
        self.account = self.objTrade.AccountNumber[0]
        self.accountFlag = self.objTrade.GoodsList(self.account, 1)
        # 미체결 조회 object
        self.objRq = win32com.client.Dispatch('CpTrade.CpTd5339')
    
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
    
    def rq5339(self, orderDicList, orderList, rqCnt=20):
        '''
        :param orderDicList:
        :param orderList:
        :param rqCnt: 요청개수(최대 20개)
        
        :type 4: 주문구분코드, {
            0:전체, 1:거래소주식, 2:장내채권, 3:코스닥주식, 4:장외단주, 5:프리보드
        }
        :type 5: 정렬구분코드, {0:순차, 1:역순}
        :type 6: 주문종가구분코드, {0:전체, 1:일반, 2:시간외종가, 4:시간외단일가}
        :type 7: 요청개수(최대 20개)
        :return:
        '''
        
        self.objRq.SetInputValue(0, self.account)
        self.objRq.SetInputValue(1, self.accountFlag[0])
        self.objRq.SetInputValue(4, "0")
        self.objRq.SetInputValue(5, "1")
        self.objRq.SetInputValue(6, "0")
        self.objRq.SetInputValue(7, rqCnt)
        
        print("[Cp5339]Start searching to orders not concluded...")
        while True:
            ret = self.objRq.BlockRequest()
            if not self._check_rq_status(): return False
            
            if ret == 2 or ret == 3:
                print(">>>Request error...", ret)
                return False
            
            # 통신 초과 요청 방지에 의한 오류인 경우
            while ret == 4: # 연속 주문 오류, 이 경우는 남은 시간동안 반드시 대기
                print(">>>", end="")
                avoid_rq_limit_warning()
                ret = self.objRq.BlockRequest()
            
            cnt = self.objRq.GetHeaderValue(5)
            print(">>>수신개수:", cnt)
            if cnt == 0: break
            
            for i in range(cnt):
                data = NotConcludedData()
                data.orderNum = self.objRq.GetDataValue(1, i)
                data.orderPrevNum = self.objRq.GetDataValue(2, i)
                data.stockCode = self.objRq.GetDataValue(3, i)
                data.stockName = self.objRq.GetDataValue(4, i)
                data.orderDesc = self.objRq.GetDataValue(5, i)
                data.amount = self.objRq.GetDataValue(6, i)
                data.price = self.objRq.GetDataValue(7, i)
                data.concAmount = self.objRq.GetDataValue(8, i)
                data.credit = self.objRq.GetDataValue(9, i)
                data.modAvaliAmount = self.objRq.GetDataValue(11, i)
                data.buyOrSell = self.objRq.GetDataValue(13, i)
                data.creditDate = self.objRq.GetDataValue(17, i)
                data.orderFlag = self.objRq.GetDataValue(19, i)
                data.orderFlagDesc = self.objRq.GetDataValue(21, i)
                
                orderDicList[data.orderNum] = data
                orderList.append(data)
            
            # 연속 처리 체크 - 다음 데이터가 없으면 중지
            if self.objRq.Continue == False:
                print(">>>No next data...")
                break
        
        return True
        