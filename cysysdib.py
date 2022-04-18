import sys
import ctypes

import time

from collections import defaultdict

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


class CpMarketEye:
    
    def __init__(self):
        
        self.objRq = win32com.client.Dispatch('CySysDib.MarketEye')

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
    
    @check_plus_status
    def request(self, codes, rqField):
        
        # Set request fields
        self.objRq.SetInputValue(0, rqField)
        self.objRq.SetInputValue(1, codes) # a stockCode or stockCodes
        self.objRq.BlockRequest()
        
        # Handle request error
        if not self._check_rq_status(): return False
        
        rst = []
        cnt = self.objRq.GetHeaderValue(2) # 종목 개수
        for c in cnt:
            rst.append([self.objRq.GetDataValue(f, c) for f in len(rqField)])
        
        return rst


class CpStockChart:
    
    def __init__(self):
        
        self.objStockChart = win32com.client.Dispatch('CySysDib.StockChart')
        self.fieldName = {
            0: ('날짜', 'date'), 1: ('시간', 'time'), 2: ('시가', 'open'), 
            3: ('고가', 'high'), 4: ('저가', 'low'), 5: ('종가', 'close'), 
            6: ('전일대비', 'diff'), 8: ('거래량', 'volume'), 9: ('거래대금', 'tradeVolume'), 
            10: ('누적체결매도수량', 'cumNegVolume'), 11: ('누적체결매수수량', 'cumPosVolume'), 
            12: ('상장주식수', 'numOfListedShares'), 13: ('시가총액', 'marketCap'), 
            14: ('외국인주문한도수량', 'limitQuantityForForeignOrders'), 
            15: ('외국인주문가능수량', 'quantityAvailableForForeignOrders'), 
            16: ('외국인현보유수량', 'holdingQuantityByForeign'), 
            17: ('외국인현보유비율', 'holdingRatioByForeign'), 18: ('수정주가일자', 'adjustedPrice'), 
            19: ('수정주가비율', 'adjustedPriceRatio'), 20: ('기관순매수', 'netBuyingInstitution'), 
            21: ('기관누적순매수', 'cumNetBuyingInstitution'), 
            22: ('등락주선', 'advanceDeclineLine'), 23: ('등락비율', 'advanceDeclineRatio'), 
            24: ('예탁금', 'deposit'), 25: ('주식회전율', 'stockTurnoverRatio'), 
            26: ('거래성립률', 'transactionRatio'), 37: ('대비부호', 'contrastSign'), 
        }
    
    def _check_rq_status(self):
        '''BlockRequest() 요청 후 통신상태 검사'''
        
        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        
        if rqStatus != 0:
            print("Request Status is abnormal[{}]{}".format(rqStatus, rqRet))
            return False
        
        return True
    
    @check_plus_status
    def request_dwm(self, stockCode, dwm, count, caller:'MainWindow', fromDate=0):
        '''차트 요청 - 최근일부터 개수 기준
        
        :param stockCode: 종목코드
        :param dwm: 'D':일봉, 'W':주봉, 'M':월봉
        :param count: 요청할 데이터 개수
        :param caller: 이 메소드를 호출한 인스턴스, 결과 데이터를 caller 의 멤버로 전달하기 위함
        :param fromDate: 요청 마지막 일
        :return: None
        '''
        
        self.objStockChart.SetInputValue(0, stockCode)
        self.objStockChart.SetInputValue(1, ord('2')) # (요청구분) - '1':기간으로 요청, '2':개수로 요청
        self.objStockChart.SetInputValue(4, count) # (요청개수)
        
        self.objStockChart.SetInputValue(5, caller.rqFields) # (요청항목)
        self.objStockChart.SetInputValue(6, dwm) # (차트 주기) - 'D':일, 'W':주, 'M':월
        self.objStockChart.SetInputValue(9, ord('1')) # (수정주가) - '0':무수정주가, '1':수정주가
        # 주가는 유무상증자, 배당, 액면분할 등이 생길 때 연속성을 잃고 단층을 보이게 된다.
        # 이 경우 이전 주가와 현재의 주가를 비교하는데 애로가 따르게 된다.
        # 따라서 주가에 연속성을 부여하기 위해 일정한 수정을 할 수 있는데, 이것을 수정주가라 한다.
        
        rcvCount = 0
        rcvData = defaultdict(list)
        while count > rcvCount:
            self.objStockChart.BlockRequest()
            self._check_rq_status()
            avoid_rq_limit_warning()
            
            rcvBatchLen = self.objStockChart.GetHeaderValue(3) # type 3: 수신개수
            rcvBatchLen = min(rcvBatchLen, count - rcvCount) # 정확히 count 개수만큼 받기 위해서
            for i in range(rcvBatchLen):
                for colIdx, col in enumerate(caller.rqColumns):
                    rcvData[col].append(self.objStockChart.GetDataValue(colIdx, i))
            
            if len(rcvData['date']) == 0:
                print(f"[{stockCode}] no data...")
                return False
            
            # rcvBatchLen 만큼 받은 데이터의 가장 오래된 date
            rcvOldestDate = rcvData['date'][-1]
            
            rcvCount += rcvBatchLen
            caller.returnStatusMsg = "{} / {}".format(rcvCount, count)
            
            # 서버가 가진 모든 데이터를 요청한 경우 break.
            # self.objStockChart.Continue 는 개수로 요청한 경우
            # count 만큼 이미 다 받았더라도 계속 1의 값을 가지고 있어서
            # while 조건문에서 count > rcvCount 를 체크해줘야 함.
            if not self.objStockChart.Continue: break
            if rcvOldestDate < fromDate: break
        
        caller.rcvData = rcvData # 받은 데이터를 caller 의 멤버로 저장
        
        return True
    
    @check_plus_status
    def request_mt(self, stockCode, mt, tickRange, count, caller:'MainWindow', fromDate=0):
        '''차트 요청 - 분간, 틱 차트
        
        :param stockCode: 종목 코드
        :param mt: 'm':분봉, 'T':틱봉
        :param tickRange: 1분봉, 5분봉, ...
        :param count: 요청할 데이터 개수
        :param caller: 이 메소드 호출한 인스턴스, 결과 데이터를 caller 의 멤버로 전달하기 위함
        :param fromDate: 요청 마지막 일
        :return:
        '''
        
        self.objStockChart.SetInputValue(0, stockCode)
        self.objStockChart.SetInputValue(1, ord('2')) # (요청구분) - '1':기간으로 요청, '2':개수로 요청
        self.objStockChart.SetInputValue(4, count) # (요청개수)
        
        self.objStockChart.SetInputValue(5, caller.rqFields) # (요청항목)
        self.objStockChart.SetInputValue(6, mt) # (차트 주기) - 'm':분봉, 'T':틱봉
        self.objStockChart.SetInputValue(7, tickRange) # (분틱차트 주기) - 'm':분봉, 'T':틱봉
        self.objStockChart.SetInputValue(9, ord('1')) # (수정주가) - '0':무수정주가, '1':수정주가
        
        rcvCount = 0
        rcvData = defaultdict(list)
        while count > rcvCount:
            self.objStockChart.BlockRequest()
            self._check_rq_status()
            avoid_rq_limit_warning()
            
            rcvBatchLen = self.objStockChart.GetHeaderValue(3) # type 3: 수신개수
            rcvBatchLen = min(rcvBatchLen, count - rcvCount) # 정확히 count 개수만큼 받기 위해서
            for i in range(rcvBatchLen):
                for colIdx, col in enumerate(caller.rqColumns):
                    rcvData[col].append(self.objStockChart.GetDataValue(colIdx, i))
            
            if len(rcvData['date']) == 0:
                print(f"[{stockCode}] no data...")
                return False
            
            # len 만큼 받은 데이터의 가장 오래된 date
            rcvOldestDate = int('{}{:04}'.format(rcvData['date'][-1], rcvData['time'][-1]))
            
            rcvCount += rcvBatchLen
            caller.returnStatusMsg = "{} / {}(maximum)".format(rcvCount, count)
            
            if not self.objStockChart.Continue: break
            if rcvOldestDate < fromDate: break
        
        # 분봉의 경우 날짜와 시간을 하나의 문자열로 합친 후 int 로 변환
        rcvData['date'] = list(map(
            lambda x, y: int('{}{:04}'.format(x, y)), 
            rcvData['date'], rcvData['time']
        ))
        del rcvData['time']
        caller.rcvData = rcvData
        
        return True
