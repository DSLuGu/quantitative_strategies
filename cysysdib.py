import sys
from PyQt5.QtWidgets import *
import win32com.client


g_objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')


def check_plus_status(originalFunc):
    '''originalFunc 콜하기 전에 PLUS 연결 상태 체크하는 데코레이터'''
    def wrapper(*args, **kwargs):
        # 연결 여부 체크
        bConnect = g_objCpCybos.IsConnect
        
        if bConnect == 0:
            print("PLUS is not propertly connected...")
            return False
        
        return originalFunc(*args, **kwargs)
    
    return wrapper


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
