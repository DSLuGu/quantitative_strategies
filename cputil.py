import os
import sys
import win32com.client
from pywinauto import application


class CpCodeMgr:
    
    def __init__(self):
        
        self.objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        
        return None
    
    def code_list(self, market):
        '''market에 해당하는 종목코드 리스트 반환하는 메소드
        
        :param market: 0: 구분없음, 1: kospi, 2: kosdaq, 3: freeboard, 4: krx, 5: konex, 
        :return: market에 해당하는 stockCode list
        '''
        
        return self.objCodeMgr.GetStockListByMarket(market)
    
    def section(self, stockCode):
        '''부구분코드를 반환하는 메소드
        
        :return: 0: 구분없음, 1: 주권, 2: 투자회사, 3: 부동산투자회사, 4: 선박투자회사, 
            5: 사회간접자본투융자회사, 6: 주식예탁증서, 7: 신수인수권증권, 8: 신주인수권증서, 
            9: 주식워런트증권, 10: 상장지수펀드(ETF), 11: 수익증권, 12: 해외ETF, 
            13: 외국주권, 14: 선물, 15: 옵션, 
        '''
        
        return self.objCodeMgr.GetStockSectionKind(stockCode)
    
    def code_to_name(self, stockCode):
        '''종목명을 반환하는 메소드'''
        
        return self.objCodeMgr.CodeToName(stockCode)
    
    def control(self, stockCode):
        '''감리구분을 반환하는 메소드
        
        :return: 0: 정상, 1: 주의, 2: 경고, 3: 위험예고, 4: 위험, 
        '''
        
        return self.objCodeMgr.GetStockControlKind(stockCode)
    
    def supervision(self, stockCode):
        '''관리구분을 반환하는 메소드
        
        :return: 0: 일반종목, 1: 관리, 
        '''
        
        return self.objCodeMgr.GetStockSupervisionKind(stockCode)
    
    def status(self, stockCode):
        '''주식상태를 반환하는 메소드
        
        :return: 0: 정상, 1: 거래정지, 2: 거래중단, 
        '''
        
        return self.objCodeMgr.GetStockStatusKind(stockCode)
    
    def listed_date(self, stockCode):
        '''상장일을 반환하는 메소드'''
        
        return self.onjCodeMgr.GetStockListedDate(stockCode)
    
    def start_time(self):
        '''장 시작 시각 반환하는 메소드'''
        
        return self.GetMarketStartTime()
    
    def end_time(self):
        '''장 마감 시각 반환하는 메소드'''
        
        return self.GetMarketEndTime()


class Creon:
    
    def __init__(self):
        
        self.objCpUtilCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        
        return None
    
    def kill_client(self):
        
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')
        
        return None
    
    def connect(self, _id, _pwd, _pwdcert):
        
        if not self.connected():
            self.disconnect()
            self.kill_client()
            app = application.Application()
            app.start(
                'C:\CREON\STARTER\coStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                    id=_id, pwd=_pwd, pwdcert:_pwdcert
                )
            )
        while not self.connected():
            time.sleep(1)
        
        return True
    
    def connected(self):
        
        bConnected = self.objCpUtilCpCybos.IsConnect
        if bConnected == 0: return False
        
        return True
    
    def disconnected(self):
        
        if self.connected():
            self.objCpUtilCpCybos.PlusDisconnect()