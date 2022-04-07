import sys
import win32com.client


class CpCodeMgr:
    
    def __init__(self):
        
        self.objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
    
    def code_list(self, market):
        '''market에 해당하는 종목코드 리스트 반환하는 메소드
        
        :param market: 1: kospi, 2:kosdaq, ...
        :return: market에 해당하는 stockCode list
        '''
        
        return self.objCodeMgr.GetStockListByMarket(market)
    
    def section_code(self, stockCode):
        '''부구분코드를 반환하는 메소드'''
        
        return self.objCodeMgr.GetStockSectionKind(stockCode)
    
    def code_name(self, stockCode):
        '''종목명을 반환하는 메소드'''
        
        return self.objCodeMgr.CodeToName(stockCode)
    
    def status(self, stockCode):
        '''주식상태를 반환하는 메소드'''
        
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
    