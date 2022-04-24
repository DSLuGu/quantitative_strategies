import os
import gc
import sys

import argparse

import sqlite3
import pandas as pd

from tqdm import trange

from cysysdib import CpStockChart
from cputil import CpCodeMgr
from utils import (
    is_time_possible_to_trade, available_latest_dttm, preformat_cjk
)


def define_argparser():
    
    p = argparse.ArgumentParser(description='Creon DataReader')
    
    p.add_argument('--dbPath', required=True, type=str)
    p.add_argument('--tickUnit', type=str, default='day', help='{1min, 5min, day, week, month}')
    
    return p.parse_args()


class CreonDataReader:
    
    def __init__(self):
        
        self.objStockChart = CpStockChart()
        self.objCodeMgr = CpCodeMgr()
        self.rcvData = dict()
        
        self.svCodeDF = pd.DataFrame()
        self.dbCodeDF = pd.DataFrame()
        
        self.svStockCodeList = self.objCodeMgr.code_list(1) + self.objCodeMgr.code_list(2)
        self.svStockNameList = list(map(self.objCodeMgr.code_to_name, self.svStockCodeList))
        self.svStockCodeDF = pd.DataFrame({
            'stockCode': self.svStockCodeList, 'stockName': self.svStockNameList
        }, columns=('stockCode', 'stockName'))
        
        # 0: ('날짜', 'date'), 1: ('시간', 'time'), 2: ('시가', 'open'), 
        # 3: ('고가', 'high'), 4: ('저가', 'low'), 5: ('종가', 'close'), 
        # 6: ('전일대비', 'diff'), 8: ('거래량', 'volume'), 9: ('거래대금', 'tradeVolume'), 
        # 10: ('누적체결매도수량', 'cumNegVolume'), 11: ('누적체결매수수량', 'cumPosVolume')
        # ==> 분, 틱 요청일 때만 제공
        # 12: ('상장주식수', 'numOfListedShares'), 13: ('시가총액', 'marketCap'), 
        # 14: ('외국인주문한도수량', 'limitQuantityForForeignOrders'), 
        # 15: ('외국인주문가능수량', 'quantityAvailableForForeignOrders'), 
        # 16: ('외국인현보유수량', 'holdingQuantityByForeign'), 
        # 17: ('외국인현보유비율', 'holdingRatioByForeign'), 18: ('수정주가일자', 'adjustedPrice'), 
        # 19: ('수정주가비율', 'adjustedPriceRatio'), 20: ('기관순매수', 'netBuyingInstitution'), 
        # 21: ('기관누적순매수', 'cumNetBuyingInstitution'), 
        # 22: ('등락주선', 'advanceDeclineLine'), 23: ('등락비율', 'advanceDeclineRatio'), 
        # 24: ('예탁금', 'deposit'), 25: ('주식회전율', 'stockTurnoverRatio'), 
        # 26: ('거래성립률', 'transactionRatio'), 37: ('대비부호', 'contrastSign'), 
        
        self.rqFields = [
            0, 1, 2, 3, 4, 
            5, 6, 8, 9, 12, 
            13, 14, 15, 16, 17, 
            18, 19, 20, 21, 22, 
            23, 24, 25, 26, 37
        ]
        self.rqColumns = [
            'date', 'time', 'open', 'high', 'low', 
            'close', 'diff', 'volume', 'tradeVolume', 'numOfListedShares', 
            'marketCap', 'limitQuantityForForeignOrders', 'quantityAvailableForForeignOrders', 
            'holdingQuantityByForeign', 'holdingRatioByForeign', 
            'adjustedPrice', 'adjustedPriceRatio', 'netBuyingInstitution', 
            'cumNetBuyingInstitution', 'advanceDeclineLine', 
            'advanceDeclineRatio', 'deposit', 'stockTurnoverRatio', 
            'transactionRatio', 'contrastSign'
        ]
    
    def update_db(self, dbPath, tickUnit='day'):
        '''
        :param dbPath: DB 파일 경로
        :param tickUnit: '1min', '5min', 'day'. 이미 dbPath 가 존재할 경우, 입력값 무시하고 기존에 사용된 값 사용
        '''
        
        # Local DB 에 저장된 종목 정보 가져와서 DataFrame 으로 저장
        con = sqlite3.connect(dbPath)
        cursor = con.cursor()
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        dbStockCodeList = cursor.fetchall()
        for i in range(len(dbStockCodeList)):
            dbStockCodeList[i] = dbStockCodeList[i][0]
        dbStockNameList = list(map(self.objCodeMgr.code_to_name, dbStockCodeList))
        
        dbLatestList = []
        for dbStockCode in dbStockCodeList:
            cursor.execute("SELECT date FROM {} ORDER BY date DESC LIMIT 1;".format(dbStockCode))
            dbLatestList.append(cursor.fetchall()[0][0])
        
        # 현재 DB 에 저장된 'date' column 의 tickUnit 확인
        if dbLatestList:
            cursor.execute("SELECT date FROM {} ORDER BY date ASC LIMIT 2;".format(dbStockCodeList[0]))
            date0, date1 = cursor.fetchall()
            
            # 날짜가 분 단위인 경우
            if date0[0] > 99999999:
                if date1[0] - date0[0] == 5: # 5분 간격인 경우
                    tickUnit = '5min'
                else: # 1분 간격인 경우
                    tickUnit = '1min'
            elif data0[0] % 100 == 0: # 월봉인 경우
                tickUnit = 'month'
            elif data0[0] % 10 == 0: # 주봉인 경우
                tickUnit = 'week'
            else: # 일봉인 경우
                tickUnit = 'day'
        
        dbCodeDF = pd.DataFrame({
            'stockCode': dbStockCodeList, 'stockName': dbStockNameList, 'uptDttm': dbLatestList
        }, columns=('stockCode', 'stockName', 'uptDttm'))
        fetchStockCodeDF = self.svStockCodeDF
        
        if not is_time_possible_to_trade():
            latestDttm = available_latest_dttm()
            if tickUnit == 'day':
                latestDate = latestDttm // 10000
            # 이미 DB 데이터가 최신인 종목들은 가져올 목록에서 제외
            alreadyUpdatedStockCodes = dbStockCodeDF.loc[
                dbStockCodeDF['uptDttm']==latestDate
            ]['stockCode'].values
            fetchStockCodeDF = fetchStockCodeDF.loc[
                fetchStockCodeDF['stockCode'].apply(lambda x: x not in alreadyUpdatedStockCodes)
            ]
        
        if tickUnit == '1min':
            count = 200000 # 서버 데이터 최대 reach 약 18.5만 이므로 (18/02/25 기준)
            tickRange = 1
        elif tickUnit == '5min':
            count = 100000
            tickRange = 5
        elif tickUnit == 'day':
            count = 10000
        elif tickUnit == 'week':
            count = 2000
        elif tickUnit == 'month':
            count = 500
        
        with sqlite3.connect(dbPath) as con:
            cursor = con.cursor()
            tqdmRange = trange(len(fetchStockCodeDF), ncols=100)
            for i in tqdmRange:
                stockCode = fetchStockCodeDF.iloc[i]
                updateStatusMsg = "[{}] {}".format(stockCode[0], stockCode[1])
                tqdmRange.set_description(preformat_cjk(updateStatusMsg, 25))
                
                fromDate = 0
                if stockCode[0] in dbStockCodeDF['stockCode'].tolist():
                    cursor.execute("SELECT date FROM {} ORDER BY date DESC LIMIT 1;".format(stockCode[0]))
                    fromDate = cursor.fetchall()[0][0]
                
                if tickUnit == 'day':
                    if self.objStockChart.request_dwm(
                        stockCode[0], ord('D'), count, self, fromDate
                    ) == False: continue
                elif tickUnit == 'week':
                    if self.objStockChart.request_dwm(
                        stockCode[0], ord('W'), count, self, fromDate
                    ) == False: continue
                elif tickUnit == 'month':
                    if self.objStockChart.request_dwm(
                        stockCode[0], ord('M'), count, self, fromDate
                    ) == False: continue
                elif tickUnit == '1min' or tickUnit == '5min':
                    if self.objStockChart.request_mt(
                        stockCode[0], ord('m'), tickRange, count, self, fromDate
                    ) == False: continue
                
                df = pd.DataFrame(self.rcvData)
                
                # 기존 DB 와 겹치는 부분 제거
                if fromDate != 0:
                    df = df.loc[:fromDate]
                    df = df.iloc[:-1]
                
                # 뒤집어서 저장 (결과적으로 date 기준 오름차순으로 저장됨)
                df.iloc[::-1].to_sql(stockCode[0], con, if_exists='append')
                
                # 메모리 overflow 방지
                del df; gc.collect()
        
        return None


def main(config):
    
    creon = CreonDataReader()
    creon.update_db(config.dbPath, config.tickUnit)
    
    return None


if __name__ == '__main__':
    
    config = define_argparser()
    main(config)
    