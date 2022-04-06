'''PER(Price Earning Ration): 주가수익비율
    - 주가가 그 기업의 1주당 수익의 몇 배가 되는지를 나타내는 지표
    - 기업이 시장에서 저평가되었는지, 고평가되었는지를 판단하는 기준
'''

'''EPS(Earnings Per Share): 주당순이익
    - 기업의 순이익(당기순이익)을 유통주식수로 나눈 수치
    - 당기순이익 규모가 늘면 높아지고, 전환사채의 주식 전환이나 증자, 분할로 인하여 주식 수가 많아지면 낮아지게 된다.
    - 분할이나 증자를 하지 않았는데 주당순이익이 시간이 지나며 계속 낮아지면 문제있는 기업이라 할 수 있다.
    - 예상 EPS와 예상 PER을 고보하면 해당 기업의 예상 주가를 구할 수 있다.
'''
import win32com.client


INST_MARKET_EYE = win32com.client.Dispatch('CpSysDib.MarketEye')


def get_per(stockCode:str):
    
    INST_MARKET_EYE.SetInputValue(0, (4, 67, 70, 111)) # (현재가, PER, EPS, 최근분기년월)
    INST_MARKET_EYE.SetInputValue(1, stockCode)
    
    INST_MARKET_EYE.BlockRequest()
    
    # curPrice = INST_MARKET_EYE.GetDataValue(0, 0)
    per = INST_MARKET_EYE.GetDataValue(1, 0)
    # eps = INST_MARKET_EYE.GetDataValue(2, 0)
    # lastQuarterYM = INST_MARKET_EYE.GetDataValue(3, 0)
    
    return per