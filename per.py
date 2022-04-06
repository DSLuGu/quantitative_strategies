'''PER(Price Earning Ration): 주가수익배수
    - 시가총액 / 당기순이익 또는 주가 / 주당순이익
    - 한 기업이 얻은 순이익 1원을 증권시장이 얼마의 가격으로 평가하고 있는가를 나타내는 수치.
    - 투자자들은 이를 척도로 서로 다른 주식의 상대적 가격을 파악.
    - 기본적으로 PER이 높으면 현 이익은 적지만 미래 수익 전망이 높아서 시장의 기대가 높은 것이라고 해석.
    - 반대로 PER이 낮으면 이익은 많지만 수익 전망이 낮아서 시장의 기대가 낮은 것이라고 해석.
'''
'''EPS(Earnings Per Share): 주당순이익
    - 기업의 순이익(당기순이익)을 유통주식수로 나눈 수치.
    - 당기순이익 규모가 늘면 높아지고, 전환사채의 주식 전환이나 증자, 분할로 인하여 주식 수가 많아지면 낮아지게 된다.
    - 분할이나 증자를 하지 않았는데 주당순이익이 시간이 지나며 계속 낮아지면 문제있는 기업이라 할 수 있다.
    - 예상 EPS와 예상 PER을 고보하면 해당 기업의 예상 주가를 구할 수 있다.
'''
'''부채비율
    - 부채 / 자기자본
    - 자산이 100원인 기업이 자기자본 50원, 부채 50원으로 구성되어 있으면 부채비율은 50/50 = 100%.
    - 기업이 자기자본이 총자산의 3분의 2(부채는 총자산의 3분의 1)일 경우 부채비율은 정확히 (1/3)/(2/3) = 50%.
'''
'''ROA(Return on Assets): 총자산순이익률
    - 순이익 / 총자산
    - 기업이 자산인 자본과 부채를 활용해 어느 정도 이익을 창출ㄹ했는지 알려주는 대표적인 수익성 지표.
    - 기업이 자산을 얼마나 효율적으로 운용했는지를 나타낸다.
'''
'''PBR(Price-Book Ratio): 주가순자산배수
    - 시가총액 / 자기자본 또는 주가 / 주당순자산
    - 기업의 순자산 1원을 증권시장이 얼마의 가격으로 평가하는지를 나타내는 수치.
    - 투자잗르은 이를 척도로 서로 다른 주식의 상대적 가격을 파악할 수 있다.
    - 기업이 얼마나 저평가되었는지 단순히 평가하고 싶을 때 PER과 함께 가장 많이 쓰이는 지표.
    - 기본적으로 PBR이 높으면(낮으면) 가진 자본은 적지만(많지만) 미래 수익 전망이 높아서(낮아서) 시장의 기대가 높은(낮은) 것이라고 해석.
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