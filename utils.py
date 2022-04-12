import unicodedata
from datetime import datetime, timedelta


def is_time_possible_to_trade():
    '''거래가능한 시간인지 확인하는 메소드
    
    :return: True:가능, False:불가능
    '''
    
    now = datetime.now()
    hm = int('{}{:02}'.format(now.hour, now.minute))
    if hm < 900 or hm > 1530: return False
    if now.weekday() >= 5: return False
    
    return True


def convert_dttm_to_int(dttm):
    
    return int(dttm.strftime('%Y%m%d%H%M'))


def available_latest_dttm():
    
    now = datetime.now()
    hm = int('{}{:02}'.format(now.hour, now.minute))
    
    # 장 중에는 최신 데이터 연속적으로 발생하므로 None 반환
    if is_time_possible_to_trade(): return None
    
    # 장 중이 아닌 경우
    latestDttm = now.replace(hour=15, minute=30)
    
    # 주말인 경우 (이외의 공휴일 체크 구현 필요)
    if now.weekday() >= 5:
        latestDttm = latestDttm - timedelta(days=now.weekday() - 4)
        return convert_dttm_to_int(latestDttm)
    
    # 주 중인 경우
    if hm > 1530: # 장 마감 후
        return convert_dttm_to_int(latestDttm)
    else: # 장 개장 전
        latestDttm = latestDttm - timedelta(days=1)
        if latestDttm.weekday() == 6:
            latestDttm = latestDttm - timedelta(days=2)
        
        return convert_dttm_to_int(latestDttm)
    
    return None


def preformat_cjk(string, width, align='<', fill=' '):
    
    count = (width - sum(1 + (unicodedata.east_asian_width(s) in "WF") for s in string))
    
    return {
        '>': lambda s: fill * count + s, 
        '<': lambda s: s + fill * count, 
        '^': lambda s: fill * (count / 2) + s + fill * (count / 2 + count % 2)
    }[align](string)
