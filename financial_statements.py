import os
import sys

import re
import json
from copy import deepcopy

import pandas as pd
from openpyxl import Workbook, load_workbook

import dart_fss as dart


CUR_DIR = os.path.dirname(os.path.abspath(__file__))


def get_dart_api_key():
    
    with open(os.path.join(CUR_DIR, 'secrets.json')) as f:
        apiKey = json.load(f)['api_key']
    
    return apiKey


def get_bs(df):
    
    if len(df) == 0: return None
    
    newColumns = []
    for i in df.columns:
        if re.search(r'\d{8}', i[0]):
            newColumns.append(i[0])
        else:
            newColumns.append(i[1])
    
    fsBS.columns = newColumns
    df = df[[c for c in newColumns if not re.search(r'class\d', c)]]
    
    years = [c for c in df.columns if re.search(r'\d{8}', c)][::-1]
    nonYears = [c for c in df.columns if not re.search(r'\d{8}', c)]
    
    rtnDF = pd.DataFrame()
    for y in years:
        appendDF = deepcopy(df[nonYears+[y]]).rename(columns={f'{y}': 'value'})
        appendDF['year'] = y
        rtnDF = pd.concat([rtnDF, appendDF])
    
    rtnDF = rtnDF[['year']+nonYears+['value']]
    
    return rtnDF


def main():
    
    dart.set_api_key(api_key=get_dart_api_key())
    
    corpList = dart.get_corp_list()
    
    # include_delisting(상장폐지 주식 포함 검색)
    corp = corpList.find_by_stock_code(stock_code='005930', include_delisting=True)
    corpInfo = corp.load()
    fs = corp.extract_fs(bgn_de=corpInfo['est_dt'])
    
    fsBS = fs.show('bs')
    # bsLabels = fs.labels['bs']
    fsIS = fs.show('is')
    # isLabels = fs.labels['is']
    fsCIS = fs.show('cis')
    # cisLabels = fs.labels['cis']
    
    get_bs(fsBS)
    
    return None


if __name__ == '__main__':
    
    main()
    