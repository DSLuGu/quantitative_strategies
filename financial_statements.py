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
    
    df.columns = newColumns
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


def get_is(df):
    
    if len(df) == 0: return None
    
    newColumns = []
    for i in df.columns:
        if re.search(r'\d{8}\-\d{8}', i[0]):
            newColumns.append(i[0][-8:])
        else:
            newColumns.append(i[1])
    
    df.columns = newColumns
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


def get_cis(df):
    
    if len(df) == 0: return None
    
    newColumns = []
    for i in df.columns:
        if re.search(r'\d{8}\-\d{8}', i[0]):
            newColumns.append(i[0][-8:])
        else:
            newColumns.append(i[1])
    
    df.columns = newColumns
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


def get_cf(df):
    
    if len(df) == 0: return None
    
    newColumns = []
    for i in df.columns:
        if re.search(r'\d{8}\-\d{8}', i[0]):
            newColumns.append(i[0][-8:])
        else:
            newColumns.append(i[1])
    
    df.columns = newColumns
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
    for corp in corpList[90000:]:
        corpInfo = corp.load()
        print(corpInfo)
        
        try:
            fs = corp.extract_fs(bgn_de=corpInfo['est_dt'])
        except RuntimeError as e:
            print(f"[{corpInfo['corp_code']}]{corpInfo['corp_name']}", e)
            continue
        
        fsBS = fs.show('bs')
        bsLabels = fs.labels['bs']
        fsBS.to_excel(f"{corpInfo['corp_code']}_before_bs.xlsx", index=False)
        fsIS = fs.show('is')
        isLabels = fs.labels['is']
        fsIS.to_excel(f"{corpInfo['corp_code']}before_is.xlsx", index=False)
        fsCIS = fs.show('cis')
        cisLabels = fs.labels['cis']
        fsCIS.to_excel(f"{corpInfo['corp_code']}before_cis.xlsx", index=False)
        fsCF = fs.show('cf')
        cfLabels = fs.labels['cf']
        fsCF.to_excel(f"{corpInfo['corp_code']}before_cf.xlsx", index=False)
        
        bsDF = get_bs(fsBS)
        bsDF.to_excel(f"{corpInfo['corp_code']}after_bs.xlsx", index=False)
        isDF = get_is(fsIS)
        isDF.to_excel(f"{corpInfo['corp_code']}after_is.xlsx", index=False)
        cisDF = get_cis(fsCIS)
        cisDF.to_excel(f"{corpInfo['corp_code']}after_cis.xlsx", index=False)
        cfDF = get_cf(fsCF)
        cfDF.to_excel(f"{corpInfo['corp_code']}after_df.xlsx", index=False)
        
    return None


if __name__ == '__main__':
    
    main()
    