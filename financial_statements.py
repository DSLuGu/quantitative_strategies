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


def get_fs(df, cols):
    
    if len(df) == 0: return None
    
    newColumns = []
    for i in df.columns:
        if re.search(r'\d{8}', i[0]):
            newColumns.append(i[0][-8:])
        else:
            newColumns.append(i[1])
    
    df.columns = newColumns
    df = df[[c for c in newColumns if not re.search(r'class\d', c) or not re.search(r'comment', c)]]
    
    years = [c for c in df.columns if re.search(r'\d{8}', c)][::-1]
    nonYears = [c for c in df.columns if not re.search(r'\d{8}', c)]
    
    rtnDF = pd.DataFrame()
    for y in years:
        appendDF = deepcopy(df[nonYears+[y]]).rename(columns={f'{y}': 'value'})
        appendDF['year'] = y
        rtnDF = pd.concat([rtnDF, appendDF])
    
    rtnDF = rtnDF[['year']+nonYears+['value']]
    for c in cols:
        if c not in rtnDF.columns: rtnDF[c] = ""
    else:
        rtnDF = rtnDF[['year']+cols+['value']]
    
    return rtnDF


def main():
    
    mainCols = ['concept_id', 'label_ko', 'label_en']
    
    dart.set_api_key(api_key=get_dart_api_key())
    
    corpList = dart.get_corp_list()
    for corp in ecorpList[90060:]:
        corpInfo = corp.load()
        print(corpInfo)
        
        try:
            fs = corp.extract_fs(bgn_de=corpInfo['est_dt'])
        except RuntimeError as e:
            print(f"[{corpInfo['corp_code']}]{corpInfo['corp_name']}", e)
            continue
        except dart.errors.errors.NotFoundConsolidated as e:
            print(f"[{corpInfo['corp_code']}]{corpInfo['corp_name']}", e)
            continue
        
        fsBS = fs.show('bs')
        bsLabels = fs.labels['bs']
        if fsBS is not None:
            fsBS.to_excel(f"{corpInfo['corp_code']}_before_bs.xlsx")
        fsIS = fs.show('is')
        isLabels = fs.labels['is']
        if fsIS is not None:
            fsIS.to_excel(f"{corpInfo['corp_code']}_before_is.xlsx")
        fsCIS = fs.show('cis')
        cisLabels = fs.labels['cis']
        if fsCIS is not None:
            fsCIS.to_excel(f"{corpInfo['corp_code']}_before_cis.xlsx")
        fsCF = fs.show('cf')
        cfLabels = fs.labels['cf']
        if fsCF is not None:
            fsCF.to_excel(f"{corpInfo['corp_code']}_before_cf.xlsx")
        
        if fsBS is not None:
            bsDF = get_fs(fsBS, mainCols)
            bsDF.to_excel(f"{corpInfo['corp_code']}_after_bs.xlsx", index=False)
        if fsIS is not None:
            isDF = get_fs(fsIS, mainCols)
            isDF.to_excel(f"{corpInfo['corp_code']}_after_is.xlsx", index=False)
        if fsCIS is not None:
            cisDF = get_fs(fsCIS, mainCols)
            cisDF.to_excel(f"{corpInfo['corp_code']}_after_cis.xlsx", index=False)
        if fsCF is not None:
            cfDF = get_fs(fsCF, mainCols)
            cfDF.to_excel(f"{corpInfo['corp_code']}_after_df.xlsx", index=False)
        
    return None


if __name__ == '__main__':
    
    main()
    