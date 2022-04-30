import os
import sys
import json
from openpyxl import Workbook, load_workbook

import pandas as pd

import dart_fss as dart


CUR_DIR = os.path.dirname(os.path.abspath(__file__))


def get_dart_api_key():
    
    with open(os.path.join(CUR_DIR, 'secrets.json')) as f:
        apiKey = json.load(f)['api_key']
    
    return apiKey


def main():
    
    dart.set_api_key(api_key=get_dart_api_key())
    
    corpList = dart.get_corp_list()
    
    
    
    return None


if __name__ == '__main__':
    
    main()
    