import os
import json
import argparse

import dart_fss as dart


CUR_DIR = os.path.dirname(os.path.realpath(__file__))


def define_argparser():
    
    p = argparse.ArgumentParser()
    
    p.add_argument('--key_fn', type=str, default='secrets.json')
    
    config = p.parse_args()
    
    return config

def main(config):
    
    with open(os.path.join(CUR_DIR, config.key_fn), 'r') as f:
        apiKey = json.load(f)['api_key']
    
    dart.set_api_key(api_key=apiKey)
    
    corpList = dart.get_corp_list()
    
    corpName = '오스템임플란트'
    
    corp = corpList.find_by_corp_name(corpName, exactly=True)[0]
    
    fs = corp.extract_fs(bgn_de='20200101')
    
    fs.save(corpName + '.xlsx')
    
    return None


if __name__ == '__main__':
    
    config = define_argparser()
    main(config)
    