# -*- coding=utf-8 -*-
import pandas as pd
import os
import time


'''
'''

# main folder path
path = str(os.path.dirname(os.path.abspath(__file__)))


def split_by_sheet(input_path=(''.join([path, '/input/']))):

    start = time.process_time()
    m, n = 0, 0
    for dirpath, dirnames, filenames in os.walk(input_path):
        
        for filename in filenames:
            m += 1
            print('BOOK', m, ':' ,filename)
            sheet_name_dict = pd.read_excel(os.path.join(dirpath, filename), None)
            sheet_name_list = sheet_name_dict.keys()
            
            for sheet_name in sheet_name_list:
                n += 1
                print('sheet', n, ':' ,sheet_name)
                df = pd.read_excel(os.path.join(dirpath, filename), sheet_name=sheet_name, skiprows=0)
                df.to_excel((''.join([path, '/output/', filename, '-', sheet_name, '.xlsx'])),sheet_name = sheet_name, index=False)
    print('Done')


if __name__ == "__main__":
    split_by_sheet()
        

