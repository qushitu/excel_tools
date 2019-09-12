#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pandas as pd
import os
import time

'''
功能：根据excel某列中不同项目拆分表格，生成多个文件；
格式：表头行必须在第一行
输入：input文件夹
输出：output文件夹
'''

def split_by_a_col_name():
    '''Split EXCEL by a Slected COLUMN'''
    
    # get pwd and time now
    path = str(os.path.dirname(os.path.abspath(__file__)))
    date = time.strftime('%Y%m%d-%H%M',time.localtime(time.time()))
	
	# def foo to read .xls*
    def read_xls(dir):
        df = pd.DataFrame()
        for dirpath, dirnames, filenames in os.walk(dir):
            for filename in filenames:
                df_i = pd.read_excel((os.path.join(dirpath, filename)), skiprows=0)
                df = df.append(df_i,ignore_index=True)
        return df
	
	# read excel
    dir = path + r'/input'
    df_in = read_xls(dir)

	# get all column name
    col_list = df_in.columns.values.tolist()
    if len(col_list) == 0:
        print('no data, please check your sheet, goodbye')
        exit(-1)

    # print column name list for user select
    for i, val in enumerate(col_list):
        print(i + 1, val)
	
    # input col No. until True
    

    while True:
        try:
            column_name = col_list[int(input('Please input COLUMN No. >>> ')) - 1]
            print('Spilt by COLUMN "' + column_name + '": >>>>>')
            break
        except:
            print('Wrong COLUMN NO., please try again.')

	# get group word name list
    df_list = df_in[column_name].drop_duplicates().dropna().tolist()

	# group by selected column
    df_base = df_in.groupby([column_name])
	
    # loop and split
    for i in range(len(df_list)):
        df_i = df_base.get_group(df_list[i])
        df_i.to_excel(path + r'/output/' + 'By_' + column_name + '_' + str(df_list[i]) + '_' + date + '.xlsx')
        print('Output >>>>> ' + 'By_' + column_name + ':' + str(i+1) + '. ' + str(df_list[i]))
		
    # print output path
    print('Split is complete, total ' +  str(len(df_list)) + ' excel, path is:')
    print(path + r'/output/' )

if __name__ == '__main__':
    split_by_a_col_name()
