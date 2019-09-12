# -*- coding=utf-8 -*-
import pandas as pd
import os
import time


'''
    1. 循环读取文件夹中excel、sheet，拼接dataframe，以当下时间建立文件夹，并将合并后的文档存储其中；
    2. 输入：建立input文件夹，将被合并excel文件放入其中， 参数填入被合并表格中的一项公共列标题；
    3. 输出：demo 完成合并：0个BOOKs, 0个 SHEETs，总计：0行，0列。耗时：0.00秒，存储至文件夹：output-11-00-01；
'''

# main folder path
path = str(os.path.dirname(os.path.abspath(__file__)))
# str out time
output_time = time.strftime("%H-%M-%S", time.localtime())




def openfile(oneofitermname, input_path=(''.join([path, '/input/']))):
    '''循环读取文件夹中excel并组合成df_rtn'''
    start = time.process_time()
    df_rtn = pd.DataFrame()
    m, n, k = 0, 0, 0
    # get file name
    for dirpath, dirnames, filenames in os.walk(input_path):

        # get sheet name
        for filename in filenames:
            m += 1
            sheet_name_dict = pd.read_excel(os.path.join(dirpath, filename), None)
            sheet_name_list = sheet_name_dict.keys()
            time_cost = '{:.2f}'.format((time.process_time() - start))
            print('Book', m, '时间:', time_cost, 'second >>',  filename)

            # get real col line number
            for sheet_name in sheet_name_list:
                n += 1
                time_cost = '{:.2f}'.format((time.process_time() - start))

                # set invalid line number
                test_line_num = 12
                df = pd.read_excel(os.path.join(dirpath, filename), nrows=test_line_num, sheet_name=sheet_name, skiprows=0)
                time_cost = '{:.2f}'.format((time.process_time() - start))
                print('Sheet', n,  '时间:', time_cost, 'second >>', sheet_name)

                # test col name, if true read and append df
                j = 0
                for i in range(len(df)):

                    # fist row, get the col name list
                    if i == 0:
                        col_list = df.columns.tolist()
                        
                        if oneofitermname in col_list:
                            k += 1
                            df = pd.read_excel(os.path.join(dirpath, filename), sheet_name=sheet_name, skiprows=(i))
                            df['book'], df['sheet'] = filename, sheet_name
                            df_rtn = df_rtn.append(df, ignore_index=True, sort = False)
                            time_cost = '{:.2f}'.format((time.process_time() - start))
                            print('扫描到有效行，添加到df', '时间:', time_cost, 'second')
                            break 

                    # from second row, get the col name list
                    if i >= 1:
                        col_list = df.iloc[(i-1):i].values.tolist()
                        col_list = col_list[0]
                        
                        if oneofitermname in col_list:
                            k += 1
                            df = pd.read_excel(os.path.join(dirpath, filename), sheet_name=sheet_name, skiprows=(i))
                            df['book'], df['sheet'] = filename, sheet_name
                            df_rtn = df_rtn.append(df, ignore_index=True, sort = False)
                            time_cost = '{:.2f}'.format((time.process_time() - start))
                            print('扫描到有效行，添加到df', '时间:', time_cost, 'second')
                            break
                    j += 1
                    print('无效行 >>', j)

    
    return  m, n, k, df_rtn





if __name__ == "__main__":
    
    try:
        os.remove(''.join([path, '/input/', '.DS_Store']))
    
    except:
        pass
    
    finally:
        # 运行
        input = input('please insert a co-col name: ')
        start = time.process_time()
        time_cost = '{:.2f}'.format((time.process_time() - start))
        print('Start:', time_cost, 'second')
        x = openfile(input)
        
        # 输出
        output_folder = os.mkdir(''.join([path, '/output-', output_time]))
        x[3].to_excel((''.join([path, '/output-', output_time, '/', 'merge_by_col_name_', input, output_time, '.xlsx'])), index=False)
        
        # 简报
        time_cost = '{:.2f}'.format((time.process_time() - start))
        print(''.join(['完成合并，累计读取', str(x[0]), '个BOOKs, ', str(x[1]), '个SHEETs，', str(x[2]) ,'个有效，', '总计：', str(x[3].shape[0]), '行，', str(x[3].shape[1]), '列。', '耗时：', time_cost, '秒，', ' 存储至文件夹：', 'output-', output_time]))