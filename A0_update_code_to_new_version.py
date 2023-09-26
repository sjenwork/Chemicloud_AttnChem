'''
This code is used to update final.py to final_newname.py

final_newname.py update code and output file name with meaningfull regular.
'''

import os
import re
import pprint
import pandas as pd

pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.max_colwidth', 30)
pd.set_option('display.max_columns', 400)
pd.options.display.width = 0


def clear_file():
    with open('final_3.py', 'w') as f:
        f.write('''
import os
import pandas as pd
import pdfplumber
# import tabula
import numpy as np

input_path = '2023_input'
processed_path = '2023_processed'

pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.max_colwidth', 30)
pd.set_option('display.max_columns', 400)
pd.options.display.width = 0 


'''
)


            
    
def read_detail_list():
    def to_str(i):
        if type(i) == str:
            return i
        else:
            return str(int(i))
    df = pd.read_excel('detail_list_2023.xlsx', sheet_name='工作表2')
    df  = df.fillna('')
    int_col = ['國內/外(0/1)', '110更新', '111更新', '112年底未更新前系統數量', '112更新']
    df[int_col] = df[int_col].applymap(to_str)
    return df

df = read_detail_list()

clear_file()
for ii, row in df.iterrows():
    # print(ii)
    # if ii > 10: continue
    
    
    #region 檢查函數 和 清單 是否一致
    def compare_func_vs_list(row):
        
        def get_func_from_py():
            # 讀取函數，要來改函數中對應的名稱
            funclist = []
            with open('final.py', 'r') as f:
                for line in f.readlines():
                    # print(line)
                    if len(funclist) == 0 and not re.match(r'^def', line):
                        continue
                    if re.match(r'^#endregion', line):
                        continue
                    
                    if re.match(r'^def', line):
                        func_name = re.findall(r'def (.*)\(', line)[0]
                        funclist.append({'func_name': func_name,  'func_content': ''})
                        funclist[-1]['func_name'] = func_name
                        funclist[-1]['func_content'] += line
                    elif re.match('^#region', line):
                        pass
                    else:
                        funclist[-1]['func_content'] += line
                        
            # funclist[-1]['func_content'] = funclist[-1]['func_content'].replace('#region', '')
            # if len(funclist[-1]['func_content'])<100:
            #     print(funclist)
            return funclist        
        

        def write_file(content):
            with open('final_2.py', 'a+') as f:
                f.write(content)
                
        
        source_name = row['source']
        funcname_no = row['python函數與對應原檔檔案編號']
        new_name = row['更新編號名稱']
        method = row['112年程式處理狀況'] + '; ' + row['112年更新方法']
        funclist = get_func_from_py()
        # 找出函數
        get_func_name = [i for i in funclist if i['func_name'] in [f'n{funcname_no}', f'n0{funcname_no}']]
        
        # 補上沒有的函數
        if len(get_func_name) == 0:
            print(source_name, ' --> ', funcname_no, ' --> ', method)
            get_func_name = [{
                    'func_name': f'n{funcname_no}', 
                    'func_content': f'''    #{method} ''',
                }]
            funclist.append(get_func_name[0])
            
        # 將函數內，新的名稱寫入
        write_fn = get_func_name[0]['func_content']
        write_fn = (
            write_fn
            .replace(f'n{funcname_no}', f'{new_name}')
            .replace(f'{funcname_no}', f'{new_name}')
            )
        write_file( f'''\n#region {new_name} {source_name}\n{write_fn}\n#endregion\n''')
        return funclist
    #endregion

    #region 檢查輸入輸出檔和清單是否一致
    
    def compare_file_vs_list(row, which='input', how='comfirm'):
        
        def read_input_and_out_file(which):
            # 讀取輸出結果，要來更改名稱
            if which == 'input':
                data = os.listdir('2023_input')
            if which == 'result':
                data = os.listdir('2023_processed')
            return data
        source_name = row['source']
        funcname_no = row['python函數與對應原檔檔案編號']
        new_name = row['更新編號名稱']
        method = row['112年程式處理狀況'] + '; ' + row['112年更新方法']       
        
        data = read_input_and_out_file(which) 
        if which == 'input':
            get_input_files = [i for i in data if i.startswith(f'{funcname_no}')]
            if how == 'comfirm':
                if len(get_input_files) == 0:
                    if (
                        ('無需處理' not in method) and 
                        ('爬蟲' not in method) and  
                        ('沿用' not in method) and 
                        ('結果沒變' not in method) and
                        ('api' not in method)
                        ):
                        print(source_name, ' --> ', funcname_no, ' --> ', method)
        
            if how == 'rename':
                # print(f'''{source_name} -> {funcname_no} -> {get_input_files}''')
                old_name = get_input_files
                new_name = [i.replace(funcname_no, new_name) for i in old_name]
                print(list(zip(old_name, new_name)))
                # 重新命名
                for old, new in zip(old_name, new_name):
                    os.rename(f'2023_input/{old}', f'2023_input/{new}')
    
        if which == 'result':
            get_input_files = [i for i in data if i.startswith(f'{funcname_no}')]
            if how == 'comfirm':
                if len(get_input_files) == 0:
                    print(f'''{source_name} -> {funcname_no} -> {get_input_files} <<-- {method}''')
    #endregion
        
    
    funclist = compare_func_vs_list(row)
    # compare_file_vs_list(row, which='input', how='rename') # 完成
    # compare_file_vs_list(row, which='result', how='comfirm')
    

