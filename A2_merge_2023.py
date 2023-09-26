import pandas as pd
import os
from utils.utils.sql import connSQL
from sqlalchemy import text
engine = connSQL('chemiPrim_Test')


pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.max_colwidth', 30)
pd.set_option('display.max_columns', 400)
pd.options.display.width = 0 

df = pd.read_excel('2023_merged/detail_list_2023.xlsx')
df = df.fillna('-')
for col in ['112年底未更新前系統數量', '110更新', '111更新', '112更新']:
    df.loc[:, col] = df[col].apply(lambda i: int(i) if type(i) != str else i) 

data_all = []
update_col = '112更新'
for irow, row in df.iterrows():
    # 取得資料及名稱後
    filename = row['資料集編號']
    fullpath = f'2023_processed/{filename}.xlsx'
    source = row['source']
    # if filename != 'A_usa_01_a': continue
    
    # 檢查檔案是否存在，若存在，則更新 欄位 `update_col` 的數量
    status = os.path.isfile(fullpath)
    if not status:
        # print(f'{fullpath} not found')
        data = pd.DataFrame([])
        len_data = None
    else:
        data = pd.read_excel(fullpath, index_col=None)
        len_data = len(data)
        df.loc[irow, update_col] = len_data
        
        if 'CASNo' not in data.columns:
            print(filename, ' no CASNo')
            if 'ChemicalChnName' in data.columns:
                print(filename, ' only ChemicalChnName')
                ChemicalChnName = tuple(data['ChemicalChnName'].tolist())
                sql_get_casno = text(
                    '''
                    select CASNoMatch, ChemiChnNameMatch from ChemiMatchMapping
                    where ChemiChnNameMatch in :ChemicalChnName
                    '''
                )             
                match_in_db = pd.read_sql(sql_get_casno, params={'ChemicalChnName': ChemicalChnName}, con=engine)
                match_in_db = match_in_db.drop_duplicates(subset=['ChemiChnNameMatch'])
                data = (
                    data
                    .merge(match_in_db, how='left', left_on='ChemicalChnName', right_on='ChemiChnNameMatch')
                    .rename(columns={'CASNoMatch': 'CASNo'})
                    .drop(columns=['ChemiChnNameMatch'])
                    )
                
                            
            elif 'ChemicalEngName' in data.columns:
                # print(data.columns, source, filename)
                print(filename, ' only ChemicalEngName')
                ChemicalEngName = tuple(data['ChemicalEngName'].tolist())
                
                sql_get_casno = text(
                    '''
                    select CASNoMatch, ChemiEngNameMatch from ChemiMatchMapping
                    where ChemiEngNameMatch in :ChemicalEngName
                    '''
                )
                match_in_db = pd.read_sql(sql_get_casno, params={'ChemicalEngName': ChemicalEngName}, con=engine)
                match_in_db = match_in_db.drop_duplicates(subset=['ChemiEngNameMatch'])
                data = (
                    data
                    .merge(match_in_db, how='left', left_on='ChemicalEngName', right_on='ChemiEngNameMatch')
                    .rename(columns={'CASNoMatch': 'CASNo'})
                    .drop(columns=['ChemiEngNameMatch'])
                    )
    data = data.fillna('-')
    
    #region 資料清洗
    
    # 移除字尾 \u3000
    for col in ['ChemicalChnName', 'ChemicalEngName']:
        if col in data.columns:
            data.loc[:, col] = data[col].apply(lambda i: i.replace('\u3000', ''))
    # print(filename, data)
    
    data_all.append(data)
    #endregion

    
    
    # 取得資料庫
    sql = text('''
        SELECT ChemicalChnName, ChemicalEngName, CASNo FROM AttnChemicalList
        WHERE source = :source
    ''')
    params = {'source': source}
    data_in_db = pd.read_sql(sql, params=params, con=engine)
    len_data_in_db = len(data_in_db)

    chnname_data_not_in_db = []
    engname_data_not_in_db = []
    casno_data_not_in_db = []
    if 'ChemicalChnName' in data.columns:
        chnname_data_not_in_db = set(data['ChemicalChnName'].tolist()).difference(set(data_in_db['ChemicalChnName'].tolist()))
    if 'ChemicalEngName' in data.columns:
        engname_data_not_in_db = set(data['ChemicalEngName'].tolist()).difference(set(data_in_db['ChemicalEngName'].tolist()))
    if 'CASNo' in data.columns:
        casno_data_not_in_db = set(data['CASNo'].tolist()).difference(set(data_in_db['CASNo'].tolist()))
    
    df.loc[irow, '更新後的中文名 不在 更新前中文名 的 數量'] = len(chnname_data_not_in_db)
    df.loc[irow, '更新後的英文名 不在 更新前英文名 的 數量'] = len(engname_data_not_in_db)
    df.loc[irow, '更新後的Casno 不在 更新前Casno 的 數量'] = len(casno_data_not_in_db)
    df.loc[irow, '112更新'] = len_data if len_data_in_db is not None else ''
    df.loc[irow, '清單數量相同'] = 'v' if len_data == len_data_in_db else ''
    

df.to_excel('2023_merged/detail_list_2023_new.xlsx', index=False)
data_all = pd.concat(data_all, ignore_index=True)
data_all.to_excel('2023_merged/AttnChemicalList@今年度整理.xlsx', index=False)
    

# 取出資料庫全部資料
sql = ''' select ChemicalChnName,ChemicalEngName,CASNo,EmsNo,Source,SourceOri,Type,Unit from AttnChemicalList '''
df2 = pd.read_sql(sql, con=engine)
df2.to_excel('2023_merged/AttnChemicalList@db20230926.xlsx', index=False)

