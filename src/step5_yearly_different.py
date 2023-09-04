import pandas as pd
import numpy as np
import difflib
import multiprocessing as mp
import os
pd.options.mode.chained_assignment = None  # default='warn'

def match_ratio(df, col1, col2):
    return difflib.SequenceMatcher(None, df[col1], df[col2]).ratio()
def init(l,search_list_,search_df_):
    global lock
    lock = l
    global search_list
    search_list = search_list_
    global search_df
    search_df = search_df_

def main():
    #先清空舊總route表
    if os.path.exists(f"diff_ratio.csv"):#首次有表頭
        os.remove("diff_ratio.csv")
    print("old emission.csv file cleared")
   
    #給今年id
    df1 = pd.read_excel("../output/Merged.xlsx", header = 0)#今年
    Unique_array = np.arange(len(df1))+1 #Unique_ID/ Unique_ship_count
    df1['id'] = (Unique_array).astype(str)#給ID

    #為 今年的表 製造 判斷欄位
    df1['judgment'] = df1['Source'].str[:]+'_'+df1['CASNo'].str[:]+'_'+df1['ChemicalChnName'].str[:]+'_'+df1['ChemicalEngName'].str[:]
    search_df =df1[['id','judgment']]
    b = df1['judgment'].astype(str).tolist()


    #mp設定
    pool_size = 4
    l = mp.Lock()
    search_list = b
    pool = mp.Pool(pool_size, initializer=init, initargs=(l,search_list,search_df,))
    # chunk_size = 50 #行數，default:1000000，初始參考設定10 * pool_size

    count = 0
    count_row = 0
    count_row_list = []
    
    # for file_chunk in pd.read_csv("../explain_and_structure/AttnList.xlsx", chunksize=chunk_size,low_memory=False) :#去年
    df0 = pd.read_excel("../explain_and_structure/AttnList.xlsx")#去年
    chunk_times = 400#月大切越多次，chunk_size越小，越無效率

    # for file_chunk in np.split(df0, len(df0) // chunk_size):#去年
    for file_chunk in np.array_split(df0,chunk_times):#去年

        if count >=0 :# 各檔案分割幾次為止(測試用), default  >=0
            # chunk_size= len(df0) //chunk_times
            # line = count * chunk_size
            # Split chunk evenly. It's better to use this method if every chunk takes similar time.
            count_row_array = pool.map(processing_chunk, np.array_split(file_chunk, pool_size))
            print(count_row_array)
            count += 1
            for count_row in count_row_array:
                # print(count_row)
                count_row_list.append(count_row[0])
                count_row = sum(count_row_list)
                print("total rows processed cumulated : " + str(count_row))
            completed_rate= count_row/len(df0)
            print(f"Processed {str(round(completed_rate,4)*100)}%")
        else:
            break
    count_row = sum(count_row_list)
    print(count_row)

    pool.close()
    pool.join()


#處理多進程chunk
def processing_chunk(chunk):
    df = chunk
    #為 去年的表 製造 判斷欄位
    df['judgment'] = df['Source'].str[:]+'_'+df['CASNo'].str[:]+'_'+df['ChemicalChnName'].str[:]+'_'+df['ChemicalEngName'].str[:]
    df['judgment'] = df['judgment'].astype(str)
    df =df[['AttnSN','judgment']]
    #獲得most_match及match_ratio
    try:
        df['most_match'] = df['judgment'].apply(lambda x: sorted(search_list, key=lambda y: difflib.SequenceMatcher(None, y, str(x)).ratio(), reverse=True)[0])
    except:
        df['most_match'] = "error"
        print("most_match produce error!!")
    
    df['match_ratio'] = df.apply(match_ratio,
                            args=('most_match', 'judgment'),
                            axis=1)
    df = pd.merge(df,search_df, left_on=['most_match'],right_on=['judgment'],how='left')
    df =df[["AttnSN","id","match_ratio","judgment_x","most_match"]]
    print(df)

    if not os.path.exists(f"diff_ratio.csv"):#首次有表頭
        lock.acquire()
        df.to_csv(f"diff_ratio.csv",index=False, sep=",", mode='w', header=True)#,index=False
        lock.release()
    else:#非首次用append
        lock.acquire()
        df.to_csv(f"diff_ratio.csv",index=False, sep=",", mode='a', header=False)#,index=False
        lock.release()

    count_row = len(df)
    return [count_row]

#go
if __name__ == "__main__":
    main()
    


# #今年的表 judgment 做成 lsit
# a = df1['id'].tolist()
# b = df1['judgment'].astype(str).tolist()
# this_year_judgment_list = [list(x) for x in zip(a, b)]
# print(this_year_judgment_list[:][0])

# #計算
# c = df['AttnSN'].tolist()
# d = df['judgment'].tolist()
# previous_year_judgment_list = [list(x) for x in zip(c, d)]
# for ele in previous_year_judgment_list:

#         match_list = sorted(b, key=lambda x: difflib.SequenceMatcher(None, x, ele[1]).ratio(), reverse=True)
#         most_match = match_list[0]
#         print(ele)
#         print(most_match)


# df= df.assign(most_match=lambda x: sorted(this_year_judgment_list, key=lambda y: difflib.SequenceMatcher(None, y, x.judgment).ratio(), reverse=True))# 判斷IMO_Number 7碼欄位

# print(df)


# this_year_judgment_list = df1[['id','judgment']].tolist()



# df = pd.merge(df[['CASNo','AttnSN']],df1[['CASNo','id']], left_on=['CASNo'],right_on=['CASNo'],how='outer')
# print(df)



# print(df)
# print(df1)

# test = 'GHS危害物質名單_336-08-3_-_octafluoroadipic acid'
# weighted_results =[]

# ratio = difflib.SequenceMatcher(None, 'GHS危害物質名單_336-08-3_-_octafluoroadipic acid', test).ratio()
# weighted_results.append(('GHS危害物質名單_336-08-3_-_octafluoroadipic acid', ratio))
# print(weighted_results)