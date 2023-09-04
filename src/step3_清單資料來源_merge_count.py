import pandas as pd

df = pd.read_excel("../explain_and_structure/清單資料來源_20220707_PC.xlsx", header = 0)
df1 = pd.read_excel("../output/Merged.xlsx", header = 0)
df2 = pd.read_excel("../explain_and_structure/AttnList.xlsx", header = 0)


df.loc[(df["單位/國家"].isnull())|(df["單位/國家"]==''),"單位/國家"]='-'

df1.loc[(df1["Unit"].isnull())|(df1["Unit"]==''),"Unit"]='-'
df1 =df1[["Source_ID",'Source',"Unit"]]
df1["Source_count"] = df1["Source"] 
df1= df1.groupby(by=["Source_ID","Source","Unit"])[["Source_count"]].count().reset_index()

Columns = {'Source': 'Source_sys',
            'Unit': 'Unit_sys'}
df2 = df2.rename(columns = Columns)
df2 = df2[Columns.values()]
df2.loc[(df2["Unit_sys"].isnull())|(df2["Unit_sys"]==''),"Unit_sys"]='-'
df2=df2[['Source_sys','Unit_sys']]
df2["Source_sys_count"] = df2["Source_sys"] 
df2= df2.groupby(by=["Source_sys",'Unit_sys'])[["Source_sys_count"]].count().reset_index()

df = pd.merge(df,df1, left_on=['編號'],right_on=['Source_ID'],how='left')
df = pd.merge(df,df2, left_on=['清單','單位/國家'],right_on=['Source_sys','Unit_sys'],how='outer')
df = df.fillna('-')
df = df.drop(columns=["Source_ID","Source","Unit"],axis=1)
df =df[["編號","國內/外","單位/國家","法令","清單","資料來源","資料來源(PC改)","檔案類型改","Source_count","Source_sys_count","Source_sys",'Unit_sys']]
df.to_excel("list_data_source_with_source_count.xlsx", index = False)

# #get source_col_compare.xlsx
# df = pd.read_excel("../output/Merged.xlsx", header = 0)
# df =df[['Source_ID','Source',"Unit"]]
# df.drop_duplicates(subset=["Source_ID"],keep="first",inplace = True)
# print(df)

# df1 = pd.read_excel("../explain_and_structure/AttnList.xlsx", header = 0)
# df1.drop_duplicates(subset=["Source"],keep="first",inplace = True)
# Columns = {'Source': 'Source_sys',
#             'Unit': 'Unit_sys'}
# df1 = df1.rename(columns = Columns)
# df1 = df1[Columns.values()]
# df1=df1[['Source_sys','Unit_sys']]

# df1 = pd.merge(df1,df, left_on=['Source_sys'],right_on=['Source'],how='outer')
# df1 = df1.sort_values(by=['Source_ID'])
# print(df1)
# df1.to_excel("source_col_compare.xlsx", index = False)

# #get source_count_compare.xlsx
# df = pd.read_excel("../output/Merged.xlsx", header = 0)
# df =df[['Source_ID','Source',"Unit"]]
# df["Source_count"] = df["Source"] 
# df= df.groupby(by=['Source_ID',"Source","Unit"])[["Source_count"]].count().reset_index()

# df1 = pd.read_excel("../explain_and_structure/AttnList.xlsx", header = 0)
# Columns = {'Source': 'Source_sys',
#             'Unit': 'Unit_sys'}
# df1 = df1.rename(columns = Columns)
# df1 = df1[Columns.values()]
# df1=df1[['Source_sys','Unit_sys']]

# df1["Source_sys_count"] = df1["Source_sys"] 
# df1= df1.groupby(by=["Source_sys",'Unit_sys'])[["Source_sys_count"]].count().reset_index()
# # print(df1)
# df1 = pd.merge(df1,df, left_on=['Source_sys'],right_on=['Source'],how='outer')
# df1 = df1.sort_values(by=['Source_ID'])

# df1=df1[["Source_ID","Source_sys","Unit_sys","Source_sys_count","Source","Unit","Source_count"]]

# df1.to_excel("source_count_compare.xlsx", index = False)

# #  sort
# df = pd.read_excel("../output/Merged.xlsx", header = 0)
# df = df.sort_values(by=['Source_ID','ChemicalEngName'])
# df.to_excel("0.xlsx", index = False)

# df_ = pd.read_excel("source_count_compare.xlsx", header = 0)

# df1 = pd.read_excel("../explain_and_structure/AttnList.xlsx", header = 0)
# df1 = pd.merge(df1,df_[["Source_sys","Source_ID"]], left_on=['Source'],right_on=['Source_sys'],how='left')
# df1 = df1.sort_values(by=['Source_ID','ChemicalEngName'])
# df1.to_excel("1.xlsx", index = False)



# df_list_data_source = pd.read_excel("../'explain_and_structure/清單資料來源_20220707_PC.xlsx'", header = 0)

# df = pd.read_excel("source_count_compare.xlsx", header = 0)


# df1 = pd.merge(df_list_data_source,df, left_on=['編號'],right_on=['Source_sys'],how='left')

# df1 = df1.sort_values(by=['Source_ID','ChemicalEngName'])
# df1.to_excel("list_data_source_with_source_count.xlsx", index = False)