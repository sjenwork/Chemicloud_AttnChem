import pandas as pd

# df = pd.read_excel("../explain_and_structure/清單資料來源_20220707_PC.xlsx", header = 0)
df1 = pd.read_excel("../output/Merged.xlsx", header = 0)
df2 = pd.read_excel("../explain_and_structure/AttnList.xlsx", header = 0)

#df1
df1['CASNo'] = df1['CASNo'].str.strip().replace(r'[ |\t|\r|\n]', '',regex =True).replace(r'(\xa0)','').replace(r'(\&)','',regex =True).replace('等','',regex =True)
df1['CASNo'] = df1['CASNo'].astype(str)
df1 =df1[['this_year_id',"CASNo",'Source',"Unit"]]
df1 = df1[(df1.CASNo !="-")]
df1 = df1[(df1.CASNo !="nan")]

#df2
Columns = {'Source': 'Source_sys',
            'Unit': 'Unit_sys',
            'AttnSN': 'AttnSN',
            'CASNo': 'CASNo',
            'ChemicalChnName':'ChemicalChnName',
            'ChemicalEngName':'ChemicalEngName'         
            }
df2 = df2.rename(columns = Columns)
df2 = df2[Columns.values()]
df2['CASNo'] = df2['CASNo'].str.strip().replace(r'[ |\t|\r|\n]', '',regex =True).replace(r'(\xa0)','').replace(r'(\&)','',regex =True).replace('等','',regex =True)
df2['CASNo'] = df2['CASNo'].astype(str)
df2=df2[['AttnSN','CASNo','Source_sys','Unit_sys']]

df = pd.merge(df2,df1, left_on=['CASNo','Source_sys','Unit_sys'],right_on=['CASNo','Source','Unit'],how='outer')
df = df.sort_values(by=["AttnSN"], ascending=True)#

df.to_excel("yearly_outer_merge.xlsx", index = False)
