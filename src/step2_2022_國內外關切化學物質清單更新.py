import os
import pandas as pd
import pdfplumber
import tabula
import pandas as pd
import numpy as np
from docx.api import Document
from win32com import client as wc
import re
from io import StringIO
from opencc import OpenCC
import aspose.words as aw
import time
def curves_to_edges(cs):
    edges = []
    for c in cs:
        edges += pdfplumber.utils.rect_to_edges(c)
    return edges
def n02():
    file = os.path.join('../2022_input', '02.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    

    Columns = {s0.columns[2]: 'ChemicalChnName',      
            s0.columns[3]: 'ChemicalEngName',
            s0.columns[5]: 'CASNo',
            }

    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('等', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\r', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\r', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

    s0.loc[s0.CASNo == '', ['CASNo']] = '-'
    s0.loc[s0['CASNo'].isnull()==True,'CASNo'] = "-"

    s0 = s0[Columns.values()]
    print(s0)


    file = os.path.join('../2022_processed', '02.xlsx')
    s0.to_excel(file, index = False)
# n02()
def n03():
    file = os.path.join('../2022_input', '03.pdf')
    pdf = pdfplumber.open(file)
    table = []
    table_settings = {"join_tolerance": 100}
    for i in range(1, 6):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' ', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('；', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('）', '', regex = False)
    s0.loc[17, 'ChemicalChnName'] =  s0.loc[17, 'ChemicalChnName'].replace('註', '')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    a = s0[s0.ChemicalChnName == ''].index
    for i in a:
        s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + s0.loc[i, 'ChemicalEngName']
        s0 = s0.drop(index = i)
    s0.CASNo = s0.CASNo.str.replace(' ', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '; ', regex = False)
    s0.loc[s0.CASNo == '', ['CASNo']] = '-'
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '03.xlsx')
    s0.to_excel(file, index = False)
    
def n04():
    file = os.path.join('../2022_input', '04.pdf')
    pdf = pdfplumber.open(file)
    table = []
    table_settings = {"join_tolerance": 100}
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[2]: 'ChemicalChnName',
               s0.columns[3]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName != '']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    for i in s0.index[::-1]:
        if pd.isna(s0.loc[i, 'ChemicalChnName']):
            s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + ' ' + s0.loc[i, 'ChemicalEngName']
            s0.drop(index = i, inplace = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '04.xlsx')
    s0.to_excel(file, index = False)

def n05():
    file = os.path.join('../2022_input', '05.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[2:], columns = table[1])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName.notna()]
    s0 = s0[s0.ChemicalEngName != '英文名稱']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' or', ';', regex = False)
    s0.loc[8, 'ChemicalEngName'] = s0.loc[8, 'ChemicalEngName'].replace('(', '; ')
    s0.loc[8, 'ChemicalEngName'] = s0.loc[8, 'ChemicalEngName'].replace(')', '')
    s0.loc[68, 'ChemicalEngName'] = s0.loc[68, 'ChemicalEngName'].replace(' (', '; ')
    s0.loc[68, 'ChemicalEngName'] = s0.loc[68, 'ChemicalEngName'].replace(')', '')
    s0.loc[70, 'ChemicalEngName'] = s0.loc[70, 'ChemicalEngName'].replace(', (', '; ')
    s0.loc[70, 'ChemicalEngName'] = s0.loc[70, 'ChemicalEngName'].replace(')', '')
    s0.loc[71, 'ChemicalEngName'] = s0.loc[71, 'ChemicalEngName'].replace('( ', '(')
    s0.loc[71, 'ChemicalEngName'] = s0.loc[71, 'ChemicalEngName'].replace(', [', '; ')
    s0.loc[71, 'ChemicalEngName'] = s0.loc[71, 'ChemicalEngName'].replace(']', '')
    s0.loc[72, 'ChemicalEngName'] = s0.loc[72, 'ChemicalEngName'].replace(' (', '; ')
    s0.loc[72, 'ChemicalEngName'] = s0.loc[72, 'ChemicalEngName'].replace(')', '')
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('(俗稱：', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(')', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('(', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.strip()
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '05.xlsx')
    s0.to_excel(file, index = False)
# def n06():
#     file = os.path.join('../2022_input', '06.pdf')
#     s0 = tabula.read_pdf(file, pages='46-49', lattice = True)
#     s0 = pd.concat(s0)
#     print(s0)

#     Columns = {s0.columns[2]: 'ChemicalChnName',      
#             s0.columns[3]: 'ChemicalEngName',
#             s0.columns[4]: 'CASNo',
#             }
        
#     s0 = s0.rename(columns = Columns)
#     s0 = s0[Columns.values()]

#     s0.CASNo = s0.CASNo.str.replace('\r', '; ', regex = False)
#     s0.CASNo = s0.CASNo.str.replace('\n', '; ', regex = False)
#     s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\r', '', regex = False)
#     s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
#     s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
#     s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

#     file = os.path.join('../2022_processed', '06.xlsx')
#     s0.to_excel(file, index = False)

def n06():
    file = os.path.join('../2022_input', '06.pdf')
    s0 = tabula.read_pdf(file, pages='34-38', lattice = True)
    s0 = pd.concat(s0)
    # print(s0)

    Columns = {s0.columns[2]: 'ChemicalChnName',      
            s0.columns[3]: 'ChemicalEngName',
            s0.columns[4]: 'CASNo',
            }
        
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0.CASNo = s0.CASNo.str.replace(' ', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('/', '; ', regex = False).replace('等','',regex =True)
    s0.CASNo = s0.CASNo.str.replace('\r', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\r', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

    #若CASNo不唯一，分成多行
    s0.CASNo = s0.CASNo.str.split(';') 
    # Merged.CASNo = Merged.CASNo.str.split(r"[;| ]")                                                              
    s0 = s0.explode('CASNo')




    print(s0)

    file = os.path.join('../2022_processed', '06.xlsx')
    s0.to_excel(file, index = False)

# n06()
def n08():
    file = os.path.join('../2022_input', '08.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[3]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]




    s0['series'] = s0['CASNo']
    for shift in range(1,5):
        s0["sf0"] =s0["ChemicalChnName"]
        s0["sf1"] =s0["ChemicalChnName"].shift(-1)
        s0["sf2"] =s0["ChemicalChnName"].shift(-2)
        s0["sf3"] =s0["ChemicalChnName"].shift(-3)
        s0["sf4"] =s0["ChemicalChnName"].shift(-4)

    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1)=="")&(s0["sf1"]!=""),'ChemicalChnName'] = s0["ChemicalChnName"]+  s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf2"]!=""),'ChemicalChnName'] = s0["ChemicalChnName"]+  s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf3"]!=""),'ChemicalChnName'] = s0["ChemicalChnName"]+  s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4)=="")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf4"]!=""),'ChemicalChnName'] = s0["ChemicalChnName"]+  s0["sf4"]

    for shift in range(1,5):
        s0["sf0"] =s0["ChemicalEngName"]
        s0["sf1"] =s0["ChemicalEngName"].shift(-1)
        s0["sf2"] =s0["ChemicalEngName"].shift(-2)
        s0["sf3"] =s0["ChemicalEngName"].shift(-3)
        s0["sf4"] =s0["ChemicalEngName"].shift(-4)
     
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1)=="")&(s0["sf1"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf2"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf3"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4)=="")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf4"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf4"]

    for shift in range(1,5):
        s0["sf0"] =s0["CASNo"]
        s0["sf1"] =s0["CASNo"].shift(-1)
        s0["sf2"] =s0["CASNo"].shift(-2)
        s0["sf3"] =s0["CASNo"].shift(-3)
        s0["sf4"] =s0["CASNo"].shift(-4)
     
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1)=="")&(s0["sf1"]!=""),'CASNo'] = s0["CASNo"]+ " " + s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf2"]!=""),'CASNo'] = s0["CASNo"]+ " " + s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf3"]!=""),'CASNo'] = s0["CASNo"]+ " " + s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4)=="")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf4"]!=""),'CASNo'] = s0["CASNo"]+ " " + s0["sf4"]

    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\r', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

    s0 = s0[(s0.CASNo != "")]
    s0 = s0[Columns.values()]
    print(s0)
    file = os.path.join('../2022_processed', '08.xlsx')
    s0.to_excel(file, index = False)

def n09():
    file = os.path.join('../2022_input', '09.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[3]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\r', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

    print(s0)
    file = os.path.join('../2022_processed', '09.xlsx')
    s0.to_excel(file, index = False)

def n10():
    file = os.path.join('../2022_input', '10.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(0,232):#232
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])

    s = s0.columns.to_series().groupby(s0.columns)
    s0.columns = np.where(s.transform('size')>1, 
                        s0.columns + s.cumcount().add(1).astype(str), 
                        s0.columns)
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {
            # s0.columns[2]: 'ChemicalChnName',      
            s0.columns[3]: 'ChemicalEngName',
            s0.columns[0]: 'CASNo',
            }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 = s0.reset_index(drop = True)
    
    s0.CASNo = s0.CASNo.str.replace('\r', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\r', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)

    print(s0)

    s0['series'] = s0['CASNo']
    for shift in range(1,5):
        s0["sf0"] =s0["ChemicalEngName"]
        s0["sf1"] =s0["ChemicalEngName"].shift(-1)
        s0["sf2"] =s0["ChemicalEngName"].shift(-2)
        s0["sf3"] =s0["ChemicalEngName"].shift(-3)
        s0["sf4"] =s0["ChemicalEngName"].shift(-4)

    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1)=="")&(s0["sf1"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf2"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf3"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4)=="")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf4"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf4"]

    s0 = s0[(s0.CASNo != "")]#
    s0 = s0[Columns.values()]
    print(s0)

    file = os.path.join('../2022_processed', '10.xlsx')
    s0.to_excel(file, index = False)

# def n10():
#     file = os.path.join('../2022_input', '10.pdf')
#     pdf = pdfplumber.open(file)
#     table = []
#     for i in range(len(pdf.pages)):
#         page = pdf.pages[i]
#         table += page.extract_table()
#     s0 = pd.DataFrame(table[1:], columns = table[0])
#     print(*enumerate(s0.columns), '\n', sep = '\n')

#     Columns = {s0.columns[0]: 'ChemicalChnName',
#                s0.columns[1]: 'ChemicalEngName',
#                s0.columns[2]: 'CASNo'
#               }
#     s0 = s0.rename(columns = Columns)
#     s0 = s0[s0.CASNo != 'CAS.NO.']
#     s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
#     s0.loc[s0.ChemicalEngName.isna(), 'ChemicalEngName'] = '-'
#     s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
#     s0 = s0[Columns.values()]

#     file = os.path.join('../2022_input', '10_2.pdf')
#     pdf = pdfplumber.open(file)
#     table = []
#     for i in range(len(pdf.pages)):
#         page = pdf.pages[i]
#         table += page.extract_table()
#     s1 = pd.DataFrame(table[1:], columns = table[0])
#     print(*enumerate(s1.columns), '\n', sep = '\n')

#     Columns = {s1.columns[1]: 'ChemicalChnName',
#                s1.columns[0]: 'ChemicalEngName',
#                s1.columns[2]: 'CASNo'
#               }
#     s1 = s1.rename(columns = Columns)
#     s1.ChemicalChnName = s1.ChemicalChnName.str.replace('\n', '', regex = False)
#     s1.ChemicalEngName = s1.ChemicalEngName.str.replace('\n', '', regex = False)
#     s1 = s1[Columns.values()]

#     file = os.path.join('../2022_input', '10_3.pdf')
#     pdf = pdfplumber.open(file)
#     table = []
#     for i in range(len(pdf.pages)):
#         page = pdf.pages[i]
#         table += page.extract_table()
#     s2 = pd.DataFrame(table[1:], columns = table[0])
#     print(*enumerate(s2.columns), '\n', sep = '\n')

#     Columns = {s2.columns[1]: 'ChemicalChnName',
#                s2.columns[0]: 'ChemicalEngName',
#                s2.columns[2]: 'CASNo'
#               }
#     s2 = s2.rename(columns = Columns)
#     a = s2[s2.CASNo == ''].index
#     for i in a:
#         s2.loc[i-1, 'ChemicalChnName'] = s2.loc[i-1, 'ChemicalChnName'] + s2.loc[i, 'ChemicalChnName']
#         s2.loc[i-1, 'ChemicalEngName'] = s2.loc[i-1, 'ChemicalEngName'] + s2.loc[i, 'ChemicalEngName']
#         s2 = s2.drop(index = i)
#     s2 = s2[s2.CASNo != 'CAS No.']
#     s2.ChemicalChnName = s2.ChemicalChnName.str.replace('\n', '', regex = False)
#     s2.ChemicalEngName = s2.ChemicalEngName.str.replace('\n', '', regex = False)
#     s2 = s2[Columns.values()]

#     df = pd.concat([s0, s1, s2], ignore_index = True)
#     file = os.path.join('../2022_processed', '10.xlsx')
#     df.to_excel(file, index = False)

def n11():
    file = os.path.join('../2022_input', '11.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[0]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName',
               s0.columns[6]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalChnName != '中文名稱 \n(學名)']
    s0 = s0[s0.ChemicalChnName.notna()]
    s0 = s0.reset_index(drop = True)
    a = s0[s0.ChemicalChnName == ''].index
    for i in a:
        s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + s0.loc[i, 'ChemicalEngName']
        s0 = s0.drop(index = i)
    s0.ChemicalChnName = s0.ChemicalChnName + '; ' + s0[s0.columns[1]]
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0[s0.columns[3]]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('；', ';', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' (', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(');', ';', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('; -', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('; -', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(':', ';', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_input', '11_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s1.columns), '\n', sep = '\n')

    Columns = {s1.columns[0]: 'ChemicalChnName',
               s1.columns[2]: 'ChemicalEngName',
               s1.columns[6]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1.ChemicalChnName != '中文名稱 \n(學名)']
    s1.ChemicalChnName = s1.ChemicalChnName + '; ' + s1[s1.columns[1]]
    s1.ChemicalEngName = s1.ChemicalEngName + '; ' + s1[s1.columns[3]]
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace('\n', '', regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('\n', '', regex = False)
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace('；', ';', regex = False)
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace('比重.+之', '', regex = True)
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace(' \(註.+\)', '', regex = True)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('; -', '', regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace(' \(Note.+\)', '', regex = True)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace(' having.+%', '', regex = True)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace(' (', '; ', n = 1, regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace(')', '', n = 1, regex = False)
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    df.drop_duplicates(subset=["ChemicalChnName"],keep="first",inplace = True)

    file = os.path.join('../2022_processed', '11.xlsx')
    df.to_excel(file, index = False)

def n12():
    file = os.path.join('../2022_processed', '12_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName',
               s0.columns[4]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    k = s0.CASNo.str.extractall('(\d+\-\d+\-\d)')
    for i in k.index.get_level_values(0).unique():
        s0.loc[i, 'CASNo'] = '; '.join(k[0][i])
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('/', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('(', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(')', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '12.xlsx')
    s0.to_excel(file, index = False)

def n13():
    file = os.path.join('../2022_input', '13.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(50, 52):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')


    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    # a = s0[s0.ChemicalChnName == ''].index
    # for i in a:
    #     s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + s0.loc[i, 'ChemicalEngName']
    #     s0 = s0.drop(index = i)

    a = [0,1,2,4,6,11,12]
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'] + "鹼"
    s0.loc[5, 'ChemicalChnName'] = "去甲麻黃鹼; 新麻黃鹼"


    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' （', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' ）', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('、', '; ', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' ;', ';', regex = False)
    s0 = s0[Columns.values()]
    print(s0)

    file = os.path.join('../2022_processed', '13.xlsx')
    s0.to_excel(file, index = False)

def n14():
    file = os.path.join('../2022_input', '14.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 =s0.reset_index(drop=True)
    
    s0.ChemicalEngName = s0.ChemicalEngName.str.lower()
    s0 = s0[~s0.duplicated()]
    
    s0.loc[(s0.ChemicalEngName.str[0:1] =="（")&(s0.ChemicalEngName.str[-1:] =="）"),'ChemicalEngName'] = s0['ChemicalEngName'].str[1:-1]
    s0.loc[(s0.ChemicalEngName.str[0:1] =="(")&(s0.ChemicalEngName.str[-1:] ==")"),'ChemicalEngName'] = s0['ChemicalEngName'].str[1:-1]
    print(s0)
    file = os.path.join('../2022_processed', '14.xlsx')
    s0.to_excel(file, index = False)
# n14()
def n15():
    file = os.path.join('../2022_input', '15.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('/', '; ', regex = False)
    s0.CASNo.fillna('-', inplace = True)
    s0.CASNo = s0.CASNo.str.replace('\n', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('/', '; ', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '15.xlsx')
    s0.to_excel(file, index = False)

def n16():
    file = os.path.join('../2022_input', '16.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
               s0.columns[4]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0['INCI名']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('/', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('/', '; ', regex = False)
    
    s0.CASNo = s0.CASNo.str.replace(';                    ', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace(';                  ', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace(',        ', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace(',       ', '; ', regex = False)

    s0.loc[(s0.CASNo.str[-2:] =="; "),'CASNo'] = s0['CASNo'].str[:-2]
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '16.xlsx')
    s0.to_excel(file, index = False)

def n17():
    file = os.path.join('../2022_input', '17.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName'}
    s0 = s0.rename(columns = Columns)
    s0['INCI名'].fillna('-', inplace = True)
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0['INCI名']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('; -', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('/', '; ', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '17.xlsx')
    s0.to_excel(file, index = False)

def n19():
    file = os.path.join('../2022_input', '19.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[2]: 'CASNo'}
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0['序號'] != '序號']
    s0 = s0[s0['序號'] != '']
    s0['其他成分名稱'] = s0['其他成分名稱'].str.replace('\n', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace(' *\n', '; ', regex = True)
    s0[['ChemicalChnName', 'ChemicalEngName']] = s0['其他成分名稱'].str.split('（', n = 1, expand = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('）', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' *\(', '; ', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(')', '', regex = False)
    s0.loc[39, 'ChemicalEngName'] = s0.loc[39, 'ChemicalChnName']
    s0.loc[40, 'ChemicalEngName'] = s0.loc[40, 'ChemicalEngName'].replace('或食用藍色 2號（', '; ')
    s0.loc[40, 'ChemicalChnName'] = s0.loc[40, 'ChemicalChnName'] + '; ' + '食用藍色 2號'
    s0.ChemicalEngName.fillna('-', inplace = True)
    s0 = s0[list(Columns.values()) + ['ChemicalChnName', 'ChemicalEngName']]

    file = os.path.join('../2022_processed', '19.xlsx')
    s0.to_excel(file, index = False)

def n20():
    file = os.path.join('../2022_input', '20.pdf')
    pdf = pdfplumber.open(file)
    page = pdf.pages[0]
    table = page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    s0['不純物'] = s0['不純物'].str.replace('\n', '', regex = False)
    s0['不純物'] = s0['不純物'].str.replace(', *', '; ', regex = True)
    s0[['ChemicalChnName', 'ChemicalEngName']] = s0['不純物'].str.split('（', n = 1, expand = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('）', '', regex = False)
    s0 = s0[['ChemicalChnName', 'ChemicalEngName']]

    file = os.path.join('../2022_processed', '20.xlsx')
    s0.to_excel(file, index = False)


def n21():
    file = os.path.join('../2022_input', '21.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(0, 10):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName'}
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalChnName.notna()]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（.+', '', regex = True)
    s0 = s0[Columns.values()]
    
    s0 = s0[s0.ChemicalChnName !="摻雜之農藥有效成分"]
    print(s0)
    file = os.path.join('../2022_processed', '21.xlsx')
    s0.to_excel(file, index = False)

def n22():
    # doc to docx 
    try:
        word = wc.Dispatch("Word.Application") # 打开word应用程序
        doc = word.Documents.Open(os.path.dirname(os.path.abspath(__file__))+"/../2022_input/22.doc") #打开word文件
        doc.SaveAs("{}x".format(os.path.dirname(os.path.abspath(__file__))+"/../2022_input/22.doc"), 12)#另存为后缀为".docx"的文件，其中参数12指docx文件
        doc.Close() #关闭原来word文件
        word.Quit()
    except:
        pass
    #docx to df
    document = Document(os.path.dirname(os.path.abspath(__file__))+"/../2022_input/22.docx")
    tables = document.tables
    s0 = pd.DataFrame()
    for table in tables:     
        for row in table.rows:
            text = [cell.text for cell in row.cells]
            s0 = s0.append([text], ignore_index=True)
    header_row = 0
    s0.columns = s0.iloc[header_row]
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0 = s0[["飼料添加物名稱","使用對象"]]
    s0["飼料添加物名稱"] = s0["飼料添加物名稱"].str.split('\n')
    s0 = s0[(s0.使用對象 !="")]
    s0 = s0.drop(columns=["使用對象"],axis=1)# 刪除判斷欄
    s0 = s0.reset_index(drop = True)

    s0 = s0.iloc[4:] 
    s0["ChemicalChnName"] = s0["飼料添加物名稱"].str[0]
    s0["ChemicalChnName"] = s0["飼料添加物名稱"].str[0].str.replace("[A-Za-z0-9\,\。\.\-]","")#re.sub("[A-Za-z0-9\,\。]", "", s0["ChemicalChnName"].apply(str))
    s0["ChemicalEngName"] = s0["飼料添加物名稱"].str[1].str.replace("[\u4e00-\u9fa5\0-9\,\。\：]","")
    s0.loc[s0["飼料添加物名稱"].str[0].str.replace("[\u4e00-\u9fa5\0-9\,\。\：]","") != s0["ChemicalEngName"],'ChemicalEngName'] = s0["飼料添加物名稱"].str[1].str.replace("[\u4e00-\u9fa5\0-9\,\。\：]","") +"; "+s0["飼料添加物名稱"].str[0].str.replace("[\u4e00-\u9fa5\0-9\,\。\：]","") 
    s0 = s0.drop(columns=["飼料添加物名稱"],axis=1)
    
    
    s0 = s0.reset_index(drop = True)
    s0.drop_duplicates(subset=["ChemicalChnName"],keep="first",inplace = True)
    s0 = s0[(s0.ChemicalEngName != "")]#

    file = os.path.join('../2022_processed', '22.xlsx')
    s0.to_excel(file, index = False)    
    print(s0)
    os.remove(os.path.dirname(os.path.abspath(__file__))+"/../2022_input/22.docx")

# def n23():#半自動
#     file = os.path.join('../2022_input', '23.pdf')
#     pdf = pdfplumber.open(file)
#     str_ = []
#     for i in range(len(pdf.pages)):
#         page = pdf.pages[i]
#         str_ += '\n' + page.extract_text()

#     str_= StringIO("".join(str_))
#     s0 = pd.read_table(str_, sep='\n')
#     s0['check'] = s0["附表四 "].str[0].str.isdigit()
#     s0 = s0[(s0.check == True)]#
    
#     s0 =s0.reset_index(drop=True)
#     s0 = s0.iloc[77: , :]
#     print(s0)

#     file = os.path.join('../2022_processed', '23.xlsx')
#     s0.to_excel(file, index = False)
# n23()



def n27():
    file = os.path.join('../2022_input', '27.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName.str.contains('Entry') != True]
    s0 = s0[Columns.values()]

    s0.ChemicalEngName = s0.ChemicalEngName.str.lower()
    s0 = s0[~s0.duplicated()]
    print(s0)
    file = os.path.join('../2022_processed', '27.xlsx')
    s0.to_excel(file, index = False)

def n28():#刪除重複的動作 可以到後面merged 再做
    file = os.path.join('../2022_input', '28.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('- *,', '', regex = True)
    s0.CASNo = s0.CASNo.str.replace(', *', '; ', regex = True)
    s0 = s0[~s0.ChemicalEngName.str.contains('available')]
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '28.xlsx')
    s0.to_excel(file, index = False)

def n29():
    file = os.path.join('../2022_input', '29.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '29.xlsx')
    s0.to_excel(file, index = False)

def n30():
    file = os.path.join('../2022_input', '30.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '30.xlsx')
    s0.to_excel(file, index = False)

def n31():
    try:
        file = os.path.join('../2022_input', '31.pdf')
        pdf = pdfplumber.open(file)
        table = []
        for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            table += page.extract_table()
        s0 = pd.DataFrame(table[2:], columns = table[1])
        print(*enumerate(s0.columns), '\n', sep = '\n')

        Columns = {s0.columns[0]: 'ChemicalEngName',
                s0.columns[1]: 'CASNo'
                }
        s0 = s0.rename(columns = Columns)
        s0 = s0[s0.CASNo != 'CAS Number']
        s0 = s0[s0.CASNo.notna()]
        s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
        s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' / ', '; ', regex = False)
        s0 = s0[Columns.values()]

        file = os.path.join('../2022_processed', '31.xlsx')
        s0.to_excel(file, index = False)
    except:
        file = os.path.join('../2022_input', '31.xlsx')
        x = pd.ExcelFile(file)
        print(*enumerate(x.sheet_names), '\n', sep = '\n')

        s0 = x.parse(sheet_name = 0, skiprows = 0)
        print(*enumerate(s0.columns), '\n', sep = '\n')
        Columns = {s0.columns[0]: 'ChemicalEngName',
                s0.columns[1]: 'CASNo'
                }
        s0 = s0.rename(columns = Columns)
        s0 = s0[Columns.values()]

        file = os.path.join('../2022_processed', '31.xlsx')
        s0.to_excel(file, index = False)

def n32():
    file = os.path.join('../2022_input', '32.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[2:], columns = table[1])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[4]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName.notna()]
    s0 = s0[s0.ChemicalEngName != 'Chemical Name']
    a = s0[s0.ChemicalEngName.str.contains(' \([A-Z]+', regex = True)].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'].replace(' (', '; ')
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-1]
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '32.xlsx')
    s0.to_excel(file, index = False)

def n33():
    file = os.path.join('../2022_processed', '33_0.xlsx')
    s0 = pd.read_excel(file, header = None)

    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' +', ' ', regex = True)
    s0.loc[35, 'ChemicalEngName'] = 'Alkyl phenol (from C5 to C9); Nonyl phenol; Octyl phenol'
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(', '; ', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(', ', ' and ', regex = False)
    a = s0[s0.ChemicalEngName.str.contains(')', regex = False)].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-1]
    file = os.path.join('../2022_processed', '33.xlsx')
    s0.to_excel(file, index = False)

# def n34(): 半自動
#     try:
#         file = os.path.join('../2022_input', '34.xlsx')
#         x = pd.ExcelFile(file)
#         print(*enumerate(x.sheet_names), '\n', sep = '\n')

#         s0 = x.parse(sheet_name = 0)
#         print(*enumerate(s0.columns), '\n', sep = '\n')
#         Columns = {s0.columns[1]: 'ChemicalEngName',
#                 s0.columns[0]: 'CASNo'
#                 }
#         s0 = s0.rename(columns = Columns)
#         s0 = s0[Columns.values()]

#         s1 = x.parse(sheet_name = 1)
#         print(*enumerate(s1.columns), '\n', sep = '\n')
#         Columns = {s1.columns[1]: 'ChemicalEngName',
#                 s1.columns[0]: 'CASNo'
#                 }
#         s1 = s1.rename(columns = Columns)
#         s1 = s1[Columns.values()]

#         df = pd.concat([s0, s1], ignore_index = True)
#         df.CASNo = df.CASNo.str.replace('*', '', regex = False)
#         df.CASNo = df.CASNo.str.replace('N.+', '-', regex = True)
#         file = os.path.join('../2022_processed', '34.xlsx')
#         df.to_excel(file, index = False)
#     except:
#         file = os.path.join('../2022_input', '34.pdf')
#         pdf = pdfplumber.open(file)
#         str_ = []
#         for i in range(len(pdf.pages)):
#             page = pdf.pages[i]
            
#             page_ex_rext = page.extract_text()
#             str_ += '\n' + page_ex_rext
            


#         str_= StringIO("".join(str_))
#         s0 = pd.read_table(str_, sep='\n',skiprows =6)

#         Columns = {s0.columns[0]: 'ChemicalEngName_CASNo',
#                 }
#         s0 = s0.rename(columns = Columns)
#         s0["ChemicalEngName_CASNo"]= s0["ChemicalEngName_CASNo"]
  
        
#         s0 = s0[s0.ChemicalEngName_CASNo !="    03/07/22 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="1 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="2 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="3 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="4 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="5 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="6 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="7 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="8 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="9 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="10 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="11 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="12 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="13 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="14 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="15 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="16 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="17 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="18 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="19 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="20 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="CAS Number   Chemical Name  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="CAS Number  Chemical Name "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added for Reporting Year 2022 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added for Reporting Year 2021 "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1990  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1991  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1994  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1995  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 2000  "]        
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1995  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Chemicals Added For Reporting Year 1995  "]
#         s0 = s0[s0.ChemicalEngName_CASNo !="Vanadium Compounds  "]
#         s0["CASNo"]=s0["ChemicalEngName_CASNo"].str.extract(r'(\d+-\d+-\d+)')


#         s0["ChemicalEngName"]= s0.apply(lambda x: x['ChemicalEngName_CASNo'].replace(str(x['CASNo']),""), axis=1)
#         s0["ChemicalEngName"] =s0["ChemicalEngName"].str.lstrip()
#         s0["ChemicalEngName"] =s0["ChemicalEngName"].str.replace(r"; \(.*\)","")
        

#         s0.loc[s0['CASNo'].isnull() == True,'CASNo'] = "-"
#         s0= s0[["CASNo","ChemicalEngName"]]
#         print(s0)
#         file = os.path.join('../2022_processed', '34.xlsx')
#         s0.to_excel(file, index = False)

# def n35():半自動
#     file = os.path.join('../2022_input', '35.pdf')
#     pdf = pdfplumber.open(file)
#     str_ = []
#     for i in range(len(pdf.pages)):
#         page = pdf.pages[i]
        
#         # print(page)
#         page_ex_rext = page.extract_text()
#         lin_list = page_ex_rext.split("\n")
        
#         for line in lin_list:
            
#             found = re.search('\d+. ',line)
#             # print(found)
            
#             if found == None:
#                 line =re.sub(r"\([a-z]\)","\n", line)
#                 print(line)
#                 str_ += line

#             else:
#                 line =re.sub(r"\([a-z]\)","\n", line)
#                 print(line)
#                 # line = re.sub(r'[(](^\w$)[)]','\n',line)
#                 str_ += '\n' + line

#     str_= StringIO("".join(str_))
#     s0 = pd.read_table(str_, sep='\n',skiprows =1)
    
#     print(*enumerate(s0.columns), '\n', sep = '\n')
#     Columns = {'Updated Schedule 1 as of May 12, 2021': "ChemicalEngName_CASNo",
#               }
#     s0 = s0.rename(columns = Columns)
#     s0["series"]=s0["ChemicalEngName_CASNo"].str.extract(r'(\d+. )')
#     s0["ChemicalEngName"]= s0.apply(lambda x: x['ChemicalEngName_CASNo'].replace(str(x['series']),""), axis=1)
#     s0= s0[["series","ChemicalEngName"]]

#     for n  in range (236,283):
#         s0.loc[n, 'ChemicalEngName'] = re.sub(r"\(.*?\)","",s0.loc[n, 'ChemicalEngName'])
#     for n  in range (106,140):
#         s0.loc[n, 'ChemicalEngName'] = s0.loc[n, 'ChemicalEngName'][4:]

#     print(s0)
#     file = os.path.join('../2022_processed', '35.xlsx')
#     s0.to_excel(file, index = False)
def n35():
    # step1,pdf to docx: https://products.aspose.com/words/python-net/conversion/pdf-to-word/ 
    # step2,copy text from docx to xlsx
    # step3,edit xlsx by the following code
    #docx to df
    file = os.path.join('../2022_input', '35.xlsx')
    x = pd.ExcelFile(file)
    # print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 0)
    # print(*enumerate(s0.columns), '\n', sep = '\n')
    print(s0)
    s0["series"]=s0["ChemicalEngName_CASNo"].str.extract(r"(\d+\.)")


    s0["ChemicalEngName"]= s0.apply(lambda x: x['ChemicalEngName_CASNo'].replace(str(x['series']),""), axis=1)
    s0= s0[["series","ChemicalEngName"]]

    for n  in range (68,129):
        s0.loc[n, 'ChemicalEngName'] = "Volatile organic compounds that participate in atmospheric photochemical reactions, excluding the following:"+s0.loc[n, 'ChemicalEngName']
        s0.loc[n, 'series'] = "-"
        if n <= 115 and n>=112:
            print("xxx")
            s0.loc[n, 'ChemicalEngName'] =s0.loc[n, 'ChemicalEngName'].replace("Volatile organic compounds that participate in atmospheric photochemical reactions, excluding the following:","")
            s0.loc[n, 'ChemicalEngName'] ="Volatile organic compounds that participate in atmospheric photochemical reactions, excluding the following:(z.18)  methyl acetate and perfluorocarbon compounds that fall into the following classes, namely"+s0.loc[n, 'ChemicalEngName']

    for n  in range (221,263):
        s0.loc[n, 'ChemicalEngName'] = re.sub(r"\(a complex.*\)","",s0.loc[n, 'ChemicalEngName'])
        s0.loc[n, 'ChemicalEngName'] = re.sub(r"\(a complex.*","",s0.loc[n, 'ChemicalEngName'])
        s0.loc[n, 'ChemicalEngName'] = re.sub(r"\(a combination.*","",s0.loc[n, 'ChemicalEngName'])
        s0.loc[n, 'ChemicalEngName'] = "The following petroleum and refinery gases:"+ s0.loc[n, 'ChemicalEngName']
    for n  in range (221,263):
        s0.loc[n, "series"]="-"


    for n  in range (143,144):
        s0.loc[n, 'ChemicalEngName'] ="The following perfluorocarbons:" + s0.loc[n, 'ChemicalEngName']
        s0.loc[n, 'series'] = "-"
    s0.loc[193, 'series'] = np.nan

    for shift in range(1,5):
        s0["sf0"] =s0["ChemicalEngName"]
        s0["sf1"] =s0["ChemicalEngName"].shift(-1)
        s0["sf2"] =s0["ChemicalEngName"].shift(-2)
        s0["sf3"] =s0["ChemicalEngName"].shift(-3)
        s0["sf4"] =s0["ChemicalEngName"].shift(-4)
     
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1).isnull()==True)&(s0["sf1"].isnull()==False),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2).isnull()==True)&(s0['series'].shift(-1).isnull()==True)&(s0["sf2"].isnull()==False),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3).isnull()==True)&(s0['series'].shift(-2).isnull()==True)&(s0['series'].shift(-1).isnull()==True)&(s0["sf3"].isnull()==False),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4).isnull()==True)&(s0['series'].shift(-3).isnull()==True)&(s0['series'].shift(-2).isnull()==True)&(s0['series'].shift(-1).isnull()==True)&(s0["sf4"].isnull()==False),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf4"]

    s0= s0[["series","ChemicalEngName"]]
    s0 = s0[(s0.series.isnull()==False)]
    s0= s0[["ChemicalEngName"]]

    s0["ChemicalEngName"] =s0["ChemicalEngName"].str.replace(" that have the molecular formula ","; ")
    s0["ChemicalEngName"] =s0["ChemicalEngName"].str.replace(" that has the molecular formula ","; ")
    s0["ChemicalEngName"] =s0["ChemicalEngName"].str.replace(", which has the molecular formula ","; ")
    s0["ChemicalEngName"] =s0["ChemicalEngName"].str.replace(", which have the molecular formula ","; ")
    



    s0 =s0.reset_index(drop =True)
    a = [64,107,189,196]
    for i in a:
        s0 = s0.drop(index = i)
        s0 =s0.reset_index(drop =True)
    

    for n  in range (64,124):
        if n>64:
            s0.loc[n, "ChemicalEngName"]=s0.loc[n, "ChemicalEngName"].replace("Volatile organic compounds that participate in atmospheric photochemical reactions, excluding the following:","")
            
            s0.loc[64, "ChemicalEngName"] = s0.loc[64, "ChemicalEngName"]+", "+s0.loc[n, "ChemicalEngName"]

    for n  in range (65,171):
        s0 = s0.drop(index = n)
    s0 =s0.reset_index(drop =True)

    for n  in range (83,123):
        s0.loc[n, "ChemicalEngName"]=s0.loc[n, "ChemicalEngName"].replace("The following petroleum and refinery gases:","")
        s0.loc[n, "ChemicalEngName"]=s0.loc[n, "ChemicalEngName"].split(" ",1)[1]

    s0 =s0.reset_index(drop =True)

    s0.loc[128, "ChemicalEngName"]=s0.loc[128, "ChemicalEngName"].split("(a complex",1)[0]
    s0 =s0.reset_index(drop =True)



    print(s0)
    file = os.path.join('../2022_processed', '35.xlsx')
    s0.to_excel(file, index = False)
# n35()
def n36():
    file = os.path.join('../2022_input', '36.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 1)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(.+', '', regex = True)
    s0.CASNo = s0.CASNo.str.replace('NA.+', '-', regex = True)
    s0 = s0[["CASNo","ChemicalEngName"]]
    try:
        file = os.path.join('../2022_processed', '36_1.xlsx')
        x2 = pd.read_excel(file, index_col = 2)
        print(*enumerate(x2.columns), '\n', sep = '\n')

        s0 = s0[~s0.duplicated(subset = ['ChemicalEngName', 'CASNo'])]
        a = x2.index
        for i in a:
            if s0.loc[i-1, 'CASNo'] == '-':
                s0.loc[i-1, 'CASNo'] = x2.loc[i, 'Substance notes content']
            else:
                s0.loc[i-1, 'CASNo'] = s0.loc[i-1, 'CASNo'] + '; ' + x2.loc[i, 'Substance notes content']
        a = s0[s0.ChemicalEngName.str[-1].str.match('\d')].index
        a = a[a > 100]
        for i in a:
            s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-2]
        a = [256, 257, 259]
        for i in a:
            s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-3]
        s0 = s0[Columns.values()]
    except:
        pass

    file = os.path.join('../2022_processed', '36.xlsx')
    s0.to_excel(file, index = False)


def n37():
    file = os.path.join('../2022_input', '37.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.loc[(s0['Synonym'].isnull()==False),'ChemicalEngName'] = s0['ChemicalEngName'] + '; ' + s0["Synonym"]
    s0.CASNo.fillna('-', inplace = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\u3000', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\u3000', '', regex = False)
    s0 = s0[Columns.values()]
    print(s0)
    file = os.path.join('../2022_processed', '37.xlsx')
    
    s0.to_excel(file, index = False)

def n38():
    file = os.path.join('../2022_input', '38.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.loc[(s0['Synonym'].isnull()==False),'ChemicalEngName'] = s0['ChemicalEngName'] + '; ' + s0["Synonym"]
    s0.CASNo.fillna('-', inplace = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\u3000', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\u3000', '', regex = False)
    s0 = s0[Columns.values()]
    print(s0)
    file = os.path.join('../2022_processed', '38.xlsx')
    
    s0.to_excel(file, index = False)

def n39():
    file = os.path.join('../2022_input', '39.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.loc[(s0['Synonym'].isnull()==False),'ChemicalEngName'] = s0['ChemicalEngName'] + '; ' + s0["Synonym"]
    s0.CASNo.fillna('-', inplace = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\u3000', '', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\u3000', '', regex = False)
    s0 = s0[Columns.values()]
    print(s0)
    file = os.path.join('../2022_processed', '39.xlsx')
    
    s0.to_excel(file, index = False)

def n40():
    # step1,just copy table from web: https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en
    # step2,copy to 40_all
    # step3,edit xlsx by the following code
    
    file = os.path.join('../2022_input', '40.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 = s0[(s0.ChemicalEngName.isnull()==False)]
    print(s0)

    file = os.path.join('../2022_processed', '40.xlsx')
    s0.to_excel(file, index = False)
def n41():
    # step1,just copy table from web: https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en
    # step2,copy to 40_all
    # step3,edit xlsx by the following code
    
    file = os.path.join('../2022_input', '41.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 = s0[(s0.ChemicalEngName.isnull()==False)]
    print(s0)

    file = os.path.join('../2022_processed', '41.xlsx')
    s0.to_excel(file, index = False)

def n42():
    # step1,just copy table from web: https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en
    # step2,copy to 40_all
    # step3,edit xlsx by the following code
    
    file = os.path.join('../2022_input', '42.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 = s0[(s0.ChemicalEngName.isnull()==False)]
    print(s0)

    file = os.path.join('../2022_processed', '42.xlsx')
    s0.to_excel(file, index = False)

def n43():
    # step1,just copy table from web: https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en
    # step2,copy to 40_all
    # step3,edit xlsx by the following code
    
    file = os.path.join('../2022_input', '43.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0 = s0[(s0.ChemicalEngName.isnull()==False)]
    print(s0)

    file = os.path.join('../2022_processed', '43.xlsx')
    s0.to_excel(file, index = False)

# def n40_43():
#     file = os.path.join('../2022_input', '40_43.csv')
#     x = pd.read_csv(file, dtype = 'string')
#     print(*enumerate(x.columns), '\n', sep = '\n')

#     Columns = {x.columns[2]: 'ChemicalEngName',
#                x.columns[0]: 'CASNo'
#               }
#     df = x.rename(columns = Columns)
#     df.loc[df.CASNo.isna(), 'CASNo'] = '-'

#     s = df[df.Properties.str.contains('Priority Assessment Chemical Substances (PACSs)', regex = False)]
#     s = s[Columns.values()]
#     file = os.path.join('../2022_processed', '40.xlsx')
#     s.to_excel(file, index = False)

#     s = df[df.Properties.str.contains('Monitoring Chemical Substances', regex = False)]
#     s = s[Columns.values()]
#     file = os.path.join('../2022_processed', '41.xlsx')
#     s.to_excel(file, index = False)

#     s = df[df.Properties.str.contains('Class I specified chemical substance', regex = False)]
#     s = s[Columns.values()]
#     file = os.path.join('../2022_processed', '42.xlsx')
#     s.to_excel(file, index = False)

#     s = df[df.Properties.str.contains('Class II specified chemical substance', regex = False) == True]
#     s = s[Columns.values()]
#     file = os.path.join('../2022_processed', '43.xlsx')
#     s.to_excel(file, index = False)

# def n44():
#     # step1,just copy table from web: https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en
#     # step2,copy to 40_all
#     # step3,edit xlsx by the following code
    
#     file = os.path.join('../2022_input', '44.xlsx')
#     x = pd.ExcelFile(file)
#     print(*enumerate(x.sheet_names), '\n', sep = '\n')

#     s0 = x.parse(sheet_name = 0, skiprows = 2)
    
#     print(s0)
#     print(*enumerate(s0.columns), '\n', sep = '\n')
#     Columns = {s0.columns[2]: 'ChemicalEngName',
#                s0.columns[1]: 'CASNo'
#               }
#     s0 = s0.rename(columns = Columns)
#     s0 = s0[Columns.values()]
#     s0 = s0[(s0.ChemicalEngName.isnull()==False)]
#     print(s0)

#     file = os.path.join('../2022_processed', '44.xlsx')
#     s0.to_excel(file, index = False)

def n44_45():
    file = os.path.join('../2022_input', '44_45.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    x = x.parse(sheet_name = 0, skiprows = 2, nrows = 472)
    print(*enumerate(x.columns), '\n', sep = '\n')
    Columns = {x.columns[2]: 'ChemicalEngName',
               x.columns[1]: 'CASNo'
              }
    x = x.rename(columns = Columns)
    a = x[(x.CASNo.isna()) | (x.ChemicalEngName.isna())].index
    for i in a[::-1]:
        if pd.isna(x.loc[i, 'CASNo']):
            x.loc[i-1, 'ChemicalEngName'] = x.loc[i-1, 'ChemicalEngName'] + '; ' + x.loc[i, 'ChemicalEngName']
        elif pd.isna(x.loc[i, 'ChemicalEngName']):
            x.loc[i-1, 'CASNo'] = x.loc[i-1, 'CASNo'] + '; ' + x.loc[i, 'CASNo']
        x = x.drop(index = i)

    s = x[Columns.values()]
    file = os.path.join('../2022_processed', '44.xlsx')
    s.to_excel(file, index = False)

    s = x[x['Specific Class I Designated Chemical Substances*'] == 'S']
    s = s[Columns.values()]
    file = os.path.join('../2022_processed', '45.xlsx')
    s.to_excel(file, index = False)

def n46():
    file = os.path.join('../2022_input', '46.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 2)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '46.xlsx')
    s0.to_excel(file, index = False)

def n47():
    # step1,get hwp file from web: https://www.law.go.kr/LSW/admRulLsInfoP.do?admRulSeq=2100000021862#AJAX
    # step2,use https://appzend.herokuapp.com/hwpviewer/ get data
    # step3,edit xlsx by the following code
        
    file = os.path.join('../2022_input', '47.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' ；', '; ')
    s0.CASNo = s0.CASNo.str.replace(',', '; ')
    print(s0)


    file = os.path.join('../2022_processed', '47.xlsx')
    s0.to_excel(file, index = False)


def n48():
    file = os.path.join('../2022_processed', '48_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[3]: 'ChemicalEngName',
               s0.columns[4]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0.drop(index = s0.index[-1])

    
    # a = s0.ChemicalChnName.fillna('').str.contains('.', regex = False).index
    # for i in a:
    #     s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'].strip('0123456789. ')
    s0.别名 = s0.别名.str.replace('；', '; ', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(';', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace(';', '; ', regex = False)
    s0.CASNo = s0.CASNo.str.replace('\n', '', regex = False)
    a = s0[s0.别名.notna()].index
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'] + '; ' + s0.loc[i, '别名']
    a = s0[s0.ChemicalChnName == '常见品种如下：'].index
    s0 = s0.drop(index = a)
    s0.ChemicalEngName.fillna('-', inplace = True)
    s0.CASNo.fillna('-', inplace = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\[[含不稳浸无粉湿在固溶按比].+\]', '', regex = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\[.+[%的称℃％]\]', '', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\([(more)(not)(available)(active)(suspended)(containing)].+\)', '', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(.+[%(type)]\)', '', regex = True)
    
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace("　", '')
    s0.loc[s0['ChemicalChnName'].str[-2:]=="; ",'ChemicalChnName'] = s0['ChemicalChnName'].str[:-2]
    s0 = s0[Columns.values()]
    

    s0["series"]=s0["ChemicalChnName"].str[:3].str.extract(r'(\d+\.)')
    s0.loc[s0["series"].isnull() ==False,'ChemicalChnName'] = s0['ChemicalChnName'].str.replace(r'(\d+\.)','', regex = True)
    s0 = s0.drop(columns=["series"],axis=1)# 刪除判斷欄
    
    file = os.path.join('../2022_processed', '48.xlsx')
    s0.to_excel(file, index = False)

def n49():
    file = os.path.join('../2022_input', '49.xlsx')
    x = pd.ExcelFile(file)
    # print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, dtype = 'string')
    # print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    print(s0.index)
    for i in s0.index:
        if pd.isna(s0.loc[i, 'CASNo']):
            s0.loc[i, 'CASNo'] = s0.loc[i, '编号']
    for i in s0.index[::-1]:
        print(i)
        if pd.isna(s0.loc[i, 'ChemicalChnName']):
            s0.loc[i-1, 'CASNo'] = s0.loc[i-1, 'CASNo'] + '; ' + s0.loc[i, 'CASNo']
            print(i)
            s0 = s0.drop(index = i)
            
    s0.CASNo = s0.CASNo.str.replace('P.+', '-', regex = True)
    s0.CASNo = s0.CASNo.str.replace('\(.+\)', '', regex = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('）', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '49.xlsx')
    s0.to_excel(file, index = False)

def n50():
    file = os.path.join('../2022_input', '50.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(2):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    s0 = s0.drop(index = 9)
    Columns = {s0.columns[2]: 'ChemicalChnName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)

    s0 = s0.fillna('')
    for i in s0.index:
        if s0.loc[i, 'ChemicalChnName'] == '':
            s0.loc[i, 'ChemicalChnName'] = s0.loc[i, '化 学 品 名 称']
    s0.loc[7, 'ChemicalChnName'] = s0.loc[7, 'ChemicalChnName'].replace('（铵）', '')
    s0 = s0[Columns.values()]

    page = pdf.pages[2]
    table = page.extract_table()
    s1 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[1]: 'ChemicalChnName',
               s1.columns[2]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    df.ChemicalChnName = df.ChemicalChnName.str.replace(' *\n', '', regex = True)
    df.CASNo = df.CASNo.str.replace(' \n', '; ', regex = False)
    df.ChemicalChnName = df.ChemicalChnName.str.replace('（.*(包括).*）', '', regex = True)
    file = os.path.join('../2022_processed', '50.xlsx')
    df.to_excel(file, index = False)

def n56_57():
    # step1,get hwp file from web: https://documents-dds-ny.un.org/doc/UNDOC/GEN/V22/026/74/PDF/V2202674.pdf?OpenElement
    # step2,edit xlsx by the following code
    
    file = os.path.join('../2022_input', '56_57.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text"
    }
    page = pdf.pages[2]
    table = page.extract_table(table_settings)
    x = pd.DataFrame(table[:])
    print(*enumerate(x.columns), '\n', sep = '\n')

    s0 = pd.DataFrame(x[0])
    Columns = {s0.columns[0]: 'ChemicalEngName'}
    s0 = s0.rename(columns = Columns)

    s0 = s0[(s0.ChemicalEngName !="")]
    s0 = s0[(s0.ChemicalEngName !="Table")]
    s0 = s0[(s0.ChemicalEngName !="The salts of the substances listed in this")]
    s0 = s0[(s0.ChemicalEngName !="Table whenever the existence of such salts")]
    s0 = s0[(s0.ChemicalEngName !="is possible.")]
                           
    s0 =s0.reset_index(drop=True)
    print(s0)

    a = [4, 9, 11, 13]
    for i in a:
        s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + ' ' + s0.loc[i, 'ChemicalEngName']
        s0 = s0.drop(index = i)
        s0 =s0.reset_index(drop=True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(“*', '; ', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('”*\)', '', regex = True)

    print(s0)

    s1 = pd.DataFrame(x[1])
    Columns = {s1.columns[0]: 'ChemicalEngName'}
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1.ChemicalEngName != '']
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('acida', 'acid', regex = False)
    s1 = s1[(s1.ChemicalEngName !="The salts of the substan")]
    s1 = s1[(s1.ChemicalEngName !="Table whenever the ex")]
    s1 = s1[(s1.ChemicalEngName !="salts is possible.")]
    s1 = s1[(s1.ChemicalEngName !="a The salts of hydroch")]
    s1 = s1[(s1.ChemicalEngName !="acid are specifically")]

    file = os.path.join('../2022_processed', '56.xlsx')
    s0.to_excel(file, index = False)
    file = os.path.join('../2022_processed', '57.xlsx')
    s1.to_excel(file, index = False)


def n58():
    # step1,get hwp file from web: https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX%3A02004R0273-20210113
    # step2,edit xlsx by the following code
    
    file = os.path.join('../2022_processed', '58_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    print(s0)
    s0 = s0[s0.CASNo.notna()]
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', ' ', regex = False)
    a = list(s0[s0.ChemicalEngName.str.contains(' (', regex = False)].index)
    a.remove(12)
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'].replace(' (', '; ')
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'].replace(')', '')
    a = s0[s0['CN designation\n(if different)'] != '\xa0'].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'] + '; ' + s0.loc[i, 'CN designation\n(if different)']
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '58.xlsx')
    s0.to_excel(file, index = False)

def n59():
    # step1,get hwp file from web: https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX%3A02004R0273-20210113
    # step2,edit xlsx by the following code
    file = os.path.join('../2022_processed', '59_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo.notna()]
    s0 = s0[Columns.values()]

    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[0]: 'ChemicalEngName',
               s1.columns[3]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1.CASNo.notna()]
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('../2022_processed', '59.xlsx')
    df.to_excel(file, index = False)

def n60():
    # step1,get hwp file from web: https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX%3A02004R0273-20210113
    # step2,edit xlsx by the following code
    file = os.path.join('../2022_processed', '60_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo.notna()]
    a = s0[s0['CN designation\n(if different)'] != '\xa0'].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'] + '; ' + s0.loc[i, 'CN designation\n(if different)']
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '60.xlsx')
    s0.to_excel(file, index = False)

def n61():
    file = os.path.join('../2022_processed', '61_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    # s0 = x.parse(sheet_name = 0, nrows = 35, header = None)
    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(*\d+\)*(\xa0)', '', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(', including:', '', regex = False)

    s0.loc[s0['ChemicalEngName'].str[-1:] =="1",'ChemicalEngName'] = s0['ChemicalEngName'].str[:-1]

    file = os.path.join('../2022_processed', '61.xlsx')
    s0.to_excel(file, index = False)

def n62():
    file = os.path.join('../2022_processed', '62_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    # s0 = x.parse(sheet_name = 0, nrows = 6, header = None)
    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(*\d+\)*(\xa0)', '', regex = True)

    file = os.path.join('../2022_processed', '62.xlsx')
    s0.to_excel(file, index = False)

def n63():
    file = os.path.join('../2022_processed', '63_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, header = None)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName'}
    s0 = s0.rename(columns = Columns)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(\d+\) ', '', regex = True)
    a = s0[s0.ChemicalEngName.str.contains('Other names')].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-1]
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(Other names.*:', ';', regex = True)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '63.xlsx')
    s0.to_excel(file, index = False)

def n64():
    file = os.path.join('../2022_processed', '64_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, header = None)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName'}
    s0 = s0.rename(columns = Columns)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(\d+\) *', '', regex = True)
    s0.loc[5, 'ChemicalEngName'] = s0.loc[5, 'ChemicalEngName'].replace(' (or', ';')
    s0.loc[5, 'ChemicalEngName'] = s0.loc[5, 'ChemicalEngName'].replace(' or', ';')
    s0.loc[5, 'ChemicalEngName'] = s0.loc[5, 'ChemicalEngName'].replace(')', '')
    s0.loc[9, 'ChemicalEngName'] = s0.loc[9, 'ChemicalEngName'].replace(' (', '; ')
    s0.loc[9, 'ChemicalEngName'] = s0.loc[9, 'ChemicalEngName'].replace(')', '')
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '64.xlsx')
    s0.to_excel(file, index = False)

def n65_67():
    file = os.path.join('../2022_processed', '65_67.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, header = None)
    s0.columns = ['ChemicalEngName']
    a = s0[s0.ChemicalEngName.str.contains(' (', regex = False)].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-1]
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' (', '; ', regex = False)

    s1 = x.parse(sheet_name = 1, header = None)
    s1.columns = ['ChemicalEngName']
    s2 = x.parse(sheet_name = 2, header = None)
    s2.columns = ['ChemicalEngName']
    df = pd.concat([s1, s2], ignore_index = True)
    
    s3 = x.parse(sheet_name = 3, header = None)
    s3.columns = ['ChemicalEngName']
    s3.loc[3, 'ChemicalEngName'] = s3.loc[3, 'ChemicalEngName'][:-1]
    s3.loc[3, 'ChemicalEngName'] = s3.loc[3, 'ChemicalEngName'].replace(' (', '; ')
    
    file = os.path.join('../2022_processed', '65.xlsx')
    s0.to_excel(file, index = False)
    file = os.path.join('../2022_processed', '66.xlsx')
    df.to_excel(file, index = False)
    file = os.path.join('../2022_processed', '67.xlsx')
    s3.to_excel(file, index = False) 

def n68():
    try:
        file = os.path.join('../2022_processed', '68_1.xlsx')
        x = pd.ExcelFile(file)
        print(*enumerate(x.sheet_names), '\n', sep = '\n')

        s0 = x.parse(sheet_name = 0, nrows = 59)
        print(*enumerate(s0.columns), '\n', sep = '\n')
        Columns = {s0.columns[0]: 'ChemicalEngName',
                s0.columns[1]: 'CASNo'
                }
        s0 = s0.rename(columns = Columns)
        s0 = s0[s0['Decision Guidance Documents'].notna()]
        s0.CASNo =  s0.CASNo.str.replace(' (*)', '', regex = False)
        s0.CASNo =  s0.CASNo.str.replace(',', ';', regex = False)
        s0.loc[24, 'CASNo'] = '-'
        s0 = s0[Columns.values()]

        file = os.path.join('../2022_input', '68_2.pdf')
        pdf = pdfplumber.open(file)
        table = []
        for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            table += page.extract_table()
        s1 = pd.DataFrame(table, columns = None)
        print(*enumerate(s1.columns), '\n', sep = '\n')

        for i in range(19, 39):
            s1.iloc[i, 0] = s1.iloc[i, 1]
            s1.iloc[i, 1] = s1.iloc[i, 2]
        Columns = {1: 'ChemicalEngName',
                0: 'CASNo'
                }
        s1 = s1.rename(columns = Columns)
        s1 = s1[s1.ChemicalEngName.notna()]
        s1 = s1[s1.ChemicalEngName != '']
        s1.loc[s1.CASNo == '', 'CASNo'] = '-'
        s1 = s1[Columns.values()]

        df = pd.concat([s0, s1], ignore_index = True)
        file = os.path.join('../2022_processed', '68.xlsx')
        df.to_excel(file, index = False)
    except:
        file = os.path.join('../2022_input', '68.xlsx')
        x = pd.ExcelFile(file)
        # print(*enumerate(x.sheet_names), '\n', sep = '\n')

        s0 = x.parse(sheet_name = 0, header = 0)
        print(*enumerate(s0.columns), '\n', sep = '\n')
        Columns = {s0.columns[0]: 'ChemicalEngName',
                s0.columns[1]: 'CASNo'
                }
        s0 = s0.rename(columns = Columns)
        s0 = s0[Columns.values()]
        s0 = s0[(s0.ChemicalEngName.isnull() ==False)&(s0.CASNo.isnull() ==False)]

        s0.CASNo = s0.CASNo.str.replace(',', ';', regex = True)
        s0.CASNo = s0.CASNo.str.replace(r" \(\*\)", '', regex = True)

        s0 =s0.reset_index(drop=True)
        a = [30, 44, 45, 46, 47]
        for i in a:
            s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'].replace(r' (', '; ')
            s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'].replace(r')', '')
            # print(s0.loc[i, 'ChemicalEngName'])

        file = os.path.join('../2022_processed', '68.xlsx')
        s0.to_excel(file, index = False)
        
def n69():
    file = os.path.join('../2022_input', '69.csv')
    s0 = pd.read_csv(file, dtype = 'string')
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo.fillna('-', inplace = True)
    s0 = s0[Columns.values()]
    print(s0)

    file = os.path.join('../2022_input', '69_2.csv')
    s1 = pd.read_csv(file)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[0]: 'ChemicalEngName',
               s1.columns[2]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1.CASNo.fillna('-', inplace = True)
    s1 = s1[Columns.values()]
    print(s1)

    file = os.path.join('../2022_input', '69_3.csv')
    s2 = pd.read_csv(file)
    print(*enumerate(s2.columns), '\n', sep = '\n')
    Columns = {s2.columns[0]: 'ChemicalEngName',
               s2.columns[2]: 'CASNo'
              }
    s2 = s2.rename(columns = Columns)
    s2.CASNo.fillna('-', inplace = True)
    s2 = s2[Columns.values()]
    print(s2)

    df = pd.concat([s0, s1, s2], ignore_index = True)
    file = os.path.join('../2022_processed', '69.xlsx')
    df.to_excel(file, index = False)

def n72():
    file = os.path.join('../2022_input', '72.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName != 'Chemicals of Interest \n(COI)']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(conc.+\)', '', regex = True)
    s0.Synonym = s0.Synonym.str.replace('[\n\[\]]', '', regex = True)
    s0.Synonym = s0.Synonym.str.replace(' or', ';', regex = False)
    a = s0[s0.Synonym != ''].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'] + '; ' + s0.loc[i, 'Synonym']
    s0.CASNo.fillna('-', inplace = True)
    s0 = s0[Columns.values()]

    file = os.path.join('../2022_processed', '72.xlsx')
    s0.to_excel(file, index = False)

def n74():
    
    file = os.path.join('../2022_input', '74.pdf')
    pdf = pdfplumber.open(file)
    str_ = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        str_ += '\n' + page.extract_text()

    str_= StringIO("".join(str_))
    s0 = pd.read_table(str_, sep='\n',skiprows = 3, header = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)

    s0["series"]=s0["ChemicalEngName"].str.extract(r'(\d+\. )')
    s0 = s0[(s0.series.isnull()==False)]
    # s0["ChemicalEngName"]=s0["ChemicalEngName"].str.extract(r'(\d+\. .+,)')
    s0["ChemicalEngName"]=s0["ChemicalEngName"].str.split(r'(\d+\. )').str[2]
    s0["ChemicalEngName"]=s0["ChemicalEngName"].str.split(r"[,|.]").str[0]
    s0 = s0[Columns.values()]
    s0 = s0.reset_index(drop=True)

    file = os.path.join('../2022_processed', '74.xlsx')
    s0.to_excel(file, index = False)


def n76():
    file = os.path.join('../2022_processed', '76_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    # s0.loc[87, 'CASNo'] = s0.loc[87, s0.columns[17]]
    # s0.loc[94:, 'CASNo'] = s0.loc[94:, s0.columns[18]]
    # s0 = s0[s0.ChemicalChnName.notna()]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\[.+\]', '', regex = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（.+）', '', regex = True)
    s0.别名 = s0.别名.str.replace('[、；]', '; ', regex = True)
    s0.CASNo.fillna('-', inplace = True)
    # print(s0)
    # a = list(s0[s0.CASNo == '-'].index)
    # a.remove(56)
    # s0 = s0.drop(index = a)    
    a = s0[s0.别名.notna()].index
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'] + '; ' + s0.loc[i, '别名']
    s0 = s0[Columns.values()]
    s0 = s0[(s0.ChemicalChnName.isnull()==False)]
    # print(s0)
    

    cc = OpenCC('s2twp')
    #簡 轉 繁
    s0['ChemicalChnName'] = s0['ChemicalChnName'].apply(lambda x:  cc.convert(x))

    s0 = s0[(s0.CASNo != "-") | (s0.ChemicalChnName == "鎂鋁粉; 鎂鋁合金粉")]
    print(s0)
    file = os.path.join('../2022_processed', '76.xlsx')
    s0.to_excel(file, index = False)

def n77():
    file = os.path.join('../2022_processed', '77_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0 = s0.drop(index = 9)
    Columns = ['ChemicalEngName', 'CASNo']
    s0[Columns] = s0[s0.columns[0]].str.split(' \(CAS RN ', expand = True)
    s0.CASNo = s0.CASNo.str.replace(')', '', regex = False)
    s0 = s0[Columns]

    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    s1 = s1.drop(index = 9)
    Columns = ['ChemicalEngName', 'CASNo']
    s1[Columns] = s1[s1.columns[0]].str.split(' \(CAS RN ', expand = True)
    s1.CASNo = s1.CASNo.str.replace(')', '', regex = False)
    # s1.CASNo = s1.CASNo.str.replace(r'\xa0(2\xa0(3 ', '', regex = False)
    # s1.CASNo = s1.CASNo.str.replace(r'\xa0(2\xa0(3 ', '', regex = False)
    s1.CASNo = s1.CASNo.str.replace(r'(\(.+)','')
    s1.CASNo = s1.CASNo.str.replace(r'(\xa0)','')

    s1 = s1[Columns]

    df = pd.concat([s0, s1], ignore_index = True)
    print(df)
    file = os.path.join('../2022_processed', '77.xlsx')
    df.to_excel(file, index = False)

def n78():
    file = os.path.join('../2022_processed', '78_1.xlsx')
    x = pd.ExcelFile(file)
    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]
    s0=s0.append({'ChemicalEngName' : 'Aniline'} , ignore_index=True)


    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[0]: 'ChemicalEngName',
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[Columns.values()]
    
    df = pd.concat([s0, s1], ignore_index = True)
    df = df[(df.ChemicalEngName != "•")]
    df.ChemicalEngName = df.ChemicalEngName.str.replace(r' \(default.*?\)', '', regex = True)
    df.ChemicalEngName = df.ChemicalEngName.str.replace(r' \(default.*', '', regex = True)
    df.ChemicalEngName = df.ChemicalEngName.str.replace(r'\(default.*', '', regex = True)
    df = df[(df.ChemicalEngName != "")]

    df = df[(df.ChemicalEngName.str[:5] != "lang=")]
    df = df.reset_index(drop=True)

    print(df)
    file = os.path.join('../2022_processed', '78.xlsx')
    df.to_excel(file, index = False)


def n79():

    file = os.path.join('../2022_input', '79.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')
    print(s0)

    file = os.path.join('../2022_input', '79_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s1.columns), '\n', sep = '\n')
    print(s1)

    file = os.path.join('../2022_input', '79_3.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s2 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s2.columns), '\n', sep = '\n')
    print(s2)

    file = os.path.join('../2022_input', '79_4.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s3 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s3.columns), '\n', sep = '\n')
    print(s3)

    file = os.path.join('../2022_input', '79_5.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s4 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s4.columns), '\n', sep = '\n')
    print(s4)


    df = pd.concat([s0, s1,s2,s3,s4], ignore_index = True)
    Columns = {df.columns[0]: 'ChemicalEngName',
               df.columns[1]: 'CASNo'
              }
    df = df.rename(columns = Columns)
    df.CASNo.fillna('-', inplace = True)
    df = df[Columns.values()]
    

    # df.loc[(df.ChemicalEngName.str[:6] == "Sodium")&(df.ChemicalEngName.str[-11:] == "uoroacetate"),'ChemicalEngName'] = "Sodium fluoroacetate"
    # df.loc[(df.ChemicalEngName.str[:11] == "Fluoroethyl")&(df.ChemicalEngName.str[-11:] == "uoroacetate"),'ChemicalEngName'] = "Fluoroethyl fluoroacetate"
    # df.loc[2, 'ChemicalEngName'] = df.loc[2, 'ChemicalEngName'].replace(' (', '; ')
    # df.loc[2, 'ChemicalEngName'] = df.loc[2, 'ChemicalEngName'].replace(')', '')
    # df.ChemicalEngName = df.ChemicalEngName.str.replace(' *', '', regex = False)

    print(df)
    file = os.path.join('../2022_processed', '79.xlsx')
    df.to_excel(file, index = False)

def n80():
    file = os.path.join('../2022_processed', '80_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalChnName != '化学品名称']
    a = s0[s0.别名.notna()].index
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'] + '; ' + s0.loc[i, '别名']
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('、', '; ', regex = False)
    s0.loc[6, 'ChemicalChnName'] = s0.loc[6, 'ChemicalChnName'].replace('; ', '、', 1)
    s0.CASNo = s0.CASNo.str.replace('（.+）', '', regex = True)
    s0.CASNo.fillna('-', inplace = True)
    s0 = s0[Columns.values()]

    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[1]: 'ChemicalChnName',
               s1.columns[2]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1.loc[11, 'ChemicalChnName'] = s1.loc[11, 'ChemicalChnName'].replace('\n（即', '; ')
    s1.loc[11, 'ChemicalChnName'] = s1.loc[11, 'ChemicalChnName'].replace('）', '')
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)

    cc = OpenCC('s2twp')
    #簡 轉 繁
    df['ChemicalChnName'] = df['ChemicalChnName'].apply(lambda x:  cc.convert(x))

    file = os.path.join('../2022_processed', '80.xlsx')
    df.to_excel(file, index = False)

def n82():
    file = os.path.join('../2022_input', '82.pdf')
    pdf = pdfplumber.open(file)
    str_ = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        str_ += '\n' + page.extract_text().replace('\r',"")
    str_= StringIO("".join(str_))
    s0 = pd.read_table(str_, sep='\n',skiprows = 2, header = 1)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalChnName',
              }
    s0 = s0.rename(columns = Columns)

    s0["ChemicalChnName"]=s0["ChemicalChnName"].str.replace(r'(\d+\、)','',regex =True)
    s0.drop(13, inplace=True)
    s0.drop(14, inplace=True)
    print(s0)

    file = os.path.join('../2022_processed', '82.xlsx')
    s0.to_excel(file, index = False)


def n83():
    file = os.path.join('../2022_input', '83.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict"
    }
    table = []
    for i in range(2):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo != 'CAS 号']
    s0 = s0[~s0.ChemicalChnName.str.contains('包括', regex = False)]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('*', '', regex = False)
    s0.loc[25, 'CASNo'] = s0.loc[25, 'CASNo'].replace(' \n', '; ')
    s0.CASNo = s0.CASNo.str.replace(' *\n.+', '', regex = True)
    s0 = s0[Columns.values()]

    cc = OpenCC('s2twp')
    #簡 轉 繁
    s0['ChemicalChnName'] = s0['ChemicalChnName'].apply(lambda x:  cc.convert(x))

    file = os.path.join('../2022_processed', '83.xlsx')
    s0.to_excel(file, index = False)

def n83():
    file = os.path.join('../2022_input', '83.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict"
    }
    table = []
    for i in range(2):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo != 'CAS 号']
    s0 = s0[~s0.ChemicalChnName.str.contains('包括', regex = False)]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('*', '', regex = False)
    s0.loc[25, 'CASNo'] = s0.loc[25, 'CASNo'].replace(' \n', '; ')
    s0.CASNo = s0.CASNo.str.replace(' *\n.+', '', regex = True)
    s0 = s0[Columns.values()]

    cc = OpenCC('s2twp')
    #簡 轉 繁
    s0['ChemicalChnName'] = s0['ChemicalChnName'].apply(lambda x:  cc.convert(x))

    file = os.path.join('../2022_processed', '83.xlsx')
    s0.to_excel(file, index = False)


def n84():
    file = os.path.join('../2022_input', '84.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict"
    }

    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')
    print(s0)

    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(r'[\t|\r|\n]', '',regex =True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('    ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('   ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('  ', ' ')
    s0.CASNo = s0.CASNo.str.replace(r'[\t|\r|\n]', '',regex =True)
    # s0.CASNo = s0.CASNo.str.replace(' & ', '; ')

    s0["check"]=s0["CASNo"].str.extract(r'(\(CAS No.\))')
    s0 = s0[(s0.check != "(CAS No.)")]
    s0 = s0[(s0.CASNo.isnull() != True)]

    s0 = s0[Columns.values()]


    print(s0)
    file = os.path.join('../2022_processed', '84.xlsx')
    s0.to_excel(file, index = False)

def n85():
    file = os.path.join('../2022_input', '85.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict"
    }

    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)
    s0 = pd.DataFrame(table[:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')
 

    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(r'[\t|\r|\n]', '',regex =True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('    ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('   ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('  ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(r'※.+', '',regex =True)

    
    s0.CASNo = s0.CASNo.str.replace(r'[\t|\r|\n]', '',regex =True)
    # s0.CASNo = s0.CASNo.str.replace(' & ', '; ')

    s0["check"]=s0["CASNo"].str.extract(r'(\(CAS No.\))')
    s0 = s0[(s0.check != "(CAS No.)")]
    # s0 = s0[(s0.CASNo.isnull() != True)]

    s0 = s0[Columns.values()]


    print(s0)
    file = os.path.join('../2022_processed', '85.xlsx')
    s0.to_excel(file, index = False)

def n86():
    file = os.path.join('../2022_input', '86.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict"
    }

    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table(table_settings)

    s0 = pd.DataFrame(table[:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')
 

    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(r'[\t|\r|\n]', '',regex =True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('    ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('   ', ' ')
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('  ', ' ')
    
    s0.CASNo = s0.CASNo.str.replace(r'[\t|\r|\n]', '',regex =True)

    s0["check"]=s0["CASNo"].str.extract(r'(\(CAS No.\))')
    s0 = s0[(s0.check != "(CAS No.)")]
    # # s0 = s0[(s0.CASNo.isnull() != True)]

    s0 = s0[Columns.values()]
    s0 = s0.reset_index(drop=True)
    s0.loc[400, 'ChemicalEngName'] = "Barium cadmium calcium chloride fluoride phosphate, antimony an d manganese-doped"
    s0 = s0[(s0.CASNo != "")]


    print(s0)
    file = os.path.join('../2022_processed', '86.xlsx')
    s0.to_excel(file, index = False)

def n87():
    file = os.path.join('../2022_input', '87.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        # "vertical_strategy": "lines_strict",
        # "horizontal_strategy": "lines_strict",
        # "join_tolerance": 100
        "intersection_tolerance": 100
                    }
    table = []
    for i in range(142,145):
        page = pdf.pages[i]
        left =  page.crop((0.0 * float(page.width), 0.0 * float(page.height), 0.5 * float(page.width), 1.0 * float(page.height)))# default:        left = page.crop((0, 0.4 * float(page.height), 0.5 * float(page.width), 0.9 * float(page.height)))
        right = page.crop((0.5 * float(page.width), 0.0 * float(page.height), 1.0 * float(page.width), 1.0 * float(page.height)))# default:        right = page.crop((0.5 * float(page.width), 0.4 * float(page.height), page.width, 0.9 * float(page.height)))
        table += left.extract_table(table_settings)
        table += right.extract_table(table_settings)
        
    s0 = pd.DataFrame(table[:])
    Columns = {s0.columns[4]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)


    s0 = s0.reset_index(drop = True)

    
    s0.loc[s0['CASNo'] =='CASNo'] = "-"
    s0['series'] = s0['CASNo']


    for shift in range(1,5):
        s0["sf0"] =s0["ChemicalEngName"]
        s0["sf1"] =s0["ChemicalEngName"].shift(-1)
        s0["sf2"] =s0["ChemicalEngName"].shift(-2)
        s0["sf3"] =s0["ChemicalEngName"].shift(-3)
        s0["sf4"] =s0["ChemicalEngName"].shift(-4)
     
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-1)=="")&(s0["sf1"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf1"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf2"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf2"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf3"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf3"]
    s0.loc[(s0['series'] !="-")&(s0['series'].shift(-4)=="")&(s0['series'].shift(-3)=="")&(s0['series'].shift(-2)=="")&(s0['series'].shift(-1)=="")&(s0["sf4"]!=""),'ChemicalEngName'] = s0["ChemicalEngName"]+ " " + s0["sf4"]

    s0 = s0[(s0.CASNo !="CAS")& (s0.CASNo !="")]
    s0 = s0[Columns.values()]

    print(s0)
    file = os.path.join('../2022_processed', '87.xlsx')
    s0.to_excel(file, index = False)

def Merged():
    print("run to merge!!")
    Columns = ['ChemicalChnName',
               'ChemicalEngName',
               'CASNo',
               'Name'
              ]
    df = pd.DataFrame(columns = Columns)
    os.chdir('../2022_processed')
  
    for dirpath, dirnames, filenames in os.walk('.'):
        print(dirnames)
        if "backup" in dirnames:
            dirnames.remove("backup")
        for name in filenames:
            if '_' not in name:
                print(name)
                df2 = pd.read_excel(name)
                df = pd.concat([df, df2], ignore_index = True)
                df.Name.fillna(name[:-5], inplace = True)
    
    
    
    os.chdir('..')
    file = os.path.join('explain_and_structure', '清單資料來源_20220707_PC.xlsx')

    df2 = pd.read_excel(file, dtype = 'string')
    
    df2['Source_link'] = df2['資料來源']
    df2.loc[df2['資料來源(PC改)'] !="",'Source_link'] = df2['資料來源(PC改)']

    Columns = {'國內/外': 'Type',
               '單位/國家': 'Unit',
               '清單': 'Source',
               '編號': 'Source_ID'  
              }
    df2 = df2.rename(columns = Columns)
    df2.Type = df2.Type.replace('國內', '0', regex = False)
    df2.Type = df2.Type.replace('國外', '1', regex = False)
    Merged = pd.merge(left = df, right = df2, how = 'left',
                      left_on = 'Name', right_on = 'Source_ID',
                      validate = 'many_to_one'
                     )
    Columns = ['Source_ID',
               'ChemicalChnName',
               'ChemicalEngName',
               'CASNo',
               'Source',
               'Type',
               'Unit',
               'Source_link'
              ]
    Merged = Merged[Columns]
    Merged = Merged.fillna('-')
    
    cc = OpenCC('s2twp')
    Merged['ChemicalChnName'] = Merged['ChemicalChnName'].apply(lambda x:  cc.convert(x))

    Merged.ChemicalEngName = Merged.ChemicalEngName.str.lower()

    Merged = Merged[~Merged.duplicated()]
    
    #若CASNo不唯一，分成多行
    Merged.CASNo = Merged.CASNo.str.split(';') 
    # Merged.CASNo = Merged.CASNo.str.split(r"[;| ]")                                                              
    Merged = Merged.explode('CASNo')

    #去前後空格
    Merged.ChemicalEngName = Merged.ChemicalEngName.str.strip().replace(r'[\t|\r|\n]', '',regex =True).replace(r'(\xa0)','')
    Merged.ChemicalChnName = Merged.ChemicalChnName.str.strip().replace(r'[\t|\r|\n]', '',regex =True).replace(r'(\xa0)','')
    Merged.CASNo = Merged.CASNo.str.strip().replace(r'[\t|\r|\n]', '',regex =True).replace(r'(\xa0)','')

    #若CASNo不唯一，分成多行
    Merged.CASNo = Merged.CASNo.str.split(' ')                                                            
    Merged = Merged.explode('CASNo')
    Unique_array = np.arange(len(Merged))+1 #Unique_ID/ Unique_ship_count
    Merged['this_year_id'] = (Unique_array).astype(str)#給ID

    file = os.path.join('output', 'Merged.xlsx')
    Merged.to_excel(file, index = False)
# Merged()