# pip install opencc-python-reimplemented
# pip install fake_useragent==1.2.1
# pip install PyPDF2
from fake_useragent import UserAgent
from bs4 import BeautifulSoup as bs
from opencc import OpenCC
import pandas as pd
import pdfplumber
import requests
import tabula
import PyPDF2
import json
import re
import os


ua = UserAgent()
headers={ "User-Agent": ua.random }


input_path = '2023_input'
processed_path = '2023_processed'


pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.max_colwidth', 30)
pd.set_option('display.max_columns', 400)
pd.options.display.width = 0 



#region D_moi_01 公共危險物品及可燃性高壓氣體
def D_moi_01():
    print('原來是odt（附表一公共危險物品之種類、分級及管制量.ODT），有表格，但格式參差不齊')


#endregion

#region D_moi_02 簡易爆裂物(IED)原料管制清單
def D_moi_02():
    print('105年會議資料，無需處理;')


#endregion

#region D_moj_01_a 毒品先驅原料
def D_moj_01_a():
    file = os.path.join(input_path, 'D_moj_01_a.PDF')
    pdfFileObj = open(file, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    text = ''
    for page in [0]:
        pageObj = pdfReader.pages[0]
        text+=pageObj.extract_text()
        
    content = [i.strip() for i in re.findall('\d、.{1,}?\s', text)]

    chinese_words = []
    english_words = []

    for item in content:
        # 提取中文部分
        chinese_match = re.search(r'、(.*[\u4e00-\u9fa5\dA-Z]+)(?=\（)', item)
        if chinese_match:
            chinese_words.append(chinese_match.group(1))
        else:
            chinese_words.append('-')
        
        # 提取英文部分
        english_match = re.search(r'\（([^（]*?)\）$', item)
        if english_match:
            english_words.append(english_match.group(1))
        else:
            english_words.append('-')
    
    df = pd.DataFrame({'ChemicalChnName': chinese_words, 'ChemicalEngName': english_words})
    df = df[~df.ChemicalEngName.str.contains('刪除')]
    df.to_excel(os.path.join(processed_path, 'D_moj_01_a.xlsx'), index = False)

#endregion

#region D_moj_01_b 毒品先驅原料
def D_moj_01_b():
    file = os.path.join(input_path, 'D_moj_01_b.PDF')
    pdfFileObj = open(file, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    print(len(pdfReader.pages))
    text = ''
    for page in range(8):
        pageObj = pdfReader.pages[page]
        text+=pageObj.extract_text()
        
    content = [i.strip() for i in re.findall('\d+?、.{1,}?\s', text)]

    chinese_words = []
    english_words = []

    for item in content:
        # 提取中文部分
        chinese_match = re.search(r'、(.*[\u4e00-\u9fa5\dA-Z]+)(?=\（)', item)
        if chinese_match:
            chinese_words.append(chinese_match.group(1))
        else:
            chinese_words.append('-')
        
        # 提取英文部分
        english_match = re.search(r'\（([^（]*?)\）$', item)
        if english_match:
            english_words.append(english_match.group(1))
        else:
            english_words.append('-')
    
    df = pd.DataFrame({'ChemicalChnName': chinese_words, 'ChemicalEngName': english_words})
    df = df[~df.ChemicalEngName.str.contains('刪除')]
    df.to_excel(os.path.join(processed_path, 'D_moj_01_b.xlsx'), index = False)

#endregion

#region D_moj_01_c 毒品先驅原料
def D_moj_01_c():
    file = os.path.join(input_path, 'D_moj_01_c.PDF')
    pdfFileObj = open(file, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    print(len(pdfReader.pages))
    text = ''
    for page in range(20):
        pageObj = pdfReader.pages[page]
        text+=pageObj.extract_text()
    text = text.replace('\n', ' ')
    
    content = [i.strip() for i in re.findall('(\d+、.*?)(?=\d+、|$)', text)]

    chinese_words = []
    english_words = []

    for item in content:
        # 提取中文部分
        chinese_match = re.search(r'、(.*[\u4e00-\u9fa5\dA-Z]+)(?=\（)', item)
        if chinese_match:
            chinese_words.append(chinese_match.group(1))
        else:
            chinese_words.append('-')
        
        # 提取英文部分
        english_match = re.search(r'\（([^（]*?)\）$', item)
        if english_match:
            english_words.append(english_match.group(1))
        else:
            english_words.append('-')
    
    df = pd.DataFrame({'ChemicalChnName': chinese_words, 'ChemicalEngName': english_words})
    df = df[~df.ChemicalEngName.str.contains('刪除')]
    df.to_excel(os.path.join(processed_path, 'D_moj_01_c.xlsx'), index = False)

#endregion

#region D_moj_01_d 毒品先驅原料
def D_moj_01_d():
    file = os.path.join(input_path, 'D_moj_01_d.PDF')
    pdfFileObj = open(file, 'rb')
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    print(len(pdfReader.pages))
    text = ''
    for page in range(4):
        pageObj = pdfReader.pages[page]
        text+=pageObj.extract_text()
    text = text.replace('\n', ' ')
    
    content = [i.strip() for i in re.findall('(\d+、.*?)(?=\d+、|$)', text)]

    chinese_words = []
    english_words = []

    for item in content:
        # 提取中文部分
        chinese_match = re.search(r'、(.*[\u4e00-\u9fa5\dA-Z]+)(?=\（)', item)
        if chinese_match:
            chinese_words.append(chinese_match.group(1))
        else:
            chinese_words.append('-')
        
        # 提取英文部分
        english_match = re.search(r'\（([^（]*?)\）$', item)
        if english_match:
            english_words.append(english_match.group(1))
        else:
            english_words.append('-')
    
    df = pd.DataFrame({'ChemicalChnName': chinese_words, 'ChemicalEngName': english_words})
    df = df[~df.ChemicalEngName.str.contains('刪除')]
    df.to_excel(os.path.join(processed_path, 'D_moj_01_d.xlsx'), index = False)

#endregion

#region D_mol_01 GHS危害物質名單
def D_mol_01():
    print("目前無法和勞動部取得最新資料，直接沿用")


#endregion

#region D_mol_02 管制性化學品
def D_mol_02():
    file = os.path.join(input_path, 'D_mol_02.xlsx')
    df = pd.read_excel(file, header=None)
    df = df.rename(columns = {0: 'ChemicalChnName'})
    df.ChemicalChnName = df.ChemicalChnName.apply(lambda i: i.split('、')[1].strip())
    file = os.path.join(processed_path, 'D_mol_02.xlsx')
    df.to_excel(file, index = False)


#endregion

#region D_mol_03_a 優先管理化學品（第2條第1款）
def D_mol_03_a():
    file = os.path.join(input_path, 'D_mol_03_a.xlsx')
    df = pd.read_excel(file)    
    df = df[['CAS No.', '英文名稱', '中文名稱']]
    df = df.rename(columns = {'CAS No.': 'CASNo', '英文名稱': 'ChemicalEngName', '中文名稱': 'ChemicalChnName'})
    file = os.path.join(processed_path, 'D_mol_03_a.xlsx')
    df.to_excel(file, index = False)

def D_mol_03_b():
    file = os.path.join(input_path, 'D_mol_03_b.xlsx')
    df = pd.read_excel(file)    
    df = df[['CAS No.', '英文名稱', '中文名稱']]
    df = df.rename(columns = {'CAS No.': 'CASNo', '英文名稱': 'ChemicalEngName', '中文名稱': 'ChemicalChnName'})
    file = os.path.join(processed_path, 'D_mol_03_b.xlsx')
    df.to_excel(file, index = False)
    
def D_mol_03_c():
    file = os.path.join(input_path, 'D_mol_03_c.xlsx')
    df = pd.read_excel(file)    
    df = df[['CAS No.', '英文名稱', '中文名稱']]
    df = df.rename(columns = {'CAS No.': 'CASNo', '英文名稱': 'ChemicalEngName', '中文名稱': 'ChemicalChnName'})
    file = os.path.join(processed_path, 'D_mol_03_c.xlsx')
    df.to_excel(file, index = False)

def D_mol_03_d():
    file = os.path.join(input_path, 'D_mol_03_d.xlsx')
    df = pd.read_excel(file)    
    df.columns = [i.strip() for i in df.columns]
    df = df[['CAS No.', '英文名稱', '中文名稱']]
    df = df.rename(columns = {'CAS No.': 'CASNo', '英文名稱': 'ChemicalEngName', '中文名稱': 'ChemicalChnName'})
    df = df.dropna(how='all')
    file = os.path.join(processed_path, 'D_mol_03_d.xlsx')
    df.to_excel(file, index = False)




#endregion

#region D_moea_01 工業局選定化學物質
def D_moea_01():
    file = os.path.join(input_path, 'D_moea_01.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {
        s0.columns[1]: 'ChemicalChnName',
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
    s0 = s0.iloc[2:].fillna(method='ffill')

    file = os.path.join(processed_path, 'D_moea_01.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region D_moea_02 先驅化學品工業原料(毒品先驅物)
def D_moea_02():
    file = os.path.join(input_path, 'D_moea_02.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {
        s0.columns[0]: 'ChemicalChnName',
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

    file = os.path.join(input_path, 'D_moea_02_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s1.columns), '\n', sep = '\n')

    Columns = {
        s1.columns[0]: 'ChemicalChnName',
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
    df = df.loc[~df.CASNo.str.contains('CAS')]

    file = os.path.join(processed_path, 'D_moea_02.xlsx')
    df.to_excel(file, index = False)


#endregion

#region D_moea_03 事業用爆炸物品名
def D_moea_03():
    file = os.path.join(input_path, 'D_moea_03.xlsx')
    df = pd.read_excel(file)
    df.columns = [i.strip() for i in df.columns]
    df = df.rename({'中文通用名稱 Chinese common name': 'ChemicalChnName', '英文名稱 English name': 'ChemicalEngName', 'CAS No.': 'CASNo'}, axis = 1)
    df = df.applymap(lambda i: i.strip() if isinstance(i, str) else i)
    file = os.path.join(processed_path, 'D_moea_03.xlsx')    
    df.to_excel(file, index = False)


#endregion

#region D_moa_01 成品農藥摻雜其他有效成分之限量基準
def D_moa_01():
    file = os.path.join(input_path, 'D_moa_01.xlsx')
    data = pd.read_excel(file, header=[0,1]).iloc[:,[1]]
    data.columns = ['ChemicalChnName']
    file = os.path.join(processed_path, 'D_moa_01.xlsx')
    data.to_excel(file, index = False)
#endregion

#region D_moa_02 農藥有害不純物之限量規格
def D_moa_02():
    file = os.path.join(input_path, 'D_moa_02.PDF')
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

    file = os.path.join(processed_path, 'D_moa_02.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moa_03 農藥其他成分之限量規格
def D_moa_03():
    file = os.path.join(input_path, 'D_moa_03.pdf')
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

    file = os.path.join(processed_path, 'D_moa_03.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moa_04 飼料添加物使用準則
def D_moa_04():
    print('比較大小確認更新狀態，目前尚未更新')
#endregion

#region D_moa_05 餌劑成品農藥摻雜其他有效成分之限量基準
def D_moa_05():
    file = os.path.join(input_path, 'D_moa_05.xlsx')
    df = pd.read_excel(file)
    df = df['摻雜之農藥有效成分']
    df = df.apply(lambda i: i.split('、')).explode().drop_duplicates()
    df = df.to_frame().rename({'摻雜之農藥有效成分': 'ChemicalChnName'}, axis = 1)
    file = os.path.join(processed_path, 'D_moa_05.xlsx')
    df.to_excel(file, index = False)
#endregion

#region D_mohw_01 化粧品中抗菌劑成分使用及限量規定清單
def D_mohw_01():
    file = os.path.join(input_path, 'D_mohw_01.xlsx')
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

    file = os.path.join(processed_path, 'D_mohw_01.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region D_mohw_02 化粧品中防腐劑成分使用及限量規定清單ㄋ
def D_mohw_02():
    file = os.path.join(input_path, 'D_mohw_02.xlsx')
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

    file = os.path.join(processed_path, 'D_mohw_02.xlsx')
    s0.to_excel(file, index = False)

#endregion

#region D_mohw_03 化粧品中禁止使用成分
def D_mohw_03():
    file = os.path.join(input_path, 'D_mohw_03.xlsx')
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

    file = os.path.join(processed_path, 'D_mohw_03.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region D_mohw_04 化粧品成分使用限制
def D_mohw_04():
    file1 = os.path.join(input_path, 'D_mohw_04.xlsx')
    df1 = pd.read_excel(file1)
    df1 = df1[['成分名', 'INCI名', 'CAS NO.']]
    file = os.path.join(processed_path, 'D_mohw_04.xlsx')
    df1.to_excel(file, index = False)
    

#endregion

#region D_mohw_05 化粧品色素成分使用限制
def D_mohw_05():
    file2 = os.path.join(input_path, 'D_mohw_05.xlsx')
    df2 = pd.read_excel(file2)
    df2 = df2[['Color Index Number/成分名', '別名']]
    # df2.to_excel('~/Desktop/tmo.xlsx', index = False)
    # 未完成

#endregion

#region D_mohw_06 食安優先加強勾稽名單
def D_mohw_06():
    print('沿用系統的;')
#endregion

#region D_mohw_07 食品添加物
def D_mohw_07():
    file = os.path.join(input_path, 'D_mohw_07.xlsx')
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
    file = os.path.join(processed_path, 'D_mohw_07.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_mohw_08 特定用途化粧品成分名稱及使用限制
def D_mohw_08():
    file3 = os.path.join(input_path, 'D_mohw_08.xlsx')
    df3 = pd.read_excel(file3)
    df3 = df3[['成分名', 'INCI名', 'CAS NO.']]
    df3 = df3.rename({'成分名': 'ChemicalChnName', 'INCI名': 'ChemicalEngName', 'CAS NO.': 'CASNo'}, axis = 1)
    file = os.path.join(processed_path, 'D_mohw_08.xlsx')
    df3.to_excel(file, index = False)
#endregion

#region D_mohw_09 第四級管制藥品原料藥（毒品先驅物）
def D_mohw_09():
    file = os.path.join(input_path, 'D_mohw_09.pdf')
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

    file = os.path.join(processed_path, 'D_mohw_09.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region D_moenv_01 列管毒性化學物質清單
def D_moenv_01():
    file = os.path.join(input_path, 'D_moenv_01.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    

    Columns = {
        # s0.columns[0]: 'no',      
        # s0.columns[1]: 'no2',      
        s0.columns[2]: 'ChemicalChnName',      
        s0.columns[3]: 'ChemicalEngName',
        s0.columns[5]: 'CASNo',
        }

    s0 = s0.rename(columns = Columns)
    s0 = s0.loc[~s0.iloc[:,0].str.contains('列管').fillna(False)]
    
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


    file = os.path.join(processed_path, 'D_moenv_01.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moenv_02 我國環境荷爾蒙建議關注清單
def D_moenv_02():
    print('以下code 2023年不採用，final檔由全哥提供。沒有input檔')
    
    if False:
        file = os.path.join(input_path, 'D_moenv_02.xlsx')
        df = pd.read_excel(file)
        
        # pdf = pdfplumber.open(file)
        s0 = tabula.read_pdf(file, pages='121-145', lattice = True)
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

        file = os.path.join(processed_path, 'D_moenv_02.xlsx')
        s0.to_excel(file, index = False)


#endregion

#region D_moenv_03 飲用水水質處理藥劑
def D_moenv_03():
    file = os.path.join(input_path, 'D_moenv_03.pdf')
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
    s0 = s0[['ChemicalChnName', 'ChemicalEngName']]
    
    s0 = s0.dropna()
    s0 = s0[s0.ChemicalEngName != '英文名稱']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.apply(lambda r : re.sub('（.*）','',r))
    s0.ChemicalChnName = s0.ChemicalChnName.apply(lambda r : re.sub('（.*\)','',r))
    
    s0 = s0[Columns.values()]

    file = os.path.join(processed_path, 'D_moenv_03.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moenv_04 應徵收土壤及地下水污染整治費之物質徵收清單
def D_moenv_D_moenv_04():
    file = os.path.join(input_path, 'D_moenv_04.PDF')
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

    file = os.path.join(processed_path, 'D_moenv_04.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moenv_05 環境用藥有效成分
def D_moenv_05():
    print(
        '''
            化學雲資料表。
            SELECT distinct [CASNo] ,[ChemicalChnName] ,[ChemicalEngName] FROM [ChemiTemp].[dbo].[TEnvironmentMdc]; 手動+程式
        ''')


#endregion

#region D_moenv_06 環境用藥禁止含有之成分
def D_moenv_06():
    file = os.path.join(input_path, 'D_moenv_06.pdf')
    pdf = pdfplumber.open(file)
    table = []
    table_settings = {"join_tolerance": 100}
    for i in range(0, 5):
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
    # s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' ', '', regex = False)
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
    s0 = s0.drop(s0[s0.ChemicalEngName == ''].index, axis=0)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('(cis:trans=', '', regex = False)

    file = os.path.join(processed_path, 'D_moenv_06.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region D_moenv_07 國內歷年食安事件
def D_moenv_07():
    print('系統直接抓資料表 FdSafetyEventChemi;')
#endregion

#region A_china_01 危險化學品目錄
def A_china_01():
    file = os.path.join(input_path, 'A_china_01.xlsx')
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
    s1 = s0.fillna('-').replace('-', None)
    s1 = s1.dropna(axis=0, how='all')
    
    file = os.path.join(processed_path, 'A_china_01.xlsx')
    s1.to_excel(file, index = False)
#endregion

#region A_china_02 易制爆危險化學品名錄（IED先驅物）
def A_china_02():
    file = os.path.join(input_path, 'A_china_02.xlsx')
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
    file = os.path.join(processed_path, 'A_china_02.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_china_03 易製毒化學品管理條例第一類（毒品先驅物）
def A_china_03():
    file = os.path.join(input_path, 'A_china_03.xlsx')
    data = pd.read_excel(file, header=None)
    data[1] = data[1].apply(lambda i: i.split('．')[1].replace('*','').strip())
    data = data.rename(columns = {0:'ChemicalChnName', 1:'CASNo'})
    
    series = 'abcde'
    for i, kind in enumerate(data.ChemicalChnName.unique()):
        tmp = data[data.ChemicalChnName == kind]
        print(tmp)
        file = os.path.join(processed_path, f'A_china_03_{series[i]}.xlsx')
        print(file)
        tmp.to_excel(file, index = False)
#endregion


#region A_china_04_a 重點監管危險化學品名錄（首批）
def A_china_04_a():
    # 可能要注意爬蟲，在公司好像無法爬
    url = 'https://www.shsmu.edu.cn/zicc/info/1044/2287.htm'
    data = pd.read_html(url)[0]
    data = data.iloc[1:, [2,4]]
    data = data.rename(columns = {2: 'ChemicalChnName', 4: 'CASNo'})
    
    file = os.path.join(processed_path, 'A_china_04_a.xlsx')
    data.to_excel(file, index = False)

#endregion

#region A_china_04_b 重點監管危險化學品名錄（第二批）
def A_china_04_b():
    file = os.path.join(input_path, 'A_china_04_b.xlsx')
    data = pd.read_excel(file).loc[:,['化学品品名', 'CAS号']].dropna()
    data = data.rename(columns = {'化学品品名': 'ChemicalChnName', 'CAS号': 'CASNo'})
    file = os.path.join(processed_path, 'A_china_04_b.xlsx')
    data.to_excel(file, index = False)
#endregion

#region A_china_05_a 優先控制化學品名錄（第一批）
def A_china_05_a():
    # 直接用爬蟲
    url = 'https://www.mee.gov.cn/gkml/hbb/bgg/201712/t20171229_428832.htm'
    data = pd.read_html(url)[2]
    data = data.iloc[1:, 1:]
    data.loc[:,2] = data.loc[:,2].str.split()
    data = data.explode(2)
    data = data.rename(columns = {1:'ChemicalChnName', 2:'CASNo'})
    data = data.fillna('')
    
    
    file = os.path.join(processed_path, 'A_china_05_a.xlsx')
    data.to_excel(file, index = False)


#endregion

#region A_china_05_b 優先控制化學品名錄（第二批）
def A_china_05_b():
    file = os.path.join(input_path, 'A_china_05_b.pdf')
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

    file = os.path.join(processed_path, 'A_china_05_b.xlsx')
    s0.to_excel(file, index = False)



#endregion

#region A_china_06 嚴格限制的有毒化學品名錄
def A_china_06():
    file = os.path.join(input_path, 'A_china_06.pdf')
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
    file = os.path.join(processed_path, 'A_china_06.xlsx')
    df.to_excel(file, index = False)


#endregion

#region A_japan_01 毒物
def A_japan_01():
    '''
    直接爬蟲
    '''
    urls = [
        'http://www.nihs.go.jp/law/dokugeki/toku.files/sheet001.html',
        'http://www.nihs.go.jp/law/dokugeki/geki.files/sheet001.html',
        'http://www.nihs.go.jp/law/dokugeki/doku.files/sheet001.html'
        ]
    n = 1
    for ii, url in enumerate(urls[:]):
        res = requests.get(url)
        res.encoding = 'Shift_JIS'
        df = pd.read_html(res.text, encoding='Shift_JIS')[0].iloc[9:-6, [3,5]]
        df = df.rename(columns = {3: 'ChemicalEngName', 5: 'CASNo'}).dropna(how='all')

        file = os.path.join(processed_path, f'A_japan_{n+ii:02d}.xlsx')
        print(file)
        df.to_excel(file, index = False)        
#endregion

#region A_japan_04_a 第１種指定(PRTR)化學物質清單
def A_japan_04_a():
    file = os.path.join(input_path, 'A_japan_04_a.xlsx')
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
    file = os.path.join(processed_path, 'A_japan_04_a.xlsx')
    s.to_excel(file, index = False)


#endregion

#region A_japan_04_b 第2種指定化學物質清單
def A_japan_04_b():
    file = os.path.join(input_path, 'A_japan_04_b.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 2)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join(processed_path, 'A_japan_04_b.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_japan_05 優先評估化學物質(PACSs)清單
def A_japan_05():
    urls = [ 
            'https://www.nite.go.jp/chem/jcheck/list7.action?category=230&request_locale=en',
            'https://www.nite.go.jp/chem/jcheck/list6.action?category=220&request_locale=en',
            'https://www.nite.go.jp/chem/jcheck/list6.action?category=211&request_locale=en',
            'https://www.nite.go.jp/chem/jcheck/list6.action?category=212&request_locale=en'
        ]

    series = 'abcde'
    for ii, url in enumerate(urls[:]):
        res = pd.read_html(url)[3]
        try:
            res = res.iloc[1:-1, [2]]
        except:
            res = res.iloc[1:-1, [1]]
        res = res.rename({2: 'ChemicalEngName'}, axis = 1)
        file = os.path.join(processed_path, f'A_japan_05_{series[ii]}.xlsx')
        print(file)
        res.to_excel(file, index = False)        

#endregion

#region A_japan_06 麻藥及影響精神藥原料（毒品先驅物）
def A_japan_06():
    print('全哥提供檔案;')


#endregion

#region A_japan_07 疑似環境荷爾蒙清單
def A_japan_07():
    print('檔案不好處理。但比較之後完全一樣不處理')
#endregion

#region A_japan_08 覺醒劑原料（毒品先驅物）
def A_japan_08():
    print(' 全哥給檔案; 人工')
#endregion

#region A_wco_01 STCE指南附錄5
def A_wco_01():
    file = os.path.join(input_path, 'A_wco_01.pdf')
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
    Columns = {
        s0.columns[4]: 'ChemicalEngName',
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
    file = os.path.join(processed_path, 'A_wco_01.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region A_canada_01_a 先驅化學品A類（毒品先驅物）
def A_canada_01():
    url = 'https://www.canada.ca/en/health-canada/services/health-concerns/controlled-substances-precursor-chemicals/precursor-chemicals/regulatory-requirements-under-controlled-drugs-substances-act.html'
    res = requests.get(url, headers=headers)
    soup = bs(res.text)
    
    classa = soup.find('h2', text='CLASS A PRECURSORS')
    classa_next = classa.parent.next_sibling
    while classa_next and not classa_next.name:
        classa_next = classa_next.next_sibling
    classa_li = [i.text for i in classa_next.select('li')]
    df1 = pd.DataFrame(classa_li).rename(columns = {0: 'ChemicalEngName'})
    
    file = os.path.join(processed_path, 'A_canada_01_a.xlsx')
    df1.to_excel(file, index = False)    
    
    classb = soup.find('h2', text='CLASS B PRECURSORS')
    classb_next = classb.parent.next_sibling
    while classb_next and not classb_next.name:
        classb_next = classb_next.next_sibling
    classb_li = [i.text for i in classb_next.select('li')]
    df2 = pd.DataFrame(classb_li).rename(columns = {0: 'ChemicalEngName'})
    file = os.path.join(processed_path, 'A_canada_01_b.xlsx')
    df2.to_excel(file, index = False)        


#endregion

#region A_canada_02 國家污染釋放清冊NPRI清單
def A_canada_02():
    url = 'https://www.canada.ca/en/environment-climate-change/services/national-pollutant-release-inventory/substances-list/threshold.html'
    res = requests.get(url=url, headers=headers)
    df = pd.read_html(res.text)[0]
    df = df[['Substance name', 'CAS RN or other substance identifier']]
    df = df.rename(columns = {'Substance name': 'ChemicalEngName','CAS RN or other substance identifier': 'CASNo'})
    df.CASNo.loc[df.CASNo.str.contains('NA -')] = '-'

    file = os.path.join(processed_path, 'A_canada_02.xlsx')
    df.to_excel(file, index = False)


#endregion

#region A_canada_03 環境保護法schedule 1之毒化物清單
    #; 這個非常麻煩，最後處理 
#endregion

#region A_canada_04 環境保護法優先評估物質
def A_canada_04():
    url1 = 'https://www.ec.gc.ca/ese-ees/default.asp?lang=En&n=95D719C5-1'
    text = requests.get(url1).text
    soup = bs(text)
    ChemicalEngName_1 = [i.text for i in soup.select('#wb-cont ul li')]
    
    url2 = 'https://www.ec.gc.ca/ese-ees/default.asp?lang=En&n=C04CA116-1'
    text = requests.get(url2).text
    soup = bs(text)
    ChemicalEngName_2 = [i.text for i in soup.select('#wb-cont ul li')]
    
    df = pd.DataFrame(ChemicalEngName_1 + ChemicalEngName_2).rename(columns = {0: 'ChemicalEngName'})
    
    
    file = os.path.join(processed_path, 'A_canada_04.xlsx')
    df.to_excel(file, index = False)
    


#endregion

#region A_usa_01_a 先驅化學品I類（毒品先驅物）
    #; 沒找到 
#endregion

#region A_usa_01_b 先驅化學品II類（毒品先驅物）
    #; 沒找到 
#endregion

#region A_usa_02 有毒物質TRI排放清單
    #; 這個非常麻煩，最後處理 
#endregion

#region A_usa_03 國土安全部化學品設施反恐標準（IED先驅物）
def A_usa_03():
    file = os.path.join(input_path, 'A_usa_03.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {
        s0.columns[0]: 'ChemicalEngName',
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

    file = os.path.join(processed_path, 'A_usa_03.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_usa_04_a 第一階段環境荷爾蒙第一批次物質最終篩選清單
def A_usa_04_a():
    ua = UserAgent()    
    headers={ "User-Agent": ua.random }       
    res = requests.get('https://www.epa.gov/endocrine-disruption/endocrine-disruptor-screening-program-tier-1-screening-determinations-and',headers=headers)
    data = pd.read_html(res.text)[0][['Chemical Name', 'CAS Number']]

    file = os.path.join(processed_path, 'A_usa_04_a.xlsx')
    data.to_excel(file, index = False)
#endregion

#region A_usa_04_b 第一階段環境荷爾蒙第二批次篩選清單
def A_usa_04_b():
    file = os.path.join(input_path, 'A_usa_04_b.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[2:], columns = table[1])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {
        s0.columns[1]: 'ChemicalEngName',
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

    file = os.path.join(processed_path, 'A_usa_04_b.xlsx')
    s0.to_excel(file, index = False)
#endregion

#region A_england_01_a 英國先驅化學品第1類（毒品先驅物）
def A_england_01_a():
    file = os.path.join(input_path, 'A_england_01.xlsx')
    data = pd.read_excel(file)
    data = data[['Substance', 'Category'] ].rename(columns={'Substance': 'ChemicalEngName'}).assign(CASNo = '')
    series = 'abcde'
    for ii, item in enumerate(data.Category.unique()):
        tmp = data[data.Category == item].loc[:, ['ChemicalEngName', 'CASNo']]
        
        file = os.path.join(processed_path, f'A_england_01_{series[ii]}.xlsx')
        tmp.to_excel(file, index=False)


#endregion

#region A_holland_01 鹿特丹PIC
def A_holland_01():
    url = 'https://informea.pops.int/Chemicals/chemicals.svc/ChemicalsAnnexIII?$callback=jQuery1124014021869977094492_1695185171271&%24inlinecount=allpages&%24format=json&%24orderby=listId%2CpicNameEnglish'
    res = requests.get(url, headers=headers).text
    data = ''.join(re.findall('^jQuery[0-9_]+\((.*)\)', res))
    data = json.loads(data)
    df = pd.DataFrame(data['value'])
    df = df[['picNameEnglish', 'cas']]
    def handldCas(i):
        i2 = i.split(',')
        return [f'{k[:-3]}-{k[-3:-1]}-{k[-1]}' for k in i2]
    
    df.cas = df.cas.apply(handldCas)
    df = df.explode('cas')
    df = df.rename(columns = {'picNameEnglish': 'ChemicalEngName', 'cas': 'CASNo'})
    
    file = os.path.join(processed_path, 'A_holland_01.xlsx')
    df.to_excel(file, index = False)


#endregion

#region A_singapore_01 武器與爆裂物法（IED先驅物）
def A_singapore_01():
    url = 'https://sso.agc.gov.sg/Act/AEA1913?ProvIds=Sc2-'
    headers={ "User-Agent": ua.random }
    res = requests.get(url, headers=headers)
    soup = bs(res.text)
    data = [i.text.split(',')[0].replace('.','') for i in soup.select('#legisContent span.pIndentTxt')]
    df = pd.DataFrame(data, columns = ['ChemicalEngName'])
    
    
    
    file = os.path.join(processed_path, 'A_singapore_01.xlsx')
    df.to_excel(file, index = False)


#endregion

#region A_eu_01 CoRAP化學物質評估清單
def A_eu_01():
    file = os.path.join(input_path, 'A_eu_01.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {
        s0.columns[0]: 'ChemicalEngName',
        s0.columns[3]: 'CASNo'
    }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join(processed_path, 'A_eu_01.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_eu_02 REACH附錄17限制清單
def A_eu_02():
    file = os.path.join(input_path, 'A_eu_02.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {
        s0.columns[0]: 'ChemicalEngName',
        s0.columns[3]: 'CASNo'
    }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.ChemicalEngName.str.contains('Entry') != True]
    s0 = s0[Columns.values()]

    s0.ChemicalEngName = s0.ChemicalEngName.str.lower()
    s0 = s0[~s0.duplicated()]
    print(s0)
    file = os.path.join(processed_path, 'A_eu_02.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_eu_03 高度關切物質(SvHC)清單
def A_eu_03():#刪除重複的動作 可以到後面merged 再做
    file = os.path.join(input_path, 'A_eu_03.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {
        s0.columns[0]: 'ChemicalEngName',
        s0.columns[3]: 'CASNo'
    }
    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('- *,', '', regex = True)
    s0.CASNo = s0.CASNo.str.replace(', *', '; ', regex = True)
    s0 = s0[~s0.ChemicalEngName.str.contains('available')]
    s0 = s0[Columns.values()]

    file = os.path.join(processed_path, 'A_eu_03.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_eu_04 歐盟PIC
def A_eu_04():
    file = os.path.join(input_path, 'A_eu_04.xlsx')
    data = pd.read_excel(file)
    
    data = data[['substance-name', 'cas-number']].rename(columns = {'substance-name': 'ChemicalEngName', 'cas-number': 'CASNo'})
    data = data.dropna()

    file = os.path.join(processed_path, 'A_eu_04.xlsx')
    data.to_excel(file, index = False)


#endregion

#region A_eu_05 歐盟先驅化學品第1類（毒品先驅物）
def A_eu_05():
    url = 'https://eur-lex.europa.eu/legal-content/EN/TXT/HTML/?uri=CELEX:02004R0273-20210113'
    
    res = pd.read_html(url)
    res_3 = res[3]
    res_5 = res[5]
    res_6 = res[6]
    
    res_3 = res_3.iloc[1:-1, [0, 3]]
    res_3 = res_3[res_3[3].str.contains('.*-.*-.*')]
    
    res_5 = res_5.iloc[1:-1, [0, 3]]
    res_5 = res_5[res_5[3].str.contains('.*-.*-.*')]
    
    res_6 = res_6.iloc[1:-1, [0, 3]]
    res_6 = res_6[res_6[3].str.contains('.*-.*-.*')]
    
    df = pd.concat([res_3, res_5, res_6]).rename(columns={0: 'ChemicalEngName', 3: 'CASNo'})


    file = os.path.join(processed_path, 'A_eu_05_a.xlsx')
    res_3.to_excel(file, index = False)
    file = os.path.join(processed_path, 'A_eu_05_b.xlsx')
    res_5.to_excel(file, index = False)
    file = os.path.join(processed_path, 'A_eu_05_c.xlsx')
    res_6.to_excel(file, index = False)
    

#endregion

#region A_eu_06 歐盟法規第98/2013號-爆裂物先驅物
def A_eu_06():
    url = 'https://eur-lex.europa.eu/legal-content/EN/TXT/HTML/?uri=CELEX:02019R1148-20190711'
    data = pd.read_html(url)
    data_2 = data[2]
    data_2 = data_2.drop(index = 0).iloc[:-1, 0]
    data_2 = data_2.apply(lambda x: x.split('(')).apply(pd.Series).rename(columns = {0: 'ChemicalEngName', 1: 'CASNo'})
    data_2.CASNo = data_2.CASNo.apply(lambda i: i.replace(')', '').replace('CAS RN ', '').strip())
    
    data_3 = (
        data[3].iloc[2:-1, 0]
        .apply(lambda x: x.split('(')).apply(pd.Series)
        .rename(columns = {0: 'ChemicalEngName', 1: 'CASNo'})
        .drop([2, 3], axis = 1)
    )
    data_3.CASNo = data_3.CASNo.apply(lambda i: i.replace(')', '').replace('CAS RN ', '').strip())
    
    df = pd.concat([data_2, data_3], ignore_index = True)

    file = os.path.join(processed_path, 'A_eu_06.xlsx')
    df.to_excel(file, index = False)



#endregion

#region A_eu_07 歐盟環境荷爾蒙物質清單
def A_eu_07():
    file = os.path.join(input_path, 'A_eu_07.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {
        s0.columns[0]: 'ChemicalEngName',
        s0.columns[3]: 'CASNo'
    }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join(processed_path, 'A_eu_07.xlsx')
    s0.to_excel(file, index = False)


#endregion

#region A_au_01 安全關注化學品
def A_au_01():
    url = 'https://www.nationalsecurity.gov.au/_layouts/15/api/Data.aspx/GetListData'
    payload = {
        "webUrl": "/",
        "listName": "SecurityConcernChemicals"
    }
    res = requests.post(url, json = payload).json()
    data = pd.DataFrame(res['d']['data'])
    df = data[['Title', 'SCCCAS']].rename(columns = {'Title': 'ChemicalEngName', 'SCCCAS': 'CASNo'})

    file = os.path.join(processed_path, 'A_au_01.xlsx')
    df.to_excel(file, index = False)


#endregion

#region A_un_01_a 國際麻醉品管制局I類（毒品先驅物）
    #; 程式無法用，手動處理 
#endregion

#region A_un_01_b 國際麻醉品管制局II類（毒品先驅物）
    #; 程式無法用，手動處理 
#endregion

#region A_korea_01 K-REACH CMR物質清單
def A_korea_01():
    print('以下 code 暫時不使用，直接比較')
    if False:
        file = os.path.join(input_path, 'A_korea_01.pdf')
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
        file = os.path.join(processed_path, 'A_korea_01.xlsx')
        s0.to_excel(file, index = False)


#endregion

#region A_korea_02 K-REACH第一批優先評估既有化學物質(PEC)清單
def A_korea_02():
    print('以下 code 暫時不使用，直接比較')
    if False:
        # step1,get hwp file from web: https://www.law.go.kr/LSW/admRulLsInfoP.do?admRulSeq=2100000021862#AJAX
        # step2,use https://appzend.herokuapp.com/hwpviewer/ get data
        # step3,edit xlsx by the following code
            
        file = os.path.join(input_path, 'A_korea_02.xlsx')
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


        file = os.path.join(processed_path, 'A_korea_02.xlsx')
        s0.to_excel(file, index = False)



#endregion

#region A_korea_03 優先管理化學物質表1
def A_korea_03():
    print('以下 code 暫時不使用，直接比較')
    if False:
        file = os.path.join(input_path, 'A_korea_03.pdf')
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
        file = os.path.join(processed_path, 'A_korea_03.xlsx')
        s0.to_excel(file, index = False)


#endregion

#region A_korea_04 優先管理化學物質表2
def A_korea_04():
    print('以下 code 暫時不使用，直接比較')
    if False:
        file = os.path.join(input_path, 'A_korea_04.pdf')
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
        file = os.path.join(processed_path, 'A_korea_04.xlsx')
        s0.to_excel(file, index = False)


#endregion
