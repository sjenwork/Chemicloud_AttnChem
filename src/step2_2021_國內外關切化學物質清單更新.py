import os
import pandas as pd
import pdfplumber

def n02():
    file = os.path.join('input', '02.csv')
    s0 = pd.read_csv(file, dtype = 'string')
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[2]: 'ChemicalChnName',
               s0.columns[3]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('、', '; ')
    s0.CASNo = s0.CASNo.str.replace('－', '-')
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '02.xlsx')
    s0.to_excel(file, index = False)

def n03():
    file = os.path.join('input', '03.pdf')
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

    file = os.path.join('processed', '03.xlsx')
    s0.to_excel(file, index = False)

def n04():
    file = os.path.join('input', '04.pdf')
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

    file = os.path.join('processed', '04.xlsx')
    s0.to_excel(file, index = False)

def n05():
    file = os.path.join('input', '05.pdf')
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

    file = os.path.join('processed', '05.xlsx')
    s0.to_excel(file, index = False)

def n08():
    file = os.path.join('input', '08.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[0]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo != 'CAS No.']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('(cid:744)', 'II', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('input', '08_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:])
    print(*enumerate(s1.columns), '\n', sep = '\n')

    Columns = {s1.columns[3]: 'ChemicalEngName',
               s1.columns[0]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1[1] != 'CAS No.']
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('\n', '', regex = False)
    a = s1[s1.CASNo.isna()].index
    for i in a:
        s1.loc[i-1, 'ChemicalEngName'] = s1.loc[i-1, 'ChemicalEngName'] + s1.loc[i, 'ChemicalEngName']
        s1 = s1.drop(index = i)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('(cid:744)', 'II', regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace(' (cid:528)', '; ', regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('(cid:529)', '', regex = False)
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('processed', '08.xlsx')
    df.to_excel(file, index = False)

def n09():
    file = os.path.join('input', '09.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:])

    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[6]: 'ChemicalChnName',
               s0.columns[3]: 'ChemicalEngName',
               s0.columns[0]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo != '']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('input', '09_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:])
    print(*enumerate(s1.columns), '\n', sep = '\n')

    Columns = {s1.columns[6]: 'ChemicalChnName',
               s1.columns[3]: 'ChemicalEngName',
               s1.columns[0]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1[1] != 'CAS No.']
    s1 = s1.reset_index()
    a = s1[s1.CASNo == ''].index
    for i in a:
        s1.loc[i-1, 'ChemicalEngName'] = s1.loc[i-1, 'ChemicalEngName'] + s1.loc[i, 'ChemicalEngName']
        s1 = s1.drop(index = i)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('\n', '', regex = False)
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace('\n', '', regex = False)
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('processed', '09.xlsx')
    df.to_excel(file, index = False)

def n10():
    file = os.path.join('input', '10.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[0]: 'ChemicalChnName',
               s0.columns[1]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[s0.CASNo != 'CAS.NO.']
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\n', '', regex = False)
    s0.loc[s0.ChemicalEngName.isna(), 'ChemicalEngName'] = '-'
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('input', '10_2.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s1 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s1.columns), '\n', sep = '\n')

    Columns = {s1.columns[1]: 'ChemicalChnName',
               s1.columns[0]: 'ChemicalEngName',
               s1.columns[2]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1.ChemicalChnName = s1.ChemicalChnName.str.replace('\n', '', regex = False)
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('\n', '', regex = False)
    s1 = s1[Columns.values()]

    file = os.path.join('input', '10_3.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(len(pdf.pages)):
        page = pdf.pages[i]
        table += page.extract_table()
    s2 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s2.columns), '\n', sep = '\n')

    Columns = {s2.columns[1]: 'ChemicalChnName',
               s2.columns[0]: 'ChemicalEngName',
               s2.columns[2]: 'CASNo'
              }
    s2 = s2.rename(columns = Columns)
    a = s2[s2.CASNo == ''].index
    for i in a:
        s2.loc[i-1, 'ChemicalChnName'] = s2.loc[i-1, 'ChemicalChnName'] + s2.loc[i, 'ChemicalChnName']
        s2.loc[i-1, 'ChemicalEngName'] = s2.loc[i-1, 'ChemicalEngName'] + s2.loc[i, 'ChemicalEngName']
        s2 = s2.drop(index = i)
    s2 = s2[s2.CASNo != 'CAS No.']
    s2.ChemicalChnName = s2.ChemicalChnName.str.replace('\n', '', regex = False)
    s2.ChemicalEngName = s2.ChemicalEngName.str.replace('\n', '', regex = False)
    s2 = s2[Columns.values()]

    df = pd.concat([s0, s1, s2], ignore_index = True)
    file = os.path.join('processed', '10.xlsx')
    df.to_excel(file, index = False)

def n11():
    file = os.path.join('input', '11.pdf')
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

    file = os.path.join('input', '11_2.pdf')
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
    file = os.path.join('processed', '11.xlsx')
    df.to_excel(file, index = False)

def n12():
    file = os.path.join('processed', '12_1.xlsx')
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

    file = os.path.join('processed', '12.xlsx')
    s0.to_excel(file, index = False)

def n13():
    file = os.path.join('input', '13.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(23, 25):
        page = pdf.pages[i]
        table += page.extract_table()
    s0 = pd.DataFrame(table[1:], columns = table[0])
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    a = s0[s0.ChemicalChnName == ''].index
    for i in a:
        s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + s0.loc[i, 'ChemicalEngName']
        s0 = s0.drop(index = i)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' （', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace(' ）', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\n', '', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('、', '; ', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' ;', ';', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '13.xlsx')
    s0.to_excel(file, index = False)

def n14():
    file = os.path.join('input', '14.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'ChemicalEngName'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '14.xlsx')
    s0.to_excel(file, index = False)

def n15():
    file = os.path.join('input', '15.xlsx')
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

    file = os.path.join('processed', '15.xlsx')
    s0.to_excel(file, index = False)

def n16():
    file = os.path.join('input', '16.xlsx')
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
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '16.xlsx')
    s0.to_excel(file, index = False)

def n17():
    file = os.path.join('input', '17.xlsx')
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

    file = os.path.join('processed', '17.xlsx')
    s0.to_excel(file, index = False)

def n19():
    file = os.path.join('input', '19.pdf')
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

    file = os.path.join('processed', '19.xlsx')
    s0.to_excel(file, index = False)

def n20():
    file = os.path.join('input', '20.pdf')
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

    file = os.path.join('processed', '20.xlsx')
    s0.to_excel(file, index = False)

def n21():
    file = os.path.join('input', '21.pdf')
    pdf = pdfplumber.open(file)
    table = []
    for i in range(1, 9):
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

    file = os.path.join('processed', '21.xlsx')
    s0.to_excel(file, index = False)

def n27():
    file = os.path.join('input', '27.xlsx')
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

    file = os.path.join('processed', '27.xlsx')
    s0.to_excel(file, index = False)

def n28():
    file = os.path.join('input', '28.xlsx')
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

    file = os.path.join('processed', '28.xlsx')
    s0.to_excel(file, index = False)

def n29():
    file = os.path.join('input', '29.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '29.xlsx')
    s0.to_excel(file, index = False)

def n30():
    file = os.path.join('input', '30.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 3)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '30.xlsx')
    s0.to_excel(file, index = False)

def n31():
    file = os.path.join('input', '31.pdf')
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

    file = os.path.join('processed', '31.xlsx')
    s0.to_excel(file, index = False)

def n32():
    file = os.path.join('input', '32.pdf')
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

    file = os.path.join('processed', '32.xlsx')
    s0.to_excel(file, index = False)

def n33():
    file = os.path.join('processed', '33_0.xlsx')
    s0 = pd.read_excel(file, header = None)

    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' +', ' ', regex = True)
    s0.loc[35, 'ChemicalEngName'] = 'Alkyl phenol (from C5 to C9); Nonyl phenol; Octyl phenol'
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(', '; ', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(', ', ' and ', regex = False)
    a = s0[s0.ChemicalEngName.str.contains(')', regex = False)].index
    for i in a:
        s0.loc[i, 'ChemicalEngName'] = s0.loc[i, 'ChemicalEngName'][:-1]
    file = os.path.join('processed', '33.xlsx')
    s0.to_excel(file, index = False)

def n34():
    file = os.path.join('input', '34.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalEngName',
               s0.columns[0]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')
    Columns = {s1.columns[1]: 'ChemicalEngName',
               s1.columns[0]: 'CASNo'
              }
    s1 = s1.rename(columns = Columns)
    s1 = s1[Columns.values()]

    df = pd.concat([s0, s1], ignore_index = True)
    df.CASNo = df.CASNo.str.replace('*', '', regex = False)
    df.CASNo = df.CASNo.str.replace('N.+', '-', regex = True)
    file = os.path.join('processed', '34.xlsx')
    df.to_excel(file, index = False)

def n36():
    file = os.path.join('input', '36.xlsx')
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

    file = os.path.join('processed', '36_1.xlsx')
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

    file = os.path.join('processed', '36.xlsx')
    s0.to_excel(file, index = False)

def n37():
    file = os.path.join('processed', '37_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('\u3000', '-', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0.Synonym
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('; \u3000', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '37.xlsx')
    s0.to_excel(file, index = False)

def n38():
    file = os.path.join('processed', '38_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0[~((s0.ChemicalEngName == '\u3000') & (s0.CASNo == '\u3000'))]
    s0.CASNo = s0.CASNo.str.replace('\u3000', '-', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0.Synonym
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\u3000', '-', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('; -', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '38.xlsx')
    s0.to_excel(file, index = False)

def n39():
    file = os.path.join('processed', '39_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[3]: 'ChemicalEngName',
               s0.columns[5]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo = s0.CASNo.str.replace('\u3000', '-', regex = False)
    s0.ChemicalEngName = s0.ChemicalEngName + '; ' + s0.Synonym
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('; \u3000', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '39.xlsx')
    s0.to_excel(file, index = False)

def n40_43():
    file = os.path.join('input', '40_43.csv')
    x = pd.read_csv(file, dtype = 'string')
    print(*enumerate(x.columns), '\n', sep = '\n')

    Columns = {x.columns[2]: 'ChemicalEngName',
               x.columns[0]: 'CASNo'
              }
    df = x.rename(columns = Columns)
    df.loc[df.CASNo.isna(), 'CASNo'] = '-'

    s = df[df.Properties.str.contains('Priority Assessment Chemical Substances (PACSs)', regex = False)]
    s = s[Columns.values()]
    file = os.path.join('processed', '40.xlsx')
    s.to_excel(file, index = False)

    s = df[df.Properties.str.contains('Monitoring Chemical Substances', regex = False)]
    s = s[Columns.values()]
    file = os.path.join('processed', '41.xlsx')
    s.to_excel(file, index = False)

    s = df[df.Properties.str.contains('Class I specified chemical substance', regex = False)]
    s = s[Columns.values()]
    file = os.path.join('processed', '42.xlsx')
    s.to_excel(file, index = False)

    s = df[df.Properties.str.contains('Class II specified chemical substance', regex = False) == True]
    s = s[Columns.values()]
    file = os.path.join('processed', '43.xlsx')
    s.to_excel(file, index = False)

def n44_45():
    file = os.path.join('input', '44_45.xlsx')
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
    file = os.path.join('processed', '44.xlsx')
    s.to_excel(file, index = False)

    s = x[x['Specific Class I Designated Chemical Substances*'] == 'S']
    s = s[Columns.values()]
    file = os.path.join('processed', '45.xlsx')
    s.to_excel(file, index = False)

def n46():
    file = os.path.join('input', '46.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, skiprows = 2)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[2]: 'ChemicalEngName',
               s0.columns[1]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '46.xlsx')
    s0.to_excel(file, index = False)

def n48():
    file = os.path.join('processed', '48_1.xlsx')
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
    a = s0.ChemicalChnName.fillna('').str.contains('.', regex = False).index
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'].strip('0123456789. ')
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
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '48.xlsx')
    s0.to_excel(file, index = False)

def n49():
    file = os.path.join('processed', '49_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, dtype = 'string')
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    for i in s0.index:
        if pd.isna(s0.loc[i, 'CASNo']):
            s0.loc[i, 'CASNo'] = s0.loc[i, '编号']
    for i in s0.index[::-1]:
        if pd.isna(s0.loc[i, 'ChemicalChnName']):
            s0.loc[i-1, 'CASNo'] = s0.loc[i-1, 'CASNo'] + '; ' + s0.loc[i, 'CASNo']
            s0 = s0.drop(index = i)
    s0.CASNo = s0.CASNo.str.replace('P.+', '-', regex = True)
    s0.CASNo = s0.CASNo.str.replace('\(.+\)', '', regex = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（', '; ', regex = False)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('）', '', regex = False)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '49.xlsx')
    s0.to_excel(file, index = False)

def n50():
    file = os.path.join('input', '50.pdf')
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
    file = os.path.join('processed', '50.xlsx')
    df.to_excel(file, index = False)

def n56_57():
    file = os.path.join('input', '56_57.pdf')
    pdf = pdfplumber.open(file)
    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text"
    }
    page = pdf.pages[2]
    table = page.extract_table(table_settings)
    x = pd.DataFrame(table[:25])
    print(*enumerate(x.columns), '\n', sep = '\n')

    s0 = pd.DataFrame(x[0])
    Columns = {s0.columns[0]: 'ChemicalEngName'}
    s0 = s0.rename(columns = Columns)
    a = [9, 11, 13]
    for i in a:
        s0.loc[i-1, 'ChemicalEngName'] = s0.loc[i-1, 'ChemicalEngName'] + ' ' + s0.loc[i, 'ChemicalEngName']
        s0 = s0.drop(index = i)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' \(“*', '; ', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('”*\)', '', regex = True)

    s1 = pd.DataFrame(x[1])
    Columns = {s1.columns[0]: 'ChemicalEngName'}
    s1 = s1.rename(columns = Columns)
    s1 = s1[s1.ChemicalEngName != '']
    s1.ChemicalEngName = s1.ChemicalEngName.str.replace('acida', 'acid', regex = False)

    file = os.path.join('processed', '56.xlsx')
    s0.to_excel(file, index = False)
    file = os.path.join('processed', '57.xlsx')
    s1.to_excel(file, index = False)

def n58():
    file = os.path.join('processed', '58_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
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

    file = os.path.join('processed', '58.xlsx')
    s0.to_excel(file, index = False)

def n59():
    file = os.path.join('processed', '59_1.xlsx')
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
    file = os.path.join('processed', '59.xlsx')
    df.to_excel(file, index = False)

def n60():
    file = os.path.join('processed', '60_1.xlsx')
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

    file = os.path.join('processed', '60.xlsx')
    s0.to_excel(file, index = False)

def n61():
    file = os.path.join('processed', '61_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, nrows = 35, header = None)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(*\d+\)*(\xa0)', '', regex = True)
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(', including:', '', regex = False)

    file = os.path.join('processed', '61.xlsx')
    s0.to_excel(file, index = False)

def n62():
    file = os.path.join('processed', '62_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0, nrows = 6, header = None)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    s0.columns = ['ChemicalEngName']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace('\(*\d+\)*(\xa0)', '', regex = True)

    file = os.path.join('processed', '62.xlsx')
    s0.to_excel(file, index = False)

def n63():
    file = os.path.join('processed', '63_1.xlsx')
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

    file = os.path.join('processed', '63.xlsx')
    s0.to_excel(file, index = False)

def n64():
    file = os.path.join('processed', '64_1.xlsx')
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

    file = os.path.join('processed', '64.xlsx')
    s0.to_excel(file, index = False)

def n65_67():
    file = os.path.join('processed', '65_67.xlsx')
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
    
    file = os.path.join('processed', '65.xlsx')
    s0.to_excel(file, index = False)
    file = os.path.join('processed', '66.xlsx')
    df.to_excel(file, index = False)
    file = os.path.join('processed', '67.xlsx')
    s3.to_excel(file, index = False) 

def n68():
    file = os.path.join('processed', '68_1.xlsx')
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

    file = os.path.join('input', '68_2.pdf')
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
    file = os.path.join('processed', '68.xlsx')
    df.to_excel(file, index = False)

def n69():
    file = os.path.join('input', '69.csv')
    s0 = pd.read_csv(file, dtype = 'string')
    print(*enumerate(s0.columns), '\n', sep = '\n')

    Columns = {s0.columns[0]: 'ChemicalEngName',
               s0.columns[2]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.CASNo.fillna('-', inplace = True)
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '69_2.xlsx')
    s1 = pd.read_excel(file)

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('processed', '69.xlsx')
    df.to_excel(file, index = False)

def n72():
    file = os.path.join('input', '72.pdf')
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

    file = os.path.join('processed', '72.xlsx')
    s0.to_excel(file, index = False)

def n76():
    file = os.path.join('processed', '76_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')
    Columns = {s0.columns[1]: 'ChemicalChnName',
               s0.columns[3]: 'CASNo'
              }
    s0 = s0.rename(columns = Columns)
    s0.loc[87, 'CASNo'] = s0.loc[87, s0.columns[17]]
    s0.loc[94:, 'CASNo'] = s0.loc[94:, s0.columns[18]]
    s0 = s0[s0.ChemicalChnName.notna()]
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('\[.+\]', '', regex = True)
    s0.ChemicalChnName = s0.ChemicalChnName.str.replace('（.+）', '', regex = True)
    s0.别名 = s0.别名.str.replace('[、；]', '; ', regex = True)
    s0.CASNo.fillna('-', inplace = True)
    a = list(s0[s0.CASNo == '-'].index)
    a.remove(56)
    s0 = s0.drop(index = a)    
    a = s0[s0.别名.notna()].index
    for i in a:
        s0.loc[i, 'ChemicalChnName'] = s0.loc[i, 'ChemicalChnName'] + '; ' + s0.loc[i, '别名']
    s0 = s0[Columns.values()]

    file = os.path.join('processed', '76.xlsx')
    s0.to_excel(file, index = False)

def n77():
    file = os.path.join('processed', '77_1.xlsx')
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
    s1.CASNo = s1.CASNo.str.replace('\xa0(2\xa0(3 ', '', regex = False)
    s1 = s1[Columns]

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('processed', '77.xlsx')
    df.to_excel(file, index = False)

def n78():
    file = os.path.join('processed', '78_1.xlsx')
    x = pd.ExcelFile(file)
    print(*enumerate(x.sheet_names), '\n', sep = '\n')

    s0 = x.parse(sheet_name = 0)
    print(*enumerate(s0.columns), '\n', sep = '\n')

    s1 = x.parse(sheet_name = 1)
    print(*enumerate(s1.columns), '\n', sep = '\n')

    df = pd.concat([s0, s1], ignore_index = True)
    file = os.path.join('processed', '78.xlsx')
    df.to_excel(file, index = False)

def n79():
    file = os.path.join('input', '79.pdf')
    pdf = pdfplumber.open(file)
    a = []
    for i in range(1, 3):
        page = pdf.pages[i]
        table = page.extract_table()
        for i in range(len(table[0])):
            a += [table[0][i].split('\n')]
    for i in a:
        for j in range(len(i)):
            i[j] = i[j].strip()
    a[5] = [''] + a[5]
    ChemicalEngName = []
    for i in range(0, 10, 2):
        ChemicalEngName += a[i]
    CASNo = []
    for i in range(1, 10, 2):
        CASNo += a[i]

    s0 = pd.DataFrame([ChemicalEngName, CASNo]).transpose()
    s0.columns = ['ChemicalEngName', 'CASNo']
    a = s0[s0.ChemicalEngName.str.fullmatch('[A-Z]')].index
    s0 = s0.drop(index = a)
    s0 = s0[s0.ChemicalEngName != '']
    s0.ChemicalEngName = s0.ChemicalEngName.str.replace(' *', '', regex = False)

    file = os.path.join('processed', '79.xlsx')
    s0.to_excel(file, index = False)

def n80():
    file = os.path.join('processed', '80_1.xlsx')
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
    file = os.path.join('processed', '80.xlsx')
    df.to_excel(file, index = False)

def n83():
    file = os.path.join('input', '83.pdf')
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

    file = os.path.join('processed', '83.xlsx')
    s0.to_excel(file, index = False)

def Merged():
    Columns = ['ChemicalChnName',
               'ChemicalEngName',
               'CASNo',
               'Name'
              ]
    df = pd.DataFrame(columns = Columns)
    os.chdir('processed')
    for dirpath, dirnames, filenames in os.walk('.'):
        for name in filenames:
            if '_' not in name:
                df2 = pd.read_excel(name)
                df = pd.concat([df, df2], ignore_index = True)
                df.Name.fillna(name[:-5], inplace = True)
    os.chdir('..')

    file = os.path.join('清單.xlsx')
    df2 = pd.read_excel(file, dtype = 'string')
    Columns = {'國內/外': 'Type',
               '單位/國家': 'Unit',
               '清單': 'Source'  
              }
    df2 = df2.rename(columns = Columns)
    df2.Type = df2.Type.replace('國內', '0', regex = False)
    df2.Type = df2.Type.replace('國外', '1', regex = False)
    Merged = pd.merge(left = df, right = df2, how = 'left',
                      left_on = 'Name', right_on = '編號',
                      validate = 'many_to_one'
                     )
    Columns = ['ChemicalChnName',
               'ChemicalEngName',
               'CASNo',
               'Source',
               'Type',
               'Unit'
              ]
    Merged = Merged[Columns]
    Merged = Merged.fillna('-')
    Merged.ChemicalEngName = Merged.ChemicalEngName.str.lower()
    Merged.ChemicalEngName = Merged.ChemicalEngName.str.strip()
    Merged.CASNo = Merged.CASNo.str.strip()
    Merged = Merged[~Merged.duplicated()]
    file = os.path.join('output', 'Merged.xlsx')
    Merged.to_excel(file, index = False)