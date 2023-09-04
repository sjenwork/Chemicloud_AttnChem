file = os.path.join('processed', '36_2.xlsx')
x3 = pd.read_excel(file)
print(*enumerate(x3.columns), '\n', sep = '\n')
x3['NPRI Substance Name'] = x3['NPRI Substance Name'].str.replace(' A', ' a', regex = True)
df = pd.merge(left = s0, right = x3, how = 'left',
              left_on = 'ChemicalEngName', right_on = 'NPRI Substance Name',
              validate = 'many_to_one'
             )
a = df[df.Synonyms.notna()].index
for i in a:
    df.loc[i, 'ChemicalEngName'] = df.loc[i, 'ChemicalEngName'] + '; ' + df.loc[i, 'Synonyms']
