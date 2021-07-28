import pandas as pd

df = pd.read_excel('ТПП.xlsx')
kalculation = pd.read_excel('Kalculation.xls')

df.fillna(0, inplace=True)
kalculation.columns = kalculation.loc[0,:]
kalculation.drop(0, inplace = True)

col = df.columns
df.iloc[:, 4] = df.iloc[:, 4].astype(str)
df.iloc[:, 0] = df.iloc[:, 0].astype(str)

dff = pd.DataFrame()
dff['name'] = range(len(df))
dff['%'] = range(len(df))

for i in range(len(df)):

    # Кількість сировини
    kalculation.кількість[2] = df['ПЕ'][i]
    kalculation.кількість[6] = df['Wh 70'][i] + df['Wh 80'][i] + df['Br 70'][i] + df['Br 80'][i]
    kalculation.кількість[7] = df['Br 70'][i] + df['Br 80'][i]
    kalculation.кількість[8] = df['Wh 70'][i] + df['Wh 80'][i]

    kalculation.кількість[11] = df['концентрат'][i]
    kalculation.кількість[12] = df['дисперсія'][i]
    kalculation.кількість[13] = df['пов-акт. зас.'][i]
    kalculation.кількість[10] = kalculation['кількість'][10:13].sum()
    kalculation.кількість[14] = df['клей'][i]

    kalculation['Сума попередня, грн'][18] = df['звв, грн'][i]
    kalculation['Сума попередня, грн'][19] = df['з/п, грн'][i]
    kalculation['Сума попередня, грн'][20] = df['опер. витр. грн'][i]
    kalculation['Сума попередня, грн'][21] = df['адмін. Витрати, грн'][i]
    kalculation['Сума попередня, грн'][22] = df['амортизація, грн'][i]
    kalculation['Сума попередня, грн'][23] = kalculation['Сума попередня, грн'][17:22].sum()

    # ==============================================================================================
    kalculation['Сума попередня, грн'][2] = df['ПЕ, грн'][i]
    kalculation['Сума попередня, грн'][3] = df['ПЕ, грн'][i]
    kalculation['Сума попередня, грн'][6] = df['Wh, грн'][i] + df['Br, грн'][i]
    kalculation['Сума попередня, грн'][7] = df['Br, грн'][i]
    kalculation['Сума попередня, грн'][8] = df['Wh, грн'][i]

    kalculation['Сума попередня, грн'][11] = df['Концентрат, грн'][i]
    kalculation['Сума попередня, грн'][12] = df['дисп., грн'][i]
    kalculation['Сума попередня, грн'][13] = df['ПАЗ, грн'][i]
    kalculation['Сума попередня, грн'][10] = kalculation['Сума попередня, грн'][10:14].sum()
    kalculation['Сума попередня, грн'][14] = df['клей, грн'][i]
    kalculation['Сума попередня, грн'][15] = kalculation['Сума попередня, грн'][10] + kalculation['Сума попередня, грн'][6] + kalculation['Сума попередня, грн'][14]

    kalculation['Сума попередня, грн'][18] = df['звв, грн'][i]
    kalculation['Сума попередня, грн'][19] = df['з/п, грн'][i]
    kalculation['Сума попередня, грн'][20] = df['опер. витр. грн'][i]
    kalculation['Сума попередня, грн'][21] = df['адмін. Витрати, грн'][i]
    kalculation['Сума попередня, грн'][22] = df['амортизація, грн'][i]
    kalculation['Сума попередня, грн'][23] = kalculation['Сума попередня, грн'][17:22].sum()

    kalculation['Сума попередня, грн'][25] = kalculation['Сума попередня, грн'][15] + kalculation['Сума попередня, грн'][3] + kalculation['Сума попередня, грн'][23]
    kalculation['Сума попередня, грн'][27] = str((100*kalculation['Сума попередня, грн'][15]/kalculation['Сума попередня, грн'][25]).round(2)) + ' %'



    name = str('{} рахунок № {} від {}, {}x{}x{}.xlsx'.format(i, df[col[3]][i], df[col[4]][i][0:10],df[col[5]][i],df[col[6]][i],df[col[7]][i]))

    dff.name[i]= name
    dff['%'][i] = kalculation['Сума попередня, грн'][27]

    kalculation.to_excel(name, index=False)
print(kalculation.to_string())