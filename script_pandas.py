import pandas as pd

df = pd.read_excel('port.xlsx',sheet_name='Лист2')
df2 = pd.read_excel('anal.xlsx',sheet_name='Лист2')
df['Столбец4'] = df['Столбец4'].dt.strftime('%d.%m.%Y')
df2['Столбец4'] = df2['Столбец4'].dt.strftime('%d.%m.%Y')
result_df = pd.concat([df, df2])

# port
# mes
mes_df = result_df.query("Столбец3 != 'Выполнено' and Столбец3 != 'Закрыто' and Столбец3 != 'Отклонено' and Столбец5 == 'MES'")
mes_df.loc[:, 'Столбец1'] = range(1, len(mes_df) + 1)
mes_df.to_excel('mes.xlsx', index=False)

# olap
olap_df = result_df.query("Столбец3 != 'Выполнено' and Столбец3 != 'Закрыто' and Столбец3 != 'Отклонено' and Столбец5 == 'OLAP'")
olap_df.loc[:, 'Столбец1'] = range(1, len(olap_df) + 1)
olap_df.to_excel('olap.xlsx', index=False)

# sti
sti_df = result_df.query("Столбец3 != 'Выполнено' and Столбец3 != 'Закрыто' and Столбец3 != 'Отклонено' and Столбец5 == 'STI'")
sti_df.loc[:, 'Столбец1'] = range(1, len(sti_df) + 1)
sti_df.to_excel('sti.xlsx', index=False)