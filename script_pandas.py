import pandas as pd

df = pd.read_excel('port.xlsx',sheet_name='Лист2')
df2 = pd.read_excel('anal.xlsx',sheet_name='Лист2')

df = df.drop(df.columns[[11, 12, 13, 14, 15, 16]], axis=1)
df2 = df2.drop(['Пустой столбец1', 'Пустой столбец2', 'Пустой столбец3', 'Пустой столбец4',
                     'Пустой столбец5', 'Пустой столбец6', 'Затраченное время'], axis=1)

df['Дата регистрации'] = df['Дата регистрации'].dt.strftime('%d.%m.%Y')
df['Дата решения'] = df['Дата решения'].dt.strftime('%d.%m.%Y')
df2['Дата регистрации'] = df2['Дата регистрации'].dt.strftime('%d.%m.%Y')
df2['Дата завершения заявки'] = df2['Дата завершения заявки'].dt.strftime('%d.%m.%Y')
result_df = pd.concat([df, df2])

# mes
mes_df = result_df.query("Статус != 'Выполнено' and Статус != 'Закрыто' and Статус != 'Отклонено' and Статус != 'Отозвано' "
                         "and Соглашение == 'MES'")
mes_df.insert(0, "№", 0)
mes_df.loc[:, '№'] = range(1, len(mes_df) + 1)
mes_df.to_excel('mes.xlsx', index=False)

# olap
olap_df = result_df.query("Статус != 'Выполнено' and Статус != 'Закрыто' and Статус != 'Отклонено' and Статус != 'Отозвано'"
                          "and Статус != 'На тестировании' and Статус != 'Планируется в релизе' and Соглашение == 'OLAP'")
olap_df.insert(0, "№", 0)
olap_df.loc[:, '№'] = range(1, len(olap_df) + 1)
olap_df.to_excel('olap.xlsx', index=False)

# sti
sti_df = result_df.query("Статус != 'Выполнено' and Статус != 'Закрыто' and Статус != 'Отклонено' and Статус != 'Отозвано' "
                         "and Соглашение == 'STI' and Организация != 'ООО\"Газпром межрегионгаз ЕЦРК\"'")
sti_df.insert(0, "№", 0)
sti_df.loc[:, '№'] = range(1, len(sti_df) + 1)
sti_df.to_excel('sti.xlsx', index=False)

# eog
eog_df = result_df.query("Статус != 'Выполнено' and Статус != 'Закрыто' and Статус != 'Отклонено' and Статус != 'Отозвано' "
                         "and Соглашение == 'STI'")
eog_df.insert(0, "N", 0)
eog_df.loc[:, '№'] = range(1, len(sti_df) + 1)
eog_df.to_excel('eog.xlsx', index=False)

# ecrk
ecrk_df = result_df.query("Статус != 'Выполнено' and Статус != 'Закрыто' and Статус != 'Отклонено' and Статус != 'Отозвано'"
                          " and Соглашение == 'STI'")
ecrk_df.insert(0, "N", 0)
ecrk_df.loc[:, '№'] = range(1, len(sti_df) + 1)
ecrk_df.to_excel('eog.xlsx', index=False)