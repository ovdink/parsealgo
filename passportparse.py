import os
import pandas as pd
import gc
import numpy as np

path = os.getcwd()
allFiles = os.listdir(path)

for file in allFiles:
    if file[-4:] == 'xlsx':
        print('\nНашли xlsx: ' + file + '\n')
        dataFile = pd.ExcelFile(file)
        dataSheet = dataFile.sheet_names
        df_new = pd.DataFrame(np.nan, index=range(200), columns=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
        df_leaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
        df_leaders = df_leaders.fillna(0)
        df_preleaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
        df_preleaders = df_leaders.fillna(0)
        df_projects = pd.DataFrame(np.nan, index=range(200), columns=['Индекс', 'ID', 'Название проекта/программы'])
        l = 0
        j = 0
        k = 0
        for sheet in dataSheet:
            if sheet[-2:] != 'ие' and sheet[-2:] != 'ЛМ':
                print('Нашли лист: ' + sheet + '\n')
                df = pd.read_excel(file, sheet)
                df = df.iloc[6:]
                df = df.reset_index(drop=True)
                df.columns = range(df.shape[1])
                df = df.fillna(0)
#                df_new = pd.DataFrame(np.nan, index=range(200), columns=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
#                df_leaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#                df_leaders = df_leaders.fillna(0)
#                df_preleaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#                df_preleaders = df_leaders.fillna(0)
#                df_projects = pd.DataFrame(np.nan, index=range(200), columns=['Индекс', 'ID', 'Название проекта/программы'])
                
                flag = 0
                flag2 = 0
#                j = 0
#                k = 0
                dolCount = 0
                for index, row in df.iterrows():
                    if row[0] == 'Наименование сервиса':
                        df_new.iloc[l, 0] = row[2]
                    if row[0] == 'Описание сервиса':
                        df_new.iloc[l, 1] = row[2]
                    if row[0] == 'Владелец':
                        df_new.iloc[l, 2] = row[2]
                    if row[0] == 'Блок':
                        df_new.iloc[l, 3] = row[2]
                    if row[0] == 'Индекс':
                        df_new.iloc[l, 4] = row[2]
                    if row[5] == 'Подразделение - владелец сервиса':
                        df_new.iloc[l, 5] = row[7]
                    if row[0] == 'Принадлежность к группе':
                        df_new.iloc[l, 6] = row[2]
                    if row[0] == 'Количество пользователей':
                        df_new.iloc[l, 7] = row[2]
                    if row[0] == 'Группировка для веб-формы опроса':
                        df_new.iloc[l, 8] = row[2]
                        l += 1
                    if row[0] == 'Должность' and dolCount == 0:
                        flag = 1
                        dolCount += 1
                    if row[0] == 0 and dolCount > 0:
                        flag = 0
                        flag2 = 0
                    if flag == 1 and row[0] != 'Должность' and row[0] != 0 and df_leaders.iloc[j, 0] == 0:
                        df_leaders.iloc[j, 0] = row[0]
                        df_leaders.iloc[j, 1] = row[3]
                        df_leaders.iloc[j, 3] = df_new.iloc[0, 4]
                        df_leaders.iloc[j, 4] = 'ГВК-13'
                        j += 1
                    if row[0] == 'Должность' and df_preleaders.iloc[0, 0] == 0 and flag == 0 and dolCount == 1:
                        print('check')
                        flag2 = 1
                    if flag2 == 1 and row[0] != 'Должность' and row[0] != 0 and df_preleaders.iloc[k, 0] == 0:
                        df_preleaders.iloc[k, 0] = row[0]
                        df_preleaders.iloc[k, 1] = row[3]
                        df_preleaders.iloc[k, 3] = df_new.iloc[0, 4]
                        df_preleaders.iloc[k, 4] = 'ГВК-13'
                        k += 1
                print('Создали файл\n')
                gc.collect()