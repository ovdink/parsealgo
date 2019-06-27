import os
import pandas as pd
#import gc
#import numpy as np
from sqlalchemy import create_engine

path = os.getcwd()
db = create_engine('sqlite:///app.db')
conn = db.connect()
allFiles = os.listdir(path)

#df_new = pd.DataFrame(np.nan, index=range(200), columns=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
#df_leaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#df_leaders = df_leaders.fillna(0)
#df_preleaders = pd.DataFrame(np.nan, index = range(300), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#df_preleaders = df_leaders.fillna(0)
#df_projects = pd.DataFrame(np.nan, index=range(200), columns=['Индекс', 'ID', 'Название проекта/программы'])
#dict_new = {'Наименование сервиса',
#             'Описание сервиса',
#             'Владелец',
#             'Блок',
#             'Индекс',
#             'Подразделение - владелец сервиса',
#             'Принадлежность к группе',
#             'Количество пользователей',
#             'Группировка для веб-формы опроса',
#             'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'}
df_new = pd.DataFrame()
dict_new = {}
dict_projects = {}
dict_leaders = {}
dict_preleaders = {}

#df_new = pd.DataFrame(columns=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
#df_new = df_new.append([0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
#print ('ДФ_НЬЮ\n', df_new)
#df_leaders = pd.DataFrame(columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#df_leaders.append([0, 1, 2, 3, 4])
#df_leaders = df_leaders.fillna(0)
#df_preleaders = pd.DataFrame(columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#df_preleaders.append([0, 1, 2, 3, 4])
#df_preleaders = df_leaders.fillna(0)
#df_projects = pd.DataFrame(columns=['Индекс', 'ID', 'Название проекта/программы'])
#df_projects.append([0, 1, 2])
#dic = {}
#cols = ['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен']
#for i in range(12,0,-1):
#    dic[str(i)] = df[cols].shift(i).add_sufix(str(i))
       
l = 0
j = 0
k = 0
n0 = n1 = n3 = n4 = n5 = n6 = n7 = n8 = 0
df_sv = pd.DataFrame()
for file in allFiles:
    if file[-4:] == 'xlsx':
        print('\nНашли xlsx: ' + file + '\n')
        dataFile = pd.ExcelFile(file)
        dataSheet = dataFile.sheet_names        
        for sheet in dataSheet:
            if sheet[-2:] != 'ие' and sheet[-2:] != 'ЛМ':
                print('Нашли лист: ' + sheet + '\n')
                df = pd.read_excel(file, sheet)
                df = df.iloc[6:]
                df = df.reset_index(drop=True)
                df.columns = range(df.shape[1])
                df = df.fillna(0)
                
                flag = 0
                flag2 = 0
                dolCount = 0
                
                #newStr = pd.Series([0, 0, 0, 0, 0, 0, 0, 0, 0, 0], index=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
                for index, row in df.iterrows():
                    #df_new = df_new.a
                    #df_new.iloc[l, 0] = df_new.append(newStr)
                    #df_new.append(newStr)
                    if row[0] == 'Наименование сервиса':
                        n0 = row[2]
                        print(sheet + ' n0 :\n', n0)
                    if row[0] == 'Описание сервиса':
                        n1 = row[2]
                        print(sheet + ' n1 :\n', n1)
                    if row[0] == 'Владелец:' or row[0] == 'Владелец':
                        n2 = row[2]
                        print(sheet + ' n2 :\n', n2)
                    if row[0] == 'Блок':
                        n3 = row[2]
                        print(sheet + ' n3 :\n', n3)
                    if row[0] == 'Индекс':
                        n4 = row[2]
                        print(sheet + ' n4 :\n', n4)
                    if row[5] == 'Подразделение - владелец сервиса':
                        n5 = row[7]
                        print(sheet + ' n5 :\n', n5)
                    if row[0] == 'Принадлежность к группе':
                        n6 = row[2]
                        print(sheet + ' n6 :\n', n6)
                    if row[0] == 'Количество пользователей':
                        n7 = row[2]  
                        print(sheet + ' n7 :\n', n7)
                    if row[0] == 'Влияние  проекта/программы на сервис:':
                        pr0 = n4
                        pr1 = row[2]
                        pr2 = row[3]
                    if row[0] == 'Группировка для веб-формы опроса':
                        n8 = row[2]
                        #l += 1
                    if row[0] == 'Должность' and dolCount == 0:
                        flag = 1
                        dolCount += 1
                    if row[0] == 0 and dolCount > 0:
                        flag = 0
                        flag2 = 0
#                    if flag == 1 and row[0] != 'Должность' and row[0] != 0 and len(dict_leaders) == 0: #df_leaders.iloc[j, 0] == 0:
#                        df_leaders.iloc[j, 0] = row[0]
#                        df_leaders.iloc[j, 1] = row[3]
#                        df_leaders.iloc[j, 3] = df_new.iloc[l, 4]
#                        df_leaders.iloc[j, 4] = 'ГВК-13'
#                        j += 1
#                        print ('строки leaders', j)
#                    if row[0] == 'Должность' and df_preleaders.iloc[k, 0] == 0 and flag == 0 and dolCount == 1:
#                        flag2 = 1
#                    if flag2 == 1 and row[0] != 'Должность' and row[0] != 0 and df_preleaders.iloc[k, 0] == 0:
#                        df_preleaders.iloc[k, 0] = row[0]
#                        df_preleaders.iloc[k, 1] = row[3]
#                        df_preleaders.iloc[k, 3] = df_new.iloc[l, 4]
#                        df_preleaders.iloc[k, 4] = 'ГВК-13'
#                        k += 1
#                        print ('строки preleaders', k)
                print('Создали файл\n')
#                gc.collect()
        dict_new = [{'Наименование сервиса': n0,
                             'Описание сервиса': n1,
                             'Владелец' : n2,
                             'Блок' : n3,
                             'Индекс': n4,
                             'Подразделение - владелец сервиса' : n5,
                             'Принадлежность к группе' : n6,
                             'Количество пользователей' : n7,
                             'Группировка для веб-формы опроса' : n8}]
        print(dict_new)
            
        dict_k = pd.DataFrame.from_dict(dict_new)
        df_sv = df_sv.append(dict_k)
    
#        df_leaders.to_sql('_leaders', con = db, if_exists='append')
#        df_new.to_sql('_general', con = db, if_exists='append')
#        df_preleaders.to_sql('_preleaders', con = db, if_exists='append')
#        df_projects.to_sql('_projects', con = db, if_exists='append')
#        
#conn.close()
#db.dispose()

            