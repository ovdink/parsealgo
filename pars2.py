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
df_project = pd.DataFrame()
df_leaders = pd.DataFrame()
df_preleaders = pd.DataFrame()


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
       
n0 = n1 = n3 = n4 = n5 = n6 = n7 = n8 = 0
l0 = l1 = l2 = l3 = l4 = 0
pl0 = pl1 = pl2 = pl3 = 0
pr0 = pr1 = pr2 = 0

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
                
                for index, row in df.iterrows():
                    if row[0] == 'Наименование сервиса':
                        n0 = row[2]
                    if row[0] == 'Описание сервиса':
                        n1 = row[2]
                    if row[0] == 'Владелец:' or row[0] == 'Владелец':
                        n2 = row[2]
                    if row[0] == 'Блок':
                        n3 = row[2]
                    if row[0] == 'Индекс':
                        n4 = row[2]
                    if row[5] == 'Подразделение - владелец сервиса':
                        n5 = row[7]
                    if row[0] == 'Принадлежность к группе':
                        n6 = row[2]
                    if row[0] == 'Количество пользователей':
                        n7 = row[2]  
                    if row[0] == 'Влияние  проекта/программы на сервис:':
                        pr0 = n4
                        pr1 = row[2]
                        pr2 = row[3]
                    if row[0] == 'Группировка для веб-формы опроса':
                        n8 = row[2]
                    if row[0] == 'Должность' and dolCount == 0:
                        flag = 1
                        dolCount += 1
                    if row[0] == 'Должность' and dolCount == 1 and flag == 0:
                        flag2 = 1
                        dolCount += 1
                    if row[0] == 0 and dolCount == 1:
                        flag = 0
                        #flag2 = 0
                    if row[0] == 0 and len(df_preleaders) != 0 and dolCount == 2:
                        flag2 = 0
                    if flag == 1 and row[0] != 'Должность' and row[0] != 0:
                        l0 = row[0]
                        l1 = row[3]
                        l3 = n4
                        l4 = 'ГВК-13'
                        dict_leaders = [{'Должность' : l0,
                                         'ФИО' : l1,
                                         'Вес КПЭ в целях (если применимо)' : l2,
                                         'Индекс' : l3,
                                         'ГВК' : l4}]
                        dToFrame2 = pd.DataFrame.from_dict(dict_leaders)
                        df_leaders = df_leaders.append(dToFrame2)
                   
                    if flag2 == 1 and row[0] != 'Должность' and row[0]!= 0:          
                        pl0 = row[0]
                        pl1 = row[3]
                        pl3 = n4
                        pl4 = 'ГВК-13'
                        dict_preleaders = [{'Должность' : pl0,
                                            'ФИО' : pl1,
                                            'Вес КПЭ в целях (если применимо)' : pl2,
                                            'Индекс' : pl3,
                                            'ГВК' : pl4}]
                        dToFrame3 = pd.DataFrame.from_dict(dict_preleaders)
                        df_preleaders = df_preleaders.append(dToFrame3)
                    #if flag2 == 0 and len(df_preleaders)
#                    elif flag2 == 1 and row[0] != 'Должность' and row[0] != 0:
#                        print(1)
                        
                df_leaders = df_leaders[df_leaders.ФИО != 0]       
                dict_project = [{'Индекс' : pr0,
                                 'ID' : pr1,
                                 'Название проекта/программы' : pr2}]
                dToFrame0 = pd.DataFrame.from_dict(dict_project)
                df_project = df_project.append(dToFrame0)
    
                dict_new = [{'Наименование сервиса': n0,
                             'Описание сервиса': n1,
                             'Владелец' : n2,
                             'Блок' : n3,
                             'Индекс': n4,
                             'Подразделение - владелец сервиса' : n5,
                             'Принадлежность к группе' : n6,
                             'Количество пользователей' : n7,
                             'Группировка для веб-формы опроса' : n8}]
                dToFrame1 = pd.DataFrame.from_dict(dict_new)
                df_new = df_new.append(dToFrame1)
                
                
                  
                    
#                dict_leaders = [{'Должность' : l0,
#                                 'ФИО' : l1,
#                                 'Вес КПЭ в целях (если применимо)' : l2,
#                                 'Индекс' : l3,
#                                 'ГВК' : l4}]
#                dToFrame2 = pd.DataFrame.from_dict(dict_leaders)
#                df_leaders = df_leaders.append(dToFrame2)
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

        df_leaders.to_sql('_leaders', con = db, if_exists='append')
        print('Заполнили таблицу leaders')
        df_new.to_sql('_general', con = db, if_exists='append')
        print('Заполнили таблицу general')
        df_preleaders.to_sql('_preleaders', con = db, if_exists='append')
        print('Заполнили таблицу preleaders')
        df_project.to_sql('_projects', con = db, if_exists='append')
        print('Заполнили таблицу projects')
        
conn.close()
print('Соединение прервано\n')
db.dispose()
            