import os
import pandas as pd
#import gc
#import numpy as np
from sqlalchemy import create_engine

path = os.getcwd()
db = create_engine('sqlite:///app.db')
conn = db.connect()
allFiles = os.listdir(path)

df_new = pd.DataFrame()
df_project = pd.DataFrame()
df_leaders = pd.DataFrame()
df_preleaders = pd.DataFrame()

dict_new = {}
dict_projects = {}
dict_leaders = {}
dict_preleaders = {}
       
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
                        dict_leaders = [{'id3' : str(l3) + '_GVK-13_1',
                                         'position' : l0,
                                         'full_name' : l1,
                                         'kpi_weight' : l2,
                                         'index' : l3,
                                         'GVK' : l4}]
                        dToFrame2 = pd.DataFrame.from_dict(dict_leaders)
                        df_leaders = df_leaders.append(dToFrame2)
                   
                    if flag2 == 1 and row[0] != 'Должность' and row[0]!= 0:          
                        pl0 = row[0]
                        pl1 = row[3]
                        pl3 = n4
                        pl4 = 'ГВК-13'
#                                            'Должность' : pl0,
#                                            'ФИО' : pl1,
#                                            'Вес КПЭ в целях (если применимо)' : pl2,
#                                            'Индекс' : pl3 + 'ГВК-13_1',
#                                            'ГВК' 
                        dict_preleaders = [{'id3' : str(pl3) + '_GVK-13_1',
                                            'position' : pl0,
                                            'full_name' : pl1,
                                            'kpi_weight' : pl2,
                                            'index' : pl3,
                                            'GVK' : pl4}]
                        dToFrame3 = pd.DataFrame.from_dict(dict_preleaders)
                        df_preleaders = df_preleaders.append(dToFrame3)
                        
                #df_leaders = df_leaders[df_leaders.full_name[-9:] != 0]       
                dict_project = [{'id3' : str(pr0) + '_GVK-13_1',
                                 'index' : pr0,
                                 'ID' : pr1,
                                 'project_name' : pr2}]
                dToFrame0 = pd.DataFrame.from_dict(dict_project)
                df_project = df_project.append(dToFrame0)
    
#                             'Наименование сервиса': n0,
#                             'Описание сервиса': n1,
#                             'Владелец' : n2,
#                             'Блок' : n3,
#                             'Индекс': n4,
#                             'Подразделение - владелец сервиса' : n5,
#                             'Принадлежность к группе' : n6,
#                             'Количество пользователей' : n7,
#                             'Группировка для веб-формы опроса' : n8}]
                dict_new = [{'id3' : str(n4) + '_GVK-13_1',
                             'service_name': n0,
                             'service_desc': n1,
                             'owner' : n2,
                             'block' : n3,
                             'index': n4,
                             'unit' : n5,
                             'membership' : n6,
                             'number_user' : n7,
                             'web_form' : n8}]
                dToFrame1 = pd.DataFrame.from_dict(dict_new)
                df_new = df_new.append(dToFrame1)
                
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
            