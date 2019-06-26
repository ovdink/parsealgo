import os
import pandas as pd
import gc
import numpy as np

path = os.getcwd()
allFiles = os.listdir(path)

#arrSheets = []

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
                #df.columns = df.iloc[1]
                #df1 = df.T
                #del df1['6']
                #t = df1.iloc[0, 1]
                #df1.columns = ['Наименование сервиса', 'Описание сервиса']
                #t = df.iterrows()
                
                #print(t)
                df_new = pd.DataFrame('nan', index=range(2), columns=['Наименование сервиса', 'Описание сервиса', 'Владелец', 'Блок', 'Индекс', 'Подразделение - владелец сервиса', 'Принадлежность к группе', 'Количество пользователей', 'Группировка для веб-формы опроса', 'КПЭ "Удовлетворенность внутренних клиентов" в проект/программу включен'])
                df_leaders = pd.DataFrame('nan', index = range(10), columns=['Должность', 'ФИО', 'Вес КПЭ в целях (если применимо)', 'Индекс', 'ГВК'])
#  df_new.iloc[0, 0] = df.iloc[1, 2]
#                df_new.iloc[0, 1] = df.iloc[3, 2]
#                df_new.iloc[0, 2] = df.iloc[4, 2]
#                df_new.iloc[0, 3] = df.iloc[5, 2]
#                df_new.iloc[0, 4] = df.iloc[6, 2]
#                df_new.iloc[0, 5] = df.iloc[8, 3]
#                df_new.iloc[0, 6] = df.iloc[9, 3]
#                df_new.iloc[0, 7] = df.iloc[10, 3]
#                df_new.iloc[0, 8] = df.iloc[11, 3]
#                df_new.iloc[0, 9] = df.iloc[14, 3]
#                df_new.iloc[0, 10] = df.iloc[21, 2]
#                df_new.iloc[0, 11] = df.iloc[23, 2]
#                x = 'dol' + str(5)
#                print(x)
                col = 0
                forDol = 5
                flag = 0
                flag2 = 0
                flag3 = 0
                j = 0
                for index, row in df.iterrows():
                    if row[0] == 'Наименование сервиса':
                        df_new.iloc[0, 0] = row[2]
                    if row[0] == 'Описание сервиса':
                        df_new.iloc[0, 1] = row[2]
                    if row[0] == 'Владелец':
                        df_new.iloc[0, 2] = row[2]
                    if row[0] == 'Блок':
                        df_new.iloc[0, 3] = row[2]
                    if row[0] == 'Индекс':
                        df_new.iloc[0, 4] = row[2]
                    if row[5] == 'Подразделение - владелец сервиса':
                        df_new.iloc[0, 5] = row[7]
                    if row[0] == 'Принадлежность к группе':
                        df_new.iloc[0, 6] = row[2]
                    if row[0] == 'Количество пользователей':
                        df_new.iloc[0, 7] = row[2]
                    if row[0] == 'Группировка для веб-формы опроса':
                        df_new.iloc[0, 8] = row[2]
                    if row[0] == 'Должность':
                        flag = flag + 1
                    if flag == 1 and row[0] != 'Должность' and row[0] != 'nan':
                        print(1)
                        df_leaders.iloc[0, 0] = row[0]
                    if row[0] == np.nan:
                        flag3 
                    #if flag > 1 and row[0]
                    flag3 = np.nansum(flag2)
                    print(flag3)
                    
                    
#                    if flag == 1 and row[0] == 'nan':
#                        flag = 0
#                    #print(forDol)
#                    col += 1
#                    forDol += 1
                    
                    
                                  
                    #index += 1
                    #print(index)
                    #break
#                    index += 1
                    #df_new.iloc[0, 1] = row[0]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
#                        df_new.iloc[0, 1] = row[2]
                        
#                for index, row in df.iterrows():
#                    if row[0] == 'ПЦП':
#                        print(1)


                
                #df_new.iloc[0, 1] = df.iloc[3, 4]
                #df.drop([6], inplace = True)
                #t = sheet['A9']
                #df_d5 = df.loc[df['Демография_5'] == 'Подразделение прямых продаж']
                #df_d5.reset_index()
                #df.to_excel(file[:-4] + sheet + '_passport' + '.xlsx')
                print('Создали файл\n')
                gc.collect()