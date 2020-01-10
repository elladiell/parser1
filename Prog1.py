import pandas as pd
import numpy as np
from datetime import date, datetime
import datetime as DT
import time
import calendar
from pandas import DataFrame

# empty dataframe where the result is
allmonths = pd.DataFrame(
    columns=['Дата', 'Qн ТМ', 'F вращ ТМ', 'ТМ Qж(тн)', 'Qж(тн)', 'Обв', 'Qж', 'Обв ХАЛ', 'Qж ТМ', 'F', 'Dшт',
             'Qж ТМ (исх)', 'Рбуф ТМ', 'Qг (АГЗУ)', 'Ртр', 'Рбуф ТМ', 'ГФ', ' Обв ТМ ручн', 'Qн ТМ ручн',
             'Pзаб(PпрTM)', 'Рзат ТМ', 'Примечание'])

dictionary = dict(Декабрь='12', Январь='01', Февраль='02', Март='03', Апрель='04', Май='05', Июнь='06', Июль='07',
                  Август='08', Сентябрь='09', Октябрь='10', Ноябрь='11')

df = pd.read_excel('C:/Users/USER/Desktop/desk/Шахматка, скв. 5301    (01.01.1993 - 01.12.2019).xlsx')

# for moving among the cells
formonths = 3
number_of_features = 21
max = int((df.shape[0] - formonths) / number_of_features)

for i in range(0, max):

    # working with month
    months = df.iloc[i * number_of_features + formonths][1]
    # print(months)

    month = months.split(' ')[0]
    year = int(months.split()[1])
    # print(year)

    # finding keys
    if month in dictionary.keys():
        month = int(dictionary[month])

    # how much days in month
    days_in_month = calendar.monthrange(year, month)[1]
    # print(days_in_month)

    new = df.iloc[i * number_of_features + formonths:i * number_of_features + number_of_features + formonths,
          6:days_in_month + 6].T
    new.columns = ['Qн ТМ', 'F вращ ТМ', 'ТМ Qж(тн)', 'Qж(тн)', 'Обв', 'Qж', 'Обв ХАЛ', 'Qж ТМ', 'F', 'Dшт',
                   'Qж ТМ (исх)', 'Рбуф ТМ', 'Qг (АГЗУ)', 'Ртр', 'Рбуф ТМ', 'ГФ', ' Обв ТМ ручн', 'Qн ТМ ручн',
                   'Pзаб(PпрTM)', 'Рзат ТМ', 'Примечание']
    time = []
    for a in range(days_in_month):
        dti = str(a + 1) + '.' + str(month) + '.' + str(year)
        time.append(dti)
    new.insert(0, 'Дата', time)

    # upload every month from excel
    allmonths = allmonths.append(new, ignore_index=True)

# new file
allmonths.to_excel('C:/Users/USER/Desktop/desk/result.xlsx', startrow=0, index=False)

# another
# hat1.to_excel('C:/Users/shevc/Desktop/desk/result.xlsx', startrow=0, index=False)
# hat2.to_excel('C:/Users/shevc/Desktop/desk/result.xlsx', startrow=0, index=False)
