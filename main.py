# Import pandas
import os.path
import pandas as pd
import pylab
# Assign spreadsheet filename to `file`
from pandas.tests.groupby.test_value_counts import df
from matplotlib import pyplot as plt

# path='D:\dk\python\covid_17.11.20.xlsx'
path = input("Введите путь ")
sh_nm = input("Введите название листа ")
xl = pd.read_excel(path, sheet_name=sh_nm)
err = 0

while err != 1:
    if (not os.path.exists(path)) and not (path == "[0-9][a-z].xlsx"):
        print("Введите корректный путь")
    else:
        err = 1


def data_about_table():
    print("названия столбцов таблицы\n", xl.columns)
    print("типы столбцов \n", xl.dtypes)
    print("число строк в таблице ", xl.shape[0])
    print("число строк в строке ", xl.shape[1])


def data_about_region(region):
    print(xl[xl['ГОРОД/РЕГИОН '] == region])


def more_number(numb):
    print(xl[xl['ЧИСЛО ВЫЗДОРОВЕВШИХ'] > numb])


def top5(str0,str1, ft):

    print(str0,xl.sort_values(by=[str1], ascending=ft)[:5])


def histogram(clm,color,tf):
    new_xl = xl.sort_values(by=[clm], ascending=tf)[:5]
    plt.barh(new_xl["ГОРОД/РЕГИОН "], new_xl[clm], color=color)


data_about_table()
reg = input("Введите регион ")
data_about_region(reg)
ent_number = input("Введите число для сравнения по выздоровевшим ")
if ent_number.isdigit():
    more_number(int(ent_number))
else:
    print("Введено не число")

# df.to_excel("D:\dk\python\covid_17.11.20.xlsx", sheet_name='Лист 1')
top5("топ-5 регионов с максимальным числом заболевших\n", 'ЧИСЛО ЗАБОЛЕВШИХ', False)
top5("топ-5 регионов с максимальным числом выздоровевших\n", 'ЧИСЛО ВЫЗДОРОВЕВШИХ', False)
top5("топ-5 регионов с максимальным числом умерших\n", 'ЧИСЛО УМЕРШИХ', False)
top5("топ-5 регионов с минимальным числом заболевших\n", 'ЧИСЛО ЗАБОЛЕВШИХ', True)
top5("топ-5 регионов с минимальным числом выздоровевших\n", 'ЧИСЛО ВЫЗДОРОВЕВШИХ', True)
top5("топ-5 регионов с минимальным числом умерших\n", 'ЧИСЛО УМЕРШИХ', True)


plt.subplot(6,1,1)
histogram("ЧИСЛО ЗАБОЛЕВШИХ","red",True)
plt.title("ТОП-5 по наибольшему числу заболевших",size=10)
plt.subplot(6,1,2)
histogram("ЧИСЛО ЗАБОЛЕВШИХ","red",False)
plt.title("ТОП-5 по наименьшему числу заболевших",size=10)
plt.subplot(6,1,3)
histogram("ЧИСЛО ВЫЗДОРОВЕВШИХ","green",True)
plt.title("ТОП-5 по наибольшему числу выздоровевших",size=10)
plt.subplot(6,1,4)
histogram("ЧИСЛО ВЫЗДОРОВЕВШИХ","green",False)
plt.title("ТОП-5 по наименьшему числу выздоровевших",size=10)
plt.subplot(6,1,5)
histogram("ЧИСЛО УМЕРШИХ","black",True)
plt.title("ТОП-5 по наибольшей смертности",size=10)
plt.subplot(6,1,6)
histogram("ЧИСЛО УМЕРШИХ","black",False)
plt.title("ТОП-5 по наименьшей смертности",size=10)
plt.show()
