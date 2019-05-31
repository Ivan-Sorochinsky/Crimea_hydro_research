import requests
import xlwt
import xlrd
# import openpyxl 2007
import subprocess
from datetime import datetime as dt

name_file = 'water_temp.xls'


# Просмотр всех исторических данных в файле
# (открыть документ excel со всеми записями)
def open_file_8(name_file):
    print("Файл " + name_file + " запущен!")
    subprocess.call(name_file, shell=True)
    print("Работа с файлом " + name_file + " закончена!")


# Поиск исследований по дате
def search_2(name_file, data_field):
    # открываем файл
    rb = xlrd.open_workbook(name_file, formatting_info=True)
    # выбираем активный лист
    sheet = rb.sheet_by_index(0)
    flag = 0
    # получаем значения всех записей таблицы
    for i in range(sheet.nrows):
        val = sheet.row_values(i)
        if data_field == val[0]:
            flag = 1
            print('Запись найдена!')
            print(val[0] + " температура воды равнялась " + str(val[1]) + " градусам.")
    if flag == 0: print("Запись отсутствует, недостаточно исторических данных!")


#
def search_data_interval_3(name_file, min_data_field, max_data_field):
    # открываем файл
    rb = xlrd.open_workbook(name_file, formatting_info=True)
    # выбираем активный лист
    sheet = rb.sheet_by_index(0)

    min = dt.strptime(min_data_field, "%d.%m.%Y")
    max = dt.strptime(max_data_field, "%d.%m.%Y")

    # Создаем новый excel файл
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Выборка_по_дате')
    j = 0  # данную переменную используем для предотвращения избыточности, новый файл будет заполнятся по порядку
    # получаем значения всех записей таблицы
    for i in range(sheet.nrows):
        if i == 0: continue  # пропускаем первою строку (текст)
        val = sheet.row_values(i)
        v = dt.strptime(val[0], "%d.%m.%Y")
        if min <= v and v <= max:
            print(val)  # результат выборки записывам в файл и открываем
            ws.write(j, 0, val[0])  # столбец А
            ws.write(j, 1, val[1])  # столбец B
            j += 1
    # сохраем рабочию книгу
    new_file = 'data_interval_' + min_data_field + ' -' + max_data_field + '.xls'
    wb.save(new_file)
    # открываем автоматически созданный файл
    open_file_8(new_file)


def filter_temp_water_4(name_file, min_temp_field, max_temp_field):
    # открываем файл
    rb = xlrd.open_workbook(name_file, formatting_info=True)
    # выбираем активный лист
    sheet = rb.sheet_by_index(0)

    min = float(min_temp_field)
    max = float(max_temp_field)

    # Создаем новый excel файл
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Выборка_по_температуре')
    j = 0  # данную переменную используем для предотвращения избыточности, новый файл будет заполнятся по порядку
    # получаем значения всех записей таблицы
    for i in range(sheet.nrows):
        if i == 0: continue  # пропускаем первою строку (текст)
        val = sheet.row_values(i)
        v = float(val[1])
        if min <= v and v <= max:
            print(val)  # результат выборки записывам в файл и открываем
            ws.write(j, 0, val[0])  # столбец А
            ws.write(j, 1, val[1])  # столбец B
            j += 1
    # сохраем рабочию книгу
    new_file = 'temp_interval_' + min_temp_field + ' -' + max_temp_field + '.xls'
    wb.save(new_file)
    # открываем автоматически созданный файл
    open_file_8(new_file)


def weather():
    url = "http://api.openweathermap.org/data/2.5/weather"
    city = "Sevastopol"
    water_temp = 18.65  # str(data["water"]["temp"])# + "'С")
    parameters = {
        'q': city,
        'appid': "778d98cf94b6609bec655b872f24b907",
        'units': 'metric',
        'lang': 'ru'
    }
    res = requests.get(url, params=parameters)
    data = res.json()
    print("Город: " + data["name"])
    print("Состояние: " + data["weather"][0]["description"])
    print("Текущая температура: " + str(data["main"]["temp"]) + "'С")
    print("Скорость ветра: " + str(data["wind"]["speed"]) + " м/с")
    print("Температура воды: " + str(water_temp) + "'С")


def read_file(name_file):
    # открываем файл
    rb = xlrd.open_workbook(name_file, formatting_info=True)

    # выбираем активный лист
    sheet = rb.sheet_by_index(0)

    # получаем значения всех записей таблицы
    for i in range(sheet.nrows):
        val = sheet.row_values(i)
        print(val)
    """
    # Для XLSX формата с 2007 года
    #открываем файл
    wb = openpyxl.load_workbook(filename = 'water_temp.xlsx')
    sheet_obj = wb.active #Выбираем активный лист таблицы
    m_row = sheet_obj.max_row

    # Выводим значения в цикле
    for cellObj in sheet_obj['A1':'B13']:
          for cell in cellObj:
                  print(cell.value)
          print('------')
    """


# open_file_8('water_temp.xls')
# search_2('water_temp.xls','05.01.2018')
# search_data_interval_3('water_temp.xls', '10.05.2018', '10.10.2018')
filter_temp_water_4('water_temp.xls', '10.05', '19.40')