import math

from setup import *
import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
import design_new  # Это наш конвертированный файл дизайна
import about
import itertools
from PyQt5.QtGui import QIcon


class ExampleApp(QtWidgets.QMainWindow, design_new.Ui_MainWindow):
    def __init__(self, parent = None):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design_new.py
        super().__init__(parent)
        self.setupUi(self)# Это нужно для инициализации нашего дизайна
        self.twoWindow = None
        # о программе
        self.action_12.triggered.connect(self.check)

        self.setWindowIcon(QIcon("1.png"))

        self.pushButton.clicked.connect(self.browse_file)  # обработка кнопки "Выбор файла"
        self.pushButton_2.clicked.connect(self.view_all_data)  # обработка кнопки "Просмотр всех исторических данных в файле"
        self.pushButton_3.clicked.connect(self.data_interval_research)  # обработка кнопки "Открыть. Интервал_исследования"
        self.pushButton_4.clicked.connect(self.temp_interval_research)  # обработка кнопки "Открыть. Интервал_температур"
        self.pushButton_5.clicked.connect(self.search_by_date)  # обработка кнопки "Поиск"
        self.pushButton_6.clicked.connect(self.math_analysis)  # обработка кнопки "Выполнить"
        self.pushButton_7.clicked.connect(self.save_file)  # обработка кнопки "Записать в файл"
        self.pushButton_8.clicked.connect(self.weather)  # обработка кнопки "Обновить"
        # self.pushButton_2.clicked.connect(self.textEdit_2.setPlainText())  # обработка кнопки
        # self.pushButton.clicked.connect(self.MyFun)  #обработка кнопки
        # self.pushButton.setCheckable(True)

    # метод инициализации окна "о программе"
    def check(self):
        # убрать комментарий если требуется автоматически закрыть предыдущее окно
        #self.close()
        self.twoWindow = TwoWindow()
        self.twoWindow.show()

    # функция выбора файла в .xls формате, возвращяет путь файла
    def browse_file(self):
        global path_file
        file = QtWidgets.QFileDialog.getOpenFileName(self,
                                             "Выберите файл",
                                             "/home","Document (*.xls)")
        # строго ограничен формат

        #path = QtWidgets.QFileDialog.getSaveFileUrl(self, "Выберите файл") СОХРАНИТЬ
        self.textBrowser.append("Файл выбран!\nРасположене выбранного файла: " + file[0])
        path_file = file[0]

    #Просмотр всех исторических данных в файле
    def view_all_data(self):
        try:
            open_file_8(path_file)
            self.textBrowser.append("Выбранный файл доступен к просмотру!")
            self.textBrowser.append("Файл " + path_file + " запущен!")
            self.textBrowser.append("Работа с файлом " + path_file + " закончена!")
        except NameError:
            self.textBrowser.append("Сначала выберите файл!")
            #date = self.dateEdit.text() # получаем значение даты из соответствующего поля
            # min_temp_field = self.doubleSpinBox.text()  # Фильтр по температуре воды "От"
            # max_temp_field = self.doubleSpinBox_2.text()  # Фильтр по температуре воды "До"
            # print(date, min_temp_field, max_temp_field)
            # self.textBrowser.append("TEST")
            #mytext = self.textEdit.toPlainText()
            #self.textEdit_2.setPlainText(replacer(mytext, self.pushButton.isChecked()))


    # обработка кнопки "Открыть. Интервал_исследования"
    def data_interval_research(self):
        try:
            min_data_field = self.dateEdit_3.text() # получаем значение даты из соответствующего поля
            max_data_field = self.dateEdit_4.text()  # получаем значение даты из соответствующего поля
            search_data_interval_3(path_file, min_data_field, max_data_field)
            self.textBrowser.append("Создан файл " + 'data_interval_' + min_data_field + ' -' + max_data_field + '.xls')
        except NameError:
            self.textBrowser.append("Сначала выберите файл в котором будет производится выборка!")

    # обработка кнопки "Открыть. Интервал_температур"
    def temp_interval_research(self):
        try:
            min_temp_field = self.doubleSpinBox.text().replace(',', '.')  # Фильтр по температуре воды "От"
            max_temp_field = self.doubleSpinBox_2.text().replace(',', '.')  # Фильтр по температуре воды "До"
            filter_temp_water_4(path_file, min_temp_field, max_temp_field)
            self.textBrowser.append("Создан файл " + 'temp_interval_' + min_temp_field + ' -' + max_temp_field + '.xls')
        except NameError:
            self.textBrowser.append("Сначала выберите файл в котором будет производится выборка!")


    # Поиск исследований по дате
    def search_by_date(self):
        try:
            date = self.dateEdit.text()  # получаем значение даты из соответствующего поля
            self.textBrowser.append("Выполняется поиск исследований по дате: " + date)
            # открываем файл
            rb = xlrd.open_workbook(path_file, formatting_info=True)
            # выбираем активный лист
            sheet = rb.sheet_by_index(0)
            flag = 0
            # получаем значения всех записей таблицы
            for i in range(sheet.nrows):
                val = sheet.row_values(i)
                if date == val[0]:
                    flag = 1
                    self.textBrowser.append('Запись найдена!')
                    self.textBrowser.append(val[0] + " температура воды равнялась " + str(val[1]) + " градусам.")
            if flag == 0: self.textBrowser.append("Запись отсутствует, недостаточно исторических данных!")

        except NameError:
            self.textBrowser.append("Сначала выберите файл в котором будет производится поиск!")


    # функция анализа и прогнозирования будующих состояний системы (любой день года)
    # в зависимости от входных данных(несколько источников разных направлений) может кореллироваться точность прогноза
    def math_analysis(self):
        try:
            # получаем значение даты из соответствующего поля
            date = self.dateEdit_2.text()  # дата для которой требуется вычислить вероятное будующее состояние
            # date.strftime('%d.%m') для поиска и сравнения будет использоваться только 2 значения (день и месяц)
            #print(date)
            # преоразуем строку в дату, форматируем дату (2 числа)
            piece_date = dt.strptime(date, "%d.%m.%Y").strftime('%d.%m')
            self.textBrowser.append("В архиве выполняется поиск исследований которые были проведены: " + piece_date) # ищем все совпадения день/месяц. Значение температуры в список.
            # открываем файл
            rb = xlrd.open_workbook(path_file, formatting_info=True)
            # выбираем активный лист
            sheet = rb.sheet_by_index(0)
            flag = 0
            count_research = 0
            sum_temp = 0
            min_temp = 100.0
            max_temp = 0
            # получаем значения всех записей таблицы
            for i in range(sheet.nrows):
                if i==0: continue # пропускаем строку с наименованием столбцов
                val = sheet.row_values(i)
                h_piece_date = dt.strptime(val[0], "%d.%m.%Y").strftime('%d.%m')
                if piece_date == h_piece_date: # сравниваем заданный день и месяц с полями всех годов
                    flag = 1
                    count_research+=1
                    print(val[1])
                    temp = float(val[1])
                    #average_temperature
                    sum_temp += temp
                    #val[1].replace(',', '.')
                    # выводится в терминал программы каждая запись
                    #self.textBrowser.append(val[0] + " температура воды равнялась " + str(val[1]) + " градусам.")
                    # поиск min/max значений температуры
                    if temp < min_temp:
                        min_temp = temp
                    if temp > max_temp:
                        max_temp = temp
            #print(min_temp, max_temp)



            if flag == 0:
                self.textBrowser.append("Записи за это число не найдены, недостаточно исторических данных!")
                return
            # average_temperature
            avg_temp = sum_temp / count_research # средняя температура за все года
            self.textBrowser.append("Выборка произведена из "+ str(count_research) +" значений.")
            self.textBrowser.append("Температура воды: " + date + " будет равняться "+ str("%.2f" %avg_temp) + " градусам!")


            # Мат часть
            # Среднее арифметическое выборки
                # avg_temp

            # Мотод Корнфельда для нахождения Абсолютной погрешности: Δ=(max-min)/2

            absolute_error = (max_temp - min_temp) / 2
            print("Абсолютная погрешность "+ str("%.2f" %absolute_error) + " градуса")
            self.textBrowser.append("Абсолютная погрешность "+ str("%.2f" %absolute_error) + " градуса")

            # Относительная погрешность = (Δ/Среднее арифметическое выборки)*100%
            relative_error = (absolute_error/avg_temp)*100
            print("Относительная погрешность " + str("%.2f" % relative_error) + "%")
            self.textBrowser.append("Относительная погрешность " + str("%.2f" % relative_error) + "%")

            # Вычислим квадраты отклонений температуры от их среднего значения:
                # при каждой итерации:
                # квадрат отклонения температуры от среднего значения = (температура - среднее значение температуры)**2
                # квадрат отклонения += квадрат отклонения температуры от среднего значения
                # count ++
            count_research_2 = 0
            square_deviation_temperature = 0
            for i in range(sheet.nrows):
                if i==0: continue # пропускаем строку с наименованием столбцов
                val = sheet.row_values(i)
                h_piece_date = dt.strptime(val[0], "%d.%m.%Y").strftime('%d.%m')
                if piece_date == h_piece_date: # сравниваем заданный день и месяц с полями всех годов
                    count_research_2+=1
                    #print(val[1])
                    temp = float(val[1])
                    # квадрат отклонения температуры от среднего значения += (температура - среднее значение температуры)**2
                    square_deviation_temperature += (temp-avg_temp)**2
                    # Дисперсия = суммируем (+=) квадраты отклонения температур и делим на n-1 (count_research_2-1)

            # Среднее арифметическое квадратов отклонения называется дисперсией
            dispersion = square_deviation_temperature/(count_research_2-1)
            print("Дисперсия = " + str("%.2f" %dispersion))
            self.textBrowser.append("Дисперсия = " + str("%.2f" %dispersion) + " (cреднее арифметическое квадратов отклонения температур)")
            # Среднеквадратическое отклонение! Определяется как квадратный корень из дисперсии случайной величины
            rms_deviation = math.sqrt(dispersion) # Среднеквадратическое отклонение
            print("Среднеквадратическое(стандартное) отклонение = " + str("%.2f" %rms_deviation) + " градуса (квадратный корень дисперсии)")
            self.textBrowser.append("Среднеквадратическое(стандартное) отклонение = " + str("%.2f" %rms_deviation) + " градуса (квадратный корень дисперсии)")
            # Результат называется стандартным отклонением на основании несмещённой оценки дисперсии. Деление на n − 1 вместо n даёт неискажённую оценку дисперсии для больших генеральных совокупностей.

            # Коэффициент вариации / степерь рассеивания данных (незначительная до 10%/средняя от 10% до 20%/значительная от 20% до 33%/ до 33% совокупность однородная, больше 33 - неоднородная)
            coefficient_variation = (rms_deviation/avg_temp)*100
            print("Коэффициент вариации = " + str("%.2f" % coefficient_variation)+ "%")
            self.textBrowser.append("Коэффициент вариации = " + str("%.2f" % coefficient_variation)+ "%")
            if coefficient_variation < 10:
                self.textBrowser.append("Cтеперь рассеивания данных незначительная до 10%, совокупность однородная")
            if coefficient_variation >= 10 and coefficient_variation <= 20:
                self.textBrowser.append("Cтеперь рассеивания данных средняя от 10% до 20%, совокупность однородная")
            if coefficient_variation > 20:
                self.textBrowser.append("Cтеперь рассеивания данных значительная от 20%, совокупность неоднородная")
            # Среднеквадратическая погрешность среднего значения Х
            rms_deviation_avg_temp = rms_deviation/math.sqrt(count_research_2) # Среднеквадратическая погрешность/корень n
            print("Стандартная ошибка средней = " + str("%.2f" % rms_deviation_avg_temp) + " градуса")
            self.textBrowser.append("Стандартная ошибка средней = " + str("%.2f" % rms_deviation_avg_temp))
            # Сохранение в excel файл
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Проноз')
            ws.write(0, 0, "Проноз на "+ str(date))
            ws.write(1, 0, "Выборка произведена из "+ str(count_research) +" значений")
            ws.write(2, 0, "Температура воды: " + date + " будет равняться "+ str("%.2f" %avg_temp) + " градусам!")
            ws.write(3, 0, "Абсолютная погрешность "+ str("%.2f" %absolute_error) + " градуса")
            ws.write(4, 0, "Относительная погрешность " + str("%.2f" % relative_error) + "%")
            ws.write(5, 0, "Дисперсия = " + str("%.2f" %dispersion) + " (cреднее арифметическое квадратов отклонения температур)")
            ws.write(6, 0, "Среднеквадратическое(стандартное) отклонение = " + str("%.2f" %rms_deviation) + " градуса (квадратный корень дисперсии)")
            ws.write(7, 0, "Коэффициент вариации = " + str("%.2f" % coefficient_variation)+ "%")
            ws.write(8, 0, "Cтеперь рассеивания данных незначительная до 10%, совокупность однородная")
            ws.write(9, 0, "Стандартная ошибка средней = " + str("%.2f" % rms_deviation_avg_temp))
            # сохраем рабочию книгу
            new_file = 'Проноз_на_' + date + '.xls'
            wb.save(new_file)

        except NameError:
            self.textBrowser.append("Сначала выберите файл исследований!")


    # метод записи прогнозов в файл
    # для каждого прогноза создается свой файл с соответствующим наименоваием
    def save_file(self):
        # Создаем новый excel файл
        path = QtWidgets.QFileDialog.getSaveFileUrl(self, "Выберите файл")
        date_2 = self.dateEdit_2.text()
        # сохраем рабочию книгу
        new_file = 'Проноз_на_' + date_2 + '.xls'
        # открываем автоматически созданный файл
        open_file_8(new_file)
        self.textBrowser.append("Файл " + new_file + " запущен!")
        self.textBrowser.append("Работа с файлом " + new_file + " закончена!")

    def weather(self):
        url = "http://api.openweathermap.org/data/2.5/weather"
        city = "Sevastopol"
        water_temp = 20.72  # str(data["water"]["temp"])# + "'С")
        parameters = {
            'q': city,
            'appid': "778d98cf94b6609bec655b872f24b907",
            'units': 'metric',
            'lang': 'ru'
        }
        res = requests.get(url, params=parameters)
        data = res.json()
        self.textBrowser_4.clear()
        self.textBrowser_4.append("Город: " + data["name"])
        self.textBrowser_4.append("Состояние: " + data["weather"][0]["description"])
        self.textBrowser_4.append("Текущая температура: " + str(data["main"]["temp"]) + "'С")
        self.textBrowser_4.append("Скорость ветра: " + str(data["wind"]["speed"]) + " м/с")
        self.textBrowser_4.append("Температура воды: " + str(water_temp) + "'С")
        self.textBrowser_4.append("_______________________\n##################\n_______________________")
        self.textBrowser_4.append("╔╗ ╔═ ╗╔ ╔╗ ╔╗ ╦ ╔╗ ╔╗ ╔╗ ║\n╚╗ ╠═ ║║ ╠╣ ╚╗ ║ ║║ ║║ ║║ ║\n╚╝ ╚═ ╚╝ ║║ ╚╝ ║ ╚╝ ╠╝ ╚╝ ╚")
        self.textBrowser.append("Информация о текущей погоде обновлена!")
        #mytext = self.textEdit.toPlainText()
        #self.textEdit_2.setPlainText(replacer(mytext, self.pushButton.isChecked()))  # вывод шифрованного текста

# инициализация 2го окна
class TwoWindow(QtWidgets.QMainWindow, about.Ui_Form):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("1.png"))
        self.setupUi(self)
        # self.pushButton.clicked.connect(self.check2)


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.setWindowIcon(QIcon('1.png'))
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()

