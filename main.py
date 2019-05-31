from setup import *
import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
import design  # Это наш конвертированный файл дизайна
import itertools
from PyQt5.QtGui import QIcon


class ExampleApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.setWindowIcon(QIcon("cryp.png"))
        self.pushButton_3.clicked.connect(self.MyFun)  # обработка кнопки XOR
        self.pushButton_4.clicked.connect(self.MyFun_1)  # обработка кнопки Подстановка
        # self.pushButton_2.clicked.connect(self.textEdit_2.setPlainText())  # обработка кнопки DeCrypt
        # self.pushButton.clicked.connect(self.MyFun)  #обработка кнопки EnCrypt
        self.pushButton.setCheckable(True)

    def MyFun(self):
        mytext = self.textEdit.toPlainText()
        cryptKey = self.textEdit_3.toPlainText()
        self.textEdit_2.setPlainText(XOR_cipher(mytext, cryptKey))  # вывод шифрованного текста

    def MyFun_1(self):
        mytext = self.textEdit.toPlainText()
        self.textEdit_2.setPlainText(replacer(mytext, self.pushButton.isChecked()))  # вывод шифрованного текста


def XOR_cipher(string, key):
    answer = []

    key = itertools.cycle(key)  # Повторяем ключ, чтобы зашифровать всю строку

    for s, k in zip(string, key):
        answer.append(chr(ord(s) ^ ord(k)))

    return ''.join(answer)

    # Функция для расшифровки точно такая же


XOR_uncipher = XOR_cipher


def replacer(string, boolean):
    textEng = 'CDABHIJEFGOPQRKLMNUVW:STZ XY;?-.,!'
    crypt = 'ЮЯЫЭЬЪШЩЦЧФХТУРСОПМНКЛ ЙЖЗДЕВГАБЁИ'  # Исходные данные таблицы подстановки / по варианту

    if boolean:
        mapping = str.maketrans(dict(zip(textEng, crypt)))
    else:
        mapping = str.maketrans(dict(zip(crypt, textEng)))
    rez = string.translate(mapping)  # 'тvоyа sтrока'
    return rez


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.setWindowIcon(QIcon('1.png'))
    window.show()  # Показываем окно

    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()