# ...
# INSTALL
# pip install openpyxl
# pip install PyQt5

# COMPILE
# pyinstaller -F -w main.py
# ...

import os
import sys
import time
import PyQt5
import PyQt5.QtWidgets
import PyQt5.QtCore
import PyQt5.QtGui
import openpyxl
import openpyxl.utils


# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # переменные, атрибуты
        self.info_for_open_file = ''
        self.info_path_open_file = ''

        self.info_extention_open_file_html = 'Файлы HTML (*.html)'
        self.info_extention_open_file_xls = 'Файлы Excel xlsx (*.xlsx)'

        self.text_empty_path_file = 'файл пока не выбран'

        # количество писем в одном пакете отправки, в штуках
        self.q_pocket = '5'
        # задержка между письмами в пакете при отправке, в секундах
        self.q_messages = '3'
        # задержка между отправками пакетов, в минутах
        self.send_delay = '5'

        # # начало диапазона поиска строк в обоих файлах
        # self.range_all_files = 'A2:'

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Рассылка почты из XLS файла на основе шаблона HTML')
        self.setGeometry(600, 200, 700, 450)

        # объекты на форме
        # HTML
        # label_html_file
        self.label_html_file = PyQt5.QtWidgets.QLabel(self)
        self.label_html_file.setObjectName('label_html_file')
        self.label_html_file.setText('1. Выберите HTML файл шаблона')
        self.label_html_file.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_html_file.setFont(font)
        self.label_html_file.adjustSize()
        self.label_html_file.setToolTip(self.label_html_file.objectName())

        # toolButton_select_html_file
        self.toolButton_select_html_file = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_html_file.setObjectName('toolButton_select_html_file')
        self.toolButton_select_html_file.setText('...')
        self.toolButton_select_html_file.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_html_file.setFixedWidth(50)
        self.toolButton_select_html_file.clicked.connect(self.select_file)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.toolButton_select_html_file.setToolTip(self.toolButton_select_html_file.objectName())

        # label_path_html_file
        self.label_path_html_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_html_file.setObjectName('label_path_html_file')
        self.label_path_html_file.setText(self.text_empty_path_file)
        self.label_path_html_file.setGeometry(PyQt5.QtCore.QRect(70, 40, 820, 16))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_html_file.setFont(font)
        self.label_path_html_file.adjustSize()
        self.label_path_html_file.setToolTip(self.label_path_html_file.objectName())

        # XLS
        # label_xls_file
        self.label_xls_file = PyQt5.QtWidgets.QLabel(self)
        self.label_xls_file.setObjectName('label_xls_file')
        self.label_xls_file.setText('2. Выберите EXCEL файл со справочником адресатов')
        self.label_xls_file.setGeometry(PyQt5.QtCore.QRect(10, 70, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_xls_file.setFont(font)
        self.label_xls_file.adjustSize()
        self.label_xls_file.setToolTip(self.label_xls_file.objectName())

        # toolButton_select_xls_file
        self.toolButton_select_xls_file = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_xls_file.setObjectName('toolButton_select_xls_file')
        self.toolButton_select_xls_file.setText('...')
        self.toolButton_select_xls_file.setGeometry(PyQt5.QtCore.QRect(10, 100, 50, 20))
        self.toolButton_select_xls_file.setFixedWidth(50)
        self.toolButton_select_xls_file.clicked.connect(self.select_file)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.toolButton_select_xls_file.setToolTip(self.toolButton_select_xls_file.objectName())

        # label_path_xls_file
        self.label_path_xls_file = PyQt5.QtWidgets.QLabel(self)
        self.label_path_xls_file.setObjectName('label_path_xls_file')
        self.label_path_xls_file.setText(self.text_empty_path_file)
        self.label_path_xls_file.setGeometry(PyQt5.QtCore.QRect(70, 100, 820, 20))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_xls_file.setFont(font)
        self.label_path_xls_file.adjustSize()
        self.label_path_xls_file.setToolTip(self.label_path_xls_file.objectName())

        # Q_POCKET
        # label_q_pocket
        self.label_q_pocket = PyQt5.QtWidgets.QLabel(self)
        self.label_q_pocket.setObjectName('label_q_pocket')
        self.label_q_pocket.setText('3. Сколько писем в одном пакете, шт.')
        self.label_q_pocket.setGeometry(PyQt5.QtCore.QRect(10, 130, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_q_pocket.setFont(font)
        self.label_q_pocket.adjustSize()
        self.label_q_pocket.setToolTip(self.label_q_pocket.objectName())

        # lineEdit_q_pocket
        self.lineEdit_q_pocket = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_q_pocket.setObjectName('lineEdit_q_pocket')
        self.lineEdit_q_pocket.setText(self.q_pocket)
        self.lineEdit_q_pocket.setGeometry(PyQt5.QtCore.QRect(10, 160, 90, 20))
        # self.lineEdit_q_pocket.setClearButtonEnabled(True)
        self.lineEdit_q_pocket.setEnabled(False)
        self.lineEdit_q_pocket.setToolTip(self.lineEdit_q_pocket.objectName())
        # self.lineEdit_q_pocket.textEdited.connect(self.check_digit)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        # 18
        gui is complete, ver
        1
        # Q_MESSAGES
        # label_q_messages
        self.label_q_messages = PyQt5.QtWidgets.QLabel(self)
        self.label_q_messages.setObjectName('label_q_messages')
        self.label_q_messages.setText('4. Задержка между письмами, сек.')
        self.label_q_messages.setGeometry(PyQt5.QtCore.QRect(10, 190, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_q_messages.setFont(font)
        self.label_q_messages.adjustSize()
        self.label_q_messages.setToolTip(self.label_q_messages.objectName())

        # lineEdit_q_messages
        self.lineEdit_q_messages = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_q_messages.setObjectName('lineEdit_q_messages')
        self.lineEdit_q_messages.setText(self.q_messages)
        self.lineEdit_q_messages.setGeometry(PyQt5.QtCore.QRect(10, 220, 90, 20))
        # self.lineEdit_q_messages.setClearButtonEnabled(True)
        self.lineEdit_q_messages.setEnabled(False)
        self.lineEdit_q_messages.setToolTip(self.lineEdit_q_pocket.objectName())
        # self.lineEdit_q_messages.textEdited.connect(self.check_digit)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        # MAIL_DELAY
        # label_mail_delay
        self.label_mail_delay = PyQt5.QtWidgets.QLabel(self)
        self.label_mail_delay.setObjectName('label_mail_delay')
        self.label_mail_delay.setText('5. Задержка между отправками пакетов, мин.')
        self.label_mail_delay.setGeometry(PyQt5.QtCore.QRect(10, 250, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_mail_delay.setFont(font)
        self.label_mail_delay.adjustSize()
        self.label_mail_delay.setToolTip(self.label_mail_delay.objectName())

        # lineEdit_mail_delay
        self.lineEdit_mail_delay = PyQt5.QtWidgets.QLineEdit(self)
        self.lineEdit_mail_delay.setObjectName('lineEdit_mail_delay')
        self.lineEdit_mail_delay.setText(self.send_delay)
        self.lineEdit_mail_delay.setGeometry(PyQt5.QtCore.QRect(10, 280, 90, 20))
        # self.lineEdit_q_messages.setClearButtonEnabled(True)
        self.lineEdit_mail_delay.setEnabled(False)
        self.lineEdit_mail_delay.setToolTip(self.lineEdit_mail_delay.objectName())
        # self.lineEdit_q_messages.textEdited.connect(self.check_digit)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        # SEND_MAIL
        # pushButton_send_mail
        self.pushButton_send_mail = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_send_mail.setObjectName('pushButton_send_mail')
        self.pushButton_send_mail.setEnabled(False)
        self.pushButton_send_mail.setText('Отправьте почту')
        self.pushButton_send_mail.setGeometry(PyQt5.QtCore.QRect(10, 310, 180, 25))
        self.pushButton_send_mail.setFixedWidth(130)
        self.pushButton_send_mail.clicked.connect(self.do_fill_data)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.pushButton_send_mail.setToolTip(self.pushButton_send_mail.objectName())

        # TEXT_STATISTICS
        # label_text_statistics
        self.label_text_statistics = PyQt5.QtWidgets.QLabel(self)
        self.label_text_statistics.setObjectName('label_text_statistics')
        self.label_text_statistics.setText('Статистика отправки:\n')
        self.label_text_statistics.setGeometry(PyQt5.QtCore.QRect(10, 340, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_text_statistics.setFont(font)
        self.label_text_statistics.adjustSize()
        self.label_text_statistics.setToolTip(self.label_text_statistics.objectName())

        # EXIT
        # button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 410, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file(self):
        # запоминание старого значения пути выбора файлов
        old_path_of_selected_html_file = self.label_path_html_file.text()
        old_path_of_selected_xls_file = self.label_path_xls_file.text()

        # определение какая кнопка выбора файла нажата
        if self.sender().objectName() == self.toolButton_select_html_file.objectName():
            self.info_for_open_file = 'Выберите HTML файл (.HTML)'
            # непосредственное окно выбора файла и переменная для хранения пути файла
            data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                                 self.info_for_open_file,
                                                                                 self.info_path_open_file,
                                                                                 self.info_extention_open_file_html)
        elif self.sender().objectName() == self.toolButton_select_xls_file.objectName():
            self.info_for_open_file = 'Выберите Excel (.XLSX)'
            # непосредственное окно выбора файла и переменная для хранения пути файла
            data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                                 self.info_for_open_file,
                                                                                 self.info_path_open_file,
                                                                                 self.info_extention_open_file_xls)

        # выбор только пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]

        # выбор где и что менять исходя из выбора пользователя
        # нажата кнопка выбора HTML файла
        if self.sender().objectName() == self.toolButton_select_html_file.objectName():
            if file_name == '':
                self.label_path_html_file.setText(old_path_of_selected_html_file)
                self.label_path_html_file.adjustSize()
            else:
                old_path_of_selected_html_file = self.label_path_html_file.text()
                self.label_path_html_file.setText(file_name)
                self.label_path_html_file.adjustSize()

        # нажата кнопка выбора XLS файла
        if self.sender().objectName() == self.toolButton_select_xls_file.objectName():
            if file_name == '':
                self.label_path_xls_file.setText(old_path_of_selected_xls_file)
                self.label_path_xls_file.adjustSize()
            else:
                old_path_of_selected_xls_file = self.label_path_xls_file.text()
                self.label_path_xls_file.setText(file_name)
                self.label_path_xls_file.adjustSize()

        # активация и деактивация объектов на форме зависящее от "выбраны ли все файлы" и "они разные"
        if self.text_empty_path_file not in (self.label_path_html_file.text(), self.label_path_xls_file.text()):
            self.pushButton_send_mail.setEnabled(True)

    # def check_digit(self, string_data):
    #     # проверка lineEdit_max_string на число
    #     flag_digit = None
    #
    #     if string_data.isdigit():
    #         flag_digit = True
    #     elif string_data == '':
    #         flag_digit = False
    #         self.lineEdit_max_string.setText('0')
    #     else:
    #         flag_digit = False
    #
    #         # информационное окно - "введите число"
    #         self.window_info = PyQt5.QtWidgets.QMessageBox()
    #         self.window_info.setWindowTitle('Только цифры')
    #         self.window_info.setText(f'Вводите только цифры!')
    #         self.window_info.exec_()
    #
    #         self.lineEdit_max_string.setText('0')
    #
    #     return flag_digit

    # событие - нажатие на кнопку заполнения файла
    def do_fill_data(self):
        pass

        # # выбор выбранных строк в списке специальностей
        # specialization_selected = [item.text() for item in self.listWidget_specialization.selectedItems()]
        #
        # # проверка на количество выбранных строк в listWidget_specialization
        # if len(specialization_selected) == 0:
        #     # информационное окно - "выберите специальность"
        #     self.window_info = PyQt5.QtWidgets.QMessageBox()
        #     self.window_info.setWindowTitle('Выберите специальности')
        #     self.window_info.setText(f'В списке специальностей ничего не выбрано,\n'
        #                              f'выберите хотя бы одну строку')
        #     self.window_info.exec_()
        # else:
        #     # считаю время заполнения
        #     time_start = time.time()
        #
        #     # открыть файл Полный и Неполный, и выбрать листы
        #     wb_full = openpyxl.load_workbook(self.label_path_full_file.text())
        #     wb_full_s = wb_full.active
        #     wb_half = openpyxl.load_workbook(self.label_path_half_file.text())
        #     wb_half_s = wb_half.active
        #
        #     # сформированные диапазоны обработки
        #     range_full_file = self.range_all_files + wb_full_s.cell(wb_full_s.max_row, wb_full_s.max_column).coordinate
        #     range_half_file = self.range_all_files + wb_half_s.cell(wb_half_s.max_row, wb_half_s.max_column).coordinate
        #     wb_full_range = wb_full_s[range_full_file]
        #     wb_half_range = wb_half_s[range_half_file]
        #
        #     # список одной строки прохода, список выбранных строк по специальностям, списки всех строк Неполного файла
        #     list_one_string = []  # временная переменная
        #     list_half_file = []  # весь Неполный файл
        #     list_filtered_string = []  # фильтрованные строки из Полного которые устраивают выбранным специальностям
        #     list_for_add = []  # список выбранных из фильтрованных для добавления в Неполный файл
        #     tuple_half_file = ()  # кортеж для хранения ФИО из Неполного файла
        #
        #     # счётчик удачных добавлений в Неполный из выбранных строк
        #     count_add_success = 0
        #
        #     # заполнение list_half_file Неполного файла
        #     for row_in_range_half in wb_half_range:
        #         # чищу список для временной строки
        #         list_one_string = []
        #
        #         # прохожу строку
        #         for cell_in_row_half in row_in_range_half:
        #             list_one_string.append(cell_in_row_half.value)
        #
        #         # все записи из Неполного файла
        #         list_half_file.append(list_one_string)
        #
        #     # количество строк "сколько хочу строк" (перевод значения в поле шага 3)
        #     count_string_want = int(self.lineEdit_max_string.text())
        #
        #     # количество строк в Неполном файле (-1 потому что верхняя строка это шапка)
        #     count_string_half = wb_half_s.max_row - 1
        #
        #     # сколько нужно добавить строк в Неполный файл, должно быть больше нуля
        #     count_string_add = count_string_want - count_string_half
        #
        #     # количество строк в отфильтрованном списке
        #     count_filter_string = len(list_filtered_string)
        #
        #     # количество строк которых будет реально добавлены в Неполный файл
        #     count_real_data_add = count_filter_string - count_string_add
        #
        #     # добавление строк в Неполный файл
        #     # если количество строк в Неполном меньше, чем хочется, то добавить разницу строк
        #     if count_string_add <= 0:
        #         # информационное окно - ""
        #         self.window_info = PyQt5.QtWidgets.QMessageBox()
        #         self.window_info.setWindowTitle('Строки')
        #         self.window_info.setText(f'Количество строк в Неполном файле\n'
        #                                  f'одинаково или больше,\n'
        #                                  f'чем число в ПУНКТЕ 3'
        #                                  # f' \n'
        #                                  # f'их разница равна {count_string_add}\n'
        #                                  # f'хочется чтобы было {count_string_want}\n'
        #                                  # f'сейчас в файле {count_string_half}\n'
        #                                  # f'надо добавить {count_string_add}\n'
        #                                  # f'могу выбрать из {count_filter_string}'
        #                                  )
        #         self.window_info.exec_()
        #     else:
        #         if count_string_add > count_filter_string:
        #             # если добавляемых больше, чем отфильтрованных, то добавлять всё из list_filtered_string
        #             # информационное окно
        #             self.window_info = PyQt5.QtWidgets.QMessageBox()
        #             self.window_info.setWindowTitle('Строки')
        #             self.window_info.setText(f'Количество строк в Полном файле по этим специальностям\n'
        #                                      f'меньше, чем число в ПУНКТЕ 3,\n'
        #                                      f'выберите ещё специальностей из списка\n'
        #                                      # f' \n'
        #                                      # f'их разница равна {count_real_data_add}\n'
        #                                      # f'хочется чтобы было {count_string_want}\n'
        #                                      # f'сейчас в файле {count_string_half}\n'
        #                                      # f'надо добавить {count_string_add}\n'
        #                                      # f'могу выбрать из {count_filter_string}'
        #                                      )
        #             self.window_info.exec_()
        #         else:
        #             # последняя строка в Неполном +2 потому, что один за прошлый вычет, а один на следующую строчку
        #             string_half_begin = (count_string_half + 1) + 1
        #             string_half_end = (count_string_half + 1) + len(list_for_add)
        #
        #             # добавление данных в эксель
        #             for string_list_for_add in list_for_add:
        #                 wb_half_s.append(string_list_for_add)
        #
        #             # сохраняю файл и закрываю оба
        #             filename_half = os.path.split(self.label_path_half_file.text())[1]
        #             wb_half.save(filename_half)
        #             wb_full.close()
        #             wb_half.close()
        #
        #             # считаю время заполнения
        #             time_finish = time.time()
        #
        #             # информационное окно о сохранении файлов
        #             self.window_info = PyQt5.QtWidgets.QMessageBox()
        #             self.window_info.setWindowTitle('Файлы')
        #             self.window_info.setText(f'Файлы сохранены и закрыты.\n'
        #                                      f'Заполнение сделано за {round(time_finish - time_start, 1)} секунд')
        #             self.window_info.exec_()

    # событие - нажатие на кнопку Выход
    @staticmethod
    def click_on_btn_exit():
        exit()


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    app_window_main = Window()
    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()
