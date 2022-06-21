# TODO
# сделать отправку данных
# сделать кнопку проверки отправки тестового письма
# сделать прогресс-бар по отправке, считать количество или время?

# ...
# INSTALL
# pip install openpyxl
# pip install PyQt5

# COMPILE
# pyinstaller -F -w main.py
# ...

# import os
import sys
import time
import PyQt5
import PyQt5.QtWidgets
import PyQt5.QtCore
import PyQt5.QtGui
import openpyxl
import openpyxl.utils


# класс получателя сообщения одной отправки
class RecipientData:
    """Класс Получателя сообщения"""

    count_Recipient = 0
    # class_suffix = 'Recipient'

    def __init__(self, rd_text_message=None):
        self.num = None
        self.fam = None
        self.im = None
        self.otch = None
        self.email = None
        self.mno_code = None
        self.text_message = rd_text_message
        self.flag_send_message = False

        RecipientData.count_Recipient += 1

    def get_all_info(self):
        # print(*self.__dict__.items())
        return f'Объект {self.get_obj_name()}, ' \
               f'{self.num = }, ' \
               f'{self.fam = }, ' \
               f'{self.im = }, ' \
               f'{self.otch = }, ' \
               f'{self.email = }, ' \
               f'{self.mno_code = }, ' \
               f'{self.text_message = }, ' \
               f'{self.flag_send_message = }'

    def get_obj_name(self):
        for glob_name, glob_val in globals().items():
            if glob_val is self:
                return glob_name

    @staticmethod
    def replace_text_message(mail_tag, tag_value, mail_text):
        mail_text = mail_text.replace('{{'+mail_tag+'}}', tag_value)
        return mail_text

    def __setattr__(self, key, value):
        if 'Recipient' in str(self.get_obj_name()):
            if key in ('num', 'fam', 'im', 'otch', 'mno_code'):
                # print('-' * 50)
                # print(f'{self.get_obj_name() = } ... {key = } ... {value = }')
                #
                # print(f' ... {self.text_message = }')
                self.text_message = self.replace_text_message(key, value, self.text_message)
                # print(f' ... {self.text_message = }')
                #
                # print()
            else:
                pass
                # print(f' ... {key = }')
        return object.__setattr__(self, key, value)


# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    """Класс главного окна"""

    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # переменные, атрибуты
        self.window_info = None
        self.info_for_open_file = ''
        self.info_path_open_file = ''

        self.info_extention_open_file_html = 'Файлы HTML (*.html; *.htm)'
        self.info_extention_open_file_xls = 'Файлы Excel xlsx (*.xlsx)'
        self.text_empty_path_file = 'файл пока не выбран'

        # количество писем в одном пакете отправки, в штуках
        self.q_pocket = '5'
        # задержка между письмами в пакете при отправке, в секундах
        self.q_messages = '3'
        # задержка между отправками пакетов, в секундах
        self.send_delay = '300'  # 5 минут

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Рассылка почты из XLS файла на основе шаблона HTML')
        self.setGeometry(600, 200, 700, 450)

        # ОБЪЕКТЫ НА ФОРМЕ
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
        self.toolButton_select_html_file.clicked.connect(self.select_file)
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
        self.toolButton_select_xls_file.clicked.connect(self.select_file)
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

        # SEND_MAIL
        # pushButton_send_mail
        self.pushButton_send_mail = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_send_mail.setObjectName('pushButton_send_mail')
        self.pushButton_send_mail.setEnabled(False)
        self.pushButton_send_mail.setText('Отправьте почту')
        self.pushButton_send_mail.setGeometry(PyQt5.QtCore.QRect(10, 310, 180, 25))
        self.pushButton_send_mail.setFixedWidth(130)
        self.pushButton_send_mail.clicked.connect(self.send_mail)
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
        self.button_exit.clicked.connect(self.click_on_btn_exit)
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file(self):
        data_of_open_file_name = None
        # запоминание старого значения пути выбора файлов
        old_path_of_selected_html_file = self.label_path_html_file.text()
        old_path_of_selected_xls_file = self.label_path_xls_file.text()

        # определение какая кнопка выбора файла нажата
        if self.sender().objectName() == self.toolButton_select_html_file.objectName():
            self.info_for_open_file = 'Выберите HTML файл (.HTML или .HTM)'
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

    # событие - нажатие на кнопку заполнения файла
    def send_mail(self):
        # считаю время "начало"
        time_start = time.time()

        # HTML ==---------------------------------
        # открываю файл HTML
        with open(self.label_path_html_file.text(), 'r') as file_html:
            all_strings_html_file = file_html.read()

        # XLS ==---------------------------------
        # открываю файл XLS и выбираю активный лист
        wb_xls = openpyxl.load_workbook(self.label_path_xls_file.text())
        wb_xls_s = wb_xls.active

        # переменные для обработки XLS
        list_replaced_words = []  # список слов для замены в HTML файле
        # chars_for_replace = '{{xxx}}'  # шаблон для замены в HTML файле
        # cell_value = None  # инициализировал переменную, а то ИДЭ ругается ))

        # получение значений ячеек из XLS файла
        for row_in_xls in range(wb_xls_s.min_row, wb_xls_s.max_row + 1):
            for col_in_xls in range(wb_xls_s.min_column, wb_xls_s.max_column + 1):
                # значение ячейки и её координаты
                cell_value = wb_xls_s.cell(row_in_xls, col_in_xls).value
                # cell_coord = wb_xls_s.cell(row_in_xls, col_in_xls).coordinate
                # print(f'{cell_coord} ... {cell_value}')

                # если первая строка, то сформировать список спецстрок из шапки, которые нужно будет искать и заменять
                # иначе обрабатывается остальные строки с данными
                if row_in_xls == 1:
                    if cell_value != 'email':  # если не колонка с почтами
                        # list_replaced_words.append(chars_for_replace.replace('xxx', cell_value))
                        list_replaced_words.append(cell_value)
                else:
                    # если первая колонка, то создаётся объект, иначе просто заполняются атрибуты из ячеек
                    if col_in_xls == wb_xls_s.min_column:
                        # создание объекта
                        globals()['Recipient' + str(wb_xls_s.cell(row_in_xls, wb_xls_s.min_column).value)] =\
                            RecipientData(rd_text_message=all_strings_html_file)

                        # заполнение первого аргумента
                        globals()['Recipient' + str(wb_xls_s.cell(row_in_xls, wb_xls_s.min_column).value)].\
                            __setattr__(wb_xls_s.cell(wb_xls_s.min_column, col_in_xls).value, str(cell_value))
                    else:
                        # заполнение остальных атрибутов по названиям колонок в верхней строке
                        globals()['Recipient' + str(wb_xls_s.cell(row_in_xls, wb_xls_s.min_column).value)].\
                            __setattr__(wb_xls_s.cell(wb_xls_s.min_column, col_in_xls).value, cell_value)

        for count_obj in range(1, RecipientData.count_Recipient + 1):
            pass
            print(f'{globals()["Recipient" + str(count_obj)].get_all_info()}')

        # time.sleep(0.1)

        # закрываю файл
        wb_xls.close()

        # считаю время "конец"
        time_finish = time.time()

        # информационное окно об окончании работы программы
        self.window_info = PyQt5.QtWidgets.QMessageBox()
        self.window_info.setWindowTitle('Окончено')
        self.window_info.setText(f'Файлы закрыты.\n'
                                 f'Отправка писем сделана за {round(time_finish - time_start, 1)} секунд.')
        self.window_info.exec_()

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
