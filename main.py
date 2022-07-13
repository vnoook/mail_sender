# TODO
# сделать проверку на повторное нажатие на Отправить
# сделать сопоставление задержек с полями на форме, с их проверкой на числа
# сделать прогресс-бар по отправке, считать количество или время?

# ...
# INSTALL
# pip install openpyxl
# pip install PyQt5

# COMPILE
# pyinstaller -F -w main.py
# ...

import sys
import time
import smtplib
import email.mime.text
import PyQt5
import PyQt5.QtWidgets
import PyQt5.QtCore
import PyQt5.QtGui
import openpyxl
import openpyxl.utils
import msc


# класс получателя сообщения одной отправки
class RecipientData:
    """Класс Получателя сообщения"""

    # количество экземпляров класса
    count_Recipient = 0

    def __init__(self, rd_text_message=None):
        self.num = None
        self.fam = None
        self.im = None
        self.otch = None
        self.email = None
        self.mno_code = None
        self.text_message = rd_text_message
        self.flag_send_message = False

        # изменение счётчика экземпляров
        RecipientData.count_Recipient += 1

    # метод получения всех значений аргументов
    def get_all_info(self):
        return f'Объект {self.get_obj_name()}, ' \
               f'{id(self)}, ' \
               f'{self.num = }, ' \
               f'{self.fam = }, ' \
               f'{self.im = }, ' \
               f'{self.otch = }, ' \
               f'{self.email = }, ' \
               f'{self.mno_code = }, ' \
               f'{self.text_message = }, ' \
               f'{self.flag_send_message = }'

    # метод получения имени экземпляра
    def get_obj_name(self):
        for glob_name, glob_val in globals().items():
            if glob_val is self:
                return glob_name

    # метод замены тегов на значения
    @staticmethod
    def replace_text_message(mail_tag, tag_value, mail_text):
        mail_text = mail_text.replace('{{' + mail_tag + '}}', tag_value)
        return mail_text

    # переопределение метода для замены тегов в почтовом сообщении
    def __setattr__(self, key, value):
        if 'Recipient' in str(self.get_obj_name()):
            if key in ('num', 'fam', 'im', 'otch', 'mno_code'):
                self.text_message = self.replace_text_message(key, value, self.text_message)
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
        self.q_pocket = 5
        # задержка между письмами в пакете при отправке, в секундах
        self.q_messages = 3
        # задержка между отправками пакетов, в секундах
        self.send_delay = 5  #  300  # 5 минут

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
        self.lineEdit_q_pocket.setText(str(self.q_pocket))
        self.lineEdit_q_pocket.setGeometry(PyQt5.QtCore.QRect(10, 160, 90, 20))
        self.lineEdit_q_pocket.setClearButtonEnabled(True)
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
        self.lineEdit_q_messages.setText(str(self.q_messages))
        self.lineEdit_q_messages.setGeometry(PyQt5.QtCore.QRect(10, 220, 90, 20))
        self.lineEdit_q_messages.setClearButtonEnabled(True)
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
        self.lineEdit_mail_delay.setText(str(self.send_delay))
        self.lineEdit_mail_delay.setGeometry(PyQt5.QtCore.QRect(10, 280, 90, 20))
        self.lineEdit_mail_delay.setClearButtonEnabled(True)
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

        # SEND_TEST_MAIL
        # pushButton_send_test_mail
        self.pushButton_send_test_mail = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_send_test_mail.setObjectName('pushButton_send_test_mail')
        self.pushButton_send_test_mail.setEnabled(False)
        self.pushButton_send_test_mail.setText('Тестовое письмо')
        self.pushButton_send_test_mail.setGeometry(PyQt5.QtCore.QRect(200, 310, 300, 25))
        self.pushButton_send_test_mail.setFixedWidth(130)
        self.pushButton_send_test_mail.clicked.connect(self.send_test_mail)
        self.pushButton_send_test_mail.setToolTip(self.pushButton_send_test_mail.objectName())

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

        # INVISIBLE
        # checkBox_inviz
        self.checkBox_inviz = PyQt5.QtWidgets.QCheckBox(self)
        self.checkBox_inviz.setObjectName('checkBox_inviz')
        self.checkBox_inviz.setGeometry(PyQt5.QtCore.QRect(10, 500, 190, 40))
        self.checkBox_inviz.clicked.connect(self.on_off_lineEdits)
        self.checkBox_inviz.setText('Хочу редактировать!')
        self.checkBox_inviz.setToolTip(self.checkBox_inviz.objectName())

    # событие - скрытие\отображение возможности редактирования полей
    def on_off_lineEdits(self):
        if self.checkBox_inviz.isChecked():
            self.lineEdit_q_pocket.setEnabled(True)
            self.lineEdit_q_messages.setEnabled(True)
            self.lineEdit_mail_delay.setEnabled(True)
        else:
            self.lineEdit_q_pocket.setEnabled(False)
            self.lineEdit_q_messages.setEnabled(False)
            self.lineEdit_mail_delay.setEnabled(False)

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

        # активация и деактивация объектов на форме зависящее от 'выбраны ли все файлы' и 'они разные'
        if self.text_empty_path_file not in (self.label_path_html_file.text(), self.label_path_xls_file.text()):
            self.pushButton_send_mail.setEnabled(True)
            self.pushButton_send_test_mail.setEnabled(True)

    # событие - нажатие на кнопку отправки почты
    def send_mail(self):
        # считаю время 'начало'
        time_start = time.monotonic()

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

        # счётчик объектов, с 0 потому что первая строка шапка и там нет обрабатываемых данных
        obj_count = 0
        # короткое обращение к объекту, утилитарная переменная
        obj_name = None

        # получение значений ячеек из XLS файла
        for row_in_xls in range(1, wb_xls_s.max_row + 1):
            for col_in_xls in range(1, wb_xls_s.max_column + 1):
                # значение ячейки
                cell_value = wb_xls_s.cell(row_in_xls, col_in_xls).value

                # если первая строка, то сформировать список спецстрок из шапки, которые нужно будет искать и заменять
                # иначе обрабатывается остальные строки с данными
                if row_in_xls == 1:
                    # считываются все, кроме email значения и вносятся для последующей теговой замены
                    if cell_value != 'email':  # если не колонка с почтами
                        list_replaced_words.append(cell_value)
                else:
                    # если первая колонка, то создаётся объект, иначе просто заполняются атрибуты из ячеек
                    if col_in_xls == 1:
                        # увеличение итерации счётчика созданных объектов
                        obj_count += 1

                        # создание объекта
                        globals()['Recipient' + str(obj_count)] = RecipientData(rd_text_message=all_strings_html_file)

                        # короткое обращение к созданному объекту
                        obj_name = globals()['Recipient' + str(obj_count)]

                        # заполнение первого аргумента
                        obj_name.__setattr__(wb_xls_s.cell(1, col_in_xls).value, str(cell_value))
                    else:
                        # заполнение остальных атрибутов по названиям колонок в верхней строке
                        obj_name.__setattr__(wb_xls_s.cell(1, col_in_xls).value, cell_value)
        # закрываю файл
        wb_xls.close()

        # # временная выдача данных перед отправкой, потом удалить!!!!!!!!!!!!!!!!
        # for count_obj in range(1, RecipientData.count_Recipient + 1):
        #     print(f'{globals()["Recipient" + str(count_obj)].get_all_info()}')
        # print()

        # участок отправки писем и ожиданий времени
        list_recipients = [x for x in range(1, RecipientData.count_Recipient + 1)]
        print()
        for recipient in range(0, RecipientData.count_Recipient, self.q_pocket):
            list_recipients_pocket = list_recipients[recipient: recipient + self.q_pocket]
            # print(f'{list_recipients_pocket = }')

            for recipient_number in list_recipients_pocket:
                print(f'{recipient_number} письмо отправляется')

                # короткое обращение к объекту
                obj_name = globals()['Recipient' + str(recipient_number)]

                # создание соединения с сервером
                smtp_link = smtplib.SMTP(msc.msc_mail_server)
                smtp_link.starttls()

                try:
                    # подключение к аккаунту
                    smtp_link.login(msc.msc_from_address, msc.msc_login_pass)
                    # создание текста письма
                    msg = email.mime.text.MIMEText(obj_name.text_message, 'html')
                    msg['From'] = msc.msc_from_address
                    msg['To'] = obj_name.email
                    msg['Subject'] = 'Проверка отправки почты HTML письмом!'
                    smtp_link.send_message(msg, msc.msc_from_address, obj_name.email)
                    smtp_link.quit()
                    print('Электронное письмо отправлено удачно!')
                    obj_name.flag_send_message = True
                except Exception as _ex:
                    print(f'{_ex}\nЭлектронное письмо не отправлено, проверьте логин-пароль!')
                print()

                if list_recipients_pocket.index(recipient_number) != len(list_recipients_pocket) - 1:
                    print('задержка в секундах между письмами', self.q_messages)
                    print()
                    time.sleep(self.q_messages)

            if len(list_recipients_pocket) == self.q_pocket:
                if RecipientData.count_Recipient not in list_recipients_pocket:
                    print('задержка в секундах между пакетами отправки', self.send_delay)
                    time.sleep(self.send_delay)
            print()

        # временная выдача данных после отправки, потом удалить!!!!!!!!!!!!!!!!
        for count_obj in range(1, RecipientData.count_Recipient + 1):
            print(f'{globals()["Recipient" + str(count_obj)].get_all_info()}')
        print()

        # считаю время 'конец'
        time_finish = time.monotonic()

        # информационное окно об окончании работы программы
        self.window_info = PyQt5.QtWidgets.QMessageBox()
        self.window_info.setWindowTitle('Окончено')
        self.window_info.setText(f'Файлы закрыты.\n'
                                 f'Отправка писем сделана за {round(time_finish - time_start, 1)} секунд.')
        self.window_info.exec_()

    # событие - нажатие на кнопку отправки тестового письма
    def send_test_mail(self):
        # открываю и читаю файл HTML
        with open(self.label_path_html_file.text(), 'r') as file_html:
            all_strings_html_file = file_html.read()

        # создание соединения с сервером
        smtp_link = smtplib.SMTP(msc.msc_mail_server)
        smtp_link.starttls()

        try:
            # подключение к аккаунту
            smtp_link.login(msc.msc_from_address, msc.msc_login_pass)
            # создание текста письма
            msg = email.mime.text.MIMEText(all_strings_html_file, 'html')
            msg['From'] = msc.msc_from_address
            msg['To'] = msc.msc_test_address
            msg['Subject'] = msc.msc_subject_text
            smtp_link.send_message(msg, msc.msc_from_address, msc.msc_test_address)
            smtp_link.quit()
            print('Электронное письмо отправлено удачно!')
            return 'Электронное письмо отправлено удачно!'
        except Exception as _ex:
            print(f'{_ex}\nЭлектронное письмо не отправлено, проверьте логин-пароль!')
            return f'{_ex}\nЭлектронное письмо не отправлено, проверьте логин-пароль!'

    # событие - нажатие на кнопку Выход
    @staticmethod
    def click_on_btn_exit():
        exit()

    # проверка строка на числовое значение - взять число из поля или взять значение по умолчанию
    @staticmethod
    def check_is_digit(data_in):
        return data_in


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
