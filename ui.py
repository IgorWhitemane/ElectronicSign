from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1000, 440)
        font = QtGui.QFont()
        font.setPointSize(8)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setStyleSheet("")
        self.centralwidget.setObjectName("centralwidget")
        self.all_people = QtWidgets.QTreeWidget(self.centralwidget)
        self.all_people.setGeometry(QtCore.QRect(5, 70, 990, 191))
        self.all_people.setObjectName("all_people")
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        self.all_people.headerItem().setForeground(0, brush)
        self.result = QtWidgets.QLabel(self.centralwidget)
        self.result.setGeometry(QtCore.QRect(250, 10, 350, 50))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.result.setFont(font)
        self.result.setStyleSheet("border: 2px solid #22222e;")
        self.result.setText("Не выбран")
        self.result.setAlignment(QtCore.Qt.AlignCenter)
        self.result.setIndent(1)
        self.result.setObjectName("result")

        font.setPointSize(10)
        self.status = QtWidgets.QLabel(self.centralwidget)
        self.status.setGeometry(QtCore.QRect(630, 10, 350, 50))
        self.status.setFont(font)
        self.status.setStyleSheet("border: 2px solid #22222e;")
        self.status.setText("Статус")
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        self.status.setIndent(1)
        self.status.setObjectName("result")

        self.btn_sender = QtWidgets.QPushButton(self.centralwidget)
        self.btn_sender.setGeometry(QtCore.QRect(10, 270, 200, 40))
        self.btn_sender.setObjectName("btn_sender")

        self.btn_new_people = QtWidgets.QPushButton(self.centralwidget)
        self.btn_new_people.setGeometry(QtCore.QRect(580, 270, 200, 30))
        self.btn_new_people.setObjectName("btn_sender")

        self.btn_log = QtWidgets.QPushButton(self.centralwidget)
        self.btn_log.setGeometry(QtCore.QRect(930, 270, 60, 22))
        self.btn_log.setObjectName("btn_log")

        self.btn_log_not_data = QtWidgets.QPushButton(self.centralwidget)
        self.btn_log_not_data.setGeometry(QtCore.QRect(930, 300, 60, 22))
        self.btn_log_not_data.setObjectName("btn_log_not_data")

        self.btn_restr_two = QtWidgets.QPushButton(self.centralwidget)
        self.btn_restr_two.setGeometry(QtCore.QRect(300, 380, 200, 22))
        self.btn_restr_two.setObjectName("btn_restr_two")

        self.btn_restr_one = QtWidgets.QPushButton(self.centralwidget)
        self.btn_restr_one.setGeometry(QtCore.QRect(300, 350, 200, 22))
        self.btn_restr_one.setObjectName("btn_restr_one")

        self.btn_writer = QtWidgets.QPushButton(self.centralwidget)
        self.btn_writer.setGeometry(QtCore.QRect(300, 320, 200, 22))
        self.btn_writer.setObjectName("btn_writer")

        self.search = QtWidgets.QLineEdit(self.centralwidget)
        self.search.setGeometry(QtCore.QRect(10, 30, 200, 30))

        self.ser_num = QtWidgets.QLineEdit(self.centralwidget)
        self.ser_num.setGeometry(QtCore.QRect(275, 275, 250, 30))

        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Отправлятор"))
        self.all_people.headerItem().setText(0, _translate("MainWindow", "ФИО"))
        self.all_people.headerItem().setText(1, _translate("MainWindow", "Пол"))
        self.all_people.headerItem().setText(2, _translate("MainWindow", "Должность"))
        self.all_people.headerItem().setText(3, _translate("MainWindow", "Отдел"))
        self.all_people.headerItem().setText(4, _translate("MainWindow", "Адрес электронной почты"))
        self.all_people.headerItem().setText(5, _translate("MainWindow", "Серия и номер паспорта"))
        self.all_people.headerItem().setText(6, _translate("MainWindow", "Кем выдан"))
        self.all_people.headerItem().setText(7, _translate("MainWindow", "Код подразделения"))
        self.all_people.headerItem().setText(8, _translate("MainWindow", "Дата выдачи"))
        self.all_people.headerItem().setText(9, _translate("MainWindow", "Дата рождения"))
        self.all_people.headerItem().setText(10, _translate("MainWindow", "Место рождения"))
        self.all_people.headerItem().setText(11, _translate("MainWindow", "Снилс"))
        self.all_people.headerItem().setText(12, _translate("MainWindow", "Инн"))
        self.all_people.headerItem().setText(13, _translate("MainWindow", "Сотовый номер"))
        self.btn_sender.setText(_translate("MainWindow", "Отправить в УЦ"))
        self.btn_new_people.setText(_translate("MainWindow", "Создать нового пользователя"))
        self.search.setPlaceholderText("Поиск...")
        self.btn_log.setText(_translate("MainWindow", "Log Date"))
        self.btn_log_not_data.setText(_translate("MainWindow", "Log"))
        self.btn_writer.setText(_translate("MainWindow", "Создать расписку"))
        self.btn_restr_one.setText(_translate("MainWindow", "Отметить в реестре выпуска"))
        self.btn_restr_two.setText(_translate("MainWindow", "Отметить в реестре ЭПЦ"))
        self.ser_num.setPlaceholderText("Написать серийный номер сертификата")
