import sys

from ui import Ui_MainWindow, Ui_Doble_Window
from func import *


class ProgSend(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(ProgSend, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.init_UI()
        self.dialog = None

    def init_UI(self):
        self.setWindowIcon(QIcon("image\\186100_900.jpg"))
        initial_filling(self)
        self.ui.all_people.itemDoubleClicked.connect(lambda: on_item_clicked(self))
        self.ui.btn_sender.clicked.connect(lambda: on_btn_sender_clicked(self))
        self.ui.search.returnPressed.connect(lambda: on_btn_search_clicked(self))
        self.ui.btn_writer.clicked.connect(lambda: write_receipt(self))
        self.ui.btn_log.clicked.connect(lambda: open_log())
        self.ui.btn_log_not_data.clicked.connect(lambda: open_log_not_data())
        self.ui.btn_restr_one.clicked.connect(lambda: write_restr_one(self))
        self.ui.btn_restr_two.clicked.connect(lambda: write_restr_two(self))
        self.ui.btn_settings.clicked.connect(self.open_dialog)

    def open_dialog(self):
        self.dialog = WindowTwo()
        self.dialog.show()


class WindowTwo(QtWidgets.QMainWindow, Ui_Doble_Window):
    def __init__(self):
        super(WindowTwo, self).__init__()
        self.ui = Ui_Doble_Window()
        self.ui.setupUI_two(self)
        self.init_Doble()

    def init_Doble(self):
        self.setWindowIcon(QIcon("image/186100_900.jpg"))
        start_config(self)
        self.ui.btn_save.clicked.connect(lambda: save_config(self))


app = QtWidgets.QApplication(sys.argv)
application = ProgSend()
application.show()
sys.exit(app.exec())
