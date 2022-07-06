from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel
from PyQt5.QtCore import pyqtSignal
import xlrd, random
import sys


class Main_window(QWidget):
    created = pyqtSignal(int, int, int, int, int)

    def __init__(self, tit1, tit2, tit3, tit4, tit5):
        super().__init__()

        self.setWindowTitle("HASHTAGS GENERATOR")

        self.setFixedHeight(550)
        self.setFixedWidth(410)

        self.setStyleSheet('background-color:#f7f1e3')

        self.tit1 = tit1
        self.tit2 = tit2
        self.tit3 = tit3
        self.tit4 = tit4
        self.tit5 = tit5


        self.create_buttons()


    def create_buttons(self):

        self.label = QLabel("Enter how many hashtags you\nwant from the given category:", self)
        self.label.setGeometry(40, 20, 500, 70)
        self.label.setFont(QFont('Times New Roman', 20))

        self.cat1_label = QLabel(self.tit1, self)
        self.cat1_label.setGeometry(120, 110, 400, 40)
        self.cat1_label.setFont(QFont('Times New Roman', 18))

        self.cat1_edit = QLineEdit(self)
        self.cat1_edit.setGeometry(60, 110, 40, 40)
        self.cat1_edit.setStyleSheet('background-color:white')

        self.cat2_label = QLabel(self.tit2, self)
        self.cat2_label.setGeometry(120, 170, 400, 40)
        self.cat2_label.setFont(QFont('Times New Roman', 18))

        self.cat2_edit = QLineEdit(self)
        self.cat2_edit.setGeometry(60, 170, 40, 40)
        self.cat2_edit.setStyleSheet('background-color:white')

        self.cat3_label = QLabel(self.tit3, self)
        self.cat3_label.setGeometry(120, 230, 400, 40)
        self.cat3_label.setFont(QFont('Times New Roman', 18))

        self.cat3_edit = QLineEdit(self)
        self.cat3_edit.setGeometry(60, 230, 40, 40)
        self.cat3_edit.setStyleSheet('background-color:white')

        self.cat4_label = QLabel(self.tit4, self)
        self.cat4_label.setGeometry(120, 290, 400, 40)
        self.cat4_label.setFont(QFont('Times New Roman', 18))

        self.cat4_edit = QLineEdit(self)
        self.cat4_edit.setGeometry(60, 290, 40, 40)
        self.cat4_edit.setStyleSheet('background-color:white')

        self.cat5_label = QLabel(self.tit5, self)
        self.cat5_label.setGeometry(120, 350, 400, 40)
        self.cat5_label.setFont(QFont('Times New Roman', 18))

        self.cat5_edit = QLineEdit(self)
        self.cat5_edit.setGeometry(60, 350, 40, 40)
        self.cat5_edit.setStyleSheet('background-color:white')

        self.submit = QPushButton("S U B M I T", self)
        self.submit.setGeometry(100, 420, 210, 50)
        self.submit.clicked.connect(self.get_cats)
        self.submit.setStyleSheet('background-color:#4a69bd;'
                                  'color: white')
        self.submit.setFont(QFont('Times New Roman', 18))

        self.quit = QPushButton("Q U I T", self)
        self.quit.setGeometry(100, 480, 210, 50)
        self.quit.setStyleSheet('background-color:#4a69bd;'
                                'color: white')
        self.quit.setFont(QFont('Times New Roman', 18))
        self.quit.clicked.connect(QApplication.instance().quit)

    def get_cats(self):

        try:
            cat1_num = (int)(self.cat1_edit.text())
            cat2_num = (int)(self.cat2_edit.text())
            cat3_num = (int)(self.cat3_edit.text())
            cat4_num = (int)(self.cat4_edit.text())
            cat5_num = (int)(self.cat5_edit.text())
            self.created.emit(cat1_num, cat2_num, cat3_num, cat4_num, cat5_num)
        except:
            pass


class Result_window(QWidget):  # tworzenie okna exit
    def __init__(self, final_list):
        super().__init__()  # wywoluje konstruktor qwidgeta

        self.final_list = final_list

        self.setWindowTitle("RESULT")  # ustawiam tytuł okna
        self.setFixedHeight(170)  # ustawiam wysokosc (nie da sie zmienic)
        self.setFixedWidth(1000)  # ustawiam szerokosc
        self.setStyleSheet('background-color:#f7f1e3')  # ustawiam kolor tła

        self.result1 = QLineEdit(self)
        self.result1.setGeometry(10, 10, 980, 150)
        for i in range(len(final_list)):
            for j in range(len(final_list[i])):
                self.result1.insert("#" + (str)(final_list[i][j]) + " ")


class Choise_window(QWidget):  # tworzenie okna exit

    choosed = pyqtSignal(str)

    def __init__(self, number, sheets_names):
        super().__init__()  # wywoluje konstruktor qwidgeta
        self.sheets_names = sheets_names
        self.number = number
        self.setWindowTitle("choose competeeon")  # ustawiam tytuł okna
        self.setFixedHeight(10 + self.number*50)  # ustawiam wysokosc (nie da sie zmienic)
        self.setFixedWidth(320)  # ustawiam szerokosc
        self.setStyleSheet('background-color:#f7f1e3')  # ustawiam kolor tła
        self.get_sheet_name()



    def get_sheet_name(self):

        self.buttom1 = QPushButton(self.sheets_names[0], self)
        self.buttom1.setGeometry(10, 10 + (50 * 0), 300, 40)
        self.buttom1.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom1.setFont(QFont('Times New Roman', 18))
        self.buttom1.clicked.connect(self.open_main_window1)


        self.buttom2 = QPushButton(self.sheets_names[1], self)
        self.buttom2.setGeometry(10, 10 + (50 * 1), 300, 40)
        self.buttom2.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom2.setFont(QFont('Times New Roman', 18))
        self.buttom2.clicked.connect(self.open_main_window2)


        self.buttom3 = QPushButton(self.sheets_names[2], self)
        self.buttom3.setGeometry(10, 10 + (50 * 2), 300, 40)
        self.buttom3.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom3.setFont(QFont('Times New Roman', 18))
        self.buttom3.clicked.connect(self.open_main_window3)


        self.buttom4 = QPushButton(self.sheets_names[3], self)
        self.buttom4.setGeometry(10, 10 + (50 * 3), 300, 40)
        self.buttom4.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom4.setFont(QFont('Times New Roman', 18))
        self.buttom4.clicked.connect(self.open_main_window4)


        self.buttom5 = QPushButton(self.sheets_names[4], self)
        self.buttom5.setGeometry(10, 10 + (50 * 4), 300, 40)
        self.buttom5.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom5.setFont(QFont('Times New Roman', 18))
        self.buttom5.clicked.connect(self.open_main_window5)


        self.buttom6 = QPushButton(self.sheets_names[5], self)
        self.buttom6.setGeometry(10, 10 + (50 * 5), 300, 40)
        self.buttom6.setStyleSheet('background-color:#4a69bd;'
                               'color: white')
        self.buttom6.setFont(QFont('Times New Roman', 18))
        self.buttom6.clicked.connect(self.open_main_window6)


    def open_main_window1(self):
        name = self.sheets_names[0]
        self.choosed.emit(name)

    def open_main_window2(self):
        name = self.sheets_names[1]
        self.choosed.emit(name)

    def open_main_window3(self):
        name = self.sheets_names[2]
        self.choosed.emit(name)

    def open_main_window4(self):
        name = self.sheets_names[3]
        self.choosed.emit(name)

    def open_main_window5(self):
        name = self.sheets_names[4]
        self.choosed.emit(name)

    def open_main_window6(self):
        name = self.sheets_names[5]
        self.choosed.emit(name)



def make_boards(Arkusz1, cat1_num, cat2_num, cat3_num, cat4_num, cat5_num):

    cat1 = []
    cat2 = []
    cat3 = []
    cat4 = []
    cat5 = []
    stale_hasz = []

    for i in range(Arkusz1.nrows):
        if Arkusz1.cell_value(i, 0) != "":
            cat1.append(Arkusz1.cell_value(i, 0))
        if Arkusz1.cell_value(i, 1) != "":
            cat2.append(Arkusz1.cell_value(i, 1))
        if Arkusz1.cell_value(i, 2) != "":
            cat3.append(Arkusz1.cell_value(i, 2))
        if Arkusz1.cell_value(i, 3) != "":
            cat4.append(Arkusz1.cell_value(i, 3))
        if Arkusz1.cell_value(i, 4) != "":
            cat5.append(Arkusz1.cell_value(i, 4))
        if Arkusz1.cell_value(i, 5) != "":
            stale_hasz.append(Arkusz1.cell_value(i, 5))

    cat1.pop(0)
    cat2.pop(0)
    cat3.pop(0)
    cat4.pop(0)
    cat5.pop(0)
    stale_hasz.pop(0)

    try:
        cat1s = random.sample(cat1, cat1_num)
        cat2s = random.sample(cat2, cat2_num)
        cat3s = random.sample(cat3, cat3_num)
        cat4s = random.sample(cat4, cat4_num)
        cat5s = random.sample(cat5, cat5_num)
    except:
        pass

    final_list = []
    final_list.append(cat1s)
    final_list.append(cat2s)
    final_list.append(cat3s)
    final_list.append(cat4s)
    final_list.append(cat5s)
    final_list.append(stale_hasz)
    return (final_list)


class Program:
    def __init__(self, qapp):
        self.main_window = None
        self.qapp = qapp
        self.result_window = None
        self.choise_window = None

    def run(self):
        filepath = "C:\\HasztagGenerator\\hasztags.xls"
        self.book = xlrd.open_workbook(filepath)
        self.choise_number = self.book.nsheets
        self.sheets = self.book.sheet_names()



        self.choise_window = Choise_window(self.choise_number, self.sheets)
        self.choise_window.show()

        self.choise_window.choosed.connect(self.clicked_buttom)

        self.qapp.exec_()



    def clicked_submit(self, cat1, cat2, cat3, cat4, cat5):

        final_list = make_boards(self.arkusz, cat1, cat2, cat3, cat4, cat5)

        self.result_window = Result_window(final_list)
        self.result_window.show()

    def clicked_buttom(self, name):

        self.arkusz = self.book.sheet_by_name(name)
        titules = self.arkusz.row_values(0)
        tit1 = titules[0]
        tit2 = titules[1]
        tit3 = titules[2]
        tit4 = titules[3]
        tit5 = titules[4]


        self.choise_window.hide()
        self.main_window = Main_window(tit1, tit2, tit3, tit4, tit5)
        self.main_window.show()
        self.main_window.created.connect(self.clicked_submit)




if __name__ == '__main__':
    app = QApplication(sys.argv)
    program = Program(app)
    sys.exit(program.run())
