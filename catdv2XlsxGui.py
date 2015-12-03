import sys
import xlsxwriter
from PyQt4 import QtGui


class Txt2Xlsx(QtGui.QWidget):

    def __init__(self):
        super(Txt2Xlsx, self).__init__()

        self.initUI()

    def initUI(self):

        btn1 = QtGui.QPushButton("Open .txt file", self)
        self.lbl1 = QtGui.QLineEdit(self)
        btn2 = QtGui.QPushButton("Save as .xlsx file", self)
        self.lbl2 = QtGui.QLineEdit(self)

        btn1.clicked.connect(self.open_text_f)
        btn2.clicked.connect(self.save_xlsx)

        main_grid = QtGui.QGridLayout()
        main_grid.setSpacing(10)

        main_grid.addWidget(btn1, 1, 0)
        main_grid.addWidget(self.lbl1, 2, 0)
        main_grid.addWidget(btn2, 3, 0)
        main_grid.addWidget(self.lbl2, 4, 0)

        self.setLayout(main_grid)

        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('CatDV to Xlsx')
        self.show()

    def open_text_f(self):
        self.text_f = QtGui.QFileDialog.getOpenFileName()
        self.lbl1.setText(self.text_f)
        self.collected_data = self.sort_catdv()

    def sort_catdv(self):
        collected = []
        tx_file = open(self.text_f, 'r')
        for i in tx_file:
            i.strip()
            i = i.split('\t')
            collected.append(i)
        return collected

    def save_xlsx(self):
        self.ask_save_as()

    def ask_save_as(self):
        self.xlsx_filename = QtGui.QFileDialog.getSaveFileName() + '.xlsx'
        self.build_xlsx(self.collected_data, self.xlsx_filename)

    def build_xlsx(self, collected, xlsx_fname):
        """Creates an xlsx workbook for any size of text output."""
        fname = str(xlsx_fname)
        row = 0
        col = 0
        workbook = xlsxwriter.Workbook(fname)
        worksheet = workbook.add_worksheet()
        for item in collected:
            for i in item:
                worksheet.write(row, col, i.rstrip())
                col += 1
                if col >= len(item):
                    col = 0
                    row += 1
        workbook.close()
        self.lbl2.setText('Saved as : {}'.format(xlsx_fname))


def main():
    app = QtGui.QApplication(sys.argv)
    tx = Txt2Xlsx()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()



