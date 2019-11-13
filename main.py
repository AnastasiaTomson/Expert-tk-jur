import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
import pyexcel as p


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Эксперт-ТК / Журналы'
        self.left = 10
        self.top = 10
        self.width = 700
        self.height = 500
        self.icon = QIcon("M.jpg")
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setWindowIcon(self.icon)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.openFileNameDialog()
        # self.saveFileDialog()

        self.show()

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "Загрузить файл", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            # print(fileName)
            self.parse_file(fileName)

    # def openFileNamesDialog(self):
    #     options = QFileDialog.Options()
    #     options |= QFileDialog.DontUseNativeDialog
    #     files, _ = QFileDialog.getOpenFileNames(self, "Открыть файлы", "",
    #                                             "All Files (*);;Python Files (*.py)", options=options)
    #     if files:
    #         print(files)

    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file, _ = QFileDialog.getSaveFileName(self, "Сохранить ", "",
                                              "All Files (*);;Text Files (*.text)", options=options)
        if file:
            print(file)

    def parse_file(self, file_name):
        if '.xlsx' not in file_name:
            p.save_book_as(file_name=file_name,
                           dest_file_name='test.xlsx')
            file_name = 'test.xlsx'
        # читаем excel-файл
        wb = openpyxl.load_workbook(file_name)

        # печатаем список листов
        sheets = wb.sheetnames
        for sheet in sheets:
            print(sheet)

        # получаем активный лист
        sheet = wb.active

        # печатаем значение ячейки A1
        print(sheet['A'].ro.value)
        # печатаем значение ячейки B1
        print(sheet['B'].value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
