import glob
import sys

from PyQt6 import QtCore, QtWidgets, QtGui
from PyQt6 import QtWebEngineWidgets, QtPdf
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtGui import QPainter, QPageSize
from jinja2 import Template
import settings


class Ui_Preview_window(object):

    def setupUi (self, Preview_window):
        Preview_window.setObjectName("Preview_window")
        Preview_window.resize(1440, 1024)
        self.verticalLayout = QtWidgets.QVBoxLayout(Preview_window)
        self.verticalLayout.setObjectName("verticalLayout")
        self.centralwidget = QtWidgets.QWidget(Preview_window)
        self.centralwidget.setObjectName("centralwidget")
        self.webEngineView = QtWebEngineWidgets.QWebEngineView(self.centralwidget)
        self.webEngineView.load(
            QtCore.QUrl().fromLocalFile(r"C:\Users\sibregion\Desktop\test\report\Новая папка\tp_curr.htm"))
        self.verticalLayout.addWidget(self.webEngineView)
        self.buttonBox = QtWidgets.QDialogButtonBox(Preview_window)
        self.buttonBox.setOrientation(QtCore.Qt.Orientation.Horizontal)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox().StandardButton.Cancel | QtWidgets.QDialogButtonBox().StandardButton.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)
        self.retranslateUi(Preview_window)
        self.buttonBox.rejected.connect(Preview_window.reject)
        QtCore.QMetaObject.connectSlotsByName(Preview_window)

    def retranslateUi (self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.buttonBox.button(QtWidgets.QDialogButtonBox().StandardButton.Cancel).setText("Назад")
        self.buttonBox.button(QtWidgets.QDialogButtonBox().StandardButton.Ok).setText("Печать")


class Preview(QtWidgets.QDialog, Ui_Preview_window):

    def __init__ (self, title = None, parent = None):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.setWindowTitle(title)
        # self.buttonBox.StandardButton.Ok.clicked.connect(self.printer)
        self.buttonBox.accepted.connect(self.print_file)

    @staticmethod
    def filling_templates (data = None):
        '''
        Заполнение html шаблонов
        :return:
        '''
        # print(fr'{os.path.dirname(os.path.abspath(__file__))}\Новая папка\templates')
        list_files = glob.glob(
            r'C:\Users\sibregion\Desktop\test\report\Новая папка\templates_curr\*[0-9].htm')  # список всех файлов с расширением .htm

        for file in list_files:
            with open(file, encoding = 'windows-1251') as f:
                read_file = f.read()
            template = Template(read_file)
            name = file.split('\\')[-1]
            with open(f"Новая папка/tp_curr.files/{name}", "w", encoding = 'windows-1251') as f:
                f.write(template.render(data = data))

    def print_file (self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        # printer.setOutputFormat()
        # printer.setOutputFileName(r'C:\Users\sibregion\Desktop\test\report\Новая папка\ТП.pdf')
        dialog = QPrintDialog(printer, self.webEngineView)
        if dialog.exec():
            self.webEngineView.setZoomFactor(1)
            self.webEngineView.page().printToPdf("test.pdf")

            # self.webEngineView.render(painter)

    def handle_paint_request (self, printer):
        painter = QtGui.QPainter(printer)
        painter.setViewport(self.webEngineView.rect())
        painter.setWindow(self.webEngineView.rect())
        self.webEngineView.render(painter)
        painter.end()

    def print_document (self):
        html = self.webEngineView.page().toHtml()
        printer = QPrinter()
        if printer.isValid():
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            # Устанавливаем параметры принтера
            printer.setPageSize(QPageSize.PageSizeId.A4Small)
            printer.setPaperSize(QPrinter.PaperSize.QPrinterA4)
            printer.setColorMode(QPrinter.ColorMode.Color)
            printer.page()
            painter = QPainter(printer)
            painter.drawText(0, 0, html)
            painter.endPage()
            printer.waitForDone()
            print(html)  # Распечатываем HTML документ на принтере
        else:
            print("Принтер недоступен")


def main ():
    if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
        QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    app = QtWidgets.QApplication(sys.argv)
    window = Preview()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
