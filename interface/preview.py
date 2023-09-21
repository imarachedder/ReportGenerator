import glob
import sys

from PyQt6 import QtCore, QtWidgets
from PyQt6 import QtWebEngineWidgets
from jinja2 import Template

from settings import path_file_html, path_templates_jinja
from write_excel_description import WriterExcelTP, WriterExcelDAD


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
            QtCore.QUrl().fromLocalFile(path_file_html))
        self.verticalLayout.addWidget(self.webEngineView)
        self.buttonBox = QtWidgets.QDialogButtonBox(Preview_window)
        self.buttonBox.setOrientation(QtCore.Qt.Orientation.Horizontal)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox().StandardButton.Cancel | QtWidgets.QDialogButtonBox().StandardButton.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)
        self.retranslateUi(Preview_window)
        # self.buttonBox.accepted.connect(Preview_window.accept)
        self.buttonBox.rejected.connect(Preview_window.reject)
        QtCore.QMetaObject.connectSlotsByName(Preview_window)

    def retranslateUi (self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.buttonBox.button(QtWidgets.QDialogButtonBox().StandardButton.Cancel).setText("Назад")


class Preview(QtWidgets.QDialog, Ui_Preview_window):

    def __init__ (self, title=None, parent=None, data=None):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super(Preview, self).__init__(parent)
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.setWindowTitle(title)
        # self.data = data
        # print(self.data)
        self.title = title
        self.buttonBox.accepted.connect(lambda : self.write_excel(data))

    def write_excel (self, data=None):
        if data is None:
            data = {}
        if self.parent().parent.tp_checkBox.isChecked():
            print('заполняю тех паспорт')

            report = WriterExcelTP(data)
            report.save_file()

        print('сохранил файл')
        # report.save_file()

    def filling_templates (self, data=None):
        '''
        Заполнение html шаблонов
        :return:
        '''
        # print(fr'{os.path.dirname(os.path.abspath(__file__))}\Новая папка\templates')
        list_files = glob.glob(path_templates_jinja)  # список всех файлов с расширением .htm

        for file in list_files:
            with open(file, encoding = 'windows-1251') as f:
                read_file = f.read()
            template = Template(read_file, autoescape = True)
            name = file.split('\\')[-1]
            with open(f"templates/tp_curr1.files/{name}", "w", encoding='windows-1251') as f:
                f.write(template.render(data=data))


def main ():
    app = QtWidgets.QApplication(sys.argv)
    window = Preview()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
