import glob
import os

from jinja2 import Environment, select_autoescape, PackageLoader, FileSystemLoader, Template

# env = Environment(
#    loader = FileSystemLoader('templates'),
#    autoescape=select_autoescape(['htm', 'html'])
# )
# print(env.get_or_select_template())
# template = env.get_template('ТП (2).htm')


# template = Template()


# with open("index.html", "w", encoding = 'utf-8') as f:
#    f.write(template.render(q = q))
# print(template.render(q=q))


from PyQt6 import QtWebEngineWidgets
from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):

    def setupUi (self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1980, 1024)
        self.verticalLayout = QtWidgets.QVBoxLayout(Dialog)
        self.verticalLayout.setObjectName("verticalLayout")
        self.centralwidget = QtWidgets.QWidget(Dialog)
        self.centralwidget.setObjectName("centralwidget")
        self.webEngineView = QtWebEngineWidgets.QWebEngineView(self.centralwidget)
        self.webEngineView.load(
            QtCore.QUrl().fromLocalFile(r'C:\Users\sibregion\Desktop\test\report\Новая папка\ТП.htm'))
        self.verticalLayout.addWidget(self.webEngineView)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setOrientation(QtCore.Qt.Orientation.Horizontal)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox().StandardButton.Cancel | QtWidgets.QDialogButtonBox().StandardButton.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)
        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi (self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))


def filling_templates():
    list_files = glob.glob(
        fr'{os.path.dirname(os.path.abspath(__file__))}\Новая папка\templates/*[0-9].htm')  # список всех файлов с расширением .htm
    q = {'client': 'OOO SibRoads',
         'name_road': 'какая то дорога',
         'year': 2066, } # для наполнения шаблонов
    for file in list_files:
        read_file = open(file, encoding = 'windows-1251').read()
        template = Template(read_file)
        name = file.split('\\')[-1]
        with open(f"Новая папка/ТПt.files/{name}", "w", encoding = 'windows-1251') as f:
            f.write(template.render(q = q))


if __name__ == "__main__":
    import sys

    #filling_templates()
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec())
