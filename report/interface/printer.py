from PyQt6.QtCore import (QCoreApplication, QEventLoop, QObject, QPointF, Qt, QUrl, pyqtSlot)
from PyQt6.QtGui import QKeySequence, QPainter, QShortcut
from PyQt6.QtPrintSupport import QPrintDialog, QPrinter, QPrintPreviewDialog
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWidgets import QApplication, QDialog, QLabel, QProgressBar, QProgressDialog


class PrintHandler(QObject):
    def __init__ (self, parent = None):
        super().__init__(parent)
        self.m_page = None
        self.m_inPrintPreview = False


    def setPage (self, page):
        assert not self.m_page
        self.m_page = page
        self.m_page.printRequested.connect(self.printPreview)

    @pyqtSlot()
    def print (self):
        # printer = QPrinter(QPrinter.HighResolution)
        printer = QPrinter()
        dialog = QPrintDialog(printer, self.m_page.view())
        if dialog.exec() != QDialog.accepted:
            return
        self.printDocument(printer)

    @pyqtSlot()
    def printPreview (self):
        if not self.m_page:
            return
        if self.m_inPrintPreview:
            return
        self.m_inPrintPreview = True
        printer = QPrinter()
        preview = QPrintPreviewDialog(printer, self.m_page.view())
        preview.paintRequested.connect(self.printDocument)
        preview.exec()
        self.m_inPrintPreview = False

    @pyqtSlot(QPrinter)
    def printDocument (self, printer):
        loop = QEventLoop()
        result = False

        def printPreview (success):
            nonlocal result
            result = success
            loop.quit()

        progressbar = QProgressDialog(self.m_page.view())
        progressbar.findChild(QProgressBar).setTextVisible(False)
        progressbar.setLabelText("Wait please...")
        progressbar.setRange(0, 0)
        progressbar.show()
        progressbar.canceled.connect(loop.quit)
        self.m_page.print(printer, printPreview)
        loop.exec()
        progressbar.close()
        if not result:
            painter = QPainter()
            if painter.begin(printer):
                font = painter.font()
                font.setPixelSize(20)
                painter.setFont(font)
                painter.drawText(QPointF(10, 25), "Could not generate print preview.")
                painter.end()


def main ():
    import sys

    QCoreApplication.setAttribute(Qt.ApplicationAttribute.AA_CompressHighFrequencyEvents)
    app = QApplication(sys.argv)
    app.setApplicationName("Previewer")

    view = QWebEngineView()
    #view.setUrl(QUrl("https://ru.stackoverflow.com/questions/1207961"))
    # self.webEngineView = QtWebEngineWidgets.QWebEngineView(self.centralwidget)
    view.setUrl(QUrl().fromLocalFile(r"C:\Users\sibregion\Desktop\test\report\Новая папка\ТП.htm"))

    #view.resize()


    handler = PrintHandler()
    handler.setPage(view.page())

    printPreviewShortCut = QShortcut(QKeySequence(Qt.Key.Key_Control + Qt.Key.Key_P), view)
    printShortCut = QShortcut(QKeySequence(Qt.Key.Key_Control + Qt.Key.Key_Shift + Qt.Key.Key_P), view)

    printPreviewShortCut.activated.connect(handler.printPreview)
    printShortCut.activated.connect(handler.print)
    view.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
