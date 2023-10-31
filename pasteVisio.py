import sys
import pyperclip
import win32com.client as win
import pathlib
from PyQt6.QtWidgets import QMessageBox
# ЗАПУСКАЕМ СНАЧАЛА ЭТОТ ЗАТЕМ MergePdfFiles.py



class ConvertVisio():
    def __init__(self, path_dir=None, filename = None):
        self.path_dir = pathlib.Path(rf"{path_dir}")
        self.out_path = rf"{path_dir}\Результат2.pdf"
        self.files = list(self.path_dir.glob(f'{filename}.xlsm')) + list(self.path_dir.glob('*.vsd'))
        self.pasteVisio()
        print(self.files)

    def pasteVisio(self):
        msg = QMessageBox()
        visio = win.Dispatch("Visio.Application")
        visio.Visible = True
        excel = win.Dispatch('Excel.Application')
        excel.Visible = True
        out_doc = visio.Documents.Add('')
        out_doc.Pages.Add('')
        counter = 1
        for i, path in enumerate(self.files, start=1):
            if path.suffix == '.vsd':
                try:
                    doc = visio.Documents.Open(path)
                    for page in doc.Pages:
                        page.CreateSelection(1).Copy()
                        out_doc.Pages.Item(counter).Paste()
                        pyperclip.copy('')
                        counter += 1
                except Exception as e:
                    visio.Quit()
                    msg.setText('Ошибка получения данных из файла Visio.\nПерезапустите программу!')
                    msg.setWindowTitle('Ошибка')
                    msg.exec()
            elif path.suffix == '.xlsm':
                try:
                    pass
                except Exception as e:
                    pass



    def set_ps(self, page, cell, value):
        page.PageSheet.Cells(cell).Formula = value

    def __convert__(self, doc, out_path):
        pdf_format = 1
        intent_print = 1
        print_all = 0
        for page in doc.Pages:
            # set_ps(page, "PageLeftMargin", "0mm")
            # set_ps(page, "PageRightMargin", "0mm")
            # set_ps(page, "PageTopMargin", "0mm")
            # set_ps(page, "PageBottomMargin", "0mm")
            page.ResizeToFitContents()
        doc.ExportAsFixedFormat(pdf_format, out_path, intent_print, print_all)
        doc.SaveAs(rf"{self.path_dir}\Результат.vsd")
        doc.Close()


def set_ps(page, cell, value):
    page.PageSheet.Cells(cell).Formula = value

def convert(doc, out_path):
    pdf_format = 1
    intent_print = 1
    print_all = 0
    for page in doc.Pages:
        # set_ps(page, "PageLeftMargin", "0mm")
        # set_ps(page, "PageRightMargin", "0mm")
        # set_ps(page, "PageTopMargin", "0mm")
        # set_ps(page, "PageBottomMargin", "0mm")
        page.ResizeToFitContents()
    doc.ExportAsFixedFormat(pdf_format, out_path, intent_print, print_all)
    doc.SaveAs(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\Результат.vsd")
    doc.Close()



basedir = pathlib.Path(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static")
out_path = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат2.pdf"
files = list(basedir.glob('ул. Голика.xlsm')) + list(basedir.glob('*.vsd'))
# ex_files = list(basedir.glob('ул. Голика.xlsm'))
print(files)
# print(ex_files)

visio = win.Dispatch("Visio.Application")
visio.Visible = True
excel = win.Dispatch("Excel.Application")
excel.Visible = True
out_doc = visio.Documents.Add("")
out_doc.Pages.Add()

counter = 1

for i, path in enumerate(files, start=1):
    if path.suffix == '.vsd':
        try:
            doc = visio.Documents.Open(path)
            for page in doc.Pages:
                page.CreateSelection(1).Copy()
                out_doc.Pages.Item(counter).Paste()
                pyperclip.copy('')
                counter += 1
            doc.Close()
            out_doc.Pages.Add()

        except Exception as e:
            visio.Quit()
            raise ValueError("Ошибка получения данных из Visio")

    elif path.suffix == '.xlsm':
        try:
            out_path2 = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат.pdf"
            doc = excel.Workbooks.Open(path)
            doc.ExportAsFixedFormat(0, out_path2, 0)
            # doc.SaveAs(path.with_suffix('.pdf'))
        except Exception as e:
            doc.Close()
            out_doc.Close()
            raise ValueError("Ошибка получения данных из Excel")

convert(out_doc, out_path)
visio.Quit()
excel.Quit()
