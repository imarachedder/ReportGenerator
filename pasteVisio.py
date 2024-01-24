import sys
import pyperclip
import win32com.client as win
import pathlib
from PyQt6.QtWidgets import QMessageBox
# ЗАПУСКАЕМ СНАЧАЛА ЭТОТ ЗАТЕМ MergePdfFiles.py



class ConvertVisio():
    def __init__(self, path_dir=None, filename = None):
        self.path_dir = pathlib.Path(rf"{path_dir}")
        print("Я уже здесь", self.path_dir)
        self.out_path = rf"{self.path_dir}\Результат2.pdf"
        self.out_path2 = rf"{self.path_dir}\Результат.pdf"
        print(filename)
        self.files = list(self.path_dir.glob(f'{filename}.xlsm')) + list(self.path_dir.glob('*.vsd'))
        print(self.files)
        self.pasteVisio()

    def set_ps(self, page, cell, value):
        page.PageSheet.Cells(cell).Formula = value

    def convert(self, doc, out_path):
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

    def pasteVisio(self):
        visio = win.Dispatch("Visio.Application")
        visio.Visible = True
        excel = win.Dispatch('Excel.Application')
        excel.Visible = True
        out_doc = visio.Documents.Add('')
        out_doc.Pages.Add()
        counter = 1
        for i, path in enumerate(self.files, start=1):
            print(path)
            if path.suffix == '.vsd':
                try:
                    doc = visio.Documents.Open(path)
                    for page in doc.Pages:
                        # page.ResizeToFitContents()
                        page.CreateSelection(1).Copy()
                        out_doc.Pages.Item(counter).Paste()
                        pyperclip.copy('')
                        counter += 1
                    doc.Close()
                    out_doc.Pages.Add()
                except Exception as e:
                    doc.Close()
                    visio.Quit()
                    excel.Quit()
                    out_doc.Close()
                    msg = QMessageBox()
                    msg.setText('Ошибка получения данных из файла Visio.\nПерезапустите программу!')
                    msg.setWindowTitle('Ошибка')
                    msg.exec()
            elif path.suffix == '.xlsm':
                try:
                    doc = excel.Workbooks.Open(path)
                    doc.ExportAsFixedFormat(0, self.out_path2, 0)
                except Exception as e:
                    doc.Close()
                    visio.Quit()
                    excel.Quit()
                    out_doc.Close()
                    msg = QMessageBox()
                    msg.setText('Ошибка получения данных из файла Excel.\nПерезапустите программу!')
                    msg.setWindowTitle('Ошибка')
                    msg.exec()
        self.convert(out_doc, self.out_path)
        visio.Quit()
        excel.Quit()
        # setting Message box window title
        msg = QMessageBox()
        msg.setText("Файл успешно сохранен")
        msg.setWindowTitle("Файл сохранен")
        msg.exec()


##################################################################################
########################### ОСНОВНОЕ РАБОЧИЙ КОД #################################
##################################################################################

#
# def set_ps(page, cell, value):
#     page.PageSheet.Cells(cell).Formula = value
#
# def convert(doc, out_path):
#     pdf_format = 1
#     intent_print = 1
#     print_all = 0
#     for page in doc.Pages:
#         # set_ps(page, "PageLeftMargin", "0mm")
#         # set_ps(page, "PageRightMargin", "0mm")
#         # set_ps(page, "PageTopMargin", "0mm")
#         # set_ps(page, "PageBottomMargin", "0mm")
#         page.ResizeToFitContents()
#     doc.ExportAsFixedFormat(pdf_format, out_path, intent_print, print_all)
#     doc.SaveAs(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\Результат.vsd")
#     doc.Close()
#
#
#
# basedir = pathlib.Path(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static")
# out_path = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат2.pdf"
# files = list(basedir.glob('ул. Голика.xlsm')) + list(basedir.glob('*.vsd'))
# # ex_files = list(basedir.glob('ул. Голика.xlsm'))
# print(files)
# # print(ex_files)
#
# visio = win.Dispatch("Visio.Application")
# visio.Visible = True
# excel = win.Dispatch("Excel.Application")
# excel.Visible = True
# out_doc = visio.Documents.Add("")
# out_doc.Pages.Add()
#
# counter = 1
#
# for i, path in enumerate(files, start=1):
#     if path.suffix == '.vsd':
#         try:
#             doc = visio.Documents.Open(path)
#             for page in doc.Pages:
#                 page.CreateSelection(1).Copy()
#                 out_doc.Pages.Item(counter).Paste()
#                 pyperclip.copy('')
#                 counter += 1
#             doc.Close()
#             out_doc.Pages.Add()
#
#         except Exception as e:
#             visio.Quit()
#             raise ValueError("Ошибка получения данных из Visio")
#
#     elif path.suffix == '.xlsm':
#         try:
#             out_path2 = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат.pdf"
#             doc = excel.Workbooks.Open(path)
#             doc.ExportAsFixedFormat(0, out_path2, 0)
#             # doc.SaveAs(path.with_suffix('.pdf'))
#         except Exception as e:
#             doc.Close()
#             out_doc.Close()
#             raise ValueError("Ошибка получения данных из Excel")
#
# convert(out_doc, out_path)
# visio.Quit()
# excel.Quit()
