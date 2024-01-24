import os
import pathlib
import win32com.client
from PyQt6.QtWidgets import QMessageBox

class MergeFiles():
    def __init__(self, path_dir = None, filename = None):
        self.path_dir = pathlib.Path(path_dir)
        self.filename = filename
        self.pdf_files_path = list(self.path_dir.glob("*.pdf"))
        self.out_file = rf"{self.path_dir}\{filename}.pdf"
        self.msg = QMessageBox()
        self.mergePdfFiles()

    def mergePdfFiles(self, pdf_files = None, out_file = None):
        """
        Объединение PDF-файлов, указанных в pdf_files, в один файл
        @:param pdf_files: Путь к PDF-файлам для объединения
        @:type pdf_files: List or str

        @:param out_file: Путь куда сохранять PDF-файл
        @:type out_file: str
        :return:
        """

        if len(self.pdf_files_path) < 2:
            self.msg.setText("Необходимо минимум 2 PDF-файла")
            self.msg.setWindowTitle("Ошибка сшивки данных")
            self.msg.exec()

        pdf_files = []
        for pdf_file_path in sorted(self.pdf_files_path):
            if not os.path.exists(pdf_file_path):
                self.msg.setWindowTitle("Ошибка файла")
                self.msg.setText(f"PDF-файл не найден: {pdf_file_path}")
                self.msg.exec()

            pdf_file = win32com.client.Dispatch('AcroExch.PDDoc')
            pdf_file.Open(pdf_file_path)
            pdf_files.append(pdf_file)

        output_pdf_file = win32com.client.Dispatch('AcroExch.PDDoc')
        output_pdf_file.Create()

        for pdf_file in pdf_files:
            num_pages = pdf_file.GetNumPages()
            print(num_pages)

            output_pdf_file.InsertPages(output_pdf_file.GetNumPages()-1, pdf_file, 0 , num_pages, 0)

        output_pdf_file.Save(1, self.out_file)
        output_pdf_file.Close()
        for pdf_file in pdf_files:
            pdf_file.Close()



#######################################################################
#################### - ОСНОВНОЙ РАБОЧИЙ КОД - #########################
#######################################################################
# def merge_pdf_files(pdf_file_paths, output_file_path):
#     """
#     Объединение PDF файлов, указанных в pdf_file_paths, в один файл
#     @param pdf_file_paths: Путь к PDF-файлом для объединения
#     @type pdf_file_paths: List or str
#
#     @param output_file_path: Путь куда сохранять PDF-файл
#     @type output_file_path: str
#     """
#     if len(pdf_file_paths) < 2:
#         raise ValueError("Необходимо указать минимум 2 PDF - файла")
#
#     pdf_files = []
#     for pdf_file_path in sorted(pdf_file_paths):
#         if not os.path.exists(pdf_file_path):
#             raise ValueError(f"PDF - файл не найден: {pdf_file_path}")
#
#         pdf_file = win32com.client.Dispatch('AcroExch.PDDoc')
#         pdf_file.Open(pdf_file_path)
#         pdf_files.append(pdf_file)
#
#     output_pdf_file = win32com.client.Dispatch('AcroExch.PDDoc')
#     output_pdf_file.Create()
#
#     for pdf_file in pdf_files:
#         num_pages = pdf_file.GetNumPages()
#         print(num_pages)
#
#         output_pdf_file.InsertPages(output_pdf_file.GetNumPages() - 1 , pdf_file, 0, num_pages, 0)
#
#         # for i in range(0, num_pages + 1):
#         #     print(i)
#         #     page = pdf_file.AcquirePage(i)
#         #     print(type(page))
#         #     output_pdf_file.InsertPages(output_pdf_file.GetNumPages(), page, 0, num_pages, 0)
#             # pdf_file.DeletePages(-1, i)
#
#         # pdf1 = win32com.client.Dispatch('AcroExch.PDDoc')
#         # pdf1.Open(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат2.pdf")
#         # pdf2 = win32com.client.Dispatch('AcroExch.PDDoc')
#         # pdf2.Open(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат.pdf")
#         #
#         # pdf1.InsertPages(pdf1.GetNumPages() - 1, pdf2, 0, pdf2.GetNumPages(), 0)
#         # pdf1.Save(1, output_files_path)
#
#     output_pdf_file.Save(1, output_file_path)
#     output_pdf_file.Close()
#
#     for pdf_file in pdf_files:
#         pdf_file.Close()
#
# # Usage
#
# basedir = pathlib.Path(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static")
# pdf_file_paths = list(basedir.glob('*.pdf'))
# output_files_path = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат3.pdf"
#
# merge_pdf_files(pdf_file_paths, output_files_path)


# pdf1 = win32com.client.Dispatch('AcroExch.PDDoc')
# pdf1.Open(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат2.pdf")
# pdf2 = win32com.client.Dispatch('AcroExch.PDDoc')
# pdf2.Open(r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static\Результат.pdf")
#
# pdf1.InsertPages(pdf1.GetNumPages()- 1, pdf2, 0, pdf2.GetNumPages(), 0)
# pdf1.Save(1, output_files_path)
# pddoc2.InsertPages(N2 - 1, pddoc1, 0, N1, 0)