# from openpyxl.chart import BarChart, Reference
import glob

import win32com.client
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

import db
from settings import path_template_excel


class WriterExcel:
    def __init__(self, data: dict = None):
        if data is None:
            self.data = {}
        self.db = db.Query()# нужно ли ?
        # self.info_window2 = window2.Window2().get_info_window2()
        # self.info_window3 = window3.Window3().get_info_from_plainTextEdit()
        self.wb = load_workbook(path_template_excel, keep_vba = True)
        self.path_dir = r"C:\Users\sibregion\Desktop\test\report\static"
       #self.img = Image(f"{self.path_dir}\схема.png")
        self.data = data
        self.page_number = 2

    def save_file (self):
        # сохранить файл
        self.wb.save(rf'{self.path_dir}\{self.data.get("название дороги", "Отчет")}.xlsm')
        #self.close_file()

    def close_file (self):
        # закрыть файл
        self.wb.close()


class WriterExcelTP(WriterExcel):
    def __init__ (self, data: dict = None):
        super().__init__(data)

        #convert_visio2svg(self.path_dir)

    def write_titular (self):
        """
        Заполняет лист 'Титульник (без рамки)'
        :return:
        """
        ws = self.wb['Титульник (без рамки)']  # выбираем лист

        ws["B4"].value = self.data.get('client')
        ws["B22"].value = self.data.get('name_road')
        ws["B31"].value = f"составлена на {self.data.get('year')} г."
        ws["B33"].value = f"Шифр:{self.data.get('cypher')}"
        ws["B52"].value = f"Омск:{self.data.get('year')}"
        ws["B41"].value = self.data.get('contractor')
        ws["B46"].value = f'{self.data.get("position_contractor")} ' \
                          f'{self.data.get("fio_contractor")}________________________'
        ws["AI41"].value = self.data.get('client')
        ws["AI46"].value = f'{self.data.get("position_client")} {self.data.get("fio_client")}________________________'

    def write_scheme (self):
        """
         Заполняет лист "схема"
        :return: None
        """
        schema = Image(f"{self.path_dir}\схема.png")
        ws = self.wb['Схема']  # выбираем лист
        schema.width = 1380
        schema.height = 800
        ws.add_image(schema, 'B5')

    def write_6 (self):
        """
        заполняет лист "6"
        :return:
        """

        ws = self.wb['6']  # выбираем лист
        n = 1  # счетчик
        # 2.1 Наименование дороги: name road
        ws["O5"].value = self.data.get('название дороги')

        for i in range(self.data.get('count_region', 1)):
            ws[f"L1{i}"].value = self.data.get('участки', 'None')[i]  # 2.2 Участок дороги: участки

        ws["S14"].value = self.data.get('протяженность дороги', None)  # 2.3 Протяженность дороги: протяженность

        # заполняет таблицу 2.3 Протяженность дороги
        for i in range(2, (self.data.get('count_region', 1) * 2) + 1, 2):
            ws[f'B2{i - 1}'].value = f'Участок {n}'
            ws[f'B2{i}'].value = self.data.get('начало дороги', None)
            ws[f'F2{i}'].value = self.data.get('конец дороги', None)
            ws[f'J2{i}'].value = self.data.get(' дороги', None)
            ws[f'N2{i}'].value = self.data.get('подъездов', None)
            ws[f'R2{i}'].value = self.data.get('дороги вместе с подъездами', None)
            ws[f'V2{i}'].value = self.data.get('обслуживаемых дорожной организацией', None)
            ws[f'AB2{i}'].value = self.data.get('находящихся в ведении города', None)
            ws[f'AG2{i}'].value = self.data.get('совмещенных', None)
            n += 1

        n = 1
        # заполняет таблицу 2.4 Наименование подъездов (обходов) и их протяженность
        ws["B37"].value = self.data.get('подъезды', None)
        # заполняет таблицу 2.5 Категория дороги (участка), подъездов
        ws["AL10"].value = self.data.get('name_road', None)

        # заполняет таблицу 2.6 Краткая историческая справка
        ws["AL33"].value = self.data.get('history_match', None)

    def write_11 (self):
        ws = self.wb['11']
        # заполнение таблицы 4.6
        ws["AU25"] = self.data.get('Автобусные остановки', '-')
        sum_sign = {'всего': 0,
                    'предупреждающие': 0,
                    'приоритета': 0,
                    'запрещающие': 0,
                    'предписывающие': 0,
                    'особых предписаний': 0,
                    'сервиса': 0,
                    'информационные': 0,
                    'дополнительной информации': 0}
        for k, v in self.data.items():
            if k[0].isdigit():
                sum_sign['всего'] += len(v['Направление'])
                if k[0] == '1':
                    sum_sign['предупреждающие'] += len(v['Направление'])
                elif k[0] == '2':
                    sum_sign['приоритета'] += len(v['Направление'])
                elif k[0] == '3':
                    sum_sign['запрещающие'] += len(v['Направление'])
                elif k[0] == '4':
                    sum_sign['предписывающие'] += len(v['Направление'])
                elif k[0] == '5':
                    sum_sign['особых предписаний'] += len(v['Направление'])
                elif k[0] == '6':
                    sum_sign['информационные'] += len(v['Направление'])
                elif k[0] == '7':
                    sum_sign['сервиса'] += len(v['Направление'])
                elif k[0] == '8':
                    sum_sign['дополнительной информации'] += len(v['Направление'])

        ws['AU30'] = sum_sign.get('всего', '-')
        ws['AU32'] = sum_sign.get('предупреждающие', '-')
        ws['AU33'] = sum_sign.get('приоритета', '-')
        ws['AU34'] = sum_sign.get('запрещающие', '-')
        ws['AU35'] = sum_sign.get('предписывающие', '-')
        ws['AU36'] = sum_sign.get('особых предписаний', '-')
        ws['AU37'] = sum_sign.get('информационные', '-')
        ws['AU38'] = sum_sign.get('сервиса', '-')
        ws['AU39'] = sum_sign.get('дополнительной информации', '-')

    def write_linear_graphs (self):
        for i in range(len(glob.glob("*.txt"))):
            print(i)
        linear_graph = Image(f"{self.path_dir}\схема.png")
        self.wb.create_sheet(f'Линейный график {i}')
        ws = self.wb.create_sheet('Students')  # выбираем лист
        self.img.width = 1380
        self.img.height = 800
        ws.add_image(self.img, 'B5')


class WriterExcelDAD(WriterExcel):
    def __init__ (self, data: dict = None):
        super().__init__(data)

    def write_titular (self):
        pass

    def write_scheme (self):
        pass

    # def write_diagrams1 (self):
    #     """
    #     Разница ширины проезжей части от требуемого значения по расстоянию
    #     """
    #     wb = Workbook()
    #     ws = wb.active
    #
    #     # данные для построения диаграмм
    #     rows = [
    #         ('Number', 'Batch 1', 'Batch 2'),
    #         (2, 10, 30),
    #         (3, 40, 60),
    #         (4, 50, 70),
    #         (5, 20, 10),
    #         (6, 10, 40),
    #         (7, 50, 30),
    #     ]
    #     for row in rows:
    #         ws.append(row)
    #
    #     # ДИАГРАММА №1
    #     # создаем объект диаграммы
    #     chart1 = BarChart()
    #     # установим тип - `вертикальные столбцы`
    #     chart1.type = "col"
    #     # установим стиль диаграммы (цветовая схема)
    #     chart1.style = 10
    #     # заголовок диаграммы
    #     chart1.title = "Столбчатая диаграмма"
    #     # подпись оси `y`
    #     chart1.y_axis.title = 'Длина выборки'
    #     # показывать данные на оси (для LibreOffice Calc)
    #     chart1.y_axis.delete = False
    #     # подпись оси `x`
    #     chart1.x_axis.title = 'Номер теста'
    #     chart1.x_axis.delete = False
    #     # выберем 2 столбца с данными для оси `y`
    #     data = Reference(ws, min_col = 2, max_col = 3, min_row = 1, max_row = 7)
    #     # теперь выберем категорию для оси `x`
    #     categor = Reference(ws, min_col = 1, min_row = 2, max_row = 7)
    #     # добавляем данные в объект диаграммы
    #     chart1.add_data(data, titles_from_data = True)
    #     # установим метки на объект диаграммы
    #     chart1.set_categories(categor)
    #     # добавим диаграмму на лист, в ячейку "A10"
    #     ws.add_chart(chart1, "A10")
    #
    #     # ДИАГРАММА №2
    #     # что бы показать типы столбчатых диаграмм, скопируем
    #     # первую диаграмму и будем менять настройки
    #     chart2 = deepcopy(chart1)
    #     # изменяем стиль
    #     chart2.style = 11
    #     # установим тип - `горизонтальные полосы`
    #     chart2.type = "bar"
    #     chart2.title = "Горизонтальные полосы"
    #     ws.add_chart(chart2, "A25")
    #
    #     # ДИАГРАММА №3
    #     chart3 = deepcopy(chart1)
    #     chart3.type = "col"
    #     chart3.style = 12
    #     # зададим группировку
    #     chart3.grouping = "stacked"
    #     # для диаграммы с группировкой,
    #     # необходимо установить перекрытие
    #     chart3.overlap = 100
    #     chart3.title = 'Сложенная диаграмма'
    #     ws.add_chart(chart3, "A40")
    #
    #     # ДИАГРАММА №4
    #     chart4 = deepcopy(chart1)
    #     chart4.type = "bar"
    #     chart4.style = 13
    #     chart4.grouping = "percentStacked"
    #     chart4.overlap = 100
    #     # отключим линии сетки
    #     chart4.y_axis.majorGridlines = None
    #     # уберем легенду
    #     chart4.legend = None
    #     chart4.title = 'Диаграмма с процентным накоплением'
    #     ws.add_chart(chart4, "A55")


def convert_visio2svg (path_dir):
    """конвертируем визио в svg"""
    visio = win32com.client.Dispatch("Visio.Application")
    doc = visio.Documents.Open(rf"{path_dir}\линейный график.vsd")

    for page in doc.Pages:
        page.Export(rf'{path_dir}\линейный график\линейный график{page.Name}.png')
    visio.Quit()


def main ():
    conn = db.Query('OMSK_CITY_2023')
    data = conn.get_tp_datas('ул. П. Некрасова')

    report = WriterExcelTP(data)
    # report.write_titular()
    # report.write_scheme()
    # report.write_11()
    # report = WriterExcelDAD()
    # report.write_diagrams1()
    report.save_file()


if __name__ == "__main__":
    import time

    start = time.time()  # точка отсчета времени
    main()
    end = time.time() - start  # собственно время работы программы
    print(end)  # вывод времени
