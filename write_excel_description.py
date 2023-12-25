# from openpyxl.chart import BarChart, Reference
import glob
import warnings

import win32com.client
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Side, Border
import db
from settings import path_template_excel, path_template_excel_application
from PyQt6.QtWidgets import QMessageBox
from settings import path_icon_app
class WriterExcel:
    def __init__(self, data: dict = None, path_template_excel = path_template_excel, path = None, data_interface = {}):
        if data is None:
            self.data = {}
        self.wb = load_workbook(path_template_excel, keep_vba = True)

        self.path_dir = path
        self.data = data
        self.data_interface = data_interface
        self.msg = QMessageBox()

    def save_file(self):
        # сохранить файл
        try:
            if len(self.data.get('название дороги')) > 50 or '/' in self.data.get('название дороги'):
                self.wb.save(rf'{self.path_dir}\{self.data.get("название дороги", "Отчет")[:40]}.xlsm')
                self.close_file()
            else:
                self.wb.save(rf'{self.path_dir}\{self.data.get("название дороги", "Отчет")}.xlsm')
                self.close_file()
        except Exception as e:
            # pass
            self.msg.setText(f"Не удалось сохранить файл {e}")
            self.msg.setWindowTitle(f"{self.data.get('название дороги')}")
            self.msg.exec()

    def close_file(self):
        # закрыть файл
        self.wb.close()

class WriterExcelTP(WriterExcel):
    # todo: глянуть измененные ячейки в новом шаблоне
    def __init__(self, data: dict = None, path = None, data_interface ={}):
        super().__init__(data = data, path=path, data_interface=data_interface)
        # self.msg = QMessageBox()
        # self.msg.setIcon(QMessageBox.Icon.Critical)
        self.write_titular()
        self.write_scheme()
        print(self.data_interface)
        print("Начал работать")
        self.write_6()
        self.write_7()
        self.write_8()
        self.write_9()
        self.write_10()
        self.write_11()
        self.write_12()
        self.write_13()
        self.write_14()
        self.write_17()
        self.write_18()

    def write_pereplet(self):
        """
        Заполняет Титульник для переплета
        :return:
        """
        ws = self.wb['Переплет'] # выбираем лист
        # ws.header_footer
        ws['B22'].value = self.data.get('название дороги')
    def write_titular(self):
        """
        Заполняет лист 'Титульник (без рамки)'
        :return:
        """
        ws = self.wb['Титульник (без рамки)']  # выбираем лист

        ws["B4"].value = self.data_interface.get('client', None)
        ws["B22"].value = ''.join(self.data.get('название дороги').split('_')[-1])
        ws["B31"].value = f"составлена на {self.data_interface.get('year', None)} г."
        ws["B33"].value = f"Шифр:{self.data_interface.get('cypher', None)}"
        ws["B52"].value = f"Омск - {self.data_interface.get('year', None)} г."
        ws["B41"].value = self.data_interface.get('contractor', None)
        ws["B46"].value = f'{self.data_interface.get("position_contractor", None)} ' \
                          f'{self.data_interface.get("fio_contractor", None)}________________________'
        ws["AI41"].value = self.data_interface.get('client', None)
        ws["AI46"].value = f'{self.data_interface.get("position_client", None)} {self.data_interface.get("fio_client", None)}' \
                           f'________________________'

    def write_scheme(self, ):
        """
         Заполняет лист "схема"
        :return: None
        """
        try:
            schema = Image(f"{self.path_dir}\схема.png")
            ws = self.wb['Схема']  # выбираем лист
            # schema.width = 1380
            # schema.height = 800
            ws.add_image(schema, 'B5')
        except Exception as e:
            self.msg.setText("Не найдена схема")
            self.msg.setWindowTitle("Ошибка")
            self.msg.exec()

    def write_6(self):
        """
        Заполняет лист "6"
        :return:
        """
        # print(self.data.get(f'участок {1}', 'весь участок').get('Ось дороги', None))
        ws = self.wb['6']  # выбираем лист
        n, i = 1, 2  # счетчик
        res = 0
        # 2.1 Наименование дороги: name road
        ws["O5"].value = ''.join(self.data.get('название дороги').split('_')[-1])
        try:
        # 2.2 Участок дороги 1, 2 и т.д., 2.3 протяженность дороги(участка) и 2.5 категория дороги(участка), подъездов
            for key, value in self.data.items():
                if key == 'название дороги':
                    ws["AL10"] = ''.join(self.data.get('название дороги').split('_')[-1])
                    continue
                else:
                    if len(self.data) > 2:
                        ws[f'B1{n - 1}'] = f'2.2 Участок дороги: {n}' if n - 1 == 0 else f'      Участок дороги: {n}'
                        ws[f'B2{i - 1}'] = f'Участок: {n}'
                    else:
                        ws[f'B1{n - 1}'] = f'2.2 Участок дороги:'
                        ws[f'B2{i - 1}'] = f'Весь участок'

                    if n % 2 != 0:
                        axis_null = self.data.get(key).get('Ось дороги', {}).get('Начало трассы', [])[i % (len(self.data) - 1)][7]
                        axis_length = self.data.get(key).get('Ось дороги', {}).get('Начало трассы', [])[i % (len(self.data) - 1)][8]
                        ws[f'B2{i}'] = f"{axis_null} + {axis_null}00"
                        ws[f'F2{i}'] = f"{axis_null} + {axis_length if axis_length > 100 else f'0{axis_length}'}"
                        ws[f'J2{i}'] = round(axis_length/1000,3)
                    else:
                        axis_null = self.data.get(key).get('Ось дороги', {}).get('Начало трассы', [])[i % (len(self.data) - 1) - 1][7]
                        axis_length = self.data.get(key).get('Ось дороги', {}).get('Начало трассы', [])[i % (len(self.data) - 1) - 1][8]
                        ws[f'B2{i}'] = f"{axis_null} + {axis_null}00"
                        ws[f'F2{i}'] = f"{axis_null} + {axis_length if axis_length > 100 else f'0{axis_length}'}"
                        ws[f'J2{i}'] = round(axis_length/1000,3)

                    ax_null = self.data.get(key).get("Ось дороги", {}).get("Начало трассы", [])[0][7]
                    ax_max = self.data.get(key).get("Ось дороги", {}).get("Начало трассы", [])[0][8]
                    ws[f'L1{n - 1}'] = f"от КМ {ax_null} + 000 до КМ 0 + {ax_max if ax_max > 100 else f'0{ax_max}'}"
                    ws[f'AW1{i - 1}'] = f"{ax_null} + {ax_null}00"
                    ws[f'BA1{i - 1}'] = f"{ax_null} + {ax_max if ax_max > 100 else f'0{ax_max}'}"

                    # todo: ПРИДУМАТЬ УНИВЕРСАЛЬНОЕ РЕШЕНИЕ К КАТЕГОРИЯМ УЧАСТКА ДОРОГИ, НАПИСАТЬ ЛОГИКУ, УСЛОВИЯ
                    try:
                        ws[f'BE1{i - 1}'] = f"{self.data.get(key).get('Граница участка дороги', '-').get('категория дорог/улиц', {})[0][0]}"
                    except Exception as e:
                        self.msg.setWindowTitle("Ошибка")
                        self.msg.setText("Ошибка данных по категории А/Д")
                        self.msg.exec()

                    res += int(self.data.get(key).get('Ось дороги', {}).get('Начало трассы', [])[0][8])
                    n += 1
                    i += 2
            else:
                ws['S14'] = f"{round(res/1000, 3)} км"

        except Exception as e:
            pass
            # self.msg.setText(f"Ошибка {e}")
            # self.msg.setWindowTitle("Ошибка - 6 лист")
            # self.msg.exec()

        # заполняет таблицу 2.4 Наименование подъездов (обходов) и их протяженность
        # ws["B37"].value = self.data.get('подъезды', {}).get('Наименование', [])

        # заполняет таблицу 2.6 Краткая историческая справка
        # ws["AL33"].value = self.data_interface.get("history_match")
        ws["AL33"].value = self.data_interface.get('history_match', None)

    def write_7 (self):
        # 2.7
        ws = self.wb['7']
        counter_distr_soder = 15  # счетчик строк для 2.7
        column_tuple = ('AX', 'AZ', 'BB', 'BD', 'BF', 'BH', 'BJ', 'BL', 'BN')  # столбцы для 2.8

        row_name_distr = 15  # счетчик строк для 2.8
        for k1, v1 in self.data.items():
            if k1 == 'название дороги':
                continue
            try:
                for idx, v2 in enumerate(v1.get('Дорожная организация', {}).get('Наименование', [])):
                    ws[f'B{counter_distr_soder}'] = self.data_interface.get('year', '')
                    ws[f'E{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Наименование', [])[idx][
                        0] if v1.get('Дорожная организация', {}).get('Наименование', []) else ''
                    ws[f'l{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Адрес', [])[idx][0] if v1.get(
                        'Дорожная организация', {}).get('Адрес', []) else ''
                    ws[f'P{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Город', [])[idx][0] if v1.get(
                        'Дорожная организация', {}).get('Город', []) else ''
                    ws[f'V{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Начало по оси', [])[idx][
                        0] if v1.get('Дорожная организация', {}).get('Начало по оси', []) else ''
                    ws[f'Y{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Конец  по оси', [])[idx][
                        0] if v1.get(
                        'Дорожная организация', {}).get('Конец  по оси', []) else ''

                    # начало и конецпо оси должны быть записаны с километровой привязкой км+м
                    start = v1.get('Дорожная организация', {}).get('Начало по оси', [])[idx][0].split('+')
                    end = v1.get('Дорожная организация', {}).get('Конец  по оси', [])[idx][0].split('+')

                    ws[f'AB{counter_distr_soder}'] = f'{((int(end[0]) - int(start[0])) * 1000 + int(end[1]) - int(start[1])) / 1000}'
                    ws[f'AK{counter_distr_soder}'] = f'=AB{counter_distr_soder}'
                    counter_distr_soder += 1
            except Exception as e:

                self.msg.setText("Ошибка заполнения таблицы 2.7")
                self.msg.setWindowTitle("Ошибка в листе 7")
                self.msg.exec()

            # 2.8 Таблица основных расстояний (в целых километрах)
            tuple_name = tuple(v1.get('Населенный пункт', {}).get('Наименование', []))
            try:
                for idx, name in enumerate(tuple_name):

                    ws[f'{column_tuple[idx]}4'] = name[0]
                    ws[f'AR{row_name_distr}'] = name[0]
                    # итератор столбцов, споймает ошибку если  населенных пунктов будет больше чем указанных столобцов
                    iter_column = iter(column_tuple[:len(tuple_name)])

                    for name1 in tuple_name:
                        '''
                        заполнение расстояний между населенными пунктами, в целых километрах. next(iter) возвращает каждый 
                        раз новый столбец
                        '''
                        ws[f'{next(iter(iter_column))}{row_name_distr}'] = abs(
                            (int(name1[-4]) - int(name[-4])) * 1000 + int(name1[-3]) - int(name[-3])) // 1000 if name1[2] - \
                                                                                                                 name[
                                                                                                                     2] != 0 else '-'
                    row_name_distr += 1
            except Exception as e:
                pass
                self.msg.setText("Ошибка заполнения таблицы 2.8")
                self.msg.setWindowTitle("Ошибка в листе 7")
                self.msg.exec()

    def write_8(self):
        """
        Расписываем экономическую характеристику
        :param data:
        :return:
        """

        # Счетчик
        n = 1

        # Выбираем лист
        ws = self.wb['8']
        # 3.1 Экономическое и административное значение дороги
        ws['B6'] = self.data_interface.get('economical_characteristic_road', '')
        # 3.2 Связь дороги с ж/д и водными путями и автомобильными дорогами
        ws['B19'] = self.data_interface.get('railway_waterway', '')
        # 3.3 Характеристика движения, его сезонность и перспектива роста
        ws['B33'] = self.data_interface.get('movement_characteristic', '')
        # 3.4 Среднесуточная интенсивность движения по данным учета

    def write_9(self):
        # todo: посмотреть правильность расчетов ширин ПЧ
        """
        Техническая характеристика
        :param res: значение макс значения оси
        :param data:
        :return:
        """

        # Функция для расчета ширины проезжей части
        def calcLengthOfTheWidthOfTheCarriageWay(res, j, key, v):
            if v == 'Ширина земляного полотна':
                result = res - int(self.data.get(key).get(v, {}).get('Ширина')[j - 1][8])
                # if j - 1 == 0:
                #     result += int(self.data.get(key).get(v, {}).get('Ширина')[j - 1][8])
                return result

            elif v == 'Ширина проезжей части':
                result = res - int(self.data.get(key).get(v, {}).get('Ширина ПЧ')[j - 1][8])
                if j - 1 == 0:
                    result += int(self.data.get(key).get(v, {}).get('Ширина ПЧ')[j - 1][8])
                return result
                # print(f"do {res}",
                #       self.data.get(key).get(v, None).get('Ширина', 'Ширина ПЧ')[
                #           j - 1][0],
                #       int(self.data.get(key).get(v, None).get('Ширина', 'Ширина ПЧ')[
                #           j - 1][8]), result, res)

        # Счетчик
        n, i = 12, 34

        ws = self.wb['9']
        ws['B7'] = self.data_interface.get("area_conditions", None)
        # 4.1 Топографические условия района проложения автомобильной дороги
        # ws['B7'] = self.data.get('area_conditioins')
        # 4.2 Ширина земляного полотна

        try:
            for key, val in self.data.items():
                if key == 'название дороги':
                    continue
                else:
                    if len(self.data) > 2:
                        ws[f'AJ1{n + 1}'] = f'Участок {i + 1}'
                    else:
                        ws[f'AJ1{n + 1}'] = f'Весь участок'
                    i += 1
                    n += 1
                    # Создаем переменные для ячеек в таблице 4.3.1 Ширина проезжей части
                    sum1, sum2, sum3, sum4, sum5, sum6 = 0, 0, 0, 0, 0, 0
                    res2, res3, res4, res5, res6, res7, res8, res9, res10, res11, res12 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                    res = val.get('Ось дороги').get('Начало трассы')[0][8]
                    # 4.2 Ширина земляного полотна
                    try:
                        for j in range(len(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])), 0, -1):
                            if float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]) <= 8.0:
                                sum1 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            elif 8.0 < float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]) < 10.0:
                                sum2 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            elif 10.0 <= float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]) < 12.0:
                                sum3 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            elif 12.0 <= float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]) < 15.0:
                                sum4 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            elif 15.0 <= float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]) < 27.5:
                                sum5 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            elif 27.5 <= float(self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][0]):
                                sum6 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина земляного полотна')
                            res = self.data.get(key).get('Ширина земляного полотна', {}).get('Ширина', [])[j - 1][8]
                            ws[f'G3{5}'].value = None if sum1 == 0 else round(sum1 / 1000, 3)
                            ws[f'K3{7}'].value = None if sum2 == 0 else round(sum2 / 1000, 3)
                            ws[f'P3{9}'].value = None if sum3 == 0 else round(sum3 / 1000, 3)
                            ws[f'U4{1}'].value = None if sum4 == 0 else round(sum4 / 1000, 3)
                            ws[f'Z4{3}'].value = None if sum5 == 0 else round(sum5 / 1000, 3)
                            ws[f'AE4{5}'].value = None if sum6 == 0 else round(sum6 / 1000, 3)
                    except Exception as e:
                        # self.msg.setText(f"Ошибка 4.2")
                        # self.msg.setWindowTitle("Ошибка в 9 листе")
                        # self.msg.exec()
                        pass


                    # 4.3 Характеристика проезжей части
                    # 4.3.1 Ширина проезжей части
                    res = self.data.get(key).get('Ось дороги').get('Начало трассы')[0][8]
                    for j in range(len(self.data.get(key).get('Ширина проезжей части', {}).get('Ширина ПЧ', [])), 0, -1):
                        if float(self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) <= 4.0:
                            res2 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 4.0 < float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 4.5:
                            res3 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 4.5 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 6.0:
                            res4 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 6.0 <= float(self.data.get(key).get('Ширина проезжей части').get('Ширина ПЧ')[j - 1][0]) < 6.6:
                            res5 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 6.6 <= float(self.data.get(key).get('Ширина проезжей части').get('Ширина ПЧ')[j - 1][0]) < 7.0:
                            res6 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 7.0 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 7.5:
                            res7 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 7.5 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 9.1:
                            res8 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 9.1 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 10.0:
                            res9 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 10.0 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]) < 15.1:
                            res10 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')
                        elif 15.1 <= float(
                                self.data.get(key).get('Ширина проезжей части', None).get('Ширина ПЧ')[j - 1][0]):
                            res11 += calcLengthOfTheWidthOfTheCarriageWay(res, j, key, 'Ширина проезжей части')

                        res = self.data.get(key).get('Ширина проезжей части').get('Ширина ПЧ')[j - 1][8]
                        ws[f'AL{n}'].value = None if res2 == 0 else round(res2 / 1000, 3)
                        ws[f'AO{n}'].value = None if res3 == 0 else round(res3 / 1000, 3)
                        ws[f'AR{n}'].value = None if res4 == 0 else round(res4 / 1000, 3)
                        ws[f'AU{n}'].value = None if res5 == 0 else round(res5 / 1000, 3)
                        ws[f'AX{n}'].value = None if res6 == 0 else round(res6 / 1000, 3)
                        ws[f'BA{n}'].value = None if res7 == 0 else round(res7 / 1000, 3)
                        ws[f'BD{n}'].value = None if res8 == 0 else round(res8 / 1000, 3)
                        ws[f'BG{n}'].value = None if res9 == 0 else round(res9 / 1000, 3)
                        ws[f'BJ{n}'].value = None if res10 == 0 else round(res10 / 1000, 3)
                        ws[f'BM{n}'].value = None if res11 == 0 else round(res11 / 1000, 3)
                    n += 1
                    i += 1
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 9 листе")
            self.msg.exec()

    def count_coating(self, v):
        """
        Расчет протяженностей типов покрытий. Для расчета нужны объекты - граница участков дороги
        @param: v
        @return: type_of_coating
        """
        capital = {'цементобетон': 0,
                   'асфальтобетон': 0,
                   'щебень/гравий, обр.вяжущий': 0,
                   'щебень/гравий': 0,
                   'грунт': 0,
                   'ж/б плиты': 0,
                   'булыжник': 0,
                   'брусчатка': 0,
                   'тротуарная плитка': 0}
        lightweight = {
            'асфальтобетон': 0,
            'органоминеральные': 0,
            'щебеночные (гравийные), обработанные вяжущим': 0,
            'щебень/гравий': 0
        }
        transition = {
            'щебень/гравий': 0,
            'Грунт и малопрочные каменные материалы, укрепленные вяжущим': 0,
            'Грунт, укрепленный различными вяжущими и местными материалами': 0,
            'Булыжный и колотый камень(мостовые)': 0
        }
        lower = {'грунт': 0,
                 'Грунт естественный': 0
                 }
        type_of_coating = {'Капитальный': capital,
                           'Облегченный': lightweight,
                           'Переходный': transition,
                           'Низший': lower
                           }

        tuple_tip = v.get('Граница участка дороги', {}).get('тип дорожной одежды', [])
        # print(tuple_tip)
        tuple_variant = v.get('Граница участка дороги', {}).get('вид покрытия', [])
        # print(tuple_variant)
        for idx, tip in enumerate(tuple_tip):
            # находим следующий километровый
            if tip == tuple_tip[-1]:
                next_tip = tuple_tip[-1]
            elif tip == tuple_tip[0]:
                next_tip = tuple_tip[1]
            else:
                next_tip = tuple_tip[idx % len(tuple_tip) + 1]
            try:
                type_of_coating[tip[0]][tuple_variant[idx][0]] += ((next_tip[-2][0] - tip[-2][0]) * 1000 + (
                        next_tip[-2][1] - tip[-2][1])) / 1000
            except KeyError:
                print(tip[0], tuple_variant[idx][0])
                print("ttttt" ,type_of_coating[tip[0]], type_of_coating[tip[0]][tuple_variant[idx][0]])
        return type_of_coating

    def write_10(self):
        ws = self.wb['10']
        column_tuple = ['AF', 'AL', 'AR', 'AX', 'BD', 'BJ']
        counter = 0
        try:
            for k1, v1 in self.data.items():
                if k1 == 'название дороги':
                    continue
                else:
                    column = column_tuple[counter]
                    if len(self.data) > 2:
                        ws[f'{column}4'] = f'Участок {counter + 1} \n {self.data_interface.get("year", None)} г.'
                    else:
                        ws[f'{column}4'] = f'{self.data_interface.get("year", None)}'
                    result = self.count_coating(v1)
                    # print(result)

                    # КАПИТАЛЬНЫЕ
                    ws[f'{column}7'] = ''
                    ws[f'{column}8'] = result.get('Капитальный').get('цементобетон') if result.get(
                        'Капитальный').get('цементобетон') > 0 else ''
                    ws[f'{column}9'] = result.get('Капитальный').get('ж/б плиты') if result.get('Капитальный').get(
                        'ж/б плиты') > 0 else ''
                    ws[f'{column}10'] = result.get('Капитальный').get('цементобетон') if result.get(
                        'Капитальный').get('цементобетон') > 0 else ''
                    ws[f'{column}11'] = result.get('Капитальный').get('цементобетон') if result.get(
                        'Капитальный').get('цементобетон') > 0 else ''
                    ws[f'{column}12'] = result.get('Капитальный').get('цементобетон') if result.get(
                        'Капитальный').get('цементобетон') > 0 else ''
                    ws[f'{column}13'] = result.get('Капитальный').get('асфальтобетон') if result.get(
                        'Капитальный').get('асфальтобетон') > 0 else ''
                    ws[f'{column}14'] = result.get('Капитальный').get('щебень/гравий, обр.вяжущий') if result.get(
                        'Капитальный').get('щебень/гравий, обр.вяжущий') > 0 else ''
                    ws[f'{column}15'] = result.get('Капитальный').get('тротуарная плитка') if result.get(
                        'Капитальный').get('тротуарная плитка') > 0 else ''
                    ws['B15'] = 'Тротуарная плитка'

                    # ОБЛЕГЧЕННЫЕ
                    ws[f'{column}19'] = result.get('Облегченный').get('асфальтобетон') \
                        if result.get('Облегченный').get('асфальтобетон') > 0 else ''
                    ws[f'{column}20'] = result.get('Облегченный').get('органоминеральные') \
                        if result.get('Облегченный').get('органоминеральные') > 0 else ''
                    ws[f'{column}21'] = result.get('Облегченный').get('щебеночные (гравийные), обработанные вяжущим') \
                        if result.get('Облегченный').get('щебеночные (гравийные), обработанные вяжущим') > 0 else ''
                    ws[f'{column}22'] = result.get('Облегченный').get('цементобетон') \
                        if result.get('Облегченный').get('щебень/гравий') > 0 else ''
                    ws[f'B22'] = 'Цементобетонные'

                    # ПЕРЕХОДНЫЕ
                    ws[f'{column}26'] = result.get('Переходный').get('щебень/гравий') \
                        if result.get('Переходный').get('щебень/гравий') > 0 else ''
                    ws[f'{column}27'] = result.get('Переходный').get(
                        'Грунт и малопрочные каменные материалы, укрепленные вяжущим') \
                        if result.get('Переходный').get(
                        'Грунт и малопрочные каменные материалы, укрепленные вяжущим') > 0 else ''
                    ws[f'{column}28'] = result.get('Переходный').get(
                        'Грунт, укрепленный различными вяжущими и местными материалами') \
                        if result.get('Переходный').get(
                        'Грунт, укрепленный различными вяжущими и местными материалами') > 0 else ''
                    ws[f'{column}29'] = result.get('Переходный').get('Булыжный и колотый камень(мостовые)') \
                        if result.get('Переходный').get('Булыжный и колотый камень(мостовые)') > 0 else ''

                    # НИЗШИЕ
                    ws[f'{column}34'] = result.get('Низший').get('грунт') \
                        if result.get('Низший').get('грунт') > 0 else ''
                    ws[f'{column}35'] = result.get('Низший').get('Грунт естественный') \
                        if result.get('Низший').get('Грунт естественный') > 0 else ''

                counter += 1
            else:
                if len(self.data) > 2:
                    column = column_tuple[counter]
                    ws[f'{column}4'] = 'Итог'

                    ws[f'{column}8'] = f'=SUM({column_tuple[0]}8:{column_tuple[counter - 1]}8)'
                    ws[f'{column}9'] = f'=SUM({column_tuple[0]}9:{column_tuple[counter - 1]}9)'
                    ws[f'{column}10'] = f'=SUM({column_tuple[0]}10:{column_tuple[counter - 1]}10)'
                    ws[f'{column}11'] = f'=SUM({column_tuple[0]}11:{column_tuple[counter - 1]}11)'
                    ws[f'{column}12'] = f'=SUM({column_tuple[0]}12:{column_tuple[counter - 1]}12)'
                    ws[f'{column}13'] = f'=SUM({column_tuple[0]}13:{column_tuple[counter - 1]}13)'
                    ws[f'{column}14'] = f'=SUM({column_tuple[0]}14:{column_tuple[counter - 1]}14)'
                    ws[f'{column}15'] = f'=SUM({column_tuple[0]}15:{column_tuple[counter - 1]}15)'

                    ws[f'{column}19'] = f'=SUM({column_tuple[0]}19:{column_tuple[counter - 1]}19)'
                    ws[f'{column}20'] = f'=SUM({column_tuple[0]}20:{column_tuple[counter - 1]}20)'
                    ws[f'{column}21'] = f'=SUM({column_tuple[0]}21:{column_tuple[counter - 1]}21)'
                    ws[f'{column}22'] = f'=SUM({column_tuple[0]}22:{column_tuple[counter - 1]}22)'

                    ws[f'{column}26'] = f'=SUM({column_tuple[0]}26:{column_tuple[counter - 1]}26)'
                    ws[f'{column}27'] = f'=SUM({column_tuple[0]}27:{column_tuple[counter - 1]}27)'
                    ws[f'{column}28'] = f'=SUM({column_tuple[0]}28:{column_tuple[counter - 1]}28)'
                    ws[f'{column}29'] = f'=SUM({column_tuple[0]}29:{column_tuple[counter - 1]}29)'

                    ws[f'{column}34'] = f'=SUM({column_tuple[0]}34:{column_tuple[counter - 1]}34)'
                    ws[f'{column}35'] = f'=SUM({column_tuple[0]}34:{column_tuple[counter - 1]}35)'
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 10 листе")
            self.msg.exec()
    def write_11 (self):
        """
        21.09.2023 таблица 4.6 заполняется
         - ограничения нужна длина
         - сигнальные столбики нет параметра количество точек
        :return:
        """
        ws = self.wb['11']
        # заполнение 11 листа
        counter = 0
        column_tuple_4_6 = ('AU', 'AX', 'BA', 'BD', 'BG', 'BJ', 'BM')
        # column_tuple_4_4 = ('B', 'E', 'L', 'S', 'Z')
        n = 16
        res_sum_curves_and_slopes = [0, 0, 0, 0]
        try:
            for k1, v1 in self.data.items():
                if k1 == 'название дороги':
                    continue
                # 4.4
                slopes_list = v1.get('Кривая', {}).get('Продольный уклон', [])
                dict_counter_and_length_slopes = {
                    'IА': [0, 0],
                    'IБ': [0, 0],
                    'IВ': [0, 0],
                    'II': [0, 0],
                    'III': [0, 0],
                    'IV': [0, 0],
                    'V': [0, 0]
                }

                curves_list = v1.get('Кривая', {}).get('R', [])
                categorys_road_list = v1.get('Граница участка дороги', {}).get('категория а/д', [])
                dict_counter_and_length_curves = {
                    'IА': [0, 0],
                    'IБ': [0, 0],
                    'IВ': [0, 0],
                    'II': [0, 0],
                    'III': [0, 0],
                    'IV': [0, 0],
                    'V': [0, 0]}
                for idx, category in enumerate(categorys_road_list):
                    # находим следующий тип покрытия
                    if category == categorys_road_list[-1]:
                        break
                        # next_category = categorys_road_list[-1]
                    elif category == categorys_road_list[0]:
                        next_category = categorys_road_list[1]
                    else:
                        next_category = categorys_road_list[idx % len(categorys_road_list) + 1]
                    for curve in curves_list:
                        # посчитать количество и протяженность кривых
                        if curve[3] == 146:
                            if category[-2] <= curve[-2] <= next_category[-2] and category[-1] <= curve[-1] <= \
                                    next_category[-1]:
                                if category[0] == 'IА' and 0.0 < abs(float(curve[0])) < 1200.0:
                                    dict_counter_and_length_curves['IА'][0] += 1
                                    dict_counter_and_length_curves['IА'][1] += curve[1]
                                    # dict_counter_and_length_curves['IА'][2].append(curve)
                                elif category[0] == 'IБ' and 0.0 < abs(float(curve[0])) < 800.0:
                                    dict_counter_and_length_curves['IБ'][0] += 1
                                    dict_counter_and_length_curves['IБ'][1] += curve[1]
                                    # dict_counter_and_length_curves['IБ'][2].append(curve)
                                elif category[0] == 'IВ' and 0.0 < abs(float(curve[0])) < 600.0:
                                    dict_counter_and_length_curves['IВ'][0] += 1
                                    dict_counter_and_length_curves['IВ'][1] += curve[1]
                                    # dict_counter_and_length_curves['IВ'][2].append(curve)
                                elif category[0] == 'II' and 0.0 < abs(float(curve[0])) < 800.0:
                                    dict_counter_and_length_curves['II'][0] += 1
                                    dict_counter_and_length_curves['II'][1] += curve[1]
                                    # dict_counter_and_length_curves['II'][2].append(curve)
                                elif category[0] == 'III' and 0.0 < abs(float(curve[0])) < 600.0:
                                    dict_counter_and_length_curves['III'][0] += 1
                                    dict_counter_and_length_curves['III'][1] += curve[1]
                                    # dict_counter_and_length_curves['III'][2].append(curve)
                                elif category[0] == 'IV' and 0.0 < abs(float(curve[0])) < 300.0:
                                    dict_counter_and_length_curves['IV'][0] += 1
                                    dict_counter_and_length_curves['IV'][1] += curve[1]
                                    # dict_counter_and_length_curves['IV'][2].append(curve)
                                elif category[0] == 'V' and 0.0 < abs(float(curve[0])) < 150.0:
                                    dict_counter_and_length_curves['V'][0] += 1
                                    dict_counter_and_length_curves['V'][1] += curve[1]
                                    # dict_counter_and_length_curves['V'][2].append(curve)
                    for slope in slopes_list:
                        # посчитать количество и протяженность продольных углов
                        if (category[-2] <= slope[-2] <= next_category[-2] and category[-1] <= slope[-1] <= next_category[
                            -1]):
                            if category[0] == 'IА' and 0.0 < abs(float(slope[0])) > 30:
                                dict_counter_and_length_slopes['IА'][0] += 1
                                dict_counter_and_length_slopes['IА'][1] += slope[1]
                                # dict_counter_and_length_slopes['IА'][2].append(slope)
                            elif category[0] == 'IБ' and 0.0 < abs(float(slope[0])) > 40:
                                dict_counter_and_length_slopes['IБ'][0] += 1
                                dict_counter_and_length_slopes['IБ'][1] += slope[1]
                                # dict_counter_and_length_slopes['IБ'][2].append(slope)
                            elif category[0] == 'IВ' and 0.0 < abs(float(slope[0])) > 50:
                                dict_counter_and_length_slopes['IВ'][0] += 1
                                dict_counter_and_length_slopes['IВ'][1] += slope[1]
                                # dict_counter_and_length_slopes['IВ'][2].append(slope)
                            elif category[0] == 'II' and 0.0 < abs(float(slope[0])) > 40:
                                dict_counter_and_length_slopes['II'][0] += 1
                                dict_counter_and_length_slopes['II'][1] += slope[1]
                                # dict_counter_and_length_slopes['II'][2].append(slope)
                            elif category[0] == 'III' and 0.0 < abs(float(slope[0])) > 50:
                                dict_counter_and_length_slopes['III'][0] += 1
                                dict_counter_and_length_slopes['III'][1] += slope[1]
                                # dict_counter_and_length_slopes['III'][2].append(slope)
                            elif category[0] == 'IV' and 0.0 < abs(float(slope[0])) > 60:
                                dict_counter_and_length_slopes['IV'][0] += 1
                                dict_counter_and_length_slopes['IV'][1] += slope[1]
                                # dict_counter_and_length_slopes['IV'][2].append(slope)
                            elif category[0] == 'V' and 0.0 < abs(float(slope[0])) > 70:
                                dict_counter_and_length_slopes['V'][0] += 1
                                dict_counter_and_length_slopes['V'][1] += slope[1]
                                # dict_counter_and_length_slopes['V'][2].append(slope)
                # if n % 2 != 0:
                #     # ws.merge_cells(f'B{n}:Z{n}')
                #     ws[f'L{n}'] = f'Участок {counter}'
                #     n += 1
                res_sum_curves_and_slopes[0] += sum(i[0] for i in dict_counter_and_length_curves.values())
                res_sum_curves_and_slopes[1] += sum(i[1] for i in dict_counter_and_length_curves.values()) / 1000
                res_sum_curves_and_slopes[2] += sum(i[0] for i in dict_counter_and_length_slopes.values())
                res_sum_curves_and_slopes[3] += sum(i[1] for i in dict_counter_and_length_slopes.values()) / 1000
                if len(self.data) > 2:
                    ws[f'B16'] = self.data_interface.get('year', None)
                    ws[f'E16'] = res_sum_curves_and_slopes[0] if res_sum_curves_and_slopes[0] > 0 else '-'
                    ws[f'L16'] = res_sum_curves_and_slopes[1] if res_sum_curves_and_slopes[1] > 0 else '-'
                    ws[f'S16'] = res_sum_curves_and_slopes[2] if res_sum_curves_and_slopes[2] > 0 else '-'
                    ws[f'Z16'] = res_sum_curves_and_slopes[3] if res_sum_curves_and_slopes[3] > 0 else '-'
                    ws[f'B{n}'] = self.data_interface.get('year', None)
                    ws[f'E{n}'] = sum(i[0] for i in dict_counter_and_length_curves.values())
                    ws[f'L{n}'] = sum(i[1] for i in dict_counter_and_length_curves.values()) / 1000
                    ws[f'S{n}'] = sum(i[0] for i in dict_counter_and_length_slopes.values())
                    ws[f'Z{n}'] = sum(i[1] for i in dict_counter_and_length_slopes.values()) / 1000
                    n += 1
                else:
                    ws[f'B16'] = self.data_interface.get('year', None)
                    ws[f'E16'] = res_sum_curves_and_slopes[0] if res_sum_curves_and_slopes[0] > 0 else '-'
                    ws[f'L16'] = res_sum_curves_and_slopes[1] if res_sum_curves_and_slopes[1] > 0 else '-'
                    ws[f'S16'] = res_sum_curves_and_slopes[2] if res_sum_curves_and_slopes[2] > 0 else '-'
                    ws[f'Z16'] = res_sum_curves_and_slopes[3] if res_sum_curves_and_slopes[3] > 0 else '-'

                # 4.6
                column = column_tuple_4_6[counter]
                # шапка участки
                if len(self.data) > 2:
                    ws[f'{column}6'] = f'Участок {counter + 1} \n {self.data_interface.get("year", None)} г.'
                else:
                    ws[f'{column}6'] = f'{self.data_interface.get("year", None)}'
                # автопавильоны капитального типа шт
                ws[f"{column}14"] = sum(1 for i in v1.get('Остановка').get('Наличие павильона') if i[0] == 'да') if v1.get(
                    'Остановка', {}).get('Наличие павильона', []) else '-'
                # площадки отдыха шт
                ws[f"{column}16"] =  v1.get('Проезжая часть').get('Назначение').count('площадка отдыха') if v1.get(
                    'Проезжая часть', {}).get('Назначение', [])[0] == 'площадка отдыха' else '-'
                # площадка для стоянок и остановок автомобилей шт
                ws[f"{column}17"] = v1.get('Проезжая часть').get('Назначение').count('парковка') if v1.get(
                    'Проезжая часть', {}).get('Назначение', [])[0]=='парковка' else '-'
                # освещение дороги км
                ws[f"{column}19"] = round(sum(float(x[1]) for x in
                                              v1.get('Опоры освещения и контактной сети').get('Статус')) / 1000,
                                          3) if v1.get(
                    'Опоры освещения и контактной сети', {}).get('Статус', []) else '-'
                # линии технологической связи кабельные км
                ws[f"{column}23"] = round(sum(float(x[1]) for x in
                                              v1.get('Подземная комуникация').get('Вид коммуникации')) / 1000, 3) if v1.get(
                    'Подземная комуникация', {}).get('Вид коммуникации', []) else '-'  # кабельные
                # линии технологической связи воздушные км
                ws[f"{column}24"] = round(sum(float(x[1]) for x in
                                              v1.get('Воздушная коммуникация').get('Вид коммуникации')) / 1000,
                                          3) if v1.get(
                    'Воздушная коммуникация', {}).get('Вид коммуникации', []) else '-'  # воздушные
                # всего км
                lep_sum = ((float(ws[f"{column}23"].value) if ws[f"{column}23"].value != '-' else 0) +
                                     (float(ws[f"{column}24"].value) if ws[f"{column}24"].value != '-' else 0))
                ws[f"{column}20"] = lep_sum if lep_sum > 0 else '-'
                # остановки шт
                ws[f"{column}25"] = len(v1.get('Остановка').get('Название остановки')) if v1.get('Остановка', None) else '-'
                # ПСП шт
                psp_sum = v1.get('Проезжая часть', {}).get('Назначение', []).count('полоса торможения') +  \
                                    v1.get('Проезжая часть', {}).get('Назначение', []).count('полоса разгона')
                ws[f"{column}26"] = psp_sum if psp_sum > 0  else '-'

                # todo: доделать ограждения
                # ограждения км
                ogr_sum = round(sum(float(x[1]) for k in
                                              ['Нестандартное ограждение', 'Пешеходное ограждение', 'Тросовое ограждение',
                                               'Типа Нью-Джерси', 'Металическое барьерное ограждение'] for x in
                                              v1.get(k, {}).get('Статус', [])) / 1000, 3)
                ws[f"{column}28"] = ogr_sum if ogr_sum > 0 else '-'   # ограждения
                # сигнальные столбики шт
                ws[f"{column}29"] = sum(x[4] for x in v1.get('Сигнальные столбики').get('Статус')) \
                    if v1.get('Сигнальные столбики', {}).get('Статус', []) else '-'

                sum_sign = {'всего': 0,
                            'предупреждающие': 0,
                            'приоритета': 0,
                            'запрещающие': 0,
                            'предписывающие': 0,
                            'особых предписаний': 0,
                            'сервиса': 0,
                            'информационные': 0,
                            'дополнительной информации': 0}

                # подсчет знаков шт
                for k, v in v1.items():
                    if k[0].isdigit():
                        # print(k, v)
                        sum_sign['всего'] += len(v['Статус']) if v['Статус'] else 0
                        if k[0] == '1':
                            sum_sign['предупреждающие'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '2':
                            sum_sign['приоритета'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '3':
                            sum_sign['запрещающие'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '4':
                            sum_sign['предписывающие'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '5':
                            sum_sign['особых предписаний'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '6':
                            sum_sign['информационные'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '7':
                            sum_sign['сервиса'] += len(v['Статус']) if v['Статус'] else 0
                        elif k[0] == '8':
                            sum_sign['дополнительной информации'] += len(v['Статус']) if v['Статус'] else 0
                # знаки шт
                ws[f'{column}30'] = sum_sign.get('всего', '-') if sum_sign.get('всего', '-') > 0 else '-'
                ws[f'{column}32'] = sum_sign.get('предупреждающие', '-') if sum_sign.get('предупреждающие', '-') > 0 else '-'
                ws[f'{column}33'] = sum_sign.get('приоритета', '-') if sum_sign.get('приоритета', '-') > 0 else '-'
                ws[f'{column}34'] = sum_sign.get('запрещающие', '-') if sum_sign.get('запрещающие', '-') > 0 else '-'
                ws[f'{column}35'] = sum_sign.get('предписывающие', '-') if sum_sign.get('предписывающие', '-') > 0 else '-'
                ws[f'{column}36'] = sum_sign.get('особых предписаний', '-') if sum_sign.get('особых предписаний', '-') > 0 else '-'
                ws[f'{column}37'] = sum_sign.get('информационные', '-') if sum_sign.get('информационные', '-') > 0 else '-'
                ws[f'{column}38'] = sum_sign.get('сервиса', '-') if sum_sign.get('сервиса', '-') > 0 else '-'
                ws[f'{column}39'] = sum_sign.get('дополнительной информации', '-') if sum_sign.get('дополнительной информации', '-') > 0 else '-'
                counter += 1
            else:

                if len(self.data) > 2:
                    # если участков несколько столбц итого
                    column = column_tuple_4_6[counter]
                    ws[f'{column}6'] = 'Итог'
                    ws[f'{column}14'] = f'=SUM({column_tuple_4_6[0]}14:{column_tuple_4_6[counter - 1]}14)'
                    ws[f'{column}16'] = f'=SUM({column_tuple_4_6[0]}16:{column_tuple_4_6[counter - 1]}16)'
                    ws[f'{column}17'] = f'=SUM({column_tuple_4_6[0]}17:{column_tuple_4_6[counter - 1]}17)'
                    ws[f'{column}19'] = f'=SUM({column_tuple_4_6[0]}19:{column_tuple_4_6[counter - 1]}19)'
                    ws[f'{column}20'] = f'=SUM({column_tuple_4_6[0]}20:{column_tuple_4_6[counter - 1]}20)'
                    ws[f'{column}23'] = f'=SUM({column_tuple_4_6[0]}23:{column_tuple_4_6[counter - 1]}23)'
                    ws[f'{column}24'] = f'=SUM({column_tuple_4_6[0]}24:{column_tuple_4_6[counter - 1]}24)'
                    ws[f'{column}25'] = f'=SUM({column_tuple_4_6[0]}25:{column_tuple_4_6[counter - 1]}25)'
                    ws[f'{column}26'] = f'=SUM({column_tuple_4_6[0]}26:{column_tuple_4_6[counter - 1]}26)'
                    ws[f'{column}28'] = f'=SUM({column_tuple_4_6[0]}28:{column_tuple_4_6[counter - 1]}28)'
                    ws[f'{column}29'] = f'=SUM({column_tuple_4_6[0]}29:{column_tuple_4_6[counter - 1]}29)'
                    ws[f'{column}30'] = f'=SUM({column_tuple_4_6[0]}30:{column_tuple_4_6[counter - 1]}30)'
                    ws[f'{column}32'] = f'=SUM({column_tuple_4_6[0]}32:{column_tuple_4_6[counter - 1]}32)'
                    ws[f'{column}33'] = f'=SUM({column_tuple_4_6[0]}33:{column_tuple_4_6[counter - 1]}33)'
                    ws[f'{column}34'] = f'=SUM({column_tuple_4_6[0]}34:{column_tuple_4_6[counter - 1]}34)'
                    ws[f'{column}35'] = f'=SUM({column_tuple_4_6[0]}35:{column_tuple_4_6[counter - 1]}35)'
                    ws[f'{column}36'] = f'=SUM({column_tuple_4_6[0]}36:{column_tuple_4_6[counter - 1]}36)'
                    ws[f'{column}37'] = f'=SUM({column_tuple_4_6[0]}37:{column_tuple_4_6[counter - 1]}37)'
                    ws[f'{column}38'] = f'=SUM({column_tuple_4_6[0]}38:{column_tuple_4_6[counter - 1]}38)'
                    ws[f'{column}39'] = f'=SUM({column_tuple_4_6[0]}39:{column_tuple_4_6[counter - 1]}39)'
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 11 листе")
            self.msg.exec()

    def write_12(self):
        ws = self.wb['12']
        # 4.7.1
        rows_avtovokzal = 11
        rows_gbdd = 11
        rows_sto = 37
        rows_hotels = 37
        try:
            for name_district, obj in self.data.items():
                if name_district == 'название дороги':
                    continue

                for idx, value in enumerate(obj.get('Здание', {}).get('Назначение', [])):

                    if value[0] in ['Автовокзалы', 'Автостанции']:
                        # 4.7.1
                        ws[f'B{rows_avtovokzal}'] = obj.get('Здание', {}).get('Наименование')[idx][0]
                        ws[f'I{rows_avtovokzal}'] = obj.get('Здание', {}).get('Адрес')[idx][0]
                        ws[f'N{rows_avtovokzal}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0]
                        rows_avtovokzal += 1
                    elif value[0] in ['пост ГИБДД', 'пост ГИБДД/КПД']:
                        # 4.7.2
                        ws[f'AJ{rows_gbdd}'] = obj.get('Здание', {}).get('Наименование')[idx][0]
                        ws[f'AS{rows_gbdd}'] = obj.get('Здание', {}).get('Адрес')[idx][0]
                        ws[f'BA{rows_gbdd}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0]
                        rows_gbdd += 1
                    elif value[0] == 'Гостиница/отель/мотель':
                        # 4.7.4
                        ws[f'B{rows_sto}'] = obj.get('Здание', {}).get('Наименование')[idx][0]
                        ws[f'I{rows_sto}'] = obj.get('Здание', {}).get('Адрес')[idx][0]
                        ws[f'N{rows_sto}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0]
                        rows_sto += 1
                    elif value[0] == 'СТО':
                        # 4.7.3
                        ws[f'AJ{rows_hotels}'] = obj.get('Здание', {}).get('Адрес')[idx][0]
                        ws[f'AS{rows_hotels}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0]
                        ws[f'BD{rows_hotels}'] = obj.get('Здание', {}).get('Наименование')[idx][0]
                        rows_hotels += 1
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 12 листе")
            self.msg.exec()

    def write_13(self):
        ws = self.wb['13']
        rows_azs = 10
        rows_car_wash = 10
        rows_ws = 37
        rows_food = 37
        try:
            for name_district, obj in self.data.items():
                if name_district == 'название дороги':
                    continue
                for idx, value in enumerate(obj.get('Здание', {}).get('Назначение', [])):
                    if value[0] == 'АЗС':
                        # 4.7.5
                        ws[f'B{rows_azs}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                            'Адрес') else ''
                        ws[f'K{rows_azs}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                             {}).get(
                            'Привязка по оси') else ''
                        ws[f'V{rows_azs}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание', {}).get(
                            'Наименование') else ''
                        rows_azs += 1
                    elif value[0] == 'Автомойка':
                        # 4.7.6
                        ws[f'AJ{rows_car_wash}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                            'Адрес') else ''
                        ws[f'AS{rows_car_wash}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                                   {}).get(
                            'Привязка по оси') else ''
                        ws[f'BD{rows_car_wash}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание',
                                                                                                                {}).get(
                            'Наименование') else ''
                        rows_car_wash += 1
                    elif value[0] == 'Общественный туалет':
                        # 4.7.7
                        ws[f'B{rows_ws}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                            'Адрес') else ''
                        ws[f'K{rows_ws}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                            {}).get(
                            'Привязка по оси') else ''
                        ws[f'V{rows_ws}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание', {}).get(
                            'Наименование') else ''
                        rows_ws += 1
                    elif value[0] == 'Кафе/столовая/ресторан':
                        # 4.7.8
                        ws[f'AJ{rows_food}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание',
                                                                                                            {}).get(
                            'Наименование') else ''
                        ws[f'AQ{rows_food}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                            'Адрес') else ''
                        ws[f'AV{rows_food}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                               {}).get(
                            'Привязка по оси') else ''
                        rows_food += 1
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 13 листе")
            self.msg.exec()

    def write_14(self):
        ws = self.wb['14']
        rows_medical = 8
        try:
            for name_district, obj in self.data.items():
                if name_district == 'название дороги':
                    continue
                for idx, value in enumerate(obj.get('Здание', {}).get('Назначение', [])):
                    if value[0] == 'Пункты первой медицинской помощи/почта/телефон':
                        # 4.7.5
                        ws[f'B{rows_medical}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                            'Адрес') else ''
                        ws[f'O{rows_medical}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                                 {}).get(
                            'Привязка по оси') else ''
                        ws[f'Y{rows_medical}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание',
                                                                                                              {}).get(
                            'Наименование') else ''
                        rows_medical += 1
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 14 листа")
            self.msg.exec()

    def write_17(self):
        """
        27.09.2023
        :return:
        """
        ws = self.wb['17']
        counter = 0
        column_tuple = ('J', 'O', 'T', 'Y', 'AD')
        cells = ('L', 'Q', 'V', 'AA', 'AF')

        pipes = {
            "Металлические": [0, 0],
            "Железобетонные": [0, 0],
            "Бетоннометаллические": [0, 0],
            "Каменные": [0, 0],
            "Деревянные": [0, 0],
            "Асбестоцементные": [0, 0],
        }
        # 4.10.2 Сводная ведомость наличия тоннелей, галерей и пешеходных переходов в разных уровнях
        types_of_structures = {
            "Тоннель (галерея)": [0, 0],
            "Пешеходный переход подземный": [0, 0],
            "Пешеходный переход надземный": [0, 0],
            "Водопропускная труба": pipes
        }

        def count_4_10_2(types_of_structures, column, cell, k):
            for key, value in types_of_structures.items():
                if self.data.get(k).get(key) == None:
                    continue
                else:
                    if key == 'Водопропускная труба':
                        for lst in self.data.get(k).get(key).get('Материал'):
                            print(lst)
                            if lst[0] == 'металл':
                                types_of_structures.get(key).get('Металлические')[0] += 1
                                types_of_structures.get(key).get('Металлические')[1] += lst[1]
                            elif lst[0] == 'ж/б':
                                types_of_structures.get(key).get('Железобетонные')[0] += 1
                                types_of_structures.get(key).get('Железобетонные')[1] += lst[1]

                        ws[f'{column}37'] = types_of_structures.get(key).get('Металлические')[0]
                        ws[f'{cell}37'] = types_of_structures.get(key).get('Металлические')[1]
                        ws[f'{column}38'] = types_of_structures.get(key).get('Железобетонные')[0]
                        ws[f'{cell}38'] = types_of_structures.get(key).get('Железобетонные')[1]

                    else:
                        result = self.data.get(k).get(key)
                        types_of_structures.get(key)[0] += 1
                        types_of_structures.get(key)[1] += result.get(list(result.keys())[0])[0][2]

                        if key == 'Тоннель (галерея)':
                            ws[f'{column}14'] = types_of_structures.get(key)[0]
                            ws[f'{cell}14'] = types_of_structures.get(key)[1]
                        elif key == 'Пешеходный переход подземный':
                            ws[f'{column}20'] = types_of_structures.get(key)[0]
                            ws[f'{cell}20'] = types_of_structures.get(key)[1]
                        elif key == 'Пешеходный переход надземный':
                            ws[f'{column}19'] = types_of_structures.get(key)[0]
                            ws[f'{cell}19'] = types_of_structures.get(key)[1]
        try:
            for k, v in self.data.items():
                if k == 'название дороги':
                    continue
                else:
                    column = column_tuple[counter]
                    cell = cells[counter]
                    if len(self.data) > 2:
                        ws[f'{column}6'] = f'Участок {counter + 1}'
                        ws[f'{column}29'] = f'Участок {counter + 1}'

                        count_4_10_2(types_of_structures, column, cell, k)
                        types_of_structures = {
                            "Тоннель (галерея)": [0, 0],
                            "Пешеходный переход подземный": [0, 0],
                            "Пешеходный переход надземный": [0, 0],
                            "Водопропускная труба": pipes
                        }
                    else:
                        # ws[f'{column}6'] = f'{self.data_interface.get("year", None)}'
                        count_4_10_2(types_of_structures, column, cell, k)
                counter += 1
        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 17 листе")
            self.msg.exec()

    def write_18(self):
        """
        Описиваем данные по 18 листу
        :return:
        """
        # todo: дописать расчет сумм съездов
        counter = 0
        counter2 = 1
        ws = self.wb['18']
        column_tuple = ('AP', 'AU', 'AZ', 'BE', 'BJ', 'BM')
        cells = ('AP', 'AR', 'AU', 'AW', 'AZ', 'BB', 'BE', 'BG')
        types = {
            "Асфальтобетонные": [0, 0],
            "Цементобетонные": [0, 0],
            "Тротуарная плитка": [0, 0],
            "Щебеночные": [0, 0],
            "Грунтовые": [0, 0],
            "Ж/б плиты": [0, 0],
        }
        # 4.10.7 Сводная ведомость тротуаров и пешеходных дорожек
        res = 0
        try:
            for k1, v1 in self.data.items():
                if k1 == 'название дороги':
                    continue
                else:
                    for idx, value in enumerate(v1.get('Тротуар', {}).get('Расположение', [])):
                        res += value[-1][1] - value[-2][1]
            if res > 0:
                ws["I30"] = 2023
                ws["I36"] = res / 1000
        except Exception as e:
            print("Тротуары не заполнены")
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 18 листе")
            self.msg.exec()
        # 4.10.8 Сводная ведомость укрепления обочин

        # todo: посмотреть расчеты съездов по нескольким учатскам
        # 4.10.9 Сводная ведомость съездов (въездов)
        try:
            for k1, v1 in self.data.items():
                if k1 == 'название дороги':
                    continue

                column = column_tuple[counter]
                cell = cells[counter2]
                if len(self.data) > 2:
                    ws[f'{column}30'] = f'Участок {counter + 1} \n {self.data_interface.get("year", None)} г.'
                else:
                    ws[f'{column}30'] = f'{self.data_interface.get("year", None)}'
                for lst in self.data.get(k1).get('Съезд', {}).get('Тип покрытия', []):
                    if lst[0] == 'асфальтобетон':
                        types.get('Асфальтобетонные')[0] += 1
                        types.get('Асфальтобетонные')[1] += lst[2]
                    elif lst[0] == 'Цементобетонные':
                        pass
                    elif lst[0] == 'тротуарная плитка':
                        ws['AI38'] = "Тротуарная плитка"
                        types.get('Тротуарная плитка')[0] += 1
                        types.get('Тротуарная плитка')[1] += lst[2]
                    elif lst[0] == 'грунт':
                        types.get('Грунтовые')[0] += 1
                        types.get('Грунтовые')[1] += lst[2]

                    elif lst[0] == 'щебень/гравий':
                        types.get('Щебеночные')[0] += 1
                        types.get('Щебеночные')[1] += lst[2]
                        pass
                    elif lst[0] == 'ж/б плиты':
                        types.get('Ж/б плиты')[0] += 1
                        types.get('Ж/б плиты')[1] += lst[2]
                    ws[f'{column}36'] = types['Асфальтобетонные'][0] if types['Асфальтобетонные'][0] > 0 else ''
                    ws[f'{cell}36'] = types['Асфальтобетонные'][1] if types['Асфальтобетонные'][1] > 0 else ''
                    ws[f'{column}37'] = types['Цементобетонные'][0] if types['Цементобетонные'][0] > 0 else ''
                    ws[f'{cell}37'] = types['Цементобетонные'][1] if types['Цементобетонные'][1] > 0 else ''
                    ws[f'{column}38'] = types['Тротуарная плитка'][0] if types['Тротуарная плитка'][0] > 0 else ''
                    ws[f'{cell}38'] = types['Тротуарная плитка'][1] if types['Тротуарная плитка'][1] > 0 else ''
                    ws[f'{column}39'] = types['Щебеночные'][0] if types['Щебеночные'][0] > 0 else ''
                    ws[f'{cell}39'] = types['Щебеночные'][1] if types['Щебеночные'][1] > 0 else ''
                    ws[f'{column}40'] = types['Грунтовые'][0] if types['Грунтовые'][0] > 0 else ''
                    ws[f'{cell}40'] = types['Грунтовые'][1] if types['Грунтовые'][1] > 0 else ''
                    ws[f'{column}41'] = types['Ж/б плиты'][0] if types['Ж/б плиты'][0] > 0 else ''
                    ws[f'{cell}41'] = types['Ж/б плиты'][1] if types['Ж/б плиты'][1] > 0 else ''
                else:
                    sd_sum = 0
                    square_sum = 0
                    for i in range(36, 42):
                        sd_sum += int(ws[f'AP{i}'])
                        square_sum += float(ws[f'AR{i}'])
                    ws['AP42'] = sd_sum
                    ws['AR42'] = square_sum

                counter += 1
                counter2 += 2

                # Расчет суммы всех показателей

                types = {
                    "Асфальтобетонные": [0, 0],
                    "Цементобетонные": [0, 0],
                    "Тротуарная плитка": [0, 0],
                    "Щебеночные": [0, 0],
                    "Грунтовые": [0, 0],
                    "Ж/б плиты": [0, 0],
                }

        except Exception as e:
            # pass
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка в 18 листе")
            self.msg.exec()

class WriterExcelDAD(WriterExcel):
    def __init__(self, data: dict = None):
        super().__init__(data)

    def write_titular(self):
        pass

    def write_scheme(self):
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

class WriterApplication(WriterExcel):
    def __init__(self, data: dict = None, path = None, data_interface=None):
        super().__init__(data=data, path_template_excel= path_template_excel_application, path=path, data_interface=data_interface)

    def save_file(self):
        # сохранить файл
        self.wb.save(
            rf"{self.path_dir}\{self.data.get('название дороги', 'Отчет')}Приложение_{'город' if self.data_interface.get('tip_passport') == 'city' else 'внегород'}.xlsm")
        self.close_file()

class WriterApplicationCityTP(WriterApplication):
    # TODO: Отловить ошибки
    def __init__(self, data: dict = None, path = None, data_interface = None):
        super().__init__(data=data, path=path, data_interface=data_interface)
        self.thin = Side(border_style = 'thin', color = '000000')
        self.thick = Side(border_style = 'thick', color = '000000')
        self.max_row_for_list = [44,86,128,170,212,254,296,338,380,422,464,506,548,590,632,674,716,758,800]
        print("#############################\nНачал формировать ведомости!\n#############################\n")
        self.write_roadway()
        self.write_separator_strip()

    def write_roadway(self):
        """ Заполнение таблиц проезжая часть"""
        ws = self.wb['ПЧ']
        row = 9
        counter_sum = 0
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                else:
                    if len(self.data) > 2:
                        ws.merge_cells(f'B{row}:K{row}')
                        ws[f'B{row}'] = k1
                        ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                        ws[f'B{row}'].alignment = Alignment(horizontal='center')
                        row += 1

                    for idx, value in enumerate(v1.get('Проезжая часть', {}).get('Название', [])):
                        if value[0] == 'основные полосы движения':
                            ws[f'B{row}'] = counter
                            ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'C{row}'] = value[-2][1]
                            ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'D{row}'] = value[-1][1]
                            ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'E{row}'] = value[-1][1] -value[-2][1]
                            ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'F{row}'] = 'оба' if v1.get('Проезжая часть', {}).get('Расположение', [])[idx][0] == 'По оси' else \
                                v1.get('Проезжая часть', {}).get('Расположение', [])[idx][0]
                            ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'G{row}'] = v1.get('Проезжая часть', {}).get('Тип покрытия', [])[idx][0]
                            ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                            ws[f'H{row}'] = value[2] # square
                            ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                            row += 1
                            counter += 1
                    counter_sum += counter
            ws[f'J{row + 2 if row + 2 > 43 else 43}'] = 'Итого протяженность (м):'
            ws[f'J{row + 2 if row + 2 >= 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'K{row + 2 if row + 2 >= 43 else 43}'] = f'=SUM(E9:E{row})'
            ws[f'K{row + 2 if row + 2 >= 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю

            ws[f'J{row + 3 if row + 3 >= 44 else 44}'] = 'Итого площадь (м²):'
            ws[f'J{row + 3 if row + 3 >= 44 else 44}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'K{row + 3 if row + 3 >= 44 else 44}'] = f'=SUM(H9:H{row})'
            ws[f'K{row + 3 if row + 3 >= 44 else 44}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю

            if counter_sum == 1:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Проезжая часть")
            self.msg.exec()


    def write_separator_strip(self):
        """
        Заполнение таблицы - Разделительные полосы
        :return:
        """

        counter_sum = 0
        row = 9
        ws = self.wb['разделительная полоса']
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'B{row}:L{row}')
                    ws[f'B{row}'] = k1
                    ws[f'B{row}'].border = Border(left = self.thin, right = self.thin, top= self.thin, bottom= self.thin)
                    ws[f'B{row}'].alignment = Alignment(horizontal = 'center')
                    row += 1
                for idx, value in enumerate(v1.get('Разделительная полоса', {}).get('Расположение', [])):

                    ws[f'B{row}'] = counter
                    ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'C{row}'] = value[-2][1]
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'D{row}'] = value[-1][1]
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'E{row}'] = value[-1][1] - value[-2][1]
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = value[0]
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}'] = v1.get('Разделительная полоса').get('Тип покрытия', [])[idx][0]
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = round(value[2] / (value[-1][1]-value[-2][1]), 2)
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = value[2]
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'K{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'L{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    row += 1
                    counter += 1
                counter_sum += counter
            ws[f'K{row + 2 if row + 2 > 43 else 43}'] = 'Итого протяженность (м):'
            ws[f'K{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'L{row + 2 if row + 2 > 43 else 43}'] = f'=SUM(E9:E{row}'
            ws[f'L{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю

            ws[f'K{row + 2 if row + 2 > 44 else 44}'] = 'Итого площадь (м²): '
            ws[f'K{row + 2 if row + 2 > 44 else 44}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'L{row + 2 if row + 2 > 44 else 44}'] = f'=SUM(I9:I{row}'
            ws[f'l{row + 2 if row + 2 > 44 else 44}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю

            if counter_sum == 1 or counter_sum == 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка-Разделительные полосы")
            self.msg.exec()

    def write_reinforces_shoulders(self):
        """
        Заполнение таблицы - наличие укрепленных обочин
        :return:
        """
        row = 8
        ws = self.wb['укреп. обочины']
        counter_sum = 0
        try:

            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'D{row}:K{row}')
                    ws[f'D{row}'] = k1
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'D{row}'].alignment = Alignment(horizontal='center')
                    row += 1
                for idx, value in enumerate(v1.get('Укрепленная часть обочины', {}).get('Расположение', [])):

                    ws[f'D{row}'] = counter
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'E{row}'] = value[-2][1]
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = value[-1][1]
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}'] = value[0]
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = value[-1][1] - value[-2][1]
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = v1.get('Укрепленная часть обочины', {}).get('Тип покрытия', [])[idx][0]
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'J{row}'] = value[2]
                    ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'K{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    row += 1
                    counter += 1
                counter_sum += counter
            ws[f'J{row + 2}'] = 'Итого протяженность (м):'
            ws[f'J{row + 2}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'K{row + 2}'] = f'=SUM(H8:H{row}'
            ws[f'K{row + 2}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю
            if counter_sum == 1 or counter_sum -- 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Укрепленные обочины")
            self.msg.exec()

    def write_exit_road(self):
        """
        Заполнение таблицы съездов
        :return:
        """
        counter_sum = 0
        row = 8
        ws = self.wb['съезды']
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'C{row}:K{row}')
                    ws[f'C{row}'] = k1
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'C{row}'].alignment = Alignment(horizontal = 'center')
                    row += 1
                for idx, value in enumerate(v1.get('Съезд', {}).get('Расположение', [])):
                    ws[f'C{row}'] = counter
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'D{row}'] = value[-2][1]

                    ws[f'E{row}'] = value[-1][1]
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = value[0]
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}'] = 'Съезд'
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = v1.get('Съезд', {}).get('Тип покрытия', [])[idx][0]
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = value[2]
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'J{row}'] = v1.get('Съезд', {}).get('Назначение съезда', [])[idx][0]
                    ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'K{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    row += 1
                    counter += 1
            ws[f'J{row + 2 if row + 2 > 43 else 43}'] = 'Итого (шт):'
            ws[f'J{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'K{row + 2 if row + 2 > 43 else 43}'] = counter_sum
            ws[f'K{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю
            if counter_sum == 1 or counter_sum == 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Съезды")
            self.msg.exec()

    def write_other_territories(self):
        """
        Заполнение таблицы съездов
        :return:
        """
        counter_sum = 0
        row = 8
        ws = self.wb['прочие территории']
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'C{row}:J{row}')
                    ws[f'C{row}'] = k1
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'C{row}'].alignment - Alignment(horizontal='center')
                    row += 1
                for idx, value in enumerate(v1.get('Проезжая часть', {}).get('Назначение', {})):
                    if value[0] in ['площадка отдыха', 'автостоянка', 'парковка', 'отстоно-разворотная площадка',
                                    'трамвайное полотно']:
                        ws[f'C{row}'] = counter
                        ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'D{row}'] = 'чет.' if v1.get('Проезжая часть').get('Расположение', [])[idx][0] == 'право' else 'нечет.'
                        ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'E{row}'] = value[-2][1]
                        ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'F{row}'] = value[-1][1]
                        ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'G{row}'] = value[0]
                        ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'H{row}'] = value[2]
                        ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'I{row}'] = ''
                        ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        ws[f'J{row}'] = ''
                        ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                        row += 1
                        counter += 1

                counter_sum += counter
            ws[f'J{row + 2 if row + 2 > 43 else 43}'] = 'Итого (шт.):'
            ws[f'J{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'K{row + 2 if row + 2 > 43 else 43}'] = counter_sum
            ws[f'K{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю
            if counter_sum <= 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Прочие территории")
            self.msg.exec()
    def write_sidewal(self):
        """
        Заполнение таблицы - тротуары
        :return: 
        """
        counter_sum = 0
        row = 8
        ws  = self.wb['тротуары']
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'D{row}:L{row}')
                    ws[f'D{row}'] = k1
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'D{row}'].alignment = Alignment(horizontal='center', vertical='center')

                    row += 1
                for idx, value in enumerate(v1.get('Тротуар', {}).get('Расположение', [])):
                    ws[f'D{row}'] = counter
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'E{row}'] = value[0]
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = value[-2][1]
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}']= value[-1][1]
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = value[-1][1] - value[-2][1] if value[-1][1] - value[-2][1] != 0 else ''
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = round(value[2] / value[-1][1] - value[-2][1], 2) if value[-1][1] - value[-2][1] != 0 else ''
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'J{row}'] = value[2]
                    ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'K{row}'] = v1.get('Тип покрытия').get('Материал покрытия', [])[idx][0]
                    ws[f'K{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'L{row}'] = ''
                    ws[f'L{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    if row in self.max_row_for_list:
                        row += 3
                    else:
                        row += 1
                    counter += 1

                counter_sum += counter
            ws[f'K{row + 2 if row + 2 > 43 else 43}'] = 'Итого (шт.): '
            ws[f'K{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'L{row + 2 if row + 2 > 43 else 43}'] = counter_sum
            ws[f'L{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю

            if counter_sum == 1 or counter_sum == 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Тротуары")
            self.msg.exec()

    def write_border(self):
        """
        Заполнение таблицы бордюры
        Должны были добавить расположение в бордюры
        :return:
        """
        row = 8
        ws = self.wb['бордюр']
        counter_sum = 0
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'B{row}:I{row}')
                    ws[f'B{row}'] = k1
                    ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
                    row += 1
                for idx, value in enumerate(v1.get('Бордюр', {}).get('Назначение', [])):

                    ws[f'B{row}'] = counter
                    ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'C{row}'] = 'чет.' if value[6] > 0 else 'нечет.'
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'D{row}'] = value[-2][1]
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'E{row}'] = value[-1][1]
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = value[-1][1] - value[-2][1]
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}'] = value[0]
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = v1.get('Бордюр', {}).get('Марка', [])[idx][0] if 0 < idx < len(v1.get('Бордюр', {}).get('Марка', [])) else ''
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = ''
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    row += 1
                    counter += 1
                counter_sum += counter
            ws[f'I{row + 2 if row + 2 > 43 else 43}'] = 'Итого (шт.):'
            ws[f'I{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'J{row + 2 if row + 2 > 43 else 43}'] = counter_sum - 2
            ws[f'J{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю
            if counter_sum <= 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Бордюры")
            self.msg.exec()

    def write_luke(self):
        """
        Заполнение таблицы - Люки
        :return:
        """

        row = 8
        ws = self.wb['люки']
        counter_sum = 0
        try:
            for k1, v1 in self.data.items():
                counter = 1
                if k1 == 'название дороги':
                    continue
                if len(self.data) > 2:
                    ws.merge_cells(f'B{row}:K{row}')
                    ws[f'B{row}'] = k1
                    ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)
                    ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
                    row += 1
                luks = v1.get('Люк смотрового колодца', {}).get('Расположение', []) + v1.get('Решетка дождеприемного колодца', {}).get('расположение', [])
                for value in luks:
                    ws[f'B{row}'] = counter
                    ws[f'B{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'C{row}'] = 'чет.' if value[6] > 0 else 'нечет.'
                    ws[f'C{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'D{row}'] = value[-2][1]
                    ws[f'D{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'E{row}'] = round(value[6], 1)
                    ws[f'E{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'F{row}'] = '+' if value[0]=='ПЧ' else ''
                    ws[f'F{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'G{row}'] = '+' if value[0] == 'Тротуар' else ''
                    ws[f'G{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'H{row}'] = '+' if value[0] == 'Газон' else ''
                    ws[f'H{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'I{row}'] = 'смотровой' if value[0] in ['Газон', 'ПЧ', 'Тротуар'] else 'ливневый'
                    ws[f'I{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'J{row}'] = ''
                    ws[f'J{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    ws[f'K{row}'] = ''
                    ws[f'K{row}'].border = Border(left=self.thin, right=self.thin, top=self.thin, bottom=self.thin)

                    row += 1
                    counter += 1
                counter_sum += counter
            ws[f'K{row + 2 if row + 2 > 43 else 43}'] = 'Итого (шт.):'
            ws[f'K{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='right') # выравнивание по правому краю

            ws[f'L{row + 2 if row + 2 > 43 else 43}'] = counter_sum - 2
            ws[f'L{row + 2 if row + 2 > 43 else 43}'].alignment = Alignment(horizontal='left') # выравнивание по левому краю
            if counter_sum >= 2:
                ws.sheet_state = 'hidden'
        except Exception as e:
            self.msg.setText(f"Ошибка {e}")
            self.msg.setWindowTitle("Ошибка - Люки")
            self.msg.exec()
class WriterApplicationNotCityTP(WriterApplication):
    def __init__(self, data : dict = None):
        super().__init__(data)
        # path_template_excel_application = ''
        self.wb = load_workbook(path_template_excel_application, keep_vba = True)

def convert_visio2svg(path_dir, name,):
    """ Конвертируем визио в svg"""
    visio = win32com.client.gencache.EnsureDispatch("Visio.Application")
    doc = visio.Documents.Open(rf"{path_dir}\{name}.vsd")
    print(rf"{path_dir}\{name}.vsd")


    for page in doc.Pages:
        print(rf'{path_dir}\{page.Name}.pdf')
        page.Export(rf"{path_dir}\{page}.png")
    visio.Quit()

def main():
    conn = db.Query('IZHEVSK_CITY_2023')
    data = conn.get_tp_datas('2620_Проезд от д. № 37 по ул. Союзной к д. № 27 по ул. Союзной')
    # report_application = WriterApplication(data)
    report = WriterExcelTP(data=data, path=r'C:\Users\sibregion\PycharmProjects\ExcelPyQT\static')
    report.save_file()

if __name__ == "__main__":
    import time
    start = time.time()  # точка отсчета времени
    main()
    end = time.time() - start  # собственно время работы программы
    print(f"{round(end, 1)} секунд")  # вывод времени
