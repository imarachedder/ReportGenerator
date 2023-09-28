# from openpyxl.chart import BarChart, Reference
import glob
import warnings

import win32com.client
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

import db
from settings import path_template_excel
import time


class WriterExcel:
    def __init__(self, data: dict = None):
        if data is None:
            self.data = {}
        # self.db = db.Query() # нужно ли ?
        # self.info_window2 = window2.Window2().get_info_window2()
        # self.info_window3 = window3.Window3().get_info_from_plainTextEdit()
        self.wb = load_workbook(path_template_excel, keep_vba = True)
        # self.path_dir = r"C:\Users\sibregion\Desktop\test\report\static"
        self.path_dir = r"C:\Users\sibregion\PycharmProjects\ExcelPyQT\static"
       #self.img = Image(f"{self.path_dir}\схема.png")
        self.data = data
        self.page_number = 2
        self.data_interface = {'year': 2023, 'count_region': 1}

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
        self.write_titular(data)
        self.write_6(data)
        self.write_7()
        self.write_9(data)
        # self.write_10(data)
        self.write_11()
        self.write_12()
        self.write_13()
        self.write_14()
        # self.write_17()
        self.write_18()
        #convert_visio2svg(self.path_dir)

    def write_titular (self, data):
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

    def write_scheme (self, data):
        """
         Заполняет лист "схема"
        :return: None
        """
        schema = Image(f"{self.path_dir}\схема.png")
        ws = self.wb['Схема']  # выбираем лист
        schema.width = 1380
        schema.height = 800
        ws.add_image(schema, 'B5')

    def write_6 (self, data):
        """
        Заполняет лист "6"
        :return:
        """
        # print(self.data.get(f'участок {1}').get('Ось дороги', None))
        ws = self.wb['6']  # выбираем лист
        n = 1  # счетчик
        # 2.1 Наименование дороги: name road
        ws["O5"].value = self.data.get('название дороги')

        # for i in range(self.data_interface.get('count_region', 1)):
        #     ws[f"L1{i}"].value = self.data.get('участки', 'None')[i]  # 2.2 Участок дороги: участки

        # 2.2 Участок дороги 1, 2 и т.д.
        if self.data_interface.get('count_region') > 1:
            for i in range(0, self.data_interface.get("count_region", 0)):
                if i == 0:
                    ws[f'B1{i}'].value = f'2.2  Участок дороги {n}'
                else:
                    ws[f'B1{i}'].value = f'         Участок дороги {n}'
                # print(self.data.get('Ось дороги', None).get('Начало трассы', 0)[i])
                # print("PECHATAYU", self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][2] )
                ws[f'L1{i}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][1]} + " \
                                     f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][2]} м."
                n += 1
            n = 1
        else:
            # ws["S14"].value = self.data.get('протяженность дороги', None)  # 2.3 Протяженность дороги: протяженность
            ws["L10"].value = f"{self.data.get(f'весь участок').get('Ось дороги', None).get('Начало трассы')[0][1]} + " \
                                 f"{self.data.get(f'весь участок').get('Ось дороги', None).get('Начало трассы')[0][2]} м."


        # 2.3 Суммарна протяженность по участку или участкам
        if self.data_interface.get("count_region") > 1:
            res = 0
            for i in range(0, self.data_interface.get("count_region", 0)):
                res += int(self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][2])
                n += 1
            n = 1
            ws["S14"].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы')[0][1]} + {res} м."
        else:
            ws["S14"].value = self.data.get(f'весь участок').get('Ось дороги', None)['Начало трассы'][0][2]

        # заполняет таблицу 2.3 Протяженность дороги
        if self.data_interface.get('count_region') > 1:
            for i in range(2, (self.data_interface.get('count_region') * 2) + 1, 2):
                # print(self.data.get('count_region') * 2)
                ws[f'B2{i - 1}'].value = f'Участок {n}'
                # print(self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[4%self.data.get('count_region')][1])
                if n % 2 != 0:
                    ws[f'B2{i}'].value = self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i%self.data_interface.get('count_region')][1]
                    ws[f'F2{i}'].value = self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i%self.data_interface.get('count_region') ][2]
                    ws[f'J2{i}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region')][1]} + " \
                                          f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region')][2]} м."
                else:
                    ws[f'B2{i}'].value = \
                    self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region') -1][1]
                    ws[f'F2{i}'].value = \
                    self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region') -1][2]
                    ws[f'J2{i}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region') -1][1]} + " \
                                         f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data_interface.get('count_region') -1][2]} м."

                # ПОКА ПРОПУСКАЕМ ЭТОТ ПУНКТ
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
        if self.data_interface.get("count_region") > 1:
            counter = 1
            ws["AL10"].value = self.data.get('название дороги')
            for i in range(0, self.data_interface.get("count_region")):
                # ДОПИСАТЬ КАТЕГОРИИ
                ws[rf'AW1{counter}'].value = f"{0} + {self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][1]}"
                ws[rf'BA1{counter}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][1]} + " \
                                        f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][2]}"
                n += 1
                counter += 2
            n = 1

        # заполняет таблицу 2.6 Краткая историческая справка
        ws["AL33"].value = "Историческая справка"
        # ws["AL33"].value = self.data.get('history_match', None)

    def write_7(self):
        # 2.7
        ws = self.wb['7']
        counter_distr_soder = 15
        for k1, v1 in self.data.items():
            if k1 == 'название дороги':
                continue
            for idx, v2 in enumerate(v1.get('Дорожная организация', {}).get('Наименование', [])):
                ws[f'B{counter_distr_soder}'] = self.data_interface.get('year', '')
                ws[f'E{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Наименование', [])[idx][
                    0] if v1.get('Дорожная организация', {}).get('Наименование', []) else ''
                ws[f'l{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Адрес', [])[idx][0] if v1.get(
                    'Дорожная организация', {}).get('Адрес', []) else ''
                ws[f'P{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Город', [])[idx][0] if v1.get(
                    'Дорожная организация', {}).get('Город', []) else ''
                ws[f'V{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Начало по оси', [])[idx][
                    0] if v1.get(
                    'Дорожная организация', {}).get('Начало по оси', []) else ''
                ws[f'Y{counter_distr_soder}'] = v1.get('Дорожная организация', {}).get('Конец  по оси', [])[idx][
                    0] if v1.get(
                    'Дорожная организация', {}).get('Конец  по оси', []) else ''
                start = v1.get('Дорожная организация', {}).get('Начало по оси', [])[idx][0].split('+')
                end = v1.get('Дорожная организация', {}).get('Конец  по оси', [])[idx][0].split('+')

                ws[
                    f'AB{counter_distr_soder}'] = f'{((int(end[0]) - int(start[0])) * 1000 + int(end[1]) - int(start[1])) / 1000}'
                counter_distr_soder += 1
        # 2.8

        column_tuple = ('AX', 'BA', 'BD', 'BJ', 'BM')

        counter = 15
        for k1, v1 in self.data.items():
            if k1 == 'название дороги':
                continue
            tuple_name = v1.get('Населенный пункт', {}).get('Наименование', [])
            for idx, name in enumerate(tuple_name):
                # находим следующий километровый
                if name == tuple_name[-1]:
                    next_name = tuple_name[-1]
                elif name == tuple_name[0]:
                    next_name = tuple_name[0]
                else:
                    next_name = tuple_name[idx % len(tuple_name) + 1]

                ws[f'{column_tuple[idx]}4'] = name[0]
                ws[f'AR{counter}'] = name[0]
                ws[f'{column_tuple[idx]}{counter}'] = next_name[2] - name[2]
                counter += 1

    def write_8(self, data):
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
        ws['B6'] = self.data.get('economical_characteristic_road')
        # 3.2 Связь дороги с ж/д и водными путями и автомобильными дорогами
        ws['B19'] = self.data.get('railway_waterway')
        # 3.3 Характеристика движения, его сезонность и перспектива роста
        ws['B33'] = self.data.get('movement_characteristic')
        # 3.4 Среднесуточная интенсивность движения по данным учета

    def write_9(self, data):
        """
        Техническая характеристика
        :param data:
        :return:
        """

        # Функция для расчета ширины проезжей части
        def calcLengthOfTheWidthOfTheCarriageWay(res, i, j):
            result = res - int(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][1])
            if j - 1 == 0:
                result += int(self.data.get(f'участок {i + 1}').get('Ширина проезжей части').get('Ширина ПЧ')[j-1][1])
            # print(f"do {res}",
            #       self.data.get(f'участок {i + 1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[
            #           j - 1][0],
            #       int(self.data.get(f'участок {i + 1}').get('Ширина проезжей части').get('Ширина ПЧ')[
            #               j - 1][1]), result, res)
            return result

        # Счетчик
        n = 1

        ws = self.wb['9']
        # 4.1 Топографические условия района проложения автомобильной дороги
        # ws['B7'] = self.data.get('area_conditioins')
        # 4.2 Ширина земляного полотна
        # 4.3 Характеристика проезжей части
        # 4.3.1 Ширина проезжей части


        if self.data_interface.get('count_region') > 1:

            # Цикл по количеству учасков
            for i in range(0, self.data_interface.get('count_region')):
                # Создаем переменные для ячеек в таблице 4.3.1 Ширина проезжей части
                res2, res3, res4, res5, res6, res7, res8, res9, res10, res11, res12 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                ws[f'AJ1{n+1}'].value = f'Участок {i+1}'
                res = self.data.get(f'участок {i + 1}').get('Ось дороги').get('Начало трассы')[0][2]
                # if i == 1:
                #     break
                for j in range(len(self.data.get(f'участок {i + 1}').get('Ширина проезжей части').get('Ширина ПЧ')), 0, -1):
                    if float(self.data.get(f'участок {i+1}').get('Ширина проезжей части').get('Ширина ПЧ')[j-1][0]) <= 4.0:
                        res2 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 4.1 < float(self.data.get(f'участок {i + 1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 4.5:
                        res3 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 4.5 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 6.0:
                        res4 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 6.0 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части').get('Ширина ПЧ')[j-1][0]) < 6.6:
                        res5 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 6.6 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части').get('Ширина ПЧ')[j-1][0]) < 7.0:
                        res6 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 7.0 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 7.5:
                        res7 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 7.5 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 9.1:
                        res8 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 9.1 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 10.0:
                        res9 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 10.0 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 15.1:
                        res10 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)
                    elif 15.1 <= float(self.data.get(f'участок {i+1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[j-1][0]) < 27.5:
                        res11 += calcLengthOfTheWidthOfTheCarriageWay(res, i, j)

                    res = self.data.get(f'участок {i + 1}').get('Ширина проезжей части').get('Ширина ПЧ')[j - 1][1]
                    ws[f'AL1{n+2}'].value = round(res2 / 1000, 3)
                    ws[f'AO1{n+2}'].value = round(res3 / 1000, 3)
                    ws[f'AR1{n+2}'].value = round(res4 / 1000, 3)
                    ws[f'AU1{n+2}'].value = round(res5 / 1000, 3)
                    ws[f'AX1{n+2}'].value = round(res6 / 1000, 3)
                    ws[f'BA1{n+2}'].value = round(res7 / 1000, 3)
                    ws[f'BD1{n+2}'].value = round(res8 / 1000, 3)
                    ws[f'BG1{n+2}'].value = round(res9 / 1000, 3)
                    ws[f'BJ1{n+2}'].value = round(res10 / 1000, 3)
                    ws[f'BM1{n+2}'].value = round(res11 / 1000, 3)
                n += 2

    def count_coating (self, v):

        """
        Расчет протяженностей типов покрытий
        :param data:
        :return:
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
            'асфальтобетонные': 0,
            'органоминеральные': 0,
            'щебеночные (гравийные), обработанные вяжущим': 0,
            'цементобетон': 0
        }
        transition = {
            'Щебеночно - гравийно - песчаные': 0,
            'Грунт и малопрочные каменные материалы, укрепленные вяжущим': 0,
            'Грунт, укрепленный различными вяжущими и местными материалами': 0,
            'Булыжный и колотый камень(мостовые)': 0
        }
        lower = {'Грунт профилированный': 0,
                 'Грунт естественный': 0}
        type_of_coating = {'Капитальный': capital,
                           'Облегченный': lightweight,
                           'Переходный': transition,
                           'Низший': lower
                           }

        tuple_tip = v.get('Граница участка дороги', None).get('тип дорожной одежды', None)
        tuple_variant = v.get('Граница участка дороги', None).get('вид покрытия', None)
        for idx, tip in enumerate(tuple_tip):
            # находим следующий километровый
            if tip == tuple_tip[-1]:
                next_tip = tuple_tip[-1]
            elif tip == tuple_tip[0]:
                next_tip = tuple_tip[1]
            else:
                next_tip = tuple_tip[idx % len(tuple_tip) + 1]
            type_of_coating[tip[0]][tuple_variant[idx][0]] += next_tip[1] - tip[1]
        return type_of_coating

    def write_10(self, data):
        # 29.08.2023 разобраться с заполнением данных
        ws = self.wb['10']
        column_tuple = ['AF', 'AL', 'AR', 'AX', 'BD', 'BJ']
        counter = 0
        # в последующем убрать

        print(self.data.keys())
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
                print(result)

                # КАПИТАЛЬНЫЕ
                ws[f'{column}7'] = ''
                ws[f'{column}8'] = result.get('Капитальный').get('цементобетон') / 1000 if result.get('Капитальный').get('цементобетон') > 0 else ''
                ws[f'{column}9'] = result.get('Капитальный').get('ж/б плиты') / 1000 if result.get('Капитальный').get('ж/б плиты') > 0 else ''
                ws[f'{column}10'] = result.get('Капитальный').get('цементобетон') / 1000 if result.get('Капитальный').get('цементобетон') > 0 else ''
                ws[f'{column}11'] = result.get('Капитальный').get('цементобетон') / 1000 if result.get('Капитальный').get('цементобетон') > 0 else ''
                ws[f'{column}12'] = result.get('Капитальный').get('цементобетон') / 1000 if result.get('Капитальный').get('цементобетон') > 0 else ''
                ws[f'{column}13'] = result.get('Капитальный').get('асфальтобетон') / 1000 if result.get('Капитальный').get('асфальтобетон') > 0 else ''
                ws[f'{column}14'] = result.get('Капитальный').get('щебень/гравий, обр.вяжущий') / 1000 if result.get('Капитальный').get('щебень/гравий, обр.вяжущий') > 0 else ''
                ws[f'{column}15'] = result.get('Капитальный').get('тротуарная плитка') / 1000 if result.get('Капитальный').get('тротуарная плитка') > 0 else ''
                ws['B15'] = 'Тротуарная плитка'

                # ОБЛЕГЧЕННЫЕ
                ws[f'{column}19'] = result.get('Облегченный').get('асфальтобетонные') / 1000\
                    if result.get('Облегченный').get('асфальтобетонные') > 0 else ''
                ws[f'{column}20'] = result.get('Облегченный').get('органоминеральные') / 1000\
                    if result.get('Облегченный').get('органоминеральные') > 0 else ''
                ws[f'{column}21'] = result.get('Облегченный').get('щебеночные (гравийные), обработанные вяжущим') / 1000\
                    if result.get('Облегченный').get('щебеночные (гравийные), обработанные вяжущим') > 0 else ''
                ws[f'{column}22'] = result.get('Облегченный').get('цементобетон') / 1000\
                    if result.get('Облегченный').get('цементобетон') > 0 else ''
                ws[f'B22'] = 'Цементобетонные'

                # ПЕРЕХОДНЫЕ
                ws[f'{column}26'] = result.get('Переходный').get('Щебеночно - гравийно - песчаные') / 1000\
                    if result.get('Переходный').get('Щебеночно - гравийно - песчаные') > 0 else ''
                ws[f'{column}27'] = result.get('Переходный').get('Грунт и малопрочные каменные материалы, укрепленные вяжущим') / 1000 \
                    if result.get('Переходный').get('Грунт и малопрочные каменные материалы, укрепленные вяжущим') > 0 else ''
                ws[f'{column}28'] = result.get('Переходный').get('Грунт, укрепленный различными вяжущими и местными материалами') / 1000 \
                    if result.get('Переходный').get('Грунт, укрепленный различными вяжущими и местными материалами') > 0 else ''
                ws[f'{column}29'] = result.get('Переходный').get('Булыжный и колотый камень(мостовые)') / 1000 \
                    if result.get('Переходный').get('Булыжный и колотый камень(мостовые)') > 0 else ''

                # НИЗШИЕ
                ws[f'{column}34'] = result.get('Низший').get('Грунт профилированный') / 1000\
                    if result.get('Низший').get('Грунт профилированный') > 0 else ''
                ws[f'{column}35'] = result.get('Низший').get('Грунт естественный') / 1000 \
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


    def write_11 (self):
        '''
       21.09.2023 таблица  4.6 заполняется, ограничения нужна длина
        :return:
        '''
        ws = self.wb['11']
        # заполнение таблицы 4.6
        counter = 0
        column_tuple = ('AU', 'AX', 'BA', 'BD', 'BG', 'BJ', 'BM')

        sum_sign = {'всего': 0,
                    'предупреждающие': 0,
                    'приоритета': 0,
                    'запрещающие': 0,
                    'предписывающие': 0,
                    'особых предписаний': 0,
                    'сервиса': 0,
                    'информационные': 0,
                    'дополнительной информации': 0}

        for k1, v1 in self.data.items():
            if k1 == 'название дороги':
                continue
            column = column_tuple[counter]
            if len(self.data) > 2:
                ws[f'{column}6'] = f'Участок {counter + 1} \n {self.data_interface.get("year", None)} г.'
            else:
                ws[f'{column}6'] = f'{self.data_interface.get("year", None)}'

            # print(v1)
            # ws[f"{column}14"] = sum(1 for i in v1.get('Остановка').get('Наличие павильона') if i[0] == 'да') if v1.get(
            #     'Остановка', {}).get('Наличие павильона', []) else '-'


            ws[f"{column}16"] = sum(
                1 for i in v1.get('Проезжая часть').get('Назначение') if i[0] == 'площадка отдыха') if v1.get(
                'Проезжая часть', {}).get('Назначение', []) else '-'

            ws[f"{column}17"] = sum(
                1 for i in v1.get('Проезжая часть').get('Назначение') if i[0] == 'парковка') if v1.get(
                'Проезжая часть', {}).get('Назначение', []) else '-'

            ws[f"{column}19"] = round(sum(int(x[2]) - int(x[1]) for x in
                                          v1.get('Опоры освещения и контактной сети').get('Статус')) / 1000,
                                      3) if v1.get(
                'Опоры освещения и контактной сети', {}).get('Статус', []) else '-'
            ws[f"{column}23"] = round(sum(int(x[2]) - int(x[1]) for x in
                                          v1.get('Подземная комуникация').get('Вид коммуникации')) / 1000, 3) if v1.get(
                'Подземная комуникация', {}).get('Вид коммуникации', []) else '-'  # кабельные

            ws[f"{column}24"] = round(sum(int(x[2]) - int(x[1]) for x in
                                          v1.get('Воздушная коммуникация').get('Вид коммуникации')) / 1000,
                                      3) if v1.get(
                'Воздушная коммуникация', {}).get('Вид коммуникации', []) else '-'  # воздушные

            ws[f"{column}20"] = ((float(ws[f"{column}23"].value) if ws[f"{column}23"].value != '-' else 0) +
                                 (float(ws[f"{column}24"].value) if ws[f"{column}24"].value != '-' else 0))  # всего

            ws[f"{column}25"] = len(v1.get('Остановка').get('Название остановки')) if v1.get('Остановка', None) else '-'
            # ПСП
            ws[f"{column}26"] = sum(
                1 for i in v1.get('Проезжая часть', {}).get('Назначение', []) if
                i[0] in ['полоса торможения', 'полоса разгона']) if v1.get(
                'Проезжая часть').get('Назначение') else '-'
            ws[f"{column}28"] = round(sum(int(x[2]) - int(x[1]) for k in
                                          ['Нестандартное ограждение', 'Пешеходное ограждение', 'Тросовое ограждение',
                                           'Типа Нью-Джерси', 'Металическое барьерное ограждение', 'Сигнальные столбики'] for x in
                                          v1.get(k, {}).get('Статус', [])) / 1000, 3)  # ограждения
            # ws[f"{column}28"] = round(sum(int(x[2]) for k in
            #                               ['Нестандартное ограждение', 'Пешеходное ограждение', 'Тросовое ограждение',
            #                                'Типа Нью-Джерси', 'Металическое барьерное ограждение', ] for x in
            #                               v1.get(k, {}).get('Статус', [])) / 1000, 3)  # ограждения
            ws[f"{column}29"] = len(v1.get('Сигнальные столбики').get('Статус')) \
                if v1.get('Сигнальные столбики', {}).get('Статус', []) else '-'

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

            ws[f'{column}30'] = sum_sign.get('всего', '-')
            ws[f'{column}32'] = sum_sign.get('предупреждающие', '-')
            ws[f'{column}33'] = sum_sign.get('приоритета', '-')
            ws[f'{column}34'] = sum_sign.get('запрещающие', '-')
            ws[f'{column}35'] = sum_sign.get('предписывающие', '-')
            ws[f'{column}36'] = sum_sign.get('особых предписаний', '-')
            ws[f'{column}37'] = sum_sign.get('информационные', '-')
            ws[f'{column}38'] = sum_sign.get('сервиса', '-')
            ws[f'{column}39'] = sum_sign.get('дополнительной информации', '-')
            counter += 1
        else:
            if len(self.data) > 2:
                column = column_tuple[counter]
                ws[f'{column}6'] = 'Итог'
                ws[f'{column}14'] = f'=SUM({column_tuple[0]}14:{column_tuple[counter - 1]}14)'
                ws[f'{column}16'] = f'=SUM({column_tuple[0]}16:{column_tuple[counter - 1]}16)'
                ws[f'{column}17'] = f'=SUM({column_tuple[0]}17:{column_tuple[counter - 1]}17)'
                ws[f'{column}19'] = f'=SUM({column_tuple[0]}19:{column_tuple[counter - 1]}19)'
                ws[f'{column}20'] = f'=SUM({column_tuple[0]}20:{column_tuple[counter - 1]}20)'
                ws[f'{column}23'] = f'=SUM({column_tuple[0]}23:{column_tuple[counter - 1]}23)'
                ws[f'{column}24'] = f'=SUM({column_tuple[0]}24:{column_tuple[counter - 1]}24)'
                ws[f'{column}25'] = f'=SUM({column_tuple[0]}25:{column_tuple[counter - 1]}25)'
                ws[f'{column}26'] = f'=SUM({column_tuple[0]}26:{column_tuple[counter - 1]}26)'
                ws[f'{column}28'] = f'=SUM({column_tuple[0]}28:{column_tuple[counter - 1]}28)'
                ws[f'{column}29'] = f'=SUM({column_tuple[0]}29:{column_tuple[counter - 1]}29)'
                ws[f'{column}30'] = f'=SUM({column_tuple[0]}30:{column_tuple[counter - 1]}30)'
                ws[f'{column}32'] = f'=SUM({column_tuple[0]}32:{column_tuple[counter - 1]}32)'
                ws[f'{column}33'] = f'=SUM({column_tuple[0]}33:{column_tuple[counter - 1]}33)'
                ws[f'{column}34'] = f'=SUM({column_tuple[0]}34:{column_tuple[counter - 1]}34)'
                ws[f'{column}35'] = f'=SUM({column_tuple[0]}35:{column_tuple[counter - 1]}35)'
                ws[f'{column}36'] = f'=SUM({column_tuple[0]}36:{column_tuple[counter - 1]}36)'
                ws[f'{column}37'] = f'=SUM({column_tuple[0]}37:{column_tuple[counter - 1]}37)'
                ws[f'{column}38'] = f'=SUM({column_tuple[0]}38:{column_tuple[counter - 1]}38)'
                ws[f'{column}39'] = f'=SUM({column_tuple[0]}39:{column_tuple[counter - 1]}39)'

    def write_12 (self):
        ws = self.wb['12']
        # 4.7.1
        rows_avtovokzal = 11
        rows_gbdd = 11
        rows_sto = 37
        rows_hotels = 37
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

    def write_13 (self):
        ws = self.wb['13']

        rows_azs = 10
        rows_car_wash = 10
        rows_ws = 37
        rows_food = 37
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
                    ws[f'I{rows_ws}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                        {}).get(
                        'Привязка по оси') else ''
                    ws[f'N{rows_ws}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание', {}).get(
                        'Наименование') else ''
                    rows_ws += 1
                elif value[0] == 'Кафе/столовая/ресторан':
                    # 4.7.8
                    ws[f'AJ{rows_food}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание',
                                                                                                        {}).get(
                        'Наименование') else ''
                    ws[f'AS{rows_food}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                        'Адрес') else ''
                    ws[f'BD{rows_food}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                           {}).get(
                        'Привязка по оси') else ''
                    rows_food += 1

    def write_14 (self):
        ws = self.wb['14']
        rows_med = 8
        for name_district, obj in self.data.items():
            if name_district == 'название дороги':
                continue

            for idx, value in enumerate(obj.get('Здание', {}).get('Назначение', [])):

                if value[0] == 'Пункты первой медицинской помощи/почта/телефон':
                    # 4.7.5
                    ws[f'B{rows_med}'] = obj.get('Здание', {}).get('Адрес')[idx][0] if obj.get('Здание', {}).get(
                        'Адрес') else ''
                    ws[f'O{rows_med}'] = obj.get('Здание', {}).get('Привязка по оси')[idx][0] if obj.get('Здание',
                                                                                                         {}).get(
                        'Привязка по оси') else ''
                    ws[f'Y{rows_med}'] = obj.get('Здание', {}).get('Наименование')[idx][0] if obj.get('Здание', {}).get(
                        'Наименование') else ''
                    rows_med += 1

    def write_17(self):
        """
        27.09.2023
        :return:
        """
        ws = self.wb['17']
        counter = 0
        column_tuple = ('J', 'O', 'T', 'Y', 'AD')
        cells = ('L','Q','V','AA','AF')

        pipes = {
            "Металлические" : [0,0],
            "Железобетонные": [0,0],
            "Бетоннометаллические": [0,0],
            "Каменные": [0,0],
            "Деревянные": [0,0],
            "Асбестоцементные": [0,0],
        }
        # 4.10.2 Сводная ведомость наличия тоннелей, галерей и пешеходных переходов в разных уровнях
        types_of_structures = {
            "Тоннель (галерея)": [0, 0],
            "Пешеходный переход подземный": [0, 0],
            "Пешеходный переход надземный": [0, 0],
            "Водопропускная труба": pipes
        }

        def count_4_10_2(types_of_structures, column, cell):
            for key, value in types_of_structures.items():
                if self.data.get(f'участок {counter + 1}').get(key) == None:
                    continue
                else:
                    if key == 'Водопропускная труба':
                        print(self.data.get(f'участок {counter + 1}').get(key))
                        for lst in self.data.get(f'участок {counter + 1}').get(key).get('Материал'):
                                if lst[0] == 'металл':
                                    types_of_structures.get(key).get('Металлические')[0] += 1
                                    types_of_structures.get(key).get('Металлические')[1] += lst[4]
                                elif lst[0] == 'ж/б':
                                    types_of_structures.get(key).get('Железобетонные')[0] += 1
                                    types_of_structures.get(key).get('Железобетонные')[1] += lst[4]

                        ws[f'{column}37'] = types_of_structures.get(key).get('Металлические')[0]
                        ws[f'{cell}37'] = types_of_structures.get(key).get('Металлические')[1]
                        ws[f'{column}38'] = types_of_structures.get(key).get('Железобетонные')[0]
                        ws[f'{cell}38'] = types_of_structures.get(key).get('Железобетонные')[1]

                    else:
                        result = self.data.get(f'участок {counter + 1}').get(key)
                        types_of_structures.get(key)[0] += 1
                        types_of_structures.get(key)[1] += result.get(list(result.keys())[0])[0][3]
                        if key == 'Тоннель (галерея)':
                            ws[f'{column}14'] = types_of_structures.get(key)[0]
                            ws[f'{cell}14'] = types_of_structures.get(key)[1]
                        elif key == 'Пешеходный переход подземный':
                            ws[f'{column}20'] = types_of_structures.get(key)[0]
                            ws[f'{cell}20'] = types_of_structures.get(key)[1]
                        elif key == 'Пешеходный переход надземный':
                            ws[f'{column}19'] = types_of_structures.get(key)[0]
                            ws[f'{cell}19'] = types_of_structures.get(key)[1]


        for k, v in self.data.items():
            if k == 'название дороги':
                continue
            else:
                column = column_tuple[counter]
                cell = cells[counter]
                if len(self.data) > 2:
                    ws[f'{column}6'] = f'Участок {counter + 1}'
                    count_4_10_2(types_of_structures, column, cell)
                    types_of_structures = {
                        "Тоннель (галерея)": [0, 0],
                        "Пешеходный переход подземный": [0, 0],
                        "Пешеходный переход надземный": [0, 0],
                        "Водопропускная труба": pipes
                    }
                else:
                    ws[f'{column}6'] = f'{self.data_interface.get("year", None)}'
                    count_4_10_2(types_of_structures, column, cell)
            counter += 1
            ""

    def write_18(self):
        """
        Описиваем данные по 18 листу
        :return:
        """
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

        # 4.10.9 Сводная ведомость съездов (въездов)
        for k1, v1 in self.data.items():
            if k1 == 'название дороги':
                continue
            else:
                column = column_tuple[counter]
                cell = cells[counter2]
                if len(self.data) > 2:
                    ws[f'{column}30'] = f'Участок {counter + 1} \n {self.data_interface.get("year", None)} г.'
                else:
                    ws[f'{column}30'] = f'{self.data_interface.get("year", None)}'
                for lst in self.data.get(k1).get('Съезд').get('Тип покрытия'):
                    if lst[0] == 'асфальтобетон':
                        types.get('Асфальтобетонные')[0] += 1
                        types.get('Асфальтобетонные')[1] += lst[3]
                    elif lst[0] == 'Цементобетонные':
                        pass
                    elif lst[0] == 'тротуарная плитка':
                        ws['AI38'] = "Тротуарная плитка"
                        types.get('Тротуарная плитка')[0] += 1
                        types.get('Тротуарная плитка')[1] += lst[3]
                    elif lst[0] == 'Щебеночные':
                        pass
                    elif lst[0] == 'Грунтовые':
                        pass
                    elif lst[0] == 'Ж/б плиты':
                        pass

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
                counter += 1
                counter2 += 2

            types = {
                "Асфальтобетонные": [0, 0],
                "Цементобетонные": [0, 0],
                "Тротуарная плитка": [0, 0],
                "Щебеночные": [0, 0],
                "Грунтовые": [0, 0],
                "Ж/б плиты": [0, 0],
            }

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
    """ Конвертируем визио в svg"""
    visio = win32com.client.Dispatch("Visio.Application")
    doc = visio.Documents.Open(rf"{path_dir}\линейный график.vsd")

    for page in doc.Pages:
        page.Export(rf'{path_dir}\линейный график\линейный график{page.Name}.png')
    visio.Quit()


def main ():
    conn = db.Query('OMSK_CITY_2023')
    data = conn.get_tp_datas('ул. Моторная')
    # data = conn.get_tp_datas('ул. Масленникова')


    report = WriterExcelTP(data)
    # report.write_titular()
    # report.write_scheme()
    # report.write_6(data)
    # report = WriterExcelDAD()
    # report.write_diagrams1()
    report.save_file()



if __name__ == "__main__":
    import time

    start = time.time()  # точка отсчета времени
    main()
    end = time.time() - start  # собственно время работы программы
    print(f"{round(end, 1)} секунд")  # вывод времени
