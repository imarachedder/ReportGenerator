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
        self.data_interface = {'year': 2023}

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
        # self.write_titular(data)
        # self.write_6(data)
        # self.write_9(data)
        # self.write_10(data)
        self.write_11()
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

        # for i in range(self.data.get('count_region', 1)):
        #     ws[f"L1{i}"].value = self.data.get('участки', 'None')[i]  # 2.2 Участок дороги: участки
        self.data['count_region'] = 2
        # 2.2 Участок дороги 1, 2 и т.д.
        if self.data.get('count_region') > 1:
            for i in range(0, self.data.get("count_region", 0)):
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
            ws["L10"].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы')[0][1]} + " \
                                 f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы')[0][2]} м."


        # 2.3 Суммарна протяженность по участку или участкам
        if self.data.get("count_region") > 1:
            res = 0
            for i in range(0, self.data.get("count_region", 0)):
                res += int(self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[0][2])
                n += 1
            n = 1
            ws["S14"].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы')[0][1]} + {res} м."
        else:
            ws["S14"].value = self.data.get(f'участок {n}').get('Ось дороги', None)['Начало трассы'][0][2]

        # заполняет таблицу 2.3 Протяженность дороги
        for i in range(2, (self.data.get('count_region') * 2) + 1, 2):
            # print(self.data.get('count_region') * 2)
            ws[f'B2{i - 1}'].value = f'Участок {n}'
            # print(self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[4%self.data.get('count_region')][1])
            if n % 2 != 0:
                ws[f'B2{i}'].value = self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i%self.data.get('count_region')][1]
                ws[f'F2{i}'].value = self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i%self.data.get('count_region') ][2]
                ws[f'J2{i}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region')][1]} + " \
                                      f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region')][2]} м."
            else:
                ws[f'B2{i}'].value = \
                self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region') -1][1]
                ws[f'F2{i}'].value = \
                self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region') -1][2]
                ws[f'J2{i}'].value = f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region') -1][1]} + " \
                                     f"{self.data.get(f'участок {n}').get('Ось дороги', None).get('Начало трассы', 0)[i % self.data.get('count_region') -1][2]} м."

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
        if self.data.get("count_region") > 1:
            counter = 1
            ws["AL10"].value = self.data.get('название дороги')
            for i in range(0, self.data.get("count_region")):
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

    def write_7(self, data):
        pass

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
            print(f"do {res}",
                  self.data.get(f'участок {i + 1}').get('Ширина проезжей части', None).get('Ширина ПЧ')[
                      j - 1][0],
                  int(self.data.get(f'участок {i + 1}').get('Ширина проезжей части').get('Ширина ПЧ')[
                          j - 1][1]), result, res)
            return result

        # Счетчик
        n = 1

        ws = self.wb['9']
        # 4.1 Топографические условия района проложения автомобильной дороги
        ws['B7'] = self.data.get('area_conditioins')
        # 4.2 Ширина земляного полотна
        # 4.3 Характеристика проезжей части
        # 4.3.1 Ширина проезжей части
        self.data['count_region'] = 2

        if self.data.get('count_region') >= 1:

            # Цикл по количеству учасков
            for i in range(0, self.data.get('count_region')):
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

    def write_10(self, data):
        # 29.08.2023 разобраться с заполнением данных
        types_of_coating = {'Капитальные': ['Цементобетонные монолитные', 'Железобетонные монолитные',
                                            'Железобетонные сборные', 'Армобетонные монолитные',
                                            'Армобетонные сборные', 'Асфальтобетонные', 'Щебеночно-мастичные'],
                            'Облегченные': ['Асфальтобетонные', 'Органоминеральные',
                                            'Щебеночные (гравийные), обработанные вяжущим'],
                            'Переходные': ['Щебеночно-гравийно-песчанные',
                                           'Грунт и малопрочные каменные материалы, укрепленные вяжущим',
                                           'Грунт, укрепленный различными вяжущими и местными материалами',
                                           'Булыжный и колотый камень (мостовые)'],
                            'Низший': ['Грунт профилированный', 'Грунт естественный']}
        print(self.data.get('участок 1').get('Граница участка дороги').get())
        n = 1
        ws = self.wb['10']
        self.data['count_region'] = 2
        if self.data.get('count_region') >= 1:
            for i in range(0, self.data.get('count_region')):
                if i == 0:
                    ws['AF4'].value = f'Участок {n}\n 2023 г.'
                elif i == 1:
                    ws['AL4'].value = f'Участок {n}\n 2023 г.'
                n += 1
            print(self.data.get('участок 1').get('Проезжая часть'))

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
            ws[f"{column}14"] = sum(1 for i in v1.get('Остановка').get('Наличие павильона') if i[0] == 'да') if v1.get(
                'Остановка', {}).get('Наличие павильона', []) else '-'

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
    # data = conn.get_tp_datas('ул. Интернациональная')
    data = conn.get_tp_datas('ул. Масленникова')


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
