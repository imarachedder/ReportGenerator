# -*- coding: utf-8 -*-
import pymssql
import pandas as pd
import pyodbc
import sys


class Query:
    def __init__(self, database=None):
        __SERVER = "SIBREGION-SRV2"
        __USER = "sibregion"
        __PASSWORD = "SibU$r2018"
        __DATABASE = str(database)
        # __DATABASE = "ZLATOUST_TEST_2021"
        print(__DATABASE)

        self.db = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                                 'Server=' + __SERVER + ';'
                                                        'Database=' + __DATABASE + ';'
                                                                                   'Trusted_Connection=yes;')
        self.db.setencoding('cp1251')
        self.cursor = self.db.cursor()  # возвращает значение стобец-строка в формате словаря ключ-значение

    def get_dad_datas(self, name: str):
        """
        Вытаскивает данные по измерениям для диагностики из базы данных

        :param name: Название автомобильной дороги из выбранной базы
        :param param_list: Хранит данные по измерения для диагностики автомобильной дороги
                            'Ось дороги', 'Колейность', 'Ровность по IRI', 'Ширина ПЧ', 'Продольный уклон'
        :return: datas_dict: Возвращает словарь с данными по измерениям
        """
        dad_datas_dict = {}

        """params_list = [
            "Граница участка дороги", "Ось дороги", "Колейность", "Ровность по IRI", "Кривая",
            "Ширина проезжей части"
        ]"""  # СТАРАЯ ВЕРСИЯ БАЗЫ

        params_list = [
            'Оценка ровности IRI', 'Глубина колеи', 'Граница участка дороги', 'Ось дороги', 'Кривая',
            'Ширина проезжей части',

        ]
        request = """
            select Road.ID_Road, Road.Name, Way.ID_Way, Way.Description, High.ID_High, High.Description,
            Attribute.ID_Attribute, Attribute.L1, Attribute.L2, Params.ID_Param, 
            Params.ValueParam, Group_Description.ID_Type_Attr, Group_Description.Item_Name
            from Road inner join Way on Road.ID_Road = Way.ID_Road
            inner join High on Way.ID_Way = High.ID_Way
            inner join Attribute on High.ID_High = Attribute.ID_High 
            inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
            inner join Group_Description on Attribute.ID_Type_Attr = Group_Description.ID_Type_Attr
            where Road.Name = ? and Group_Description.Item_Name = ?
        """
        for num, param in enumerate(params_list):
            self.cursor.execute(request, (name, param))
            datas_dict = {
                'ID_Road': [], 'Name': [], 'ID_Way': [], 'Description': [], 'ID_High': [], 'Description_uch1': [],
                'ID_Attribute': [], 'L1-Начало': [], 'L2-Конец': [],
                'ID_Param': [], 'ValueParam': [], 'ID_Type_Attr': [], 'Item_Name': []
            }
            # print(self.cursor.fetchone())

            for row in self.cursor.fetchall():
                # datas_dict['ID_Road'] = row[0]
                datas_dict['Name'] = row[1]
                # datas_dict['ID_Way'].append(row[2])
                # datas_dict['Description'].append(row[3])  # направление движения по записи
                # datas_dict['ID_High'].append(row[4])
                # datas_dict['Description_uch1'].append(row[5])
                # datas_dict['ID_Attribute'].append(row[6])
                datas_dict['L1-Начало'].append(row[7])
                datas_dict['L2-Конец'].append(row[8])
                datas_dict['ID_Param'].append(row[9])
                datas_dict['ValueParam'].append(row[10])
                datas_dict['ID_Type_Attr'].append(row[11])
                datas_dict['Item_Name'].append(row[12])
            dad_datas_dict[param] = datas_dict

        for key, value in dad_datas_dict.items():
            print(key, value)

    def get_tp_datas(self, road_name):
        """
        13.07.2023 г. Обновил запросы, и списов параметров
        Вытаскивает данные по техническому паспорту из базы данных
        :return:
        """
        res = {'название дороги': f'{road_name}', }
        request_for_items = "select Item_Name from Group_Description"
        item_list = ['Ось дороги', 'Граница участка дороги', 'Километровые знаки', 'Остановка',
                     'Опоры освещения и контактной сети', 'Проезжая часть']  # 'Граница участка дороги'
        request = """
            select Road.ID_Road, Road.Name, Way.Description,High.Description, Way.ID_Way,
            Params.ID_Param, Group_Description.Item_Name, Types_Description.Param_Name, Params.ValueParam,
            Attribute.L1, Attribute.L2, dbo.CalcSquare(Image_Points) as Square
            from Road inner join Way on Road.ID_Road = Way.ID_Road
            inner join High on Way.ID_Way = High.ID_Way
            inner join Attribute on High.ID_High = Attribute.ID_High
            inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
            inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
            inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
            where Road.Name = ? 

        """  # and Group_Description.Item_Name = ?
        # for i, item in enumerate(item_list):
        self.cursor.execute(request, road_name)  # item
        for param in self.cursor.fetchall():

            print(param)

            if param[6] in res:
                if param[7] in res.get(param[6]):
                    res.get(param[6]).get(param[7]).append(param[8::])
                else:
                    res.get(param[6]).update({param[7]: [param[8::]]})
            else:
                res.update({param[6]: {param[7]: [param[8::]]}})

        # сортирует в словаре координаты по возрастанию
        for _, value in res.items():
            if type(value) == dict:
                for val in value.values():
                    val.sort(key=lambda x: x[1])

        return res

    def close_db(self):
        return self.db.close()


def databases():
    request = "select name from sys.databases"
    db = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server}; Server=SIBREGION-SRV2; Trusted_Connection=yes;')
    print()
    db_list = [i[0] for i in db.cursor().execute(request).fetchall()[4:]]

    db.close()
    return db_list


def test(data):
    list_start_end_road = data.get('Ось дороги', {}).get('Начало трассы', [])
    list_start_end_km_sign = data.get('Километровые знаки').get('Значение в прямом направлении')
    l = {}
    # min(enumerate(a), key = lambda x: abs(x[1] - 11.5))
    for num_road in list_start_end_road:
        for num_sign in list_start_end_km_sign:
            if -1000 < num_road[1] - num_sign[1] < 1000:
                print(num_road, 'start',  num_sign)
                #l['start'] += (num_road[1] - num_sign[1])
                if -1000 < num_road[2] - num_sign[2] < 1000:
                    print(num_road, 'end', num_sign)

def main():
    df = Query(database="Testovaya")
    # df.get_dad_datas(name='Adigeya-Maykop')
    res = df.get_tp_datas(road_name="P-254")
    print(res)
    test(res)
    # print(df.get_dad_datas(name='Adigeya-Maykop'))


if __name__ == "__main__":
    main()
