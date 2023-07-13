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


    def get_databases(self):
        request = "select name from sys.databases"
        self.cursor.execute(request)
        self.db_list = [i[0] for i in self.cursor.fetchall()[4:]]
        return self.db_list

    def set_road_name(self):
        '''
            Вытаскивает из базы столбцы с названиями а/д
        :return: self.road_names -> dict
        '''
        self.cursor.execute("SELECT ID_Road, Name FROM Road")
        self.road_names = {'№': [], 'ID_Road': [], 'Название': []}
        for row in self.cursor.fetchall():
            # self.road_names['№'].append(row[0])
            self.road_names['ID_Road'].append(row[0])
            self.road_names['Название'].append(row[1])
            # print(row[1])
        print(self.road_names['Название'])
        return self.road_names['Название']

    def get_dad_datas(self, name):
        """ 29.06.2023 """
        """
        Вытаскиваем данные из бд и формируем словарь по атрибутам
        :param attr_list: хранит в себе названия атрибутов
        :param name: название автомобильной дороги 
        :return: возвращаем словарь с данными по диагностике
        """
        dad_datas_dict = {}

        attr_list = [
            'Оценка ровности IRI', 'Глубина колеи', 'Граница участка дороги', 'Ось дороги', 'Кривая',
            'Ширина проезжей части'
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

        for num, attr in enumerate(attr_list):
            self.cursor.execute(request, (name, attr))
            datas_dict = {
                'Name': [], 'Description_uch1': [], 'L1-Начало': [], 'L2-Конец': [],
                'ID_Param': [], 'ValueParam': [], 'ID_Type_Attr': [], 'Item_Name': []
            }
            for row in self.cursor.fetchall():
                """ заполяяем список данными из базы """

                # datas_dict['ID_Road'] = row[0]
                datas_dict['Name'] = row[1]
                # datas_dict['ID_Way'].append(row[2])
                # datas_dict['Description'].append(row[3])  # направление движения по записи
                # datas_dict['ID_High'].append(row[4])
                datas_dict['Description_uch1'].append(row[5])
                # datas_dict['ID_Attribute'].append(row[6])
                datas_dict['L1-Начало'].append(row[7])
                datas_dict['L2-Конец'].append(row[8])
                datas_dict['ID_Param'].append(row[9])
                datas_dict['ValueParam'].append(row[10])
                datas_dict['ID_Type_Attr'].append(row[11])
                datas_dict['Item_Name'].append(row[12])
            dad_datas_dict[attr] = datas_dict

        for key, value in dad_datas_dict.items():
            print(key, value)
        return dad_datas_dict

    def close_db(self):
        self.db.close()


def databases():
    request = "select name from sys.databases"
    db = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server}; Server=SIBREGION-SRV2; Trusted_Connection=yes;')
    print()
    db_list = [i[0] for i in db.cursor().execute(request).fetchall()[4:]]

    db.close()
    return db_list

# databases()
