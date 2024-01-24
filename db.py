# # -*- coding: utf-8 -*-
# import sys
# from time import sleep
#
# import pyodbc
# from icecream import ic
#
# import settings
#
#
# class Query:
#     def __init__ (self, database = settings.database):
#         __SERVER = settings.server
#         __USER = settings.user
#         __PASSWORD = settings.password
#         __DATABASE = str(database)
#         # print(__DATABASE)
#         self.db = pyodbc.connect('Driver={SQL Server};'
#                                  'Server=' + __SERVER + ';'
#                                                         'Database=' + __DATABASE + ';'
#                                                                                    'Trusted_Connection=yes;')
#         self.db.setencoding('cp1251')
#         self.cursor = self.db.cursor()  # возвращает значение стобец-строка в формате словаря ключ-значение
#
#     def __enter__ (self):
#         return self
#
#     def __exit__(self, exc_type, exc_val, exc_tb):
#         self.db.close()
#
#     def get_tp_datas (self, road_name):
#         """
#         последовательность объектов:
#         {название дороги: название дороги,
#         участок_дороги:{
#             объект:{
#             свойство объекта: [(значение, длина, площадь, id параметра, количество точек, минимальное отклонение,
#             максимальное отклонение, начало_метры, конец_метры, (начало_километры,начало_метры),
#             (конец_километры, конец_метры)), ....], ...}, ...}, ...}
#         :return:
#         """
#
#         res_km = {}
#         #Road.ID_Road,
#         request_km = """ select  Road.Name,  High.Description, Group_Description.Item_Name,
#                      Types_Description.Param_Name, Params.ValueParam, dbo.CalcLength(ATTRIBUTE.Image_Points) as Length,
#                      dbo.CalcSquare(ATTRIBUTE.Image_Points) as Square, Params.ID_Param, dbo.GetCountPoint(Image_Points),
#                      dbo.Amin(Image_Points) as Amin, dbo.Amax(Image_Points) as Amax,
#                     Attribute.L1, Attribute.L2
#                     from Road inner join Way on Road.ID_Road = Way.ID_Road
#                     inner join High on Way.ID_Way = High.ID_Way
#                     inner join Attribute on High.ID_High = Attribute.ID_High
#                     inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
#                     inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
#                     inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
#                     where Road.Name = ? and Group_Description.Item_Name = 'Километровые знаки' """
#         self.cursor.execute(request_km, road_name)
#         for param in self.cursor.fetchall():
#             # последовательность объектов:[(значение, длина, площадь, id параметра, количество точек, минимальное отклонение,
#             # максимальное отклонение, начало_метры, конец_метры, (начало_километры,начало_метры), (конец_километры, конец_метры)), ....]
#             #print(param)
#             if param[1] in res_km:
#
#                 if param[2] in res_km.get(param[1]):
#                     if param[3] in res_km.get(param[1]).get(param[2]):
#                         res_km.get(param[1]).get(param[2]).get(param[3]).append(param[4::])
#                     else:
#                         res_km.get(param[1]).get(param[2]).update({param[3]: [param[4::]]})
#                 else:
#                     res_km.get(param[1]).update({param[2]: {param[3]: [param[4::]]}})
#             else:
#                 res_km.update({param[1]: {param[2]: {param[3]: [param[4::]]}}})
#
#         self.sort_dict_binding(res_km)
#
#         res = {'название дороги': f'{road_name}', }
#
#         request = """
#                     select  Road.Name,  High.Description, Group_Description.Item_Name,
#                     Types_Description.Param_Name, Params.ValueParam, dbo.CalcLength(ATTRIBUTE.Image_Points) as Length,
#                     dbo.CalcSquare(ATTRIBUTE.Image_Points) as Square, Params.ID_Param, dbo.GetCountPoint(Image_Points),
#                     dbo.Amin(Image_Points) as Amin, dbo.Amax(Image_Points) as Amax,
#                     Attribute.L1, Attribute.L2
#                     from Road inner join Way on Road.ID_Road = Way.ID_Road
#                     inner join High on Way.ID_Way = High.ID_Way
#                     inner join Attribute on High.ID_High = Attribute.ID_High
#                     inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
#                     inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
#                     inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
#                     where Road.Name = ?
#
#                 """  # and Group_Description.Item_Name = ?
#
#         self.cursor.execute(request, road_name)  # item
#
#         for param in self.cursor.fetchall():
#
#             tuple_km = tuple(
#                     res_km.get(param[1], {}).get('Километровые знаки', {}).get('Значение в прямом направлении', []))
#
#             if tuple_km:
#                 km = self.convert_m_to_km(param, tuple_km)
#             else:
#                 km = ((0, param[-2]), (0, param[-1]))
#             coordinates = (*param[4::], *km)
#
#
#             #print(param, coordinates)
#
#             if param[1] in res:
#
#                 if param[2] in res.get(param[1]):
#                     if param[3] in res.get(param[1]).get(param[2]):
#
#                         res.get(param[1]).get(param[2]).get(param[3]).append(coordinates)
#                     else:
#                         res.get(param[1]).get(param[2]).update({param[3]: [coordinates]})
#                 else:
#                     res.get(param[1]).update({param[2]: {param[3]: [coordinates]}})
#             else:
#                 res.update({param[1]: {param[2]: {param[3]: [coordinates]}}})
#
#         # сортирует в словаре координаты по возрастанию
#         self.sort_dict_binding(res)
#         res = dict(sorted(res.items()))
#         return res
#
#     def get_databases (self):
#         """Получить список баз данных"""
#         request = "select name from sys.databases"
#         self.cursor.execute(request)
#         db_list = [i[0] for i in self.cursor.fetchall()[4:]]
#
#         return db_list
#
#     def get_list_roads (self) -> list:
#         """
#               Вытаскивает из базы столбцы с названиями дорог
#               :return: self.road_names -> dict
#               """
#         self.cursor.execute("SELECT ID_Road, Name FROM Road")
#         self.road_names = [row[1] for row in self.cursor.fetchall()]
#         # for row in self.cursor.fetchall():
#         #     self.road_names.append(row[1])
#             # print(row[1])
#         #print(self.road_names)
#
#         return self.road_names
#
#     def get_dad_datas (self, name):
#         """ 29.06.2023 """
#         """
#         Вытаскиваем данные из бд и формируем словарь по атрибутам
#         :param attr_list: хранит в себе названия атрибутов
#         :param name: название автомобильной дороги
#         :return: возвращаем словарь с данными по диагностике
#         """
#         res = {'название дороги': f'{name}', }
#         # Road.ID_Road,
#         request = """select  Road.Name, Way.Description,High.Description, Way.ID_Way,
#             Params.ID_Param, Group_Description.Item_Name, Types_Description.Param_Name, Params.ValueParam,
#             Attribute.L1, Attribute.L2, dbo.CalcSquare(ATTRIBUTE.Image_Points) as Square
#             from Road inner join Way on Road.ID_Road = Way.ID_Road
#             inner join High on Way.ID_Way = High.ID_Way
#             inner join Attribute on High.ID_High = Attribute.ID_High
#             inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
#             inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
#             inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
#             where Road.Name = ? """
#         # and Group_Description.Item_Name = ?
#
#         # for _, attr in enumerate(attr_list):
#         self.cursor.execute(request, name)
#
#         for param in self.cursor.fetchall():
#             """ заполяяем список данными из базы """
#             print(param)
#             if param[6] in res:
#                 if param[7] in res.get(param[6]):
#                     res.get(param[6]).get(param[7]).append(param[8::])
#                 else:
#                     res.get(param[6]).update({param[7]: [param[8::]]})
#             else:
#                 res.update({param[6]: {param[7]: [param[8::]]}})
#
#         return res
#
#     @classmethod
#     def databases (cls):
#         request = "select name from sys.databases"
#         db = pyodbc.connect(f'Driver={settings.driver}; Server={settings.server}; Trusted_Connection=yes;')
#
#         db_list = [i[0] for i in db.cursor().execute(request).fetchall()[4:]]
#         # print(db_list)
#         # db.close()
#         return db_list
#
#     def convert_m_to_km (self, param, list_km):
#         '''
#         переводит метры в километры с привязкой к километровым знакам
#         :param param: объект из базы данных
#         :param list_km: список километровых знаков
#         :param distance_km: список дистанций между километровыми знаками
#         :return: привязки начала и конца объекта
#         '''
#         # частный случай километровые знаки
#         if param[2] == 'Километровые знаки':
#             return (int(param[-7]), 0), (int(param[-7]), 0)
#         # если точечный объект
#         if param[-2] == param[-1]:
#             # print('start==end')
#
#             for idx_km, num_sign in enumerate(list_km):
#                 # находим следующий километровый
#                 if num_sign == list_km[-1]:
#                     next_km = list_km[-1]
#                 elif num_sign == list_km[0]:
#                     next_km = list_km[1]
#                 else:
#                     next_km = list_km[idx_km % len(list_km) + 1]
#
#                 if num_sign[-2] <= param[-2] < next_km[-2]:
#                     return (int(num_sign[0]), param[-2] - num_sign[-2]), (int(num_sign[0]), param[-1] - num_sign[-1])
#                 elif param[-2] >= list_km[-1][-2]:
#                     return (int(list_km[-1][0]), param[-2] - list_km[-1][-2]), (int(list_km[-1][0]), param[-1] - list_km[-1][-1])
#                 # elif tmp[-2] <= param[-2] < num_sign[-2]:
#                 #     return (tmp[0], param[-2] - tmp[-2]), (tmp[0], param[-1] - tmp[-1])
#                 elif param[-2] < num_sign[-2]:
#                     return (int(num_sign[0]), param[-2] - num_sign[-2]), (int(num_sign[0]), param[-1] - num_sign[-1])
#                 # elif next_km[-2] <= param[-2]:
#                 #     continue
#                 #     # return (next_km[0], param[-2] - next_km[-2]), (next_km[0], param[-1] - next_km[-1])
#                 else:
#                     continue
#         # если объект линейный или площадной
#         elif param[-2] != param[-1]:
#             start_km = 0
#             end_km = 0
#             start_m = 0  # начало
#             end_m = 0  # конец
#             idx = 0  # индекс километрового start
#             # ищем start
#             for idx_km, num_sign in enumerate(list_km):
#                 if num_sign == list_km[-1]:
#                     next_km = list_km[-1]
#                     idx = idx_km
#                 elif num_sign == list_km[0]:
#                     next_km = list_km[1]
#                     idx = idx_km
#                 else:
#                     next_km = list_km[idx_km % len(list_km) + 1]
#                     idx = idx_km
#                 if num_sign[-2] <= param[-2] < next_km[-2] or param[-2] < num_sign[-2]:
#                     start_km = num_sign[0]
#                     start_m = param[-2] - num_sign[-2]
#                     break
#                 elif param[-2] > list_km[-1][-2]:
#                     start_km = list_km[-1][0]
#                     start_m = param[-2] - list_km[-1][-2]
#                     break
#                 else:
#                     continue
#             # ищем end начиная с idx благодаря этому отсекаются уже пройденные километровые
#             for idx_km, num_sign in enumerate(list_km[idx:]):
#                 if num_sign == list_km[-1]:
#                     next_km = list_km[-1]
#                 elif num_sign == list_km[0]:
#                     next_km = list_km[1]
#                 else:
#                     next_km = list_km[idx % len(list_km) + 1]
#                 if num_sign[-1] <= param[-1] < next_km[-1] or param[-1] < num_sign[-1]:
#                     end_km = num_sign[0]
#                     end_m = param[-1] - num_sign[-1]
#                     break
#                 elif param[-1] > list_km[-1][-1]:
#                     end_km = list_km[-1][0]
#                     end_m = param[-1] - list_km[-1][-1]
#                     break
#
#                 else:
#                     continue
#             return (int(start_km), start_m), (int(end_km), end_m)
#
#     def sort_dict_binding (self, res):
#         '''
#         сортирует словарь с объектами по первой метровой привязке
#         :param res:
#         :return: res
#         '''
#
#         for _, value in res.items():
#             if type(value) == dict:
#                 for i, val in value.items():
#                     if type(val) == dict:
#                         for elem in val.values():
#                             elem.sort(key = lambda x: (x[-2], x[-1]))
#                                 #elem.sort(key = lambda x: (x[-2][0], x[-2][1]))
#
#     def close_db (self):
#         self.cursor.close()
#         self.db.close()
#
#
# # def count_coating (data):
# #     type_of_coating = {'Капитальный': {},
# #                        'Облегченный': {},
# #                        'Переходный': {},
# #                        'Низший': {}}
# #
# #     for k, v in data.items():
# #         if k == 'название дороги':
# #             continue
# #         tuple_tip = v.get('Граница участка дороги', None).get('тип дорожной одежды', None)
# #         tuple_variant = v.get('Граница участка дороги', None).get('вид покрытия', None)
# #         for idx_km, num_sign in enumerate(tuple_tip):
# #             # находим следующий километровый
# #             if num_sign == tuple_tip[-1]:
# #                 next_km = tuple_tip[-1]
# #             elif num_sign == tuple_tip[0]:
# #                 next_km = tuple_tip[1]
# #             else:
# #                 next_km = tuple_tip[idx_km % len(tuple_tip) + 1]
# #
# #             type_of_coating[num_sign[0]][tuple_variant[idx_km][0]] = next_km[1] - num_sign[1]
# #         print(type_of_coating)
# #
# #         return 0
#
#
#
# def main ():
#     #db = Query('ZAP2023')  # FKU_VOLGO_VYATSK_1 OMSK_CITY_2023 Query('IZHEVSK_CITY_2023')
#     with Query('OMSK_CITY_2023') as db:
#         databases = db.get_list_roads()
#         data = db.get_tp_datas('ул. Интернациональная') #'2477_Проезд от ул. Воровского до МБДОУ «Детский Сад № 92» (Воровского, 151а)'
#     print(databases)
#     print(data)
#
#
#
#
#
# if __name__ == '__main__':
#     main()


# -*- coding: utf-8 -*-
import pymssql
import pandas as pd
import pyodbc
import sys
import settings

# стоянки, парковки смотеть в проезжей части, счтиается количество
# Опоры освещения и контактные сети, считается длина
# воздушная коммуникация считается сумма длина
# Остановки количество
# ПСП считаются длина, сумма
# ограждения, считаются сумма длин
# Сигнальные столбики, шт.
# Дорожные знаки, количество


class Query:
    def __init__(self, database=None):
        __SERVER = settings.server
        __USER = settings.user
        __PASSWORD = settings.password
        __DATABASE = str(database)
        print(__DATABASE)
        self.db = pyodbc.connect('Driver={SQL Server};'
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
        """
            Вытаскивает из базы столбцы с названиями а/д
        :return: self.road_names -> dict
        """
        self.cursor.execute("SELECT ID_Road, Name FROM Road")
        self.road_names = {'№': [], 'ID_Road': [], 'Название': []}
        for row in self.cursor.fetchall():
            # self.road_names['№'].append(row[0])
            self.road_names['ID_Road'].append(row[0])
            self.road_names['Название'].append(row[1])
            # print(row[1])
        print(self.road_names['Название'])
        return self.road_names['Название']

    def get_tp_datas (self, road_name):
        """
        последовательность объектов:
        {название дороги: название дороги,
        участок_дороги:{
            объект:{
            свойство объекта: [(значение, длина, площадь, id параметра, количество точек, минимальное отклонение,
            максимальное отклонение, начало_метры, конец_метры, (начало_километры,начало_метры),
            (конец_километры, конец_метры)), ....], ...}, ...}, ...}
        :return:
        """

        res_km = {}
        # Road.ID_Road,
        request_km = """ select  Road.Name,  High.Description, Group_Description.Item_Name, 
                     Types_Description.Param_Name, Params.ValueParam, dbo.CalcLength(ATTRIBUTE.Image_Points) as Length, 
                     dbo.CalcSquare(ATTRIBUTE.Image_Points) as Square, Params.ID_Param, dbo.GetCountPoint(Image_Points),
                     dbo.Amin(Image_Points) as Amin, dbo.Amax(Image_Points) as Amax,
                    Attribute.L1, Attribute.L2
                    from Road inner join Way on Road.ID_Road = Way.ID_Road
                    inner join High on Way.ID_Way = High.ID_Way
                    inner join Attribute on High.ID_High = Attribute.ID_High
                    inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
                    inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
                    inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
                    where Road.Name = ? and Group_Description.Item_Name = 'Километровые знаки' """
        self.cursor.execute(request_km, road_name)
        for param in self.cursor.fetchall():
            # последовательность объектов:[(значение, длина, площадь, id параметра, количество точек, минимальное отклонение,
            # максимальное отклонение, начало_метры, конец_метры, (начало_километры,начало_метры), (конец_километры, конец_метры)), ....]
            # print(param)
            if param[1] in res_km:

                if param[2] in res_km.get(param[1]):
                    if param[3] in res_km.get(param[1]).get(param[2]):
                        res_km.get(param[1]).get(param[2]).get(param[3]).append(param[4::])
                    else:
                        res_km.get(param[1]).get(param[2]).update({param[3]: [param[4::]]})
                else:
                    res_km.get(param[1]).update({param[2]: {param[3]: [param[4::]]}})
            else:
                res_km.update({param[1]: {param[2]: {param[3]: [param[4::]]}}})

        self.sort_dict_binding(res_km)

        res = {'название дороги': f'{road_name}', }

        request = """
                    select  Road.Name,  High.Description, Group_Description.Item_Name, 
                    Types_Description.Param_Name, Params.ValueParam, dbo.CalcLength(ATTRIBUTE.Image_Points) as Length, 
                    dbo.CalcSquare(ATTRIBUTE.Image_Points) as Square, Params.ID_Param, dbo.GetCountPoint(Image_Points),
                    dbo.Amin(Image_Points) as Amin, dbo.Amax(Image_Points) as Amax,
                    Attribute.L1, Attribute.L2
                    from Road inner join Way on Road.ID_Road = Way.ID_Road
                    inner join High on Way.ID_Way = High.ID_Way
                    inner join Attribute on High.ID_High = Attribute.ID_High
                    inner join Params on Attribute.ID_Attribute = Params.ID_Attribute
                    inner join Types_Description on Params.ID_Param = Types_Description.ID_Param
                    inner join Group_Description on Types_Description.ID_Type_Attr = Group_Description.ID_Type_Attr
                    where Road.Name = ? 

                """  # and Group_Description.Item_Name = ?

        self.cursor.execute(request, road_name)  # item

        for param in self.cursor.fetchall():

            tuple_km = tuple(
                res_km.get(param[1], {}).get('Километровые знаки', {}).get('Значение в прямом направлении', []))

            if tuple_km:
                km = self.convert_m_to_km(param, tuple_km)
            else:
                km = ((0, param[-2]), (0, param[-1]))
            coordinates = (*param[4::], *km)

            # print(param, coordinates)

            if param[1] in res:

                if param[2] in res.get(param[1]):
                    if param[3] in res.get(param[1]).get(param[2]):

                        res.get(param[1]).get(param[2]).get(param[3]).append(coordinates)
                    else:
                        res.get(param[1]).get(param[2]).update({param[3]: [coordinates]})
                else:
                    res.get(param[1]).update({param[2]: {param[3]: [coordinates]}})
            else:
                res.update({param[1]: {param[2]: {param[3]: [coordinates]}}})

        # сортирует в словаре координаты по возрастанию
        self.sort_dict_binding(res)
        res = dict(sorted(res.items()))
        return res

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

    def convert_m_to_km (self, param, list_km):
        '''
        переводит метры в километры с привязкой к километровым знакам
        :param param: объект из базы данных
        :param list_km: список километровых знаков
        :param distance_km: список дистанций между километровыми знаками
        :return: привязки начала и конца объекта
        '''
        # частный случай километровые знаки
        if param[2] == 'Километровые знаки':
            return (int(param[-7]), 0), (int(param[-7]), 0)
        # если точечный объект
        if param[-2] == param[-1]:
            # print('start==end')

            for idx_km, num_sign in enumerate(list_km):
                # находим следующий километровый
                if num_sign == list_km[-1]:
                    next_km = list_km[-1]
                elif num_sign == list_km[0]:
                    next_km = list_km[1]
                else:
                    next_km = list_km[idx_km % len(list_km) + 1]

                if num_sign[-2] <= param[-2] < next_km[-2]:
                    return (int(num_sign[0]), param[-2] - num_sign[-2]), (int(num_sign[0]), param[-1] - num_sign[-1])
                elif param[-2] >= list_km[-1][-2]:
                    return (int(list_km[-1][0]), param[-2] - list_km[-1][-2]), (
                    int(list_km[-1][0]), param[-1] - list_km[-1][-1])
                # elif tmp[-2] <= param[-2] < num_sign[-2]:
                #     return (tmp[0], param[-2] - tmp[-2]), (tmp[0], param[-1] - tmp[-1])
                elif param[-2] < num_sign[-2]:
                    return (int(num_sign[0]), param[-2] - num_sign[-2]), (int(num_sign[0]), param[-1] - num_sign[-1])
                # elif next_km[-2] <= param[-2]:
                #     continue
                #     # return (next_km[0], param[-2] - next_km[-2]), (next_km[0], param[-1] - next_km[-1])
                else:
                    continue
        # если объект линейный или площадной
        elif param[-2] != param[-1]:
            start_km = 0
            end_km = 0
            start_m = 0  # начало
            end_m = 0  # конец
            idx = 0  # индекс километрового start
            # ищем start
            for idx_km, num_sign in enumerate(list_km):
                if num_sign == list_km[-1]:
                    next_km = list_km[-1]
                    idx = idx_km
                elif num_sign == list_km[0]:
                    next_km = list_km[1]
                    idx = idx_km
                else:
                    next_km = list_km[idx_km % len(list_km) + 1]
                    idx = idx_km
                if num_sign[-2] <= param[-2] < next_km[-2] or param[-2] < num_sign[-2]:
                    start_km = num_sign[0]
                    start_m = param[-2] - num_sign[-2]
                    break
                elif param[-2] > list_km[-1][-2]:
                    start_km = list_km[-1][0]
                    start_m = param[-2] - list_km[-1][-2]
                    break
                else:
                    continue
            # ищем end начиная с idx благодаря этому отсекаются уже пройденные километровые
            for idx_km, num_sign in enumerate(list_km[idx:]):
                if num_sign == list_km[-1]:
                    next_km = list_km[-1]
                elif num_sign == list_km[0]:
                    next_km = list_km[1]
                else:
                    next_km = list_km[idx % len(list_km) + 1]
                if num_sign[-1] <= param[-1] < next_km[-1] or param[-1] < num_sign[-1]:
                    end_km = num_sign[0]
                    end_m = param[-1] - num_sign[-1]
                    break
                elif param[-1] > list_km[-1][-1]:
                    end_km = list_km[-1][0]
                    end_m = param[-1] - list_km[-1][-1]
                    break

                else:
                    continue
            return (int(start_km), start_m), (int(end_km), end_m)

    def sort_dict_binding (self, res):
        '''
        сортирует словарь с объектами по первой метровой привязке
        :param res:
        :return: res
        '''

        for _, value in res.items():
            if type(value) == dict:
                for i, val in value.items():
                    if type(val) == dict:
                        for elem in val.values():
                            elem.sort(key = lambda x: (x[-2], x[-1]))

    def close_db(self):
        self.db.close()


def databases():
    request = "select name from sys.databases"
    db = pyodbc.connect(f'Driver={settings.driver}; Server=SIBREGION-SRV2; Trusted_Connection=yes;')
    print()
    db_list = [i[0] for i in db.cursor().execute(request).fetchall()[4:]]
    # print(db_list)

    db.close()
    return db_list

def main ():
    db = Query('OMSK_CITY_2023')  # FKU_VOLGO_VYATSK_1
    list_roads = databases()
    # print(list_roads)
    # data = db.get_dad_datas('P-254')  # Р-176 "Вятка" Чебоксары - Йошкар-Ола - Киров - Сыктывкар
    data_test = db.get_tp_datas('ул. Интернациональная')
    # print(data_test)


    # test(data)
    # print(data)
    db.close_db()
    # with open(rf'{data.get("название дороги","отчет")}.txt', 'w', encoding = 'utf-8') as file:
    #     for key, val in data.items():
    #         file.write(f'{key}:{val}\n')


if __name__ == '__main__':
    main()
