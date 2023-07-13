# -*- coding: utf-8 -*-
import pyodbc


class Query:
    def __init__ (self, database = ''):
        self.list_database = None
        __SERVER = "SIBREGION-SRV2"
        __USER = "sibregion"
        __PASSWORD = "SibU$r2018"
        self.database = database
        # self.db = pyodbc.connect(DRIVER={'ODBC Driver 18 for SQL Server'}, server=__SERVER, database=self.database,)
        self.db = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                                 'Server=' + __SERVER + ';'
                                                        'Database=' + self.database + ';'
                                                                                      'Trusted_Connection=yes;')
        self.db.setencoding('cp1251')
        self.cursor = self.db.cursor()  # создаем курсор

    def get_iri (self):
        request = """
            select High.Description, Attribute.L1, Attribute.L2, Group_Description.Item_Name, Params.ValueParam
            from High JOIN Attribute ON High.ID_High = Attribute.ID_High
            JOIN Group_Description On Group_Description.ID_Type_Attr = Attribute.ID_Type_Attr
            JOIN Params ON Params.ID_Attribute = Attribute.ID_Attribute WHERE High.ID_High = ? AND Item_Name = ?
        """
        self.cursor.execute(request, ('1', 'Ровность по IRI'))
        for row in self.cursor.fetchall():
            print(row, '\t')

    def get_rutting (self):
        return

    def get_description (self):
        return

    def get_road_name (self) -> dict:
        '''
            Вытаскивает из базы столбцы с названиями а/д
        :return: self.road_names -> dict
        '''
        self.cursor.execute("SELECT ID_Way, Description FROM High")
        self.road_names = {'№': [], 'Направление': [], 'Дорога': []}
        for row in self.cursor.fetchall():
            # self.road_names['№'].append(row[0])
            self.road_names['Направление'].append(row[0])
            self.road_names['Дорога'].append(row[1])
        return self.road_names

    def get_length_axis (self):
        """
            Вытаскивает кортеж данных из базы
            Название а/д дороги, Протяженность, Название объекта - Ось дороги
        :return:
        """
        request = """
            SELECT High.Description, Attribute.L1, Attribute.L2, Group_Description.Item_Name FROM High JOIN  
            Attribute ON High.ID_High = Attribute.ID_High JOIN Group_Description ON Attribute.ID_Type_Attr = 
            Group_Description.ID_Type_Attr WHERE High.ID_High=? AND Item_Name=?;
            """
        # request = "SELECT * FROM High JOIN Group_Description on High.ID_High = Attribute.ID_High"
        self.cursor.execute(request, ("1", "Ось дороги"))

        for col in self.cursor.fetchall():
            print(col)

    def get_list_database (self) -> list:
        request = """
            SELECT name  FROM sys.databases;  
            """
        self.cursor.execute(request)
        self.list_database = [i[0] for i in self.cursor.fetchall()[4:]]
        return self.list_database

    def get_list_roads (self) -> list:

        self.cursor.execute("SELECT  Description FROM High")
        self.list_roads = [i[0] for i in self.cursor.fetchall()]
        print(self.list_roads)
        return self.list_roads

    def close_db (self):
        self.db.close()


def main ():
    conn = Query()


if __name__ == '__main__':
    main()
