# This is a sample Python script.
import db
from interface import connection


def main ():
    main_window = connection.main()
    con_db = db.Query(main_window.window.get_name_database())
    # print(noname_con_db.database)
    con_db.database = main_window.window.get_name_database()
    con_db.close_db()


if __name__ == '__main__':
    main()
