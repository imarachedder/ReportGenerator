import db
from interface import connection


def main():
    main_window = connection.main()
    main_window.exec()



if __name__ == '__main__':
    main()

