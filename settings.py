import os
from pathlib import Path
# root path
ROOT_PATH = Path(__file__).parent

PATH_STATIC = Path.joinpath(ROOT_PATH, 'static')
PATH_TEMPLATE = Path.joinpath(ROOT_PATH, 'templates')
# static path's

path_icon_app = Path.joinpath(PATH_STATIC, 'icon.png') #rf'{ROOT_PATH}\static\icon.png'
path_file_name_info = r'info.json'
path_icon_done = Path.joinpath(PATH_STATIC,'check_green.png') #rf'{ROOT_PATH}\static\check_green.png'
path_icon_not_done = Path.joinpath(PATH_STATIC,'check_red.png')
path_logo = Path.joinpath(PATH_STATIC, 'logo.png')
# template path's

path_template_excel_dad = Path.joinpath(PATH_TEMPLATE, 'Диагностика.xlsm')
path_template_excel = Path.joinpath(PATH_TEMPLATE, 'Паспорт_Рамки2(Рамки).xlsx')
path_template_excel_application = Path.joinpath(PATH_TEMPLATE,'Приложение_Город_Рамки.xlsx')
path_templates_jinja = Path.joinpath(PATH_TEMPLATE, 'templates_curr\*[0-9].html')
path_file_html = Path.joinpath(PATH_TEMPLATE, 'tp_curr1.htm')

# DB
server = 'SIBREGION-SRV2'
database = 'Testovaya'
user = 'sibregion'
password = 'SibU$r2018'
driver = '{SQL Server}'
