import os

# root path
ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
PATH_STATIC = os.path.join(ROOT_PATH, 'static')
PATH_TEMPLATE = os.path.join(ROOT_PATH, 'templates')
# static path's

path_icon_app = os.path.join(ROOT_PATH, 'static','icon.png') #rf'{ROOT_PATH}\static\icon.png'
path_file_name_info = r'info.json'
path_icon_done = os.path.join(PATH_STATIC,'check_green.png') #rf'{ROOT_PATH}\static\check_green.png'
path_icon_not_done = os.path.join(PATH_STATIC,'check_red.png')
path_logo = os.path.join(PATH_STATIC, 'logo.png')
# template path's

# path_template_excel = rf'{ROOT_PATH}\templates\tp_curr.xlsx'
path_template_excel = os.path.join(PATH_TEMPLATE, 'Паспорт_Рамки2(Рамки).xlsx')
path_template_excel_application = os.path.join(PATH_TEMPLATE,'Приложение_Город_Рамки.xlsx')
path_templates_jinja = os.path.join(PATH_TEMPLATE, 'templates_curr\*[0-9].html')
path_file_html = os.path.join(PATH_TEMPLATE, 'tp_curr1.htm')

# DB
server = 'SIBREGION-SRV2'
database = 'Testovaya'
user = 'sibregion'
password = 'SibU$r2018'
driver = '{SQL Server}'
