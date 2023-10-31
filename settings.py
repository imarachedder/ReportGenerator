import os

# root path
root_path = os.path.dirname(os.path.abspath(__file__))

# static path's
path_icon_app = rf'{root_path}\static\icon.png'
path_file_name_info = r'info.json'
path_icon_done = rf'{root_path}\static\check_green.png'
path_icon_not_done = rf'{root_path}\static\check_red.png'

# template path's
path_template_excel = rf'{root_path}\templates\tp_curr.xlsx'
path_templates_jinja = rf'{root_path}\templates\templates_curr\*[0-9].html'
path_template_excel_application = rf'{root_path}\templates\tp_vedomost.xlsx'
path_file_html = rf'{root_path}\templates\tp_curr1.htm'

# DB
server = 'SIBREGION-SRV2'
database = 'Testovaya'
user = 'sibregion'
password = 'SibU$r2018'
driver = '{SQL Server Native Client 11.0}'
