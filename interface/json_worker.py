import os
import json

from settings import path_file_name_info


class JsonWorker:
    """
    Класс для работы с json файлом
    """
    @staticmethod
    def write_json_file_info(data: dict):
        """
        Функция записи json в файл
        :param data: словарь для записи в файл info.json
        :return:
        """
        with open(path_file_name_info, "w", encoding="utf-8", ) as write_file:
            json.dump(data, write_file, indent=4, ensure_ascii=False)

    @staticmethod
    def read_json_file_info():
        """
        Функция чтения json из файла
        :return: data
        """
        if os.path.exists(path_file_name_info):
            with open(path_file_name_info, 'r', encoding='utf-8') as file:
                str_dict = file.read()
                if not str_dict:
                    return {}
                data = json.loads(str_dict, strict=False)
                return data
        else:
            return {}

