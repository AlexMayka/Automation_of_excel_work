"""
Скрипт для автоматического импорта данных из файлов:
    1. Определение номеров нужных столбцов в файле
    2. Определение номера строки, с которого берется нужная информация
    3. Получение данных из файлов

Алгоритм:
    1. Открытие файла excel (pywin32)
    2. Нахождение номеров колонок: Местоположение, Скважина, Объект Разработки/Пласт, при наличии (pywin32)
    3. Нахожждение номера строки, с которой начинаются основные данные (pywin32)
    4. Закрытие файла excel (pywin32)
    4. Получение данных из файла, начиная с найденной строки и с нужных столбцов (pandas)
"""

import pandas
import pandas as pd
import win32com.client
import os
from common_functions import open_excel_file_win32, close_excel


def search_data_file(path_file=None):
    """
    Поиск столбцов и номеров в выбранном файле
    :param path_file: полный адрес к файлу
    :return: start_row: стартовая строка в файле, с которой начинаются данные
             dict_column_file: номара колонок с нужными данными
    """

    def search_for_titles() -> (int, dict):
        """
        Поиск стартовых толбцов в файле
        :return: row: номер строки, на которой расположены заголовок
                 dict_column_file: словарь с номерами колонок
        """
        try:
            for row in range(1, ws.UsedRange.Rows.Count):
                dict_column_file = {}
                for col in range(1, ws.UsedRange.Columns.Count):
                    if "мест" in str(ws.Cells(row, col).Value).lower() or "м-е" in str(
                            ws.Cells(row, col).Value).lower():
                        if "Месторождение" in dict_column_file:
                            continue
                        dict_column_file['Месторождение'] = col

                    elif "скв" in str(ws.Cells(row, col).Value).lower():
                        if "Скважина" in dict_column_file:
                            continue
                        dict_column_file['Скважина'] = col

                    elif "объек" in str(ws.Cells(row, col).Value).lower() or "пласт" in str(
                            ws.Cells(row, col).Value).lower():
                        if "Объект Разработки" in dict_column_file:
                            continue
                        dict_column_file['Объект Разработки'] = col

                    if len(dict_column_file) == 3:
                        return row, dict_column_file

                if len(dict_column_file) in [2, 3]:
                    return row, dict_column_file

            if len(dict_column_file) not in [2, 3]:
                raise IOError(f"Программа не смогла найти заголоки")

        except Exception as ERROR_search_for_titles:
            close_excel(xl_app, wb, False)
            raise IOError(f"Ошибка при поиске заголовков: {ERROR_search_for_titles}")

    def search_data_begin() -> int:
        """
        Поиск строки, с которых начинаются нужные данные
        :return: row: строка, с которых начинаются нужные данные
        """
        try:
            start_col = dict_column_file['Месторождение']
            row = start_row
            while row == 1 or 'None' in list(map(lambda x: str(x), ws.Range(ws.Cells(row, start_col),
                                                                            ws.Cells(row + 5, start_col)))):
                row += 1

            return row + 1 if str(ws.Cells(row, start_col)).isdigit() else row

        except Exception as ERROR_search_data_begin:
            close_excel(xl_app, wb, False)
            raise IOError(f"Ошибка при поиске данных: {ERROR_search_data_begin}")

    xl_app, wb, ws = open_excel_file_win32(path_file)
    start_row, dict_column_file = search_for_titles()
    start_row = search_data_begin()
    close_excel(xl_app, wb, False)
    return start_row, dict_column_file


def get_data_pandas(path_file: str = None, start_row: int = None, column_head_file: dict = None) -> pandas.DataFrame:
    """
    Получение данных с помощью pandas
    :param path_file: путь к файлу
    :param start_row: стартовая строка
    :param column_head_file: словарь с номерами колонок
    :return: dataframe с данными
    """
    if not path_file or not start_row or not column_head_file:
        raise IOError("Не переданы параметры")

    try:
        # импорт данным с пропуском строк
        input_data_df = pd.read_excel(path_file, skiprows=start_row - 1, header=None)
        column_head_file = {key: value - 1 for key, value in column_head_file.items()}

        # фильтрация по столбцам и отсеченик не нужных столбцов
        if len(column_head_file) == 2:
            input_data_df = input_data_df[
                (input_data_df[column_head_file['Месторождение']].isna() == False)
                & (input_data_df[column_head_file['Скважина']].isna() == False)][column_head_file.values()]
        else:
            input_data_df = input_data_df[
                (input_data_df[column_head_file['Месторождение']].isna() == False)
                & (input_data_df[column_head_file['Скважина']].isna() == False)
                & (input_data_df[column_head_file['Объект Разработки']].isna() == False)][column_head_file.values()]

        # переименование заголовков
        column_head_file = {value: key for key, value in column_head_file.items()}
        input_data_df = input_data_df.rename(columns=column_head_file)

        return input_data_df

    except Exception as ERROR_get_data_pandas:
        raise IOError(f"Ошибка при чтении файла pandas: {ERROR_get_data_pandas}")


def get_data_from_excel(path_file_excel: pandas.DataFrame = None) -> object:
    """
    Главная функция работы скрипта
    :param path_file_excel: путь к файлу
    :return: dataframe с данными
    """
    if not path_file_excel:
        raise IOError("Не корректный путь к файлу")

    start_row, column_head_file = search_data_file(path_file_excel)
    data_df = get_data_pandas(path_file_excel, start_row, column_head_file)
    return data_df


if __name__ == "__main__":
    path_file_excel = f"{os.getcwd()}\\data\\ШАХМАТКА 11.2022.xls"
    data_df = get_data_from_excel(path_file_excel)
