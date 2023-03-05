"""
Скрипт для закрашивание ячеек внутри файла excel
Алгоритм:
    1. Открытие файла excel (pywin32)
    2. Перебор данных входящего dataframe и данных файлов, сравнение строк между собой (pywin32)
    3. Совпадающие строки закрашиваются (pywin32)
    4. Закрытие файла (pywin32)
"""

import pandas as pd
import win32com.client
import os
from common_functions import open_excel_file_win32, close_excel


def color_row(xl_app, wb, ws, row_color_df, color):
    """
    Закрашивание ячеек в файле
    :param xl_app: COM объект pywin32
    :param wb: Рабочий файл (класс pywin32)
    :param ws: Рабочий лист (класс pywin32)
    :param row_color_df: dataframe с данными, которые надо закрасить
    :param color: цвет, в который будет закрашена ячейка
    """
    try:
        for row_file in range(1, ws.UsedRange.Rows.Count + 1):
            row_value_list = list(
                map(lambda x: x.Value, ws.Range(ws.Cells(row_file, 1), ws.Cells(row_file, ws.UsedRange.Columns.Count))))

            for index, row_color in row_color_df.iterrows():
                if (row_color['Месторождение'] in row_value_list and row_color['Скважина'] in row_value_list
                        and row_color['Объект Разработки'] in row_value_list):
                    ws.Range(ws.Cells(row_file, 1), ws.Cells(row_file, ws.UsedRange.Columns.Count)).Interior.Color = 255

    except Exception as ERROR_color_ro:
        close_excel(xl_app, wb, False)
        raise IOError(f"Ошибка при закрашивании файлов: {ERROR_color_ro}")


def color_row_in_file(row_color_df, path_file, color):
    """
    Главная функция для закрашивания файлов
    :param row_color_df: dataframe с данными, которые надо закрасить
    :param path_file: путь к файлу
    :param color: цвет, в который будет зарашен файл
    """
    if not path_file or not color:
        raise IOError(f"Неккоректные параметры")

    xl_app, wb, ws = open_excel_file_win32(path_file)
    color_row(xl_app, wb, ws, row_color_df, color)
    close_excel(xl_app, wb, True)
