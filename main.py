"""
Основной скрипт работы скрипта.
Выполнение заданий
"""

import pandas
import pandas as pd
import win32com.client
import os

from get_data_from_excel import get_data_from_excel
from color_row_in_file import color_row_in_file


def check_report_fond(*excel_files: set) -> pandas.DataFrame:
    """
    Проверка уникальных строк среди файлов
    (Первый переданный dataframe будет основным и будет сравниваться с остальными данными)
    :param excel_files: перечень файлов в dataframe
    :return: df с уникальными строками
    """
    if not excel_files or len(excel_files) < 2:
        raise IOError("Не переданы аргументы")

    excel_files = list(excel_files)
    not_inner_row_df = excel_files[0].copy()

    for excel_file in excel_files[1:]:
        # Левый join между dataframe
        not_inner_row_df = not_inner_row_df.merge(excel_file, how='left', indicator=True)

        # Отсечение сторк, которые находятся между данными
        not_inner_row_df = not_inner_row_df[not_inner_row_df['_merge'] == 'left_only'].drop(columns='_merge')
        not_inner_row_df = not_inner_row_df.drop_duplicates()
    return not_inner_row_df


def task_one(dict_file: dict):
    """
    Выполнение первого задания
    в файле ОТЧЕТ - подсветить красным строку, если строка не найдена хотя бы в одном из 4х файлов

    :param dict_file: список с данными файлов
    """

    print('Первое задание: ', end='')
    report_row_df = check_report_fond(dict_file['report']['df'],
                                      dict_file['trds']['df'],
                                      dict_file['fond']['df'],
                                      dict_file['shakhmatka_11']['df'],
                                      dict_file['shakhmatka_12']['df'])

    color_row_in_file(report_row_df, dict_file['report']['path'], 255)
    print(u'\u2713')


def task_two(dict_file: dict):
    """
    Выполнения второго задания
    в каждом из 4х фалов - подсветить красным строку, если строка не найдена в файле "ОТЧЕТ"

    :param dict_file: список с данными файлов
    """

    print('Второе задание: ', end='')
    for key in dict_file.keys():
        check_df = check_report_fond(dict_file[key]['df'], dict_file['report']['df'])
        color_row_in_file(check_df, dict_file[key]['path'], 255)
    print(u'\u2713', end='\n\n')


if __name__ == '__main__':
    # запись в переменные пути файлов
    path_report = f"{os.getcwd()}\\data\\ОТЧЕТ 01.2023.xlsx"
    path_trds = f"{os.getcwd()}\\data\\TRDS 12.2022.xlsx"
    path_fond = f"{os.getcwd()}\\data\\Fond_ESP 12.2022.xlsx"
    path_shakhmatka_11 = f"{os.getcwd()}\\data\\ШАХМАТКА 11.2022.xls"
    path_shakhmatka_12 = f"{os.getcwd()}\\data\\ШАХМАТКА 12.2022.xls"

    # импорт данных из файлов
    print('Импорт файла Отчет: ', end='')
    report_df = get_data_from_excel(path_report)
    print(u'\u2713')

    print('Импорт файла trds: ', end='')
    trds_df = get_data_from_excel(path_trds)
    print(u'\u2713')

    print('Импорт файла fond: ', end='')
    fond_df = get_data_from_excel(path_fond)
    print(u'\u2713')

    print('Импорт файла shakhmatka_11: ', end='')
    shakhmatka_11 = get_data_from_excel(path_shakhmatka_11)
    print( u'\u2713')

    print('Импорт файла shakhmatka_12: ', end='')
    shakhmatka_12 = get_data_from_excel(path_shakhmatka_12)
    print(u'\u2713', end='\n\n')

    # запись в словарь данных о файлах
    dict_file = {
        'report': {'df': report_df, 'path': path_report},
        'trds': {'df': trds_df, 'path': path_trds},
        'fond': {'df': fond_df, 'path': path_fond},
        'shakhmatka_11': {'df': shakhmatka_11, 'path': path_shakhmatka_11},
        'shakhmatka_12': {'df': shakhmatka_12, 'path': path_shakhmatka_12},
    }

    task_one(dict_file)
    task_two(dict_file)
    print('Конец работы программы')
