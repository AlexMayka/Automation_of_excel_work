import win32com.client
import os


def open_excel_file_win32(path_file):
    """
    Открытие файла с помощью pywin32
    :return: xl_app: COM объект pywin32
             work_book: Рабочий файл (класс pywin32)
             work_sheet: Рабочий лист (класс pywin32)
    """
    xl_app = win32com.client.Dispatch("Excel.Application")
    xl_app.Visible = 0
    work_book = xl_app.Workbooks.Open(path_file)
    work_sheet = work_book.Worksheets(1)

    if xl_app == False or work_book == False or work_sheet == False:
        raise IOError(f'Ошибка при открытии файла')

    return xl_app, work_book, work_sheet


def close_excel(xl_app=None, wb=None, savechanges=False):
    """
    Закрытие файла excel (с сохранением результатов)
    :param xl_app: COM объект pywin32
    :param wb: Рабочий файл (класс pywin32)
    :param wb: Сохранение изменений
    :return:
    """

    if xl_app and wb:
        wb.Close(SaveChanges=savechanges)
        xl_app.Quit()
    else:
        raise IOError("Ошибка при закрытии файла excel")


if __name__ == '__main__':
    path_file = f'{os.getcwd()}\\data\\ШАХМАТКА 11.2022.xls'
    xl, wb, ws = open_excel_file_win32(path_file)
    close_excel(xl, wb, True)

