import xlrd as xl_rd


def open_excel_file(xl_file_path: str):
    xl_file = xl_rd.open_workbook(xl_file_path)

    return xl_file
