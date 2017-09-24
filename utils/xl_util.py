import xlrd as xl_rd


def open_excel_file(xl_file_path: str):
    xl_file = xl_rd.open_workbook(xl_file_path)

    xl_sheet = xl_file.sheet_by_index(0)

    values = []

    for row in range(xl_sheet.nrows):
        col_val = []
        for col in range(xl_sheet.ncols):
            value = xl_sheet.cell(row, col).value
            try:
                value = str(int(value))
            except ReferenceError:
                print("something wrong with input data")

            col_val.append(value)
        values.append(col_val)

    print(values)

    return xl_file
