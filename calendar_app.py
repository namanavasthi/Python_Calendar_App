import argparse
from utils import xl_util


def main(excel_file: str):
    print("this is the excel file : " + excel_file)
    file_return_value = xl_util.open_excel_file(xl_file_path=excel_file)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-xl", "--excel", help="Excel File")
    args = parser.parse_args()
    excel_file_name = args.excel
    main(excel_file=excel_file_name)
