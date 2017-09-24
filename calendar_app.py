import argparse
from utils import xl_util


def main(excel_file: str, name_to_find: str, address_to_send: str, path_save_at: str):
    print("this is the excel file : " + excel_file)
    file_return_value = xl_util.open_excel_file(xl_file_path=excel_file, input_name=name_to_find,
                                                mail_address=address_to_send, save_path=path_save_at)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-xl", "--excel", help="Excel File")
    parser.add_argument("-n", "--name", help="Name to Search")
    parser.add_argument("-m", "--mail", help="Mail ID to Bind Calendar Events to")
    parser.add_argument("-p", "--path", help="Where to Save Cal Event")
    args = parser.parse_args()
    excel_file_name = args.excel
    name_to_search = args.name
    mail_id = args.mail
    save_path_input = args.path
    main(excel_file=excel_file_name, name_to_find=name_to_search, address_to_send=mail_id, path_save_at=save_path_input)
