import argparse
from utils import xl_util

from tkinter import *
from tkinter import ttk

__author__ = "NAMAN AVASTHI"

root = Tk()

root.geometry('{}x{}'.format(800, 450))

root.wm_title("iCal APP by Naman Avasthi")

row = 2
col = 2

label1 = Label(root, text="Excel File Location : ", font=("Helvetica", 16))
E1 = Entry(root)

label2 = Label(root, text="Name to Search : ", font=("Helvetica", 16))
E2 = Entry(root)

label3 = Label(root, text="Mail address : ", font=("Helvetica", 16))
E3 = Entry(root)

label4 = ttk.Label(root, text="Where To Save Path : ", font=("Helvetica", 16))
E4 = ttk.Entry(root)


def quit_program():
    root.destroy()
    root.quit()


def destroy_ui_components():
    E1.destroy()
    E2.destroy()
    E3.destroy()
    E4.destroy()
    label1.destroy()
    label2.destroy()
    label3.destroy()
    label4.destroy()
    submit.destroy()


def get_date():
    excel_file = E1.get()
    name_to_find = E2.get()
    address_to_send = E3.get()
    path_save_at = E4.get()

    file_return_value = xl_util.open_excel_file(xl_file_path=excel_file, input_name=name_to_find,
                                                mail_address=address_to_send, save_path=path_save_at)

    if file_return_value == 0:
        destroy_ui_components()
        issue = ttk.Label(root, text="Something went Wrong! Please Check Inputs Again", font=("Helvetica", 16))
        issue.grid(row=2, column=2)
        resolve = ttk.Label(root, text="Contact Developer for possible solutions at navasthi@usc.edu",
                            font=("Helvetica", 16))
        resolve.grid(row=3, column=2)
    else:
        destroy_ui_components()
        success = Label(root, text="SUCCESS! Calendar Events created at " + path_save_at, font=("Helvetica", 16))
        success.grid(row=2, column=2)


submit = Button(root, text="Submit", command=get_date)
exit_button = Button(root, text="Exit", command=quit_program)

label1.grid(row=2, column=2)
E1.grid(row=row, column=col + 1)
label2.grid(row=row + 1, column=col)
E2.grid(row=row + 1, column=col + 1)
label3.grid(row=row + 2, column=col)
E3.grid(row=row + 2, column=col + 1)
label4.grid(row=row + 3, column=col)
E4.grid(row=row + 3, column=col + 1)
submit.grid(row=row + 4, column=col)
exit_button.grid(row=row + 5, column=col)

root.mainloop()


def main(excel_file: str, name_to_find: str, address_to_send: str, path_save_at: str):
    print("this is the excel file : " + excel_file)


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
