import argparse
from utils import xl_util

from tkinter import *

__author__ = "NAMAN AVASTHI"

root = Tk()

root.geometry("600x400")

root.title("Naman Avasthi iCal APP")

label1 = Label(root, text="Excel File Location : ")
E1 = Entry(root, bd=2)

label2 = Label(root, text="Name to Search : ")
E2 = Entry(root, bd=2)

label3 = Label(root, text="Mail address : ")
E3 = Entry(root, bd=2)

label4 = Label(root, text="Where To Save Path : ")
E4 = Entry(root, bd=2)


def quit_program():
    root.destroy()
    root.quit()


def destry_ui_components():
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
        issue = Label(root, text="Something went Wrong! Please Check Inputs Again")
        issue.pack()
        resolve = Label(root, text="Contact Developer for possible solutions at navasthi@usc.edu")
        resolve.pack()
    else:
        destry_ui_components()
        success = Label(root, text="SUCCESS! Calendar Events created at " + path_save_at)
        success.pack(side=TOP)


submit = Button(root, text="Submit", command=get_date)
exit_button = Button(root, text="Exit", command=quit_program)

label1.grid(row=0,column=0,sticky=W)
E1.pack()
label2.pack()
E2.pack()
label3.pack()
E3.pack()
label4.pack()
E4.pack()
submit.pack(side=BOTTOM)
exit_button.pack(side=BOTTOM)
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
