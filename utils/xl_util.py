import xlrd as xl_rd

from icalendar import Calendar, Event

import datetime
import pytz

import os


def open_excel_file(xl_file_path: str, input_name: str, mail_address: str, save_path: str):
    days_list = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    try:
        xl_file = xl_rd.open_workbook(xl_file_path)

        xl_sheet = xl_file.sheet_by_index(0)

        for row in range(xl_sheet.nrows):

            # day parameter
            col = xl_sheet.cell(row, 0).value

            if col in days_list:

                # date parameter in mm/dd/yy format in excel file
                shift_day = xl_sheet.cell(row, 1).value
                # convert to python compatible format
                py_date = xl_rd.xldate.xldate_as_datetime(shift_day, 0)

                for rows_in_day in range(row + 1, xl_sheet.nrows):

                    # is the time value or day of the week value
                    day_value = xl_sheet.cell(rows_in_day, 0).value

                    if day_value in days_list:
                        break
                    else:
                        #gives the shift position
                        location_field = xl_sheet.cell(rows_in_day, 1).value

                        # this for loop is to check names in last 3 columns
                        for cols in range(3, 6):

                            # get those column values and check with given input name
                            name_field = xl_sheet.cell(rows_in_day, cols).value

                            if name_field == input_name:

                                # manipulation of time value, i.e. HH:MMAM/PM - HH:MMAM/PM
                                day_value = day_value.split(" - ")
                                date_start = day_value[0]
                                date_end = day_value[1]

                                hours_start = date_start.split(":")[0]
                                hours_check = (date_start.split(":")[1])[2]
                                if hours_check == 'P':
                                    if hours_start != "12":
                                        hours_start = str(int(hours_start) + 12)
                                    minutes_start = (date_start.split(":")[1]).split("PM")[0]
                                else:
                                    minutes_start = (date_start.split(":")[1]).split("AM")[0]

                                hours_end = date_end.split(":")[0]
                                hours_check = (date_end.split(":")[1])[2]
                                if hours_check == 'P':
                                    if hours_end != "12":
                                        hours_end = str(int(hours_end) + 12)
                                    minutes_end = (date_end.split(":")[1]).split("PM")[0]
                                else:
                                    minutes_end = (date_end.split(":")[1]).split("AM")[0]


                                # CALENDAR EVENT CREATION

                                cal = Calendar()
                                event = Event()

                                la = pytz.timezone("America/Los_Angeles")

                                event.add('summary', 'Lib Sec Shift')

                                event.add('location',location_field)

                                # datetime format = (yyyy,m,day,hrs,mins,secs)

                                event.add('dtstart',
                                          la.localize(
                                              datetime.datetime(py_date.year, py_date.month, py_date.day,
                                                                int(hours_start), int(minutes_start), 0)))


                                event.add('dtend',
                                          la.localize(
                                              datetime.datetime(py_date.year, py_date.month, py_date.day,
                                                                int(hours_end), int(minutes_end), 0)))

                                event.add('dtstamp',
                                          la.localize(datetime.datetime(py_date.year, py_date.month, py_date.day,
                                                                        0, 10, 0)))


                                cal.add('attendee',mail_address)

                                cal.add_component(event)

                                save_file = str(col) + str(day_value) + ".ics"

                                f = open(os.path.join(save_path, save_file), 'wb')
                                f.write(cal.to_ical())
                                f.close()
    except FileNotFoundError:
        return 0

    return 1
