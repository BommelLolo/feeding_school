from report_data import CLASS_NAME_TABLE_COL, CLASS_REPORT_SIGNS_DICT, CLASS_TEACHER, CLASS_WAS_NOT
from cell_formats import *


def set_worksheet_classes(self, pupil, days):
    """Settings for class report"""
    # fit to printed area
    self.fit_to_pages(1, 1)
    # paper type A4
    self.set_paper(9)
    # center the printed page horizontally
    self.center_horizontally()
    # margins    set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])
    self.set_margins(0.7, 0.7, 0.73, 0.92)
    # width of cells
    self.set_column_pixels("A:A", 28)
    self.set_column_pixels("B:B", 160)
    # width for working days
    if len(days) <= (ord("Z") - ord("C")):
        end_letter = chr(ord("C") + len(days) - 1)
        address = "C:" + end_letter
    else:
        shift = len(days) - (ord("Z") - ord("C"))
        end_letter = chr(ord("A") + shift - 1)
        address = "C:" + "A" + end_letter
    self.set_column_pixels(address, 34)
    # width last cells after working days
    for set_width in range(3):
        if address[0] == "Z" or len(address) == 4:
            address = "AA:AA"
        elif len(address) == 5:
            end_letter = chr(ord(address[4]) + 1)
            address = "A" + end_letter + ":A" + end_letter
        else:
            end_letter = chr(ord(end_letter) + 1)
            address = end_letter + ":" + end_letter
        self.set_column_pixels(address, 87)
        set_width += 1

    # height of cells

    self.set_row_pixels(0, 40)
    self.set_row_pixels(1, 40)
    self.set_row_pixels(2, 20)
    self.set_row_pixels(3, 30)
    self.set_row_pixels(4, 30)
    # Set row height depending on number of pupils
    set_row = 5
    for set_pupil in range(set_row, pupil+set_row+1):
        self.set_row_pixels(set_pupil, 36)
    set_row += pupil
    self.set_row_pixels(set_row, 40)
    self.set_row_pixels(set_row+1, 54)
    self.set_row_pixels(set_row+2, 20)
    self.set_row_pixels(set_row+3, 40)
    self.set_row_pixels(set_row+4, 40)
    self.set_row_pixels(set_row+5, 40)
    self.set_row_pixels(set_row+6, 40)
    return set_row


"""Розрахувати кількість дітей, які харчуються"""


def cell_value_check(value: str, values: tuple, cost: float) -> any:
    """Check the value in the cell"""
    if value == values[0]:
        value = cost
    elif value == values[2]:
        value = CLASS_WAS_NOT
    return value


def pupil_missed_days(value: str, values: tuple) -> int:
    """if pupil missed day, then add 1"""
    res = 0
    if value != values[0]:
        res = 1
    return res


def pupil_child_days(value: str, values: tuple) -> int:
    """if pupil was this day, then add 1"""
    res = 0
    if value == values[0]:
        res = 1
    return res


def draw_class_report_title(book, sheet, data, days, row=0, col=0):
    """Creating the header of the template."""
    # write 1 row "Звіт"
    sheet.merge_range(row, col, row, col+4+len(days), data[0], title_format2(book))
    row += 1

    # write 2 row
    sheet.merge_range(row, col, row, col+4+len(days), data[1], title_format2(book))
    row += 2

    # make table
    sheet.merge_range(row, col, row+1, col, CLASS_NAME_TABLE_COL[0], text_box_center_wrap_format2(book))
    sheet.merge_range(row, col+1, row+1, col+1, CLASS_NAME_TABLE_COL[1], text_box_center_wrap_format2(book))

    # Write days. Divide them on 2 rows by weeks
    col = 0
    j = days[1]
    k = 1
    numbers_format = ['', text_box_center_nums_top_format, text_box_center_nums_bot_format]
    for day in range(len(days)):
        if int(days[day]) - j <= 1:
            k = k
        else:
            k = -k
        sheet.merge_range(row, col + 2, row + 1, col + 2, days[day], numbers_format[k](book))
        col += 1
        j = int(days[day])

    # write names for last columns
    col = 0
    sheet.merge_range(row, col+2+len(days), row+1, col+2+len(days),
                      CLASS_NAME_TABLE_COL[2], text_box_center_wrap_format2(book))
    sheet.merge_range(row, col+3+len(days), row+1, col+3+len(days),
                      CLASS_NAME_TABLE_COL[3], text_box_center_wrap_format2(book))
    sheet.merge_range(row, col+4+len(days), row+1, col+4+len(days),
                      CLASS_NAME_TABLE_COL[4], text_box_center_wrap_format2(book))

    return row


def draw_class_report_child_days(book, sheet, data, day_price, row=0, col=0):
    """Fill bottom cells of table"""
    # write cells "Всього дітоднів" та "Сума"
    sheet.merge_range(row, col, row, col+1, "Всього дітоднів", classes_results_names_format(book))
    sheet.merge_range(row+1, col, row+1, col+1, "Сума", classes_results_names_format(book))
    col = 2
    # write results for each working day in rows "Всього дітоднів" and "Сума"
    for day in data:
        sheet.write(row, col, data[day], classes_results_child_days_format(book))

        sheet.write(row+1, col, data[day] * day_price, classes_results_sum_format(book))
        col += 1


def draw_class_report_signs(book, sheet, data, row, col=2):
    """Creating the footer of the template."""
    row += 1
    # make signs list
    for k, v in CLASS_REPORT_SIGNS_DICT.items():
        sheet.merge_range(row, col, row, col+7, k, classes_signs_format(book))
        sheet.write(row, col+12, v, classes_signs_format(book))
        row += 1
    # make sign of class teacher
    sheet.merge_range(row, col, row, col + 7, CLASS_TEACHER, classes_signs_format(book))
    sheet.write(row, col + 12, data, classes_signs_format(book))