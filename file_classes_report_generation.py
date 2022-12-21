from report_data import CLASS_NAME_TABLE_COL, CLASS_REPORT_SIGNS_DICT, CLASS_TEACHER
from cell_formats import *


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


    # # write 1 row "Звіт"
    # sheet.merge_range(row, col, row, col+4+len(days), data[0], title_format2(book))
    # row += 1
    #
    # # write 2 row
    # sheet.merge_range(row, col, row, col+4+len(days), data[1], title_format2(book))
    # row += 2
    #
    # # make table
    # sheet.merge_range(row, col, row+1, col, CLASS_NAME_TABLE_COL[0], text_box_center_wrap_format2(book))
    # sheet.merge_range(row, col+1, row+1, col+1, CLASS_NAME_TABLE_COL[1], text_box_center_wrap_format2(book))
