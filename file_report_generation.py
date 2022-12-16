import copy
from report_data import *

from cell_formats import title_format, date_format, default_format, \
                        text_box_wrap_format, text_box_center_wrap_format, \
                        text_box_center_wrap_num_format, text_box_center_blue_format, \
                        text_box_center_bold_format, text_box_center_bold_num_format


# write down sums for every class to the cell
def write_classes_sums(self, in_row, in_col, sheet, book):
    for idx, (form, numbers) in enumerate(self.items()):
        sheet.write(in_row, in_col, form, text_box_center_wrap_format(book))
        tmp_col = in_col + 1
        n = 0
        for number in numbers:
            if n < 2:
                sheet.write(in_row, tmp_col, number, text_box_center_blue_format(book))
            elif n == 3:
                sheet.write(in_row, tmp_col, number, text_box_center_wrap_num_format(book))
            else:
                sheet.write(in_row, tmp_col, number, text_box_center_wrap_format(book))
            tmp_col += 1
            n += 1
        in_row += 1
    return in_row, in_col


# write sums for same classes
# def write_all_sums(self, in_row, in_col, sheet, book):
#     n = 0
#     for numbers in self:
#         if n == 4:
#             sheet.write(in_row, in_col, numbers, text_box_center_bold_num_format(book))
#         else:
#             sheet.write(in_row, in_col, numbers, text_box_center_bold_format(book))
#         n += 1
#         in_col += 1
#     sheet.write(in_row, in_col, " ", text_box_center_bold_format(book))
#     return in_row, in_col


# def calc_sums(self, row_name: str):
#     pupils_sum = 0
#     pupils_eat = 0
#     pupils_days = 0
#     for form in self:
#         pupils_sum += int(self[form][0])
#         pupils_eat += int(self[form][1])
#         pupils_days += int(self[form][2])
#     # print(pupils_sum, pupils_eat, pupils_days)
#     sum_classes = [row_name, pupils_sum, pupils_eat, pupils_days, pupils_days * PRICE]
#     # print(sum_classes)
#     return sum_classes




def draw_report_title(book, sheet, data, row=0, col=0):
    """Creating the header of the template."""
    # write 1 row "Звіт"
    sheet.merge_range(row, col, row, col+5, data[0], title_format(book))
    row += 1

    # write 2 row
    sheet.merge_range(row, col, row, col+5, data[1], title_format(book))
    row += 1

    # write 3 row
    month_year = data[2]
    sheet.merge_range(row, col, row, col+6,
                      month_year, date_format(book))
    row += 2

    # make table
    sheet.merge_range(row, col, row + 1, col, NAME_TABLE_COL[0],
                      text_box_center_wrap_format(book))
    sheet.merge_range(row, col + 1, row + 1, col + 1, NAME_TABLE_COL[1],
                      text_box_center_wrap_format(book))
    sheet.merge_range(row, col + 2, row + 1, col + 2, NAME_TABLE_COL[2],
                      text_box_center_wrap_format(book))
    sheet.merge_range(row, col + 3, row + 1, col + 3, NAME_TABLE_COL[3],
                      text_box_center_wrap_format(book))
    sheet.merge_range(row, col + 4, row + 1, col + 4, NAME_TABLE_COL[4],
                      text_box_center_wrap_format(book))
    sheet.merge_range(row, col + 5, row + 1, col + 5, NAME_TABLE_COL[5],
                      text_box_center_wrap_format(book))

    row += 2
    col = 0
    return row


"""Здесь был расчет сумм"""
# sum_1_classes = calc_sums(data_1_classes, all_1_classes)
# print(sum_1_classes)
# sum_2_classes = calc_sums(data_2_classes, all_2_classes)
# print(sum_2_classes)
# sum_3_classes = calc_sums(data_3_classes, all_3_classes)
# print(sum_3_classes)
# sum_4_classes = calc_sums(data_4_classes, all_4_classes)
# print(sum_4_classes)
#
# all_sums = copy.copy(sum_1_classes)
# all_sums[0] = result_name
#
# for i in range(1, 5):
#     all_sums[i] = all_sums[i] + sum_2_classes[i] + sum_3_classes[i] + sum_4_classes[i]
#
# print(all_sums)

# # write data for all 1 classes
#   write_classes_sums(data_1_classes, row, col, worksheet, workbook)
#   row += len(data_1_classes)
#   # write sums for 1 classes
#   write_all_sums(sum_1_classes, row, col, worksheet, workbook)
#   row += 1
#
#   # write data for all 2 classes
#   write_classes_sums(data_2_classes, row, col, worksheet, workbook)
#   row += len(data_2_classes)
#   # write sums for 2 classes
#   write_all_sums(sum_2_classes, row, col, worksheet, workbook)
#   row += 1
#
#   # write data for all 3 classes
#   write_classes_sums(data_3_classes, row, col, worksheet, workbook)
#   row += len(data_3_classes)
#   # write sums for 3 classes
#   write_all_sums(sum_3_classes, row, col, worksheet, workbook)
#   row += 1
#
#   # write data for all 4 classes
#   write_classes_sums(data_4_classes, row, col, worksheet, workbook)
#   row += len(data_4_classes)
#   # write sums for 4 classes
#   write_all_sums(sum_4_classes, row, col, worksheet, workbook)
#
#   for i in range(6):
#       sheet.write(row, col+i, ' ', text_box_center_wrap_format(book))
#
#   # write resulting sums
#   row += 1
#   write_all_sums(all_sums, row, col)


def draw_report_signs(book, sheet, row, col=0):
    """Creating the footer with signatures."""
    row += 1
    # make signs list
    for k, v in REPORT1_SIGNS_DICT.items():
        row += 1
        sheet.merge_range(row, col+1, row, col+2, k, default_format(book))
        sheet.write(row, col+5, v, default_format(book))
