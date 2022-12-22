from report_data import NAME_TABLE_COL, REPORT1_SIGNS_DICT, FORMS_TITLE
from cell_formats import *


def set_worksheet_main(self):
    """"Settings for general report"""
    # page orientation as landscape
    self.set_landscape()
    # fit to printed area
    self.fit_to_pages(1, 1)
    # paper type A4
    self.set_paper(9)
    # center the printed page horizontally
    self.center_horizontally()
    # margins    set_margins([left=0.7,] right=0.7,] top=0.75,] bottom=0.75]]])
    self.set_margins(0.7, 0.7, 0.73, 0.92)
    # width of cells
    self.set_column("A:A", 25)
    self.set_column("B:B", 35)
    self.set_column("C:E", 37)
    self.set_column("F:F", 90)
    # height of cells
    self.set_row(0, 43.5)
    self.set_row(1, 43.5)
    self.set_row(2, 43.5)
    self.set_row(3, 43.5)
    for s_row in range(4, 50):
        self.set_row(s_row, 41)


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


def calc_report_table_forms_sums(data):
    """Calculate sums for all same years classes"""
    # write data for all 1 classes
    forms = {FORMS_TITLE[0]: [0, 0, 0, 0, ' '],
             FORMS_TITLE[1]: [0, 0, 0, 0, ' '],
             FORMS_TITLE[2]: [0, 0, 0, 0, ' '],
             FORMS_TITLE[3]: [0, 0, 0, 0, ' ']
             }

    for form in data:
        if '1' in form:
            for i in range(0, 4):
                forms[FORMS_TITLE[0]][i] += data[form][i]
        elif '2' in form:
            for i in range(0, 4):
                forms[FORMS_TITLE[1]][i] += data[form][i]
        elif '3' in form:
            for i in range(0, 4):
                forms[FORMS_TITLE[2]][i] += data[form][i]
        elif '4' in form:
            for i in range(0, 4):
                forms[FORMS_TITLE[3]][i] += data[form][i]
    return forms


def draw_report_table_forms(book, sheet, data, form, row, col=0):
    """Fill cells for each class with data"""
    sheet.write(row, col, form, text_box_center_wrap_format(book))
    sheet.write(row, col + 1, data[0], text_box_center_blue_format(book))
    sheet.write(row, col + 2, data[1], text_box_center_blue_format(book))
    sheet.write(row, col + 3, data[2], text_box_center_wrap_format(book))
    sheet.write(row, col + 4, data[3], text_box_center_wrap_num_format(book))
    sheet.write(row, col + 5, data[4], text_box_left_wrap_format(book))
    row += 1
    return row


def draw_report_table_sums(book, sheet, data, form, row, col=0):
    """Fill cells for each same year classes with sum data"""
    n = 0
    sheet.write(row, col, form, text_box_center_bold_format(book))
    col += 1
    for numbers in data[form]:
        if n == 4:
            sheet.write(row, col, numbers, text_box_center_bold_num_format(book))
        else:
            sheet.write(row, col, numbers, text_box_center_bold_format(book))
        n += 1
        col += 1
    row += 1
    return row


def draw_report_signs(book, sheet, row, col=0):
    """Creating the footer with signatures."""
    row += 1
    # make signs list
    for k, v in REPORT1_SIGNS_DICT.items():
        row += 1
        sheet.merge_range(row, col+1, row, col+2, k, default_format(book))
        sheet.write(row, col+4, v, default_format(book))
