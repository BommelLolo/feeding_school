import pandas as pd
import warnings
import xlsxwriter
from datetime import date
# from report_data import *
from file_report_generation import draw_report_title, draw_report_signs
from file_classes_report_generation import draw_class_report_title, \
    draw_class_report_child_days, draw_class_report_signs
from report_data import CLASS_WAS_NOT, LIST_REPORT_NAME
import cell_formats


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


def cell_value_check(value: str, values: tuple) -> any:
    """Check the value in the cell"""
    if value == values[0]:
        value = price
    elif value == values[2]:
        value = CLASS_WAS_NOT
    return value


def pupil_missed_days(value: str, values: tuple) -> int:
    """if pupil missed day, then add 1"""
    res = 0
    if value == values[1]:
        res = 1
    return res


def pupil_child_days(value: str, values: tuple) -> int:
    """if pupil was this day, then add 1"""
    res = 0
    if value == values[0]:
        res = 1
    return res


if __name__ == "__main__":

    try:
        xls = pd.ExcelFile('Харчування.xlsx')
        # ignore UserWarnings (about DataValidation and Conditional Formatting)
        warnings.simplefilter(action='ignore', category=UserWarning)
    except FileNotFoundError:
        print("File could not be found.")

    # Receive sheet names and quantity
    feeding_sheet_names = tuple(xls.sheet_names)
    feeding_sheet_numbers = len(feeding_sheet_names)
    print('Sheets:', feeding_sheet_names, feeding_sheet_numbers)

    # Receive data from "Service" sheet in xlsx
    sheet_xls_name = "Service"

    # Receive variants what cell contains
    service = xls.parse(sheet_xls_name, header=None)
    cells_const_values = tuple(service.values[1:, 0])
    print('cells_const_values:', cells_const_values)

    # Receive data from "Загальні налаштування" sheet in xlsx
    sheet_xls_name = "Загальні налаштування"

    # Receive general settings
    settings_from_xls = xls.parse(sheet_xls_name)
    year_report = int(settings_from_xls.values[0, 0])
    month_report = settings_from_xls.values[0, 1]
    price = float(settings_from_xls.values[0, 2])
    school = settings_from_xls.values[0, 3]

    # Receive database from class' sheets in xlsx
    classes_list = []
    db_all_classes = {}
    for sheet_xls_name in feeding_sheet_names:
        if sheet_xls_name != "Service" and sheet_xls_name != "Загальні налаштування":
            classes_list.append(sheet_xls_name)
            class_data = xls.parse(sheet_xls_name)
            db_all_classes[sheet_xls_name] = class_data

    # Create a workbook for report generation
    today = date.today().strftime("(%d.%m.%Y)")
    file_report_name = 'Звіт_з_харчування_за_' + str(month_report) + '_' + str(year_report) + str(today)
    file_report_xlsx = file_report_name + ".xlsx"
    workbook = xlsxwriter.Workbook(file_report_xlsx)

    # Create the first worksheet
    worksheet = workbook.add_worksheet(LIST_REPORT_NAME)
    set_worksheet_main(worksheet)

    teachers = {}  # dict for class teachers

    # Create other worksheets and add information from input file
    for form in classes_list:
        # Get column names
        class_columns = tuple(db_all_classes[form].columns)
        # Define working days
        working_days = []
        # Define summary of child days and missed days
        child_days_form, missed_days_form = 0, 0

        for c in class_columns:
            if type(c) is int:
                working_days.append(c)

        # Define number of pupils
        pupil_quantity = (len(db_all_classes[form][1]))
        # Add new sheet for class
        worksheet2 = workbook.add_worksheet(form)
        # Set worksheet for class
        row_fin_pupils = set_worksheet_classes(worksheet2, pupil_quantity, working_days)
        # Filling class report title
        class_zvit = [f"Звіт з харчування учнів {form} класу {school}",
                      f"за    {month_report}      {year_report}     року"]
        # Draw table with column names and get current row after report title
        temp_row = draw_class_report_title(workbook, worksheet2, class_zvit, working_days)
        # Define class teachers
        teachers.setdefault(form, db_all_classes[form].values[0, 0])

        # Write all cells for pupils per working days in one class w/o class teacher
        pupils_col = {}
        temp_col = 0
        temp_row += 2
        working_days = []
        child_days_dict = {}

        for c in class_columns:
            # write dict of column names for each class
            pupils_col.setdefault(c, db_all_classes[form][c])
            cur_row = temp_row
            # skip class teacher column
            if c == class_columns[0]:
                temp_col -= 1
            # print numbers
            elif c == class_columns[1]:  # Numbers
                # pupils_col.setdefault(c, db_all_classes[form][c])
                for i in list(pupils_col[c]):
                    worksheet2.write(cur_row, temp_col, i, cell_formats.pupils_number_format(workbook))
                    cur_row += 1
            elif c == class_columns[2]:  # Pupil name
                # pupils_col.setdefault(c, db_all_classes[form][c])
                for i in list(pupils_col[c]):
                    worksheet2.write(cur_row, temp_col, i, cell_formats.pupils_format(workbook))
                    cur_row += 1
            else:
                for i in list(pupils_col[c]):
                    # go around fault of Nan value writing. Change Nan to ''
                    i = cell_value_check(i, cells_const_values)
                    try:
                        worksheet2.write(cur_row, temp_col, i,
                                         cell_formats.text_pupil_cells_center_format(workbook))
                        cur_row += 1
                    except TypeError:
                        i = ''
                        worksheet2.write(cur_row, temp_col, i,
                                         cell_formats.text_pupil_cells_center_format(workbook))

                    # calc child days for every working day
                    child_days = len(db_all_classes[form][db_all_classes[form][c] == cells_const_values[0]])
                    child_days_dict.setdefault(c, child_days)
                    # make dataframe with child days for all classes
                    calc_df = pd.DataFrame({form: child_days_dict})
            temp_col += 1
        # write results for each working day in rows "Всього дітоднів" and "Сума"
        col = 0
        draw_class_report_child_days(workbook, worksheet2, dict(calc_df[form]), price, row_fin_pupils, col)
        # For all rows with pupils calc missed and child days
        for r in db_all_classes[form].index:
            sum_missed = 0
            sum_child = 0
            # Go for rows
            for q in range(3, len(class_columns)):
                if type(q) is int:
                    """
                    FOR FUTURE:
                    MAKE MERGE OF EQUAL CELLS IN ROW FOR INDIVIDUAL EDUCATION, IN AND OUT
                    """
                    # Calc missed and child days for each pupil
                    sum_missed += pupil_missed_days(db_all_classes[form].iloc[r, q], cells_const_values)
                    sum_child += pupil_child_days(db_all_classes[form].iloc[r, q], cells_const_values)
            # Write to class report missed and child days for each pupil, and money
            worksheet2.write(5 + r, len(class_columns) - 1, sum_missed,
                             cell_formats.pupils_number_format(workbook))
            worksheet2.write(5 + r, len(class_columns), sum_child,
                             cell_formats.pupils_number_format(workbook))
            worksheet2.write(5 + r, len(class_columns) + 1, sum_child * price,
                             cell_formats.pupils_child_price_format(workbook))
            # Calc sum for missed and child days for class
            missed_days_form += sum_missed
            child_days_form += sum_child

        # Write to class report missed and child days for each pupil, and money + draw free cells
        worksheet2.write(5 + pupil_quantity, len(class_columns) - 1, missed_days_form,
                         cell_formats.pupils_number_format(workbook))
        worksheet2.write(5 + pupil_quantity + 1, len(class_columns) - 1, '',
                         cell_formats.pupils_number_format(workbook))

        worksheet2.write(5 + pupil_quantity, len(class_columns), child_days_form,
                         cell_formats.pupils_number_format(workbook))
        worksheet2.write(5 + pupil_quantity + 1, len(class_columns), '',
                         cell_formats.pupils_number_format(workbook))

        worksheet2.write(5 + pupil_quantity, len(class_columns) + 1, '',
                         cell_formats.pupils_child_price_format(workbook))
        worksheet2.write(5 + pupil_quantity + 1, len(class_columns) + 1, child_days_form * price,
                         cell_formats.pupils_child_price_format(workbook))

        # Draw signs in the bottom of sheets for each class
        draw_class_report_signs(workbook, worksheet2, teachers[form], 5 + pupil_quantity + 2)

    """"
    Продолжить тут формирование отчетов по классам
    + Заполнить строки с учениками
    + просчитать все суммы и внести в отчеты по классам
    - добавить подписи
    """

        # calc missed days, child days, sum for every pupil

    # for sheet_name in classes_list:

    # fill the first worksheet
    zvit = ["Звіт",
            f"з харчування учнів 1-4 класів {school}",
            f"за    {month_report}      {year_report}     року"]

    temp_row = 0
    temp_row = draw_report_title(workbook, worksheet, zvit, row=temp_row)

    draw_report_signs(workbook, worksheet, temp_row)

    workbook.close()
