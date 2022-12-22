TITLE_SIZE = 36
DATE_SIZE = 28
REGULAR_SIZE = 24
REGULAR_FONT = "Arial Cyr"

TITLE_SIZE2 = 20
REGULAR_SIZE2 = 12
REGULAR_FONT2 = "Arial Cyr"
CELLS_SIZE = 8
CHILD_DAYS_SIZE = 10


# format settings for different cell types in .xlsx
def title_format(self):
    title_f = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': TITLE_SIZE,
        'align': 'center',
        'valign': 'vcenter',
    })
    return title_f


def date_format(self):
    date_f = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': DATE_SIZE,
        'align': 'center',
        'valign': 'vcenter',
    })
    return date_f


def default_format(self):
    default = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'valign': 'vcenter',
    })
    return default


def text_box_wrap_format(self):
    text_box_wrap = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'justify',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True
    })
    return text_box_wrap


def text_box_center_wrap_format(self):
    text_box_center_wrap = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True
    })
    return text_box_center_wrap


def text_box_left_wrap_format(self):
    text_box_left_wrap = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'left',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True
    })
    return text_box_left_wrap


def text_box_center_wrap_num_format(self):
    text_box_center_wrap = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True,
        'num_format': '0.00'
    })
    return text_box_center_wrap


def text_box_center_blue_format(self):
    text_box_center_blue = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True,
        'bg_color': '#C5D9F1'
    })
    return text_box_center_blue


def text_box_center_bold_format(self):
    text_box_center_bold = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True,
        'bold': True
    })
    return text_box_center_bold


def text_box_center_bold_num_format(self):
    text_box_center_bold = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True,
        'bold': True,
        'num_format': '0.00'
    })
    return text_box_center_bold


# format settings for different cell types in .xlsx for class' sheets
def title_format2(self):
    title_f2 = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': TITLE_SIZE2,
        'align': 'center',
        'valign': 'vcenter',
    })
    return title_f2


def pupils_format(self):
    pupils_f = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'left',
        'valign': 'vcenter',
        'border': True
    })
    return pupils_f


def pupils_number_format(self):
    pupils_number = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'bold': True
    })
    return pupils_number


def text_box_center_wrap_format2(self):
    text_box_center_wrap2 = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True
    })
    return text_box_center_wrap2


def text_box_center_nums_top_format(self):
    text_box_center_nums_top = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'center',
        'valign': 'top',
        'border': True,
    })
    return text_box_center_nums_top


def text_box_center_nums_bot_format(self):
    text_box_center_nums_bot = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'center',
        'valign': 'bottom',
        'border': True,
    })
    return text_box_center_nums_bot


def text_pupil_cells_center_format(self):
    text_pupil_cells_center = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': CELLS_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'num_format': '0.00'
    })
    return text_pupil_cells_center


def classes_results_names_format(self):
    classes_results_names = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'left',
        'valign': 'vcenter',
        'border': True,
        'bold': True
    })
    return classes_results_names


def classes_results_child_days_format(self):
    classes_results_child_days = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': CHILD_DAYS_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'bold': False
    })
    return classes_results_child_days


def classes_results_sum_format(self):
    classes_results_sum = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': CHILD_DAYS_SIZE,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'bold': False,
        'rotation': 90,
        'num_format': '0.00'
    })
    return classes_results_sum


def pupils_child_price_format(self):
    pupils_child_price = self.add_format({
        'font_name': REGULAR_FONT2,
        'font_size': REGULAR_SIZE2,
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'bold': True,
        'num_format': '0.00'
    })
    return pupils_child_price


def classes_signs_format(self):
    classes_signs = self.add_format({
        'font_name': REGULAR_FONT,
        'font_size': REGULAR_SIZE2,
        'align': 'left',
        'valign': 'vcenter',
    })
    return classes_signs