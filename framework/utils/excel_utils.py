from framework.geometry.rect import Rect
from xlsxwriter.utility import xl_rowcol_to_cell

book_formats = {}
default_format = {
    "align": "center",
    "valign": "vcenter",
    "border": 1,
    "font_size": 12,
}


def set_default_format(default_fmt: dict):
    default_format.update(default_fmt)

def init_format(workbook, format_name, format_dict):
    fmt = default_format.copy()
    fmt.update(format_dict)
    book_formats[format_name] = workbook.add_format(fmt)

def set_table_column(sheet, cel, name, size):
    sheet.set_column(cel[1], cel[1], size)
    sheet.write(cel[0], cel[1], name, book_formats['title'])

def get_sum_range_formula(rect_range : Rect):
    return f'=SUM({get_range_notation(rect_range)})'

def get_avg_weighted_formula(data_range : Rect, weight_range : Rect):

    data_range_str = get_range_notation(data_range)
    weights_range_str = get_range_notation(weight_range)
    return f'=SUMPRODUCT( {data_range_str},{weights_range_str} ) / SUM({weights_range_str})'

def get_avg_range_formula(rect_range : Rect):
    return f'=AVERAGE({get_range_notation(rect_range)})'

def get_range_notation(rect_range : Rect):
    return f'{get_cell_notation(*rect_range.min)}:{get_cell_notation(*rect_range.max)}'

def get_cell_notation(row, col):
    return xl_rowcol_to_cell(row, col)