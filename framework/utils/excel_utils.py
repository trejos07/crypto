from framework.geometry.rect import Rect
from xlsxwriter.utility import xl_rowcol_to_cell

book_formats = {}
default_format_properties = {
    "align": "center",
    "valign": "vcenter",
    "border": 1,
    "font_size": 12,
}

def set_default_format(workbook, default_fmt: dict):
    default_format_properties.update(default_fmt)
    init_format(workbook, 'default', default_format_properties)

def init_format(workbook, format_name, properties):
    fmt_properties = default_format_properties.copy()
    fmt_properties.update(properties)
    book_formats[format_name] = workbook.add_format(fmt_properties)

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

def get_cell_notation(row, col, row_abs=False, col_abs=False):
    return xl_rowcol_to_cell(row, col, row_abs, col_abs)

def get_format_properties(format_name):
    fmt = book_formats[format_name]
    dft_fmt = book_formats['default']
    properties = [f[4:] for f in dir(fmt) if f[0:4] == 'set_']
    return {key : value for key, value in fmt.__dict__.items() if key in properties and dft_fmt.__dict__[key] != value}

def combine_formats(workbook, format_names : list, format_name : str):
    combined_format = {}
    for fmt_name in format_names:
        fmt_properties = get_format_properties(fmt_name)
        combined_format.update(fmt_properties)

    init_format(workbook, format_name, combined_format)