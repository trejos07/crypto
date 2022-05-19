from ast import Lambda
from locale import currency
import os
import math
from binance.client import Client
import pandas as pd
import xlsxwriter
from framework.geometry.rect import Rect
import framework.utils.excel_utils as excel_utils


file_name = 'spot-trade-history.xlsx'
pairs = ['BTCUSDT', 'ETHUSDT', 'SOLUSDT', 'ADAUSDT', 'DOTUSDT', 'FTMUSDT']

def main():
    api_key = os.getenv('BNCAK')
    api_secret = os.getenv('BNCSK')
    client = Client(api_key, api_secret)

    workbook = xlsxwriter.Workbook(file_name)
    excel_utils.set_default_format({"align": "center", "valign": "vcenter", "border": 1, "font_size": 12, })
    init_formats(workbook)

    for pair in pairs:
        current_price = round(float(client.get_avg_price(symbol='BTCUSDT')['price']), 2)
        trades = get_trade_data(client, pair)
        create_pair_sheet(workbook, pair, trades, current_price)

    workbook.close()
    os.startfile(file_name)

def get_trade_data(client, pair):
    trades = client.get_my_trades(symbol=pair)
    trades = [parse_trade_data(trade) for trade in trades]
    return trades

def parse_trade_data(trade_data):
    side = 'BUY' if trade_data['isBuyer'] else 'SELL'
    date_time = pd.to_datetime(trade_data['time'], unit='ms')
    price = float(trade_data['price'])
    quantity = float(trade_data['qty'])
    return Trade(trade_data["symbol"], side, price, quantity, date_time)

def create_pair_sheet(workbook, pair, trades = None, current_price = None):
    sheet = workbook.add_worksheet(pair)

    sheet.merge_range(0, 0, 0, 8, pair, excel_utils.book_formats['title'])
    sheet.merge_range(1, 0, 1, 8, current_price, excel_utils.book_formats['currency-2'])
    sheet.merge_range(2, 0, 2, 8, 'Trades', excel_utils.book_formats['title'])

    excel_utils.set_table_column(sheet, (3, 0), 'Date', 16)
    excel_utils.set_table_column(sheet, (3, 1), 'Side', 10)
    excel_utils.set_table_column(sheet, (3, 2), 'Price', 14)
    excel_utils.set_table_column(sheet, (3, 3), 'Quantity', 12)
    excel_utils.set_table_column(sheet, (3, 4), 'Cost', 10)
    excel_utils.set_table_column(sheet, (3, 5), 'Total Quantity', 18)
    excel_utils.set_table_column(sheet, (3, 6), 'Total Cost', 14)
    excel_utils.set_table_column(sheet, (3, 7), 'Profit', 10)
    excel_utils.set_table_column(sheet, (3, 8), 'Profit %', 10)

    table_start_row = 4
    total_quantity = 0
    total_cost = 0

    for i, trade in enumerate(trades):
        
        total_quantity += trade.quantity
        total_cost += trade.cost

        sheet.write(i + table_start_row, 0, trade.time.strftime("%d/%m/%Y"), excel_utils.book_formats["default"])
        sheet.write(i + table_start_row, 1, trade.side, excel_utils.book_formats["default"])
        sheet.write(i + table_start_row, 2, trade.price, excel_utils.book_formats["currency"])	
        sheet.write(i + table_start_row, 3, trade.quantity, excel_utils.book_formats["default"])
        sheet.write(i + table_start_row, 4, trade.cost, excel_utils.book_formats["currency"])
        sheet.write(i + table_start_row, 5, total_quantity, excel_utils.book_formats["default"])
        sheet.write(i + table_start_row, 6, total_cost, excel_utils.book_formats["currency"])
        sheet.write(i + table_start_row, 7, "-", excel_utils.book_formats["default"])
        sheet.write(i + table_start_row, 8, "-", excel_utils.book_formats["default"])

    count = len(trades)
    table_end_row = table_start_row + count - 1

    price_range = Rect(table_start_row, 2, count - 1, 0)
    quantity_range = Rect(table_start_row, 3, count - 1, 0)
    cost_range = Rect(table_start_row, 4, count - 1, 0)

    sheet.write(table_end_row + 1, 2, excel_utils.get_avg_weighted_formula(price_range, quantity_range), excel_utils.book_formats["currency"])
    sheet.write(table_end_row + 1, 3, excel_utils.get_sum_range_formula(quantity_range), excel_utils.book_formats["default"])
    sheet.write(table_end_row + 1, 4, excel_utils.get_avg_weighted_formula(cost_range, quantity_range), excel_utils.book_formats["currency"])

    color_range = excel_utils.get_range_notation(Rect(table_start_row, 0, count - 1, 8))
    side_column_start = excel_utils.get_cell_notation(table_start_row, 1)

    sheet.conditional_format(color_range,  {"type": "formula", "criteria": f'=${side_column_start}="BUY"', "format": excel_utils.book_formats["green"]})
    sheet.conditional_format(color_range,  {"type": "formula", "criteria": f'=${side_column_start}="SELL"', "format": excel_utils.book_formats["red"]})

    return sheet


def init_formats(workbook):
    excel_utils.init_format(workbook, 'default', {})
    excel_utils.init_format(workbook, 'title', {'bold': True, 'font_size': 14, 'bg_color': '#555555', 'font_color': '#ffffff'})
    excel_utils.init_format(workbook, 'subtitle', {'bold': True, 'font_size': 14,})
    excel_utils.init_format(workbook, 'date', {'num_format': 'dd/mm/yyyy'})
    excel_utils.init_format(workbook, 'currency', {'num_format': '$#,##0.00'})
    excel_utils.init_format(workbook, 'currency-2', {'num_format': '$#,##0.00', 'bold': True, 'font_size': 14})
    excel_utils.init_format(workbook, 'green', {'font_color': '#006100', 'bg_color': '#C6EFCE'})
    excel_utils.init_format(workbook, 'red', {'font_color': '#9C0006', 'bg_color': '#FFC7CE'})

def get_number_format(x):

    def int_len(n):
        if int(n) == 0:
            return 1

        return int(math.log10(n)) + 1

    def float_len(x):
        decimal_part = x % 1
        return len(str(decimal_part).replace('0.', ''))

    int_l = int_len(x)
    float_l = float_len(x)

    fmt = []
    if int_l > 3:
        fmt.append('#,##0')
    else: 
        fmt.append('0' * int_l)

        if int_l < 3:
            if float_l > 0:
                fmt.append('.')

            if int_l > 2 or float_l > 13:
                float_l = min(float_l, 2)

            fmt.append('0' * float_l)

    return ''.join(fmt)

def Average(l, selector = lambda x: x):
    return sum(selector(x) for x in l) / len(l)



class PairData:
    def __init__(self, symbol, current_price):
        self.symbol = symbol
        self.current_price = current_price

class Trade:
    def __init__(self, symbol, side, price, qty, time):
        self.symbol = symbol
        self.side = side
        self.price = price
        self.quantity = qty if side == 'BUY' else -qty
        self.time = time
        self.cost = self.quantity * self.price

    def __str__(self):
        return f'{self.symbol} {self.side} {self.price} {self.quantity} {self.time}'

    def __repr__(self):
        return f'Trade({self.symbol}, {self.side}, {self.price}, {self.quantity}, {self.time})'


if __name__ == '__main__':
    main()