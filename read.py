# -*- coding:utf-8 -*-

import os
import copy

from openpyxl import Workbook
from openpyxl import load_workbook


class Reader(object):
    def __init__(self, filename):
        self.workbook = load_workbook(filename)
        self.sheet_count = len(self.workbook.worksheets)
        self.buy_sheet = None
        self.pay_sheet = None
        self.buy_title_row = None
        self.pay_title_row = None

    def read(self):
        for i in range(0, self.sheet_count, 2):
            self.read_sheet(i)

    def read_sheet(self, index):
        self.buy_sheet = self.workbook.worksheets[index]
        self.pay_sheet = self.workbook.worksheets[index + 1]
        buy_sheet = self.buy_sheet
        pay_sheet = self.pay_sheet

        buyno = buy_sheet.columns.next()
        rows = []
        for row in buy_sheet.rows:
            rows.append(row)
        row_num = 0
        self.buy_title_row = self.tran(rows[0])
        buylist = []
        for no in buyno[1:]:
            row_num += 1
            row = rows[row_num]
            temp = {'buyno': no.value, 'row': self.tran(row)}
            buylist.append(temp)

        payno = self.pay_sheet.columns.next()
        rows = []
        for row in pay_sheet.rows:
            rows.append(row)
        row_num = 0
        self.pay_title_row = self.tran(rows[0])
        paylist = []
        for no in payno[1:]:
            row_num += 1
            row = rows[row_num]
            temp = {'payno': no.value, 'row': self.tran(row)}
            paylist.append(temp)

        buy_emptyrow = copy.copy(self.buy_title_row)
        for i in range(0, len(buy_emptyrow)):
            buy_emptyrow[i] = ''

        pay_emptyrow = copy.copy(self.pay_title_row)
        for i in range(0, len(pay_emptyrow)):
            pay_emptyrow[i] = ''

        inner_list = []
        left_list = []
        right_list = []

        for buy in buylist:
            temp = {'buy': buy, 'pay': {'payno': '', 'row': pay_emptyrow}}
            for pay in paylist:
                if buy['buyno'].__class__ == pay['payno'].__class__ and buy['buyno'].strip() == pay['payno'].strip():
                    temp['pay'] = pay
                    inner_list.append(temp)
                    break
            left_list.append(temp)

        for pay in paylist:
            temp = {'buy': {'buy': '', 'row': buy_emptyrow}, 'pay': pay}
            for buy in buylist:
                if buy['buyno'].__class__ == pay['payno'].__class__ and buy['buyno'].strip() == pay['payno'].strip():
                    temp['pay'] = pay
                    break
            right_list.append(temp)

        self.save('out' + str(index) + '.xlsx', '订单与支付都有'.decode('utf8'), inner_list)
        self.save('out' + str(index) + '.xlsx', '所有订单+支付'.decode('utf8'), left_list)
        self.save('out' + str(index) + '.xlsx', '订单+所有支付'.decode('utf8'), right_list)

    def save(self, filename, sheetname, outlist):
        if os.path.exists(filename):
            wb = load_workbook(filename)
        else:
            wb = Workbook()
            wb.remove(wb.worksheets[0])
        sheet = wb.create_sheet(sheetname)

        col_i = 1
        row_i = 1
        for title in self.buy_title_row:
            sheet.cell(column=col_i, row=row_i).value = title
            col_i += 1

        for title in self.pay_title_row:
            sheet.cell(column=col_i, row=row_i).value = title
            col_i += 1

        col_i = 1
        row_i += 1

        for item in outlist:
            for ce in item['buy']['row']:
                sheet.cell(column=col_i, row=row_i).value = ce
                col_i += 1

            for ce in item['pay']['row']:
                sheet.cell(column=col_i, row=row_i).value = ce
                col_i += 1

            col_i = 1
            row_i += 1

        wb.save(filename)
        wb.close()

    def tran(self, row):
        ret = []
        for ce in row:
            ret.append(ce.value)
        return ret


if __name__ == "__main__":
    reader = Reader('in.xlsx')
    reader.read()
