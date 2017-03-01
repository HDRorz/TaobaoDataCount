# -*- coding:utf-8 -*-

from openpyxl import Workbook
from openpyxl import cell
from openpyxl import load_workbook


class Reader(object):

    def __init__(self, filename):
        self.workbook = load_workbook(filename)
        wb = Workbook()
        self.buy_sheet = wb.create_sheet()
        self.pay_sheet = wb.create_sheet()

    def read(self):
        return None

    def read_sheet(self, index):
        self.buy_sheet = self.workbook.worksheets[index]
        self.pay_sheet = self.workbook.worksheets[index + 1]

        buyno = self.buy_sheet.columns[0]
        buylist = []
        for i in range(1,len(buyno)) :
            temp = {'buyno': buyno[i], 'row': self.buy_sheet.rows[i]}
            buylist.append(temp)

        payno = self.pay_sheet.columns[0]
        paylist = []
        for i in range(1, len(payno)):
            temp = {'payno': payno[i], 'row': self.pay_sheet.rows[i]}
            paylist.append(temp)

        inner_list = []
        left_list = []
        right_list = []

        for buy in buylist :
            temp = {'buy': buy, 'pay' : None}
            for pay in paylist :
                if buy['buyno'] == pay['payno'] :
                    temp['pay'] = pay
                    inner_list.append(temp)
                    continue
            left_list.append(temp)

        for pay in paylist:
            temp = {'buy': None, 'pay': pay}
            for buy in buylist:
                if buy['buyno'] == pay['payno']:
                    temp['pay'] = pay
                    continue
            right_list.append(temp)

