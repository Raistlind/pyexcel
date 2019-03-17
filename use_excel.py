#!/usr/bin/env python
# -*- encoding: utf-8 -*-
"""
@File    :   use_excel.py    
@Contact :   Johnd0712@hotmail.com
@License :   (C)Copyright 2017-2018, Krynn.cn

@Modify Time         @Author      @Version    @Desciption
-----------------    ---------    --------    -----------
2019-03-17 16:28      Raistlind    1.0         None
"""

# import lib
from datetime import datetime

from openpyxl import Workbook
from openpyxl.drawing.image import Image


class ExcelUtils():

    def __init__(self):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws_two = self.wb.create_sheet('第一个表单')
        self.ws_three = self.wb.create_sheet()

    def do_sth(self):
        self.ws['A1'] = 66
        self.ws['A2'] = '测试'
        self.ws['A3'] = datetime.now()

        img = Image('./static/130402161P2_300.jpg')
        self.ws.add_image(img, 'B1')
        self.wb.save('./static/test.xlsx')


if __name__ == '__main__':
    client = ExcelUtils()
    client.do_sth()
