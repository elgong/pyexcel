"""
    openpyxl 的封装，有很多函数尽可能的使用原始的名字
"""

import os
import time
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from .Config import *

class ExcelUtil(object):

    wb = Workbook()
    # 工作簿的名字
    wb_name = None

    def __init__(self, work_book_name=None, sheet_list=["sheet1"]):
        """
        :param work_book_name:  excel 名
        :param sheet_list:  sheet 名的列表
        """
        # 类型检查 -保证传入列表
        if not isinstance(sheet_list, list):
            raise Exception("waring:  sheet_list need  list[] ")

        if work_book_name is not None:
            self.wb_name = work_book_name
            # 创建 sheet
            for sheetName in sheet_list:
                self.wb.create_sheet(sheetName)

            # 删除默认生成的 Sheet
            self.wb.remove(self.wb["Sheet"])

            print("创建sheet列表： ", end="  ")
            print(self.wb.sheetnames)



    # 添加表
    def add_sheet(self, sheet_list=[]):
        for sheetName in sheet_list:
            self.wb.create_sheet(sheetName)

    # cell 常规设置（表，坐标，值，背景色，字体，边框，对齐方式）
    def set_cell(self, sheet, row, col, value=None, fill=None, font=None, border=None, alignment="center" ):

        cell = self.wb[sheet].cell(row=row, column=col)

        # 设置内容
        if value is not None:
            cell.value = value

        # 填充 背景颜色
        if fill is not None:
            cell.fill = cell_background[fill]

        # 设置字体
        if font is not None:
            cell.font = font_styles[font]

        # 设置单元格 右边框样式
        if border == "right":
            cell.border = cell_border[border]

        # 对齐方式
        if alignment is not None:
            cell.alignment = cell_alignment[alignment]

        return cell

    def set_title(self):
        pass

    # 调整列宽度
    def set_col_weight(self, sheet, col, width=15):
        sheet = self.wb[sheet]
        sheet.column_dimensions[get_column_letter(col)].width = width

    # def set_row_weight(self, sheet, row, width=15):
    #     sheet = self.wb[sheet]
    #     sheet.[get_column_letter(row)].width = width


    # 合并单元格

    def merge_cells(self, sheet, start_row, start_column, end_row, end_column):
        sheet = self.wb[sheet]
        sheet.merge_cells(start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column)

    # 自动保存文件
    def save(self):
        try:
            if os.path.exists(self.wb_name):
                os.remove(self.wb_name)
            self.wb.save(self.wb_name)
            print("save success")
        except PermissionError as e:
            print("Permission denied!, Wait close the old file")
            print("You have 5s to close")
            time.sleep(5)
            self.wb.save(self.wb_name)
        finally:
            print("save success")












