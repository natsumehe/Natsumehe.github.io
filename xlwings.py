#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import xlwings


class ToolExcel(object):
    __file_name = "workbook.xlsx"
    __sheet_name = "Sheet1"

    # 新建工作簿
    @staticmethod
    def workbook_new(file_name: str = __file_name):

        # 工作簿文件路径
        workbook_file_path = os.path.join(os.getcwd(), "workbook", file_name)
        # 工作簿当前目录
        workbook_dir_path = os.path.dirname(workbook_file_path)

        # 如果不存在目录路径,就创建
        if not os.path.exists(workbook_dir_path):
            # 创建工作簿路径,makedirs可以创建级联路径
            os.makedirs(workbook_dir_path)

        # 如果不存在,Excel工作簿文件,就创建工作簿
        if not os.path.exists(workbook_file_path):
            # 打开Excel程序，APP程序(即Excel程序)不可见，只打开不新建工作薄，屏幕更新关闭
            app = xlwings.App(visible=False, add_book=False)
            # Excel工作簿显示警告,不显示
            app.display_alerts = False
            # 工作簿屏幕更新,不更新
            app.screen_updating = False
            # 创建工作簿
            wb = app.books.add()
            # 保存工作簿，若未指定路径，保存在当前工作目录。
            wb.save(workbook_file_path)
            # 关闭工作簿
            wb.close()
            # 退出Excel
            app.quit()

    # 读取工作簿全部内容,返回二维列表
    @staticmethod
    def workbook_read(file_name=__file_name, sheet_name=__sheet_name):
        # 工作簿文件路径
        workbook_file_path = os.path.join(os.getcwd(), "workbook", file_name)
        # 如果文件存在,就执行
        if os.path.exists(workbook_file_path):
            # 打开Excel程序，APP程序(即Excel程序)不可见，只打开不新建工作薄，屏幕更新关闭
            app = xlwings.App(visible=False, add_book=False)
            # Excel工作簿显示警告,不显示
            app.display_alerts = False
            # 工作簿屏幕更新,不更新
            app.screen_updating = False
            # 打开工作簿
            wb = app.books.open(workbook_file_path)
            # 获取活动的工作表
            sheet = wb.sheets[sheet_name]

            # 获取已编辑的矩形区域,最底部且最右侧的单元格
            last_cell = sheet.used_range.last_cell
            # 最大行数
            last_row = last_cell.row
            # 最大列数
            last_col = last_cell.column

            """
            # 读取二维列表
            # 注释:如果含有 .options(expand='table').value 参数,空值隔断的部分,不会被读取
            # data = sheet.range((1, 1), (last_row, last_col)).options(expand='table').value
            """

            # 读取二维列表
            data = sheet.range((1, 1), (last_row, last_col)).value

            # 关闭工作簿
            wb.close()
            # 退出Excel
            app.quit()
            return data

    # 写入二维列表,追加模式
    @staticmethod
    def workbook_append(data: list = None, file_name=__file_name, sheet_name=__sheet_name):
        # 工作簿文件路径
        workbook_file_path = os.path.join(os.getcwd(), "workbook", file_name)

        # 如果工作簿不存在,就创建工作簿
        if not os.path.exists(workbook_file_path):
            ToolExcel.workbook_new()

        # 如果文件存在,就执行
        if os.path.exists(workbook_file_path):
            # 打开Excel程序，APP程序(即Excel程序)不可见，只打开不新建工作薄，屏幕更新关闭
            app = xlwings.App(visible=False, add_book=False)
            # Excel工作簿显示警告,不显示
            app.display_alerts = False
            # 工作簿屏幕更新,不更新
            app.screen_updating = False
            # 打开工作簿
            wb = app.books.open(workbook_file_path)
            # 获取活动的工作表
            sheet = wb.sheets[sheet_name]

            # 获取已编辑的矩形区域,最底部且最右侧的单元格
            last_cell = sheet.used_range.last_cell
            # 最大行数
            last_row = last_cell.row

            # 写入二维列表,追加模式
            sheet.range((last_row + 1, 1)).options(expand='table').value = data

            # # 保存文件,保存以后重新读取单元格,重新获取所有活动区域的cell.
            # # 是否保存, 有待考证?
            # wb.save()

            # 获取已编辑的矩形区域,最底部且最右侧的单元格
            last_cell = sheet.used_range.last_cell
            # 最大行数
            last_row = last_cell.row
            # 最大列数
            last_col = last_cell.column
            # 在range中,cell的大小自适应
            sheet.range((1, 1), (last_row, last_col)).columns.autofit()

            # 保存文件
            wb.save()
            # 关闭工作簿
            wb.close()
            # 退出Excel
            app.quit()

    # 写入二维列表,重写模式
    @staticmethod
    def workbook_rewrite(data: list = None, file_name=__file_name, sheet_name=__sheet_name):
        # 工作簿文件路径
        workbook_file_path = os.path.join(os.getcwd(), "workbook", file_name)

        # 如果工作簿不存在,就创建工作簿
        if not os.path.exists(workbook_file_path):
            ToolExcel.workbook_new()

        # 如果文件存在,就执行
        if os.path.exists(workbook_file_path):
            # 打开Excel程序，APP程序(即Excel程序)不可见，只打开不新建工作薄，屏幕更新关闭
            app = xlwings.App(visible=False, add_book=False)
            # Excel工作簿显示警告,不显示
            app.display_alerts = False
            # 工作簿屏幕更新,不更新
            app.screen_updating = False
            # 打开工作簿
            wb = app.books.open(workbook_file_path)
            # 获取活动的工作表
            sheet = wb.sheets[sheet_name]

            # 清除sheet的内容和格式
            sheet.clear()

            # 写入二维列表,重写模式
            sheet.range("A1").options(expand='table').value = data

            # 获取已编辑的矩形区域,最底部且最右侧的单元格
            last_cell = sheet.used_range.last_cell
            # 最大行数
            last_row = last_cell.row
            # 最大列数
            last_col = last_cell.column
            # 所有range的大小自适应
            sheet.range((1, 1), (last_row, last_col)).columns.autofit()

            # 保存文件
            wb.save()
            # 关闭工作簿
            wb.close()
            # 退出Excel
            app.quit()

