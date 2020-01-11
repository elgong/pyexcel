from pyexcel.ExcelExport import *

excel_path = "./output/"
excel_name = "example.xlsx"

# 写入Excel 的逻辑
def main():

    # excel 处理对象
    excelUtil = ExcelUtil(excel_path + excel_name)

    # 添加sheet-2
    new_sheet_list = ["sheet-1", "sheet-2", "sheet-3"]
    excelUtil.add_sheet(new_sheet_list)

    # 对 sheet 的处理
    # 循环处理吧, 最后三个表达到一样的效果
    # 1. 设置主标题
    base_row = 1
    base_col = 1
    new_sheet_list = ["sheet-1"]
    for sheet_name in new_sheet_list:
        excelUtil.merge_cells(sheet_name, base_row, base_col, base_row, base_row+4)
        excelUtil.set_cell(sheet_name, 1, 1, value="Example"+sheet_name, cell_fill="yellow", cell_alignment="center", font_bold=True, font_size=18)

    # 2. 二级列标
    base_row = 2
    base_col = 2
    level_2_col_list = ["col-1", "col-2", "col-3", "col-4"]
    for sheet_name in new_sheet_list:
        for col_offset, value in enumerate(level_2_col_list):
            excelUtil.set_cell(sheet_name, base_row, base_col+col_offset, value=value, cell_fill="dark_blue", cell_alignment="center")
            # 调整列间距
            excelUtil.set_col_weight(sheet_name, base_col+col_offset, 20)

    # 3. 一级行标
    base_row = 3
    base_col = 1
    level_2_row_list = ["row-1", "row-2", "row-3", "row-4"]
    for sheet_name in new_sheet_list:
        for row_offset, value in enumerate(level_2_row_list):
            excelUtil.set_cell(sheet_name, base_row + row_offset, base_col, value=value, cell_fill="dark_blue", cell_alignment="center")
            # 调整行间距
            excelUtil.set_col_weight(sheet_name, base_row + row_offset, 15)

    base_col = 2
    base_row = 3
    # data
    for offset in range(4):
        excelUtil.set_cell("sheet-1", base_row+offset, base_col, value=offset+5)

    excelUtil.draw_pie3D(
                    "sheet-1", data_position=[2, 2, 3, 6], label_position=[1, 1, 3, 6],
                    display_position="A10", title="Bar Chart", x_title="display_x",
                    y_title="display_y"
              )
    excelUtil.save()

if __name__ == "__main__":

    main()





