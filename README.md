# pyexcel

> 我在 cis** 实习期间使用过 `openpyxl` 库, 使用时感觉有很多功能比较零碎, 比如在单元格 `cell` 和字体格式 `Font`, 不能直接通过源代码快速推断出该怎么通过 `Font`类配置字体，官方文档我也很难推断出代码到底该怎么写（承认我有点弱）。
>`pyexcel` 库是我对 `openpyxl` 的一层封装，希望能将常用的操作组合在一起, 给开发带来方便。

## 依赖环境
- openpyxl = 2.6.1
- python = 3.7

## 模块解析
- `Styles.py`
    - 样式相关代码（字体样式，单元格背景填充样式）
  
- `ExcelExport.py`

## 简单测试
    # 1. excel 工作簿处理对象
    excelUtil = ExcelUtil("./example.xlsx")

    # 2. 通过列表创建工作表
    new_sheet_list = ["sheet-1", "sheet-2", "sheet-3"]
    excelUtil.add_sheet(new_sheet_list)
    
    # 3. 指定位置写入指定内容
    excelUtil.set_cell(
                   "sheet-1", 
                   col=1, 
                   row=1, 
                   value="Text", 
                   cell_fill="yellow", 
                   cell_alignment="center", 
                   font_bold=True, 
                   font_size=18
              )
    # 4. 保存          
    excelUtil.save()
    
## API

- 添加工作表

`add_sheet(self, sheet_list=[])`

- 设置单元格

        set_cell(
                    sheet, # 工作表名
                    row,  # 单元格所在行
                    col,  # 单元格所在列
                    value=None,  # 要写入的值
                    cell_fill=None,  # 背景是否填充
                    cell_border=None,  # 边框
                    cell_alignment="center",  # 对齐方式， 默认居中对齐
                    font_type=u'Calibri',  # 字体 类型
                    font_size=12,   # 字体大小
                    font_bold=False,  # 字体是否加粗
                    font_italic=False,  # 字体 斜体
                    font_strike=False,  # 字体下划线
                    font_color="black"  # 字体颜色
              )


- 调整行高

        set_row_height(
                    sheet,  # 工作表名
                    row,   # 行数
                    height=15  # 默认行高15
                )

- 调整列宽

        set_col_weight(
                    sheet,  # 表名
                    col,    # 列数
                    width=15  # 默认列宽 15
               )

- 合并单元格（需要指定起始，终止的行和列）
        merge_cells(
                sheet,     # 表名
                start_row, 
                start_column, 
                end_row, 
                end_column
               )

- 2D 柱状图

- 3D 饼状图
`draw_pie3D`

- 保存工作簿
`save()`




