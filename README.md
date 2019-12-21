# pyexcel

## 开发环境
- openpyxl = 2.6.1
- python = 3.7

## 常用 API

`   

    add_sheet(self, sheet_list=[])

    # cell 常规设置（表，坐标，值，背景色，字体，边框，对齐方式）
    set_cell(self, sheet, row, col, value=None, fill=None, font=None, border=None, alignment="center" )

    set_title(self)

    # 调整列宽度
    set_col_weight(self, sheet, col, width=15)
    set_row_weight(self, sheet, row, width=15)

    # 合并单元格
    merge_cells(self, sheet, start_row, start_column, end_row, end_column)
        
    # 自动保存文件
    def save(self):
`

