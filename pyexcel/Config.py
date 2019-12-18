
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

# cell 背景 颜色参数字典
"""
fill_type="solid"       # 纯色填充方式
start_color='FF0000FF'  # 前景色
end_color='FF000000'    # 背景色
"""
cell_background = {
    None: PatternFill(fill_type="none", start_color='FFFFFF', end_color='FF000000'),
    "red": PatternFill(fill_type="solid", start_color='FF4500', end_color='FF000000'),
    "green":  PatternFill(fill_type="solid", start_color='7FFF00', end_color='FF000000'),
    "blue": PatternFill(fill_type="solid", start_color='1E90FF', end_color='FF000000'),
    "yellow": PatternFill(fill_type="solid", start_color='FFD700', end_color='FF000000'),
}


# 字体
"""
bold  加粗
italic  斜体
"""
font_styles = {
    None:  Font(u'宋体', size=11, bold=False, italic=False, strike=True, color='000000'),
    "white": Font(u'宋体', size=11, bold=True, italic=False, strike=False, color='FFFFFF'),
}


# 单元格对齐方式
"""
horizontal   水平方向
vertical    垂直方向
"""
cell_alignment = {
    None: Alignment(horizontal='center', vertical='center', text_rotation=0, wrap_text=False, shrink_to_fit=False,
                    indent=0),
}


border = Border(left=Side(border_style=None,
                           color='FF000000'),
                right=Side(border_style=None,
                           color='FF000000'),
                top=Side(border_style=None,
                          color='FF000000'),
                 bottom=Side(border_style=None,
                             color='FF000000'),
                 diagonal=Side(border_style=None,
                               color='FF000000'),
                 diagonal_direction=0,
                 outline=Side(border_style=None,
                             color='FF000000'),
                vertical=Side(border_style=None,
                              color='FF000000'),
                horizontal=Side(border_style=None,
                                color='FF000000')
               )

number_format = 'General'
protection = Protection(locked=True,
                        hidden=False)