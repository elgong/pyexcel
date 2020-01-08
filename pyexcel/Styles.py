"""
    log：
    2020/1/8
    重构了格式控制代码。
    通过Styles 类提供所有的格式风格。
"""
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


class Styles(object):

    """ 预设参数 """
    # 背景颜色配置
    cell_background_parameter = {
        None: PatternFill(fill_type="none", start_color='FFFFFF', end_color='FF000000'),

        "red": PatternFill(fill_type="solid", start_color='FF4500', end_color='FF000000'),

        "green": PatternFill(fill_type="solid", start_color='7FFF00', end_color='FF000000'),

        "blue": PatternFill(fill_type="solid", start_color='1E90FF', end_color='FF000000'),

        "yellow": PatternFill(fill_type="solid", start_color='FFD700', end_color='FF000000'),

        "light_blue": PatternFill(fill_type='solid', fgColor='C9D9EE'),
        "dark_blue": PatternFill(fill_type='solid', fgColor='5981B8')

        # 这里添加其他背景颜色
    }

    def __init__(self):
        pass

    def cell_background(self, background_color=None):
        try:
            return self.cell_background_parameter[background_color]
        except:
            raise Exception("bad background_color !")




