"""
    log：
    2020/1/8
    重构了格式控制代码。
    通过Styles 类提供所有的格式风格。
"""
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


class Styles(object):

    """ 预设参数区 """
    # 背景颜色配置
    # fill_type = "solid"  # 纯色填充方式
    # start_color = 'FF0000FF'  # 前景色
    # end_color = 'FF000000'  # 背景色
    cell_background_parameter = {
        None: PatternFill(fill_type="none", start_color='FFFFFF', end_color='FF000000'),

        "red": PatternFill(fill_type="solid", start_color='FF4500', end_color='FF000000'),

        "green": PatternFill(fill_type="solid", start_color='7FFF00', end_color='FF000000'),

        "blue": PatternFill(fill_type="solid", start_color='1E90FF', end_color='FF000000'),

        "yellow": PatternFill(fill_type="solid", start_color='FFD700', end_color='FF000000'),

        "light_blue": PatternFill(fill_type='solid', fgColor='C9D9EE'),

        "dark_blue": PatternFill(fill_type='solid', fgColor='5981B8')
        # 在这里添加其他背景颜色，颜色表可查 http://www.365jz.com/article/24452
    }

    def __init__(self):
        pass

    # 单元格背景颜色
    def set_cell_background(self, background_color=None):
        try:
            return self.cell_background_parameter[background_color]
        except:
            raise Exception("Background_color Error!")

    # 字体风格
    def set_font(self, _font=u'Calibri', _size=12, _bold=False, _italic=False, _strike=False, _color="black"):
        """
        功能: 字体风格设置, 字体，大小，加粗，斜体， 下划线，颜色等
            参数名: [可选参数,]
            _font: [u'宋体', u'Calibri']
            _size:    字体大小
            _bold:    是否加粗
            _italic:  斜体
            _strike: 下划线
            _color:  字体颜色
        """
        if _color == "black":  # 默认白色
            _color_para = "000000"
        elif _color == "white":
            _color_para ="FFFFFF"
        else:
            raise Exception("Font color Error !")

        return Font(_font, size=_size, bold=_bold, italic=_italic, strike=_strike, color=_color_para)

        # 单元格对齐方式
        # """
        # horizontal   水平方向
        # vertical    垂直方向
        # """
        # cell_alignment = {
        #     "center": Alignment(
        #         horizontal='center', vertical='center', text_rotation=0, wrap_text=False, shrink_to_fit=False,
        #         indent=0
        #     ),
        #
        # }


        # font_styles = {
        #
        #     # """ 标题字体 """
        #     "title": Font(u'Calibri', size=26, bold=True, italic=False, strike=False, color='000000'),
        #
        #     # """  内容字体 """
        #     # 黑色  宋体
        #     "black": Font(u'宋体', size=11, bold=False, italic=False, strike=False, color='000000'),
        #
        #     # 黑色  宋体 加粗
        #     "black_bold": Font(u'宋体', size=11, bold=True, italic=False, strike=False, color='000000'),
        #
        #     # 白色  宋体
        #     "white": Font(u'宋体', size=11, bold=False, italic=False, strike=False, color='FFFFFF'),
        #     # 白色  宋体  加粗
        #     "white_bold": Font(u'宋体', size=11, bold=True, italic=False, strike=False, color='FFFFFF'),
        #
        #     # 这里添加其他样式字体
        # }



