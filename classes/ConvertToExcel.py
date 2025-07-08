from openpyxl import workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
import re

class HtmlTableToEcelConverter:
    CSS_COLOR_NAMES = {
        "black": "000000",
        "silver": "C0C0C0",
        "gray": "808080",
        "white": "FFFFFF",
        "maroon": "800000",
        "red": "FF0000",
        "purple": "800080",
        "fuchsia": "FF00FF",
        "green": "008000",
        "lime": "00FF00",
        "olive": "808000",
        "yellow": "FFFF00",
        "navy": "000080",
        "blue": "0000FF",
        "teal": "008080",
        "aqua": "00FFFF"
    }
    def __init__(self, table_data):
        self.table_data = table_data
        self.wb = Workbook()
        self.sheet = self.wb.active
        self.current_row = 1


    def apply_styles(self, cell, style_str):
        if not style_str:
            return
        style_dict = self.parse_style(style_str)
        if 'width' in style_dict:
            width_str = style_dict['width']
            try:
                width = int(width_str.replace('px', '').strip()) / 7
                amount_columns = len(self.table_data['tbody']['rows'][0]['cells'])
                width = width / amount_columns
                col_letter = cell.column_letter
                self.sheet.column_dimensions[col_letter].width = width
            except ValueError:
                pass

        if 'text-align' in style_dict:
            cell.alignment = Alignment(horizontal=style_dict['text-align'])
        if 'font-weight' in style_dict and style_dict['font-weight'] == 'bold':
            cell.font = Font(bold=True)
        if 'color' in style_dict:
            color_val = style_dict['color'].strip().lower()
            if color_val.startswith('#'):
                hex_color = color_val.lstrip('#')
            elif color_val in self.CSS_COLOR_NAMES:
                hex_color = self.CSS_COLOR_NAMES[color_val]
            else:
                hex_color = None
            if hex_color and re.fullmatch(r'[0-9a-fA-F]{6}', hex_color):
                font_color = 'FF' + hex_color.upper()
                old_font = cell.font or Font()
                cell.font = Font(
                    name=old_font.name,
                    size=old_font.size,
                    bold=old_font.bold,
                    italic=old_font.italic,
                    underline=old_font.underline,
                    color=font_color
                )
        if 'background-color' in style_dict:
            color = style_dict['background-color'].replace('#', '')
            if re.fullmatch(r'[0-9a-fA-F]{6}', color):
                excel_color = 'FF' + color.upper()
                cell.fill = PatternFill(start_color=excel_color, patternType="solid")
        if 'border' in style_dict:
            border_str = style_dict['border']
            parts = border_str.split()
            if len(parts) >= 3:
                width_str, style_str, color_str = parts[:3]
                try:
                    width_px = int(width_str.replace('px', ''))
                except ValueError:
                    width_px = 1

                if width_px <= 1:
                    border_style = "thin"
                elif width_px <= 2:
                    border_style = "medium"
                else:
                    border_style = "thick"

                color = color_str.lstrip('#')
                if len(color) == 6:
                    color = "FF" + color.upper()
                else:
                    color = "FF000000"

                side = Side(border_style=border_style, color=color)
                cell.border = Border(left=side, right=side, top=side, bottom=side)

    def parse_style(self, style_str):
        style_dict = {}
        for part in style_str.split(';'):
            if ':' in part:
                key, value = part.strip().split(':', 1)
                style_dict[key.strip()] = value.strip()
        return style_dict

    def set_col_widths(self):
        for index, col in enumerate(self.table_data['colgroup'], start=1):
            style = col.get('style')
            width = None
            if style:
                styles = self.parse_style(style)
                width_str = styles.get('width')
                if width_str and width_str.endswith('px'):
                    try:
                        width = int(width_str.replace('px', '').strip()) / 7
                    except ValueError:
                        pass
            if width:
                col_letter = get_column_letter(index)
                self.sheet.column_dimensions[col_letter].width = width

    def add_styles_to_section(self, section):
        section_data = self.table_data[section]
        section_style = section_data.get('style', None)

        for row in section_data.get('rows', []):
            row_style = row.get('style', None)
            for col_idx, cell_data in enumerate(row['cells'], start=1):
                cell = self.sheet.cell(row=self.current_row, column=col_idx)
                cell.value = cell_data['text']

                self.apply_styles(cell, self.table_data['table_style'])
                self.apply_styles(cell, section_style)
                self.apply_styles(cell, row_style)
                self.apply_styles(cell, cell_data.get('style'))

            self.current_row += 1


    def convert(self, output_file='test.xlsx'):
        self.set_col_widths()
        self.add_styles_to_section('thead')
        self.add_styles_to_section('tbody')
        self.add_styles_to_section('tfoot')
        self.wb.save(output_file)