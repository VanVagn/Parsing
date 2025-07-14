from openpyxl import workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.styles.builtins import total
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
import re

class HtmlTableToExcelConverter:
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

    APPLICABLE_CSS_PROPERTIES = {
        'font-weight', 'font-style', 'text-decoration', 'color', 'background-color',
        'text-align', 'vertical-align', 'border'
    }


    def __init__(self, table_data):
        self.table_data = table_data
        self.wb = Workbook()
        self.sheet = self.wb.active
        self.current_row = 1

    def merge_styles(self, *styles):
        merged = {}
        for style_str in styles:
            if style_str is None:
                style_str = ''
            style_dict = self.parse_style(style_str)
            for key, value in style_dict.items():
                if key in self.APPLICABLE_CSS_PROPERTIES:
                    merged[key] = value
        return merged

    def apply_styles(self, cell, style_dict):
        if not style_dict or not isinstance(style_dict, dict):
            return

        # Выравнивание текста
        horizontal = style_dict.get('text-align', None)
        vertical = style_dict.get('vertical-align', None)

        if vertical == 'middle':
            vertical = 'center'

        alignment_kwargs = {}
        if horizontal:
            alignment_kwargs['horizontal'] = horizontal
        if vertical:
            alignment_kwargs['vertical'] = vertical
        if alignment_kwargs:
            cell.alignment = Alignment(**alignment_kwargs)

        # Стиль текста
        bold = style_dict.get('font-weight', None)
        italic = style_dict.get('font-style', None)
        underline = style_dict.get('text-decoration', None)
        if underline:
            underline = 'single' if 'underline' in underline else None

        # Цвет текста
        color_val = style_dict.get('color', "").strip().lower()
        if color_val.startswith('#'):
            hex_color = color_val.lstrip('#')
        elif color_val in self.CSS_COLOR_NAMES:
            hex_color = self.CSS_COLOR_NAMES[color_val]
        else:
            hex_color = None
        font_color = None
        if hex_color and re.fullmatch(r'[0-9a-fA-F]{6}', hex_color):
            font_color = 'FF' + hex_color.upper()

        old_font = cell.font or Font()
        cell.font = Font(
            name=old_font.name,
            size=old_font.size,
            bold=(bold == "bold") if bold is not None else old_font.bold,
            italic=(italic == "italic") if italic is not None else old_font.italic,
            underline=underline if underline is not None else old_font.underline,
            color=font_color if font_color else old_font.color
        )

        # Фон ячейки
        if 'background-color' in style_dict:
            bg_color = style_dict['background-color'].replace('#', '').strip()
            end_color = self.expand_short_hex(bg_color)
            if re.fullmatch(r'[0-9a-fA-F]{6}', end_color):
                excel_color = 'FF' + end_color.upper()
                cell.fill = PatternFill(start_color=excel_color, end_color=excel_color, patternType="solid")

        # Обработка границ
        if 'border' in style_dict:
            border_str = style_dict['border']
            parts = border_str.strip().split()

            width_px = 1
            border_style = "thin"
            color = "FF000000"

            for part in parts:
                if part.endswith("px"):
                    try:
                        width_px = int(part.replace("px", ""))
                    except ValueError:
                        width_px = 1
                elif part.lower() in ["solid", "dashed", "dotted", "double"]:
                    continue  # можно обработать, если нужно
                elif part.startswith("#") or part.lower() in self.CSS_COLOR_NAMES:
                    color_val = part.lower()
                    if color_val.startswith('#'):
                        color_val = color_val.lstrip('#')
                    if color_val in self.CSS_COLOR_NAMES:
                        color_val = self.CSS_COLOR_NAMES[color_val]
                    if re.fullmatch(r'[0-9a-fA-F]{6}', color_val):
                        color = "FF" + color_val.upper()

            if width_px <= 1:
                border_style = "thin"
            elif width_px <= 2:
                border_style = "medium"
            else:
                border_style = "thick"

            side = Side(border_style=border_style, color=color)
            cell.border = Border(left=side, right=side, top=side, bottom=side)

    def parse_style(self, style_str):
        style_dict = {}
        if not style_str:
            return style_dict
        for part in style_str.split(';'):
            if ':' in part:
                key, value = part.strip().split(':', 1)
                style_dict[key.strip()] = value.strip()
        return style_dict

    def expand_short_hex(self, hex_color):
        if len(hex_color) == 3:
            return ''.join([c * 2 for c in hex_color])
        elif len(hex_color) == 6:
            return hex_color
        return hex_color

    def set_col_widths(self):
        colgroup = self.table_data.get('colgroup', [])
        table_style = self.parse_style(self.table_data.get('table_style', ''))

        table_width_px = None
        width_str = table_style.get('width', '')
        if width_str.endswith('px'):
            try:
                table_width_px = int(width_str.replace('px', '').strip())
            except ValueError:
                pass

        body_rows = self.table_data.get('tbody', {}).get('rows', [])
        known_widths = {}
        max_col_index = 0

        for row in body_rows:
            cells = row.get('cells', [])
            col_cursor = 0
            for cell in cells:
                colspan = int(cell.get('colspan', 1))
                style = self.parse_style(cell.get('style', ''))
                width_str = style.get('width', '')

                if width_str.endswith('px'):
                    try:
                        px = int(width_str.replace('px', '').strip())
                        per_col_px = px / colspan
                        for offset in range(colspan):
                            col_idx = col_cursor + offset
                            if col_idx not in known_widths:
                                known_widths[col_idx] = per_col_px
                    except ValueError:
                        pass

                col_cursor += colspan
                max_col_index = max(max_col_index, col_cursor)

        for col_idx in range(max_col_index):
            if col_idx not in known_widths and col_idx < len(colgroup):
                col_style = self.parse_style(colgroup[col_idx].get('style') or '')
                col_width_str = col_style.get('width', '')
                if col_width_str.endswith('px'):
                    try:
                        px = int(col_width_str.replace('px', '').strip())
                        known_widths[col_idx] = px
                    except ValueError:
                        pass

        unknown_indexes = [i for i in range(max_col_index) if i not in known_widths]
        if table_width_px and unknown_indexes:
            total_known = sum(known_widths.values())
            remaining_px = max(table_width_px - total_known, 0)
            per_col_px = remaining_px / len(unknown_indexes)
            for idx in unknown_indexes:
                known_widths[idx] = per_col_px

        for col_idx in range(max_col_index):
            px = known_widths.get(col_idx)
            if px:
                excel_width = (px - 5) / 7
                col_letter = get_column_letter(col_idx + 1)
                self.sheet.column_dimensions[col_letter].width = excel_width

    def add_styles_to_section(self, section):
        section_data = self.table_data[section]
        section_style = section_data.get('style', None)

        rowspan_map = {}

        for row in section_data.get('rows', []):
            row_style = row.get('style', None)
            row_height = None

            # height из стиля строки
            if row_style:
                row_style_dict = self.parse_style(row_style)
                height_str = row_style_dict.get('height')
                if height_str and height_str.endswith('px'):
                    try:
                        px = int(height_str.replace('px', '').strip())
                        row_height = px * 0.75  # Перевод в pt
                    except ValueError:
                        pass

            if row_height:
                self.sheet.row_dimensions[self.current_row].height = row_height

            excel_col_idx = 1

            while (self.current_row, excel_col_idx) in rowspan_map:
                cell_info = rowspan_map[(self.current_row, excel_col_idx)]
                cell = self.sheet.cell(row=self.current_row, column=excel_col_idx)
                cell.value = cell_info['text']
                self.apply_styles(cell, cell_info['style'])

                self.sheet.merge_cells(
                    start_row=self.current_row,
                    end_row=self.current_row + cell_info['remaining'],
                    start_column=excel_col_idx,
                    end_column=excel_col_idx
                )
                del rowspan_map[(self.current_row, excel_col_idx)]
                excel_col_idx += 1

            for cell_data in row['cells']:
                colspan = int(cell_data.get('colspan', 1))
                rowspan = int(cell_data.get('rowspan', 1))
                cell_style = cell_data.get('style', "")
                cell = self.sheet.cell(row=self.current_row, column=excel_col_idx)
                cell.value = cell_data.get('text', "")


                if section == 'thead' and 'font-weight' not in cell_style:
                    if cell_style:
                        cell_style += "; "
                    cell_style += "font-weight: bold"

                merged_style = self.merge_styles(
                    self.table_data.get('table_style', ""),
                    section_style,
                    row_style,
                    cell_style
                )
                self.apply_styles(cell, merged_style)

                # Объединение по горизонтали
                if colspan > 1:
                    self.sheet.merge_cells(
                        start_row=self.current_row,
                        end_row=self.current_row,
                        start_column=excel_col_idx,
                        end_column=excel_col_idx + colspan - 1
                    )

                # Объединение по вертикали
                if rowspan > 1:
                    for offset in range(rowspan):
                        target_row = self.current_row + offset
                        for offset_col in range(colspan):
                            if offset != 0:
                                rowspan_map[(target_row, excel_col_idx + offset_col)] = {
                                    'text': cell_data.get('text', ""),
                                    'style': merged_style,
                                    'remaining': rowspan - offset - 1
                                }
                    self.sheet.merge_cells(
                        start_row=self.current_row,
                        end_row=self.current_row + rowspan - 1,
                        start_column=excel_col_idx,
                        end_column=excel_col_idx + colspan - 1
                    )

                excel_col_idx += colspan

            self.current_row += 1

    def convert(self, output_file='test.xlsx'):
        self.set_col_widths()
        self.add_styles_to_section('thead')
        self.add_styles_to_section('tbody')
        self.add_styles_to_section('tfoot')
        self.wb.save(output_file)