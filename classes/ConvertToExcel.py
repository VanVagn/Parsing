from openpyxl import workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.styles.builtins import total
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.workbook import Workbook
import re
import webcolors

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

        self.apply_alignment(cell, style_dict)
        self.apply_font(cell, style_dict)
        self.apply_background(cell, style_dict)
        self.apply_border(cell, style_dict)

    # Выравнивание
    def apply_alignment(self, cell, style_dict):
        horizontal = style_dict.get('text-align')
        vertical = style_dict.get('vertical-align')

        if vertical == 'middle':
            vertical = 'center'

        alignment = Alignment(
            horizontal=horizontal or 'general',
            vertical=vertical or 'bottom',
            wrap_text=True
        )
        cell.alignment = alignment

    # Стиль текста
    def apply_font(self, cell, style_dict):
        bold = style_dict.get('font-weight', None)
        italic = style_dict.get('font-style', None)
        underline = style_dict.get('text-decoration', None)

        # Переделываем под формат openpyxl
        if underline:
            underline = 'single' if 'underline' in underline else None

        # Цвет текста
        color = style_dict.get('color') or ""
        color_val = color.strip().lower()
        try:
            if color_val.startswith('#'):
                hex_color = color_val.lstrip('#')
            else:
                hex_color = webcolors.name_to_hex(color_val).lstrip('#')
        except ValueError:
            hex_color = None
        font_color = None
        if hex_color and re.fullmatch(r'[0-9a-fA-F]{6}', hex_color):
            font_color = 'FF' + hex_color.upper()

        # Собираем новый Font
        old_font = cell.font or Font()
        cell.font = Font(
            name=old_font.name,
            size=old_font.size,
            bold=bold if bold is not None else old_font.bold,
            italic=italic if italic is not None else old_font.italic,
            underline=underline if underline is not None else old_font.underline,
            color=font_color if font_color else old_font.color
        )


    # Фон
    def apply_background(self, cell, style_dict):

        if 'background-color' in style_dict:
            raw_color = style_dict['background-color'].strip().lower()
            try:
                if raw_color.startswith('#'):
                    color = raw_color.lstrip('#')
                else:
                    color = webcolors.name_to_hex(raw_color).lstrip('#')
            except ValueError:
                color = 'FFFFFF'  # fallback: white
            end_color = self.expand_short_hex(color)
            if re.fullmatch(r'[0-9a-fA-F]{6}', end_color):
                excel_color = 'FF' + end_color.upper()
                cell.fill = PatternFill(start_color=excel_color, patternType="solid")

    # Границы

    def apply_border(self, cell, style_dict):
        sheet = self.sheet

        def parse_border_part(border_value):

            parts = border_value.strip().split()
            if len(parts) < 3:
                return None

            # Толщина
            width_str, style_str, color_str = parts[:3]
            try:
                width_px = int(width_str.replace('px', ''))
            except ValueError:
                width_px = 1

            if width_px <= 1:
                border_style = 'thin'
            elif width_px == 2:
                border_style = 'medium'
            else:
                border_style = 'thick'

            # Стиль линии
            line_style_map = {
                'solid': border_style,
                'dashed': 'dashed',
                'dotted': 'dotted',
                'double': 'double',
            }
            border_style_final = line_style_map.get(style_str.lower(), 'thin')

            # Цвет
            try:
                if color_str.startswith('#'):
                    hex_color = color_str.lstrip('#')
                else:
                    hex_color = webcolors.name_to_hex(color_str).lstrip('#')
            except ValueError:
                hex_color = '000000'  # default black

            # Расширяем короткий hex
            if len(hex_color) == 3:
                hex_color = ''.join([c * 2 for c in hex_color])

            color = 'FF' + hex_color.upper()
            return Side(border_style=border_style_final, color=color)

        # Начальные стороны
        borders = {'left': None, 'right': None, 'top': None, 'bottom': None}

        # Общий border
        if 'border' in style_dict:
            common = parse_border_part(style_dict['border'])
            for side in borders:
                borders[side] = common

        # Переопределения
        for css_key, side_name in {
            'border-left': 'left',
            'border-right': 'right',
            'border-top': 'top',
            'border-bottom': 'bottom'
        }.items():
            if css_key in style_dict:
                specific = parse_border_part(style_dict[css_key])
                borders[side_name] = specific

        # Применение к объединённым или обычным ячейкам
        merged_ranges = sheet.merged_cells.ranges
        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            if min_row <= cell.row <= max_row and min_col <= cell.column <= max_col:
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        target_cell = sheet.cell(row=r, column=c)
                        target_cell.border = Border(
                            left=borders['left'] if c == min_col else None,
                            right=borders['right'] if c == max_col else None,
                            top=borders['top'] if r == min_row else None,
                            bottom=borders['bottom'] if r == max_row else None
                        )
                return

        # Если обычная ячейка
        cell.border = Border(
            left=borders['left'],
            right=borders['right'],
            top=borders['top'],
            bottom=borders['bottom']
        )

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
        rowspan_tracker = {}

        for row in section_data.get('rows', []):
            excel_col_idx = 1

            while (self.current_row, excel_col_idx) in rowspan_tracker:
                rowspan_tracker[(self.current_row, excel_col_idx)] -= 1
                if rowspan_tracker[(self.current_row, excel_col_idx)] == 0:
                    del rowspan_tracker[(self.current_row, excel_col_idx)]
                excel_col_idx += 1

            row_style = row.get('style', None)
            for cell_data in row['cells']:
                colspan = int(cell_data.get('colspan', 1))
                rowspan = int(cell_data.get('rowspan', 1))
                cell = self.sheet.cell(row=self.current_row, column=excel_col_idx)
                cell_style = cell_data.get('style', "")
                cell.value = cell_data.get('text', "")

                if colspan > 1:
                    self.sheet.merge_cells(
                        start_row=self.current_row,
                        start_column=excel_col_idx,
                        end_row=self.current_row,
                        end_column=excel_col_idx + colspan - 1
                    )

                if rowspan > 1:
                    self.sheet.merge_cells(
                        start_row=self.current_row,
                        start_column=excel_col_idx,
                        end_row=self.current_row + rowspan - 1,
                        end_column=excel_col_idx + colspan - 1
                    )

                    for r in range(1, rowspan):
                        for c in range(colspan):
                            key = (self.current_row + r, excel_col_idx + c)
                            rowspan_tracker[key] = rowspan_tracker.get(key, 0) + 1



        # Жирный шрифт по умолчанию
                if section == 'thead':
                    if 'font-weight' not in cell_style:
                        if cell_style:
                            cell_style += "; "
                        cell_style += "font-weight: bold"

                merged_style = self.merge_styles(
                    self.table_data.get('table_style', ""),
                    section_style,
                    row_style,
                    cell_style
                )

                cell = self.get_master_cell(cell)
                self.apply_styles(cell, merged_style)
                excel_col_idx += colspan


            self.current_row += 1

    def get_master_cell(self, cell):
        for merged_range in self.sheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            if min_row <= cell.row <= max_row and min_col <= cell.column <= max_col:
                return self.sheet.cell(row=min_row, column=min_col)
        return cell


    def convert(self, output_file='test.xlsx'):
        self.set_col_widths()
        self.add_styles_to_section('thead')
        self.add_styles_to_section('tbody')
        self.add_styles_to_section('tfoot')
        self.wb.save(output_file)


        # сделать учет автовысоты по контенту, чтобы текст нормально отображался. параметр fid
        # настройка границ ячейки: цвет, толщина обрамления, стороны(все, одна) - все кроме отдельных границ ячеек
        # библиотека конвертации цветов: red -> #f44336 - Сделано
        # поддержка форматирования текста внутри ячейки "<b>Этот</b> текст" чистка inline тегов(b,i) - Есть библиотека xlsxwriter,  но невозможно работать одновременно со стилями и в ней, и в open pyxl
        # добавление скриптов VBA функция allert showMessage Библиотеки ес1ть, но только на винду