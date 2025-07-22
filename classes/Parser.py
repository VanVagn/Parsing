from html.parser import HTMLParser

class MyParser(HTMLParser):
    def __init__(self, target_class=None):
        super().__init__()
        self.target_class = target_class
        self.in_target_table = False
        self.current_section = None
        self.table_data = {
            'table_style': None,
            'thead': {'style': None, 'rows': []},
            'tbody': {'style': None, 'rows': []},
            'tfoot': {'style': None, 'rows': []},
            'colgroup': []
        }
        self.current_row = {'style': None, 'cells': []}
        self.current_cell_content = ""
        self.inside_cell = False
        self.cell_style = None
        self.cell_colspan = 1
        self.cell_rowspan = 1
        self.last_char_was_space = False  # Флаг для контроля пробелов
        self.in_colgroup = False

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        if tag == 'table':
            table_class = attrs_dict.get('class')
            if self.target_class is None or table_class == self.target_class:
                self.in_target_table = True
                self.table_data['table_style'] = attrs_dict.get('style', None)
        if not self.in_target_table:
            return

        if tag == 'colgroup':
            self.in_colgroup = True
        elif self.in_colgroup and tag == 'col':
            self.table_data['colgroup'].append({
                'style': attrs_dict.get('style', None)
            })
        elif tag in ('thead', 'tbody', 'tfoot'):
            self.current_section = tag
            self.table_data[tag]['style'] = attrs_dict.get('style', None)
        elif tag == 'tr':
            self.current_row = {'style': attrs_dict.get('style', None), 'cells': []}
        elif tag in ('td', 'th'):
            self.cell_style = attrs_dict.get('style', None)
            self.cell_colspan = int(attrs_dict.get('colspan', '1'))
            self.cell_rowspan = int(attrs_dict.get('rowspan', '1'))
            self.current_cell_content = ""
            self.inside_cell = True

        # Вставляем пробел для разделения там, где это логично
        if self.inside_cell and tag in ('br', 'p', 'div', 'span', 'li'):
            # Добавляем пробел, если предыдущий символ не пробел
            if not self.last_char_was_space and self.current_cell_content and not self.current_cell_content.endswith(' '):
                self.current_cell_content += ' '
                self.last_char_was_space = True

    def handle_data(self, data):
        if self.in_target_table and self.inside_cell:
            if data.strip():
                # Добавляем данные и сбрасываем флаг пробела
                self.current_cell_content += data
                self.last_char_was_space = False
            else:
                # Если данные пустые или пробелы, добавляем один пробел, если нет уже пробела в конце
                if not self.last_char_was_space and self.current_cell_content and not self.current_cell_content.endswith(' '):
                    self.current_cell_content += ' '
                    self.last_char_was_space = True

    def handle_endtag(self, tag):
        if not self.in_target_table:
            return

        # При закрытии тегов-обёрток (p, div, br, span, li) тоже вставляем пробел
        if self.inside_cell and tag in ('br', 'p', 'div', 'span', 'li'):
            if not self.last_char_was_space and self.current_cell_content and not self.current_cell_content.endswith(' '):
                self.current_cell_content += ' '
                self.last_char_was_space = True

        if tag in ('td', 'th'):
            cell = {
                'style': self.cell_style,
                'text': self.current_cell_content.strip(),
                'colspan': self.cell_colspan,
                'rowspan': self.cell_rowspan
            }
            self.current_row['cells'].append(cell)
            self.inside_cell = False
            self.last_char_was_space = False  # сбрасываем флаг после ячейки

        elif tag == 'tr':
            if self.current_section is None:
                self.current_section = "tbody"
            self.table_data[self.current_section]['rows'].append(self.current_row)
        elif tag == 'colgroup':
            self.in_colgroup = False
        elif tag in ('thead', 'tbody', 'tfoot'):
            self.current_section = None
        elif tag == 'table':
            self.in_target_table = False
