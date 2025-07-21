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
            'colgroup' : []
        }
        self.current_row = {
            'style': None,
            'cells': []
        }
        self.current_cell_content = []
        self.inside_cell = False
        self.in_colgroup = False
        self.cell_style = None
        self.cell_colspan = 1
        self.cell_rowspan = 1

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)

        if tag == 'table':
            table_class = attrs_dict.get('class')
            if self.target_class is None or table_class == self.target_class:
                self.in_target_table = True
                style = attrs_dict.get('style', '')
                if 'border' in attrs_dict:
                    border_value = attrs_dict['border']
                    if border_value == '1':
                        style += '; border: 1px solid black'

                self.table_data['table_style'] = style if style else None
        if not self.in_target_table:
            return

        # todo посмотреть приоритет, возможно надо переделать
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
            self.current_row = {
                'style': attrs_dict.get('style', None),
                'cells': []
            }
        elif tag in ('td', 'th'):
            self.cell_style = attrs_dict.get('style', None)
            self.cell_colspan = int(attrs_dict.get('colspan', '1'))
            self.cell_rowspan = int(attrs_dict.get('rowspan', '1'))
            self.current_cell_content = ""
            self.inside_cell = True


    def handle_data(self, data):
        if self.in_target_table and self.inside_cell:
            self.current_cell_content += data.strip()

    def handle_endtag(self, tag):
        if not self.in_target_table:
            return

        if tag in ('td', 'th'):
            cell = {
                'style': self.cell_style,
                'text': self.current_cell_content,
                'colspan': self.cell_colspan,
                'rowspan': self.cell_rowspan

            }
            if tag == 'th':
                if cell['style']:
                    cell['style'] += "; "
                cell['style'] += "font-weight: bold"
            self.current_row['cells'].append(cell)
            self.inside_cell = False
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








