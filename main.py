import requests

from classes.ConvertToExcel import HtmlTableToEcelConverter
from classes.Parser import MyParser

url = "https://ru.onlinemschool.com/math/formula/sine_table/"
path = "test_html.html"
response = requests.get(url)
# html = response.text
# table_class = "oms_mnt1"

# parser = MyParser(table_class)

with open('html/test_html.html', 'r', encoding='utf-8') as file:
    html = file.read()
parser = MyParser()

parser.feed(html)
converter = HtmlTableToEcelConverter(parser.table_data)
converter.convert("excelFiles/test.xlsx")
print(parser.table_data['table_style'])
# k = len(parser.table_data['tbody']['rows'])
# for i in range(k):
#     j = len(parser.table_data['tbody']['rows'][i]['cells'])
#     for m in range(j):
#         print(i+1, ' ', parser.table_data['tbody']['rows'][i]['cells'][m])


