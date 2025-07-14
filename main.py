import requests

from classes.ConvertToExcel import HtmlTableToExcelConverter
from classes.Parser import MyParser

#url = "https://ru.onlinemschool.com/math/formula/sine_table/"
path = "test3.html"
#response = requests.get(url)


with open('html/test3.html', 'r', encoding='utf-8') as file:
    html = file.read()
parser = MyParser()

parser.feed(html)
converter = HtmlTableToExcelConverter(parser.table_data)
converter.convert("excelFiles/test.xlsx")



