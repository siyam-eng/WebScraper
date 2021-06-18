from bs4 import BeautifulSoup
import requests
import requests.exceptions
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import random

HEADERS_LIST = [
'Mozilla/5.0 (Windows; U; Windows NT 6.1; x64; fr; rv:1.9.2.13) Gecko/20101203 Firebird/3.6.13',
'Mozilla/5.0 (compatible, MSIE 11, Windows NT 6.3; Trident/7.0; rv:11.0) like Gecko',
'Mozilla/5.0 (Windows; U; Windows NT 6.1; rv:2.2) Gecko/20110201',
'Opera/9.80 (X11; Linux i686; Ubuntu/14.10) Presto/2.12.388 Version/12.16',
'Mozilla/5.0 (Windows NT 5.2; RW; rv:7.0a1) Gecko/20091211 SeaMonkey/9.23a1pre'
]


def find_code(soup, code):
    """Finds all codes matching with the given code"""
    codes = soup.findAll(text=re.compile(f'{code}'))
    codes2 = []
    def find_indexes(str, ch):
        indexes = []
        for i, ltr in enumerate(str):
            if ltr == ch:
                indexes.append(i)
        return indexes
    for c in codes:
        code2 = re.search(f'[A-Z-/]*{code}[A-Z0-9a-z-/]*', str(c))[0]
        code2 = code2.lstrip('-') if code2.startswith('-') else code2
        code2 = code2.lstrip('/') if code2.startswith('/') else code2
        if '/' in code2 or '-' in code2:
            # limiting the lenth of the code upto 3rd seperator
            if code2.count('-') > 4:
                third_hyphen = find_indexes(code2, '-')[4]
                code2 = code2[: third_hyphen]
            if code2.count('/') > 3:
                third_slash = find_indexes(code2, '/')[3]
                code2 = code2[: third_slash]
            if code2.count('/') > 2 and '-' in code2:
                code2 = code2[: code2.index('-')]


            codes2.append(code2)
    return codes2


def get_codes(url, soup, code_generator):
    for code in code_generator:
        # codes containing present code on page
        present_code = find_code(soup, code)
        for code in present_code:
            code_type = '/' if '/' in code else '-'
            yield {'url': url, 'code': code, 'type': code_type, 'length': len(code)}


def main():
    FILE_PATH = 'webpages.xlsx'
    NEW_URL_STARTING_ROW = 2

    wb = load_workbook(FILE_PATH)
    webpages = wb['Webpages']
    code_lookups = wb['Code_Lookups']
    codes = wb.create_sheet('Codes') if 'Codes' not in wb.sheetnames else wb['Codes']

    font = Font(color="000000", bold=True)
    bg_color = PatternFill(fgColor='E8E8E8', fill_type='solid')

    # editing the users sheet
    codes_columns = zip(('A',  'B', 'C', 'D'), ('Web Page', 'Code Type', 'Code', 'Lenth'))
    for col, value in codes_columns:
        cell = codes[f'{col}1']
        cell.value = value
        cell.font = font
        cell.fill = bg_color
        codes.freeze_panes = cell

        # fixing the column width
        codes.column_dimensions[col].width = 20

    def webpage_urls_generator():
        start = NEW_URL_STARTING_ROW
        for row in range(start, webpages.max_row + 1):
            cell = webpages[f'A{row}']
            if cell.value:
                yield cell.value 

    def codes_lookups_generator():
        start = 2
        for row in range(start, code_lookups.max_row + 1):
            cell = code_lookups[f'A{row}']
            if cell.value:
                yield cell.value 

    for webpage in webpage_urls_generator():
        header = {'User-Agent': random.choice(HEADERS_LIST), 'X-Requested-With': 'XMLHttpRequest'}
        response = requests.get(webpage, headers=header)
        if not response.ok:
            response = requests.get(webpage)
        soup = BeautifulSoup(response.text,"html.parser")

        code_generator = codes_lookups_generator()
        for code in get_codes(webpage, soup, code_generator):
            print(code)
            codes.append((
                code['url'],
                code['code'],
                code['type'],
                code['length'],
            ))
    wb.save(FILE_PATH)


if __name__ == '__main__':
    main()

