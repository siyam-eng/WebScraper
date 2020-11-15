from requests.api import request
from requests_html import HTMLSession
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


url = "https://www.terveys.gsk.fi/fi-fi/_adverse-effect-reportage"
WORD_LIST = ['review',]
session = HTMLSession()
r = session.get(url)
wb = load_workbook('webpages_inputdata.xlsx')


# Returns True if any of the given texts is found
def find_text(response, word_list):
    html = response.html.html
    found = None
    
    for word in word_list:
        result = html.lower().find(word.lower())
        if not result == -1:
            found = word
            break
    
    return found


# finds input labels and its text
def find_input_labels(response):
    html = response.html
    labels = html.find('label')
    for label in labels:
        print(label.text)


def find_inputs(response):
    html = response.html
    inputs = html.find('input')
    visible_inputs = [inp for inp in inputs if 'type' in inp.attrs and inp.attrs['type'] != 'hidden']
    return len(visible_inputs) > 1




# Prepares the excel sheets and names the columns
def customize_excel_sheet():
    global wb
    output = wb.create_sheet('Output') if 'Output' not in wb.sheetnames else wb['Output']
    
    font = Font(color="000000", bold=True)
    bg_color = PatternFill(fgColor='E8E8E8', fill_type='solid')

    # editing the output sheet
    output_column = zip(('A',  'B', 'C'), ('URL', 'Keyword Found', 'Can Input'))
    for col, value in output_column:
        cell = output[f'{col}1']
        cell.value = value
        cell.font = font
        cell.fill = bg_color
        output.freeze_panes = cell

        # fixing the column width
        output.column_dimensions[col].width = 20


# Generates the input links
def generate_input_urls():
    global wb
    inputs = wb['Input']
    for row in range(2, inputs.max_row + 1):
        # generates the links one by one
        if value := inputs[f"A{row}"].value:
            yield value


def get_data(response):
    can_input = find_inputs(response)
    keyword_found = find_text(response, WORD_LIST)
    