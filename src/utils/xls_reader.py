import openpyxl as pyxl


class xlsReader():
    workbook = pyxl.load_workbook('Elaw.xlsx')
    sheet = workbook.active
