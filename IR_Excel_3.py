# import the main libraries
from colorama import init # for style
from colorama import Fore, Back, Style # for style
init()
from openpyxl.styles import Font, Color, colors
from openpyxl import load_workbook

def re_converted(file):
    '''The function converts the .xls file to .xlsx'''
    pass

def format_check(file_name):
    ''' Format check'''
    file = file_name + ".xlsx"

    active_excel = load_workbook(filename=file, data_only=True) #getting data
    active_sheet = active_excel.active #active sheet

    #check
    if active_sheet["B4"].value == "Номенклатура":
        if active_sheet["D4"].value == "ХарактеристикаНоменклатуры":
            if active_sheet["G4"].value == "Единица":
                if active_sheet["J4"].value == "НоменклатураТипИзделия":
                    if active_sheet["K4"].value == "Штрихкод":
                        if active_sheet["L4"].value == "ЦенаРозничная":
                            if active_sheet["M4"].value == "Количество":
                                return True


def rezult_check():
    '''Print rezult'''
    r = format_check(input("Введите название файла: "))
    if r == True:
        print("ФАЙЛ ПРИНЯТ И СООТВЕТСТВУЕТ ФОРМАТУ!")
    else:
        print("ФАЙЛ НЕ СООТВЕТСТВУЕТ ФОРМАТУ!")

rezult_check()

def write():
    '''

    '''
    pass
