# import the main libraries
from colorama import init # for style
from colorama import Fore, Back, Style # for style
init()
from openpyxl import load_workbook
import pyexcel as p

def re_converted(file):
    '''The function save the .xls file to .xlsx'''
    print(str(file) + ".xls ФАЙЛ ПРИНЯТ ДЛЯ КОНВЕРТАЦИИ!")
    file_xls = file + ".xls"
    file_xlsx = file + ".xlsx"
    p.save_book_as(file_name=file_xls, dest_file_name=file_xlsx)
    print(file_xls + " СКОНВЕРТИРОВАН В .xlsx ")
    return file_xlsx

def format_check(file_name):
    ''' Format check'''
    #передаем файл для конвертации
    file_xlsx = re_converted(file_name)

    active_excel = load_workbook(filename=file_xlsx, data_only=True)  # getting data
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
    name = input("Введите название файла: ")
    name_xlsx = name + ".xlsx"
    r = format_check(name)

    if r == True:
        print("ФАЙЛ ПРИНЯТ И СООТВЕТСТВУЕТ ФОРМАТУ!")

        active_excel = load_workbook(filename=name_xlsx, data_only=True)  # getting data
        active_sheet = active_excel.active  # active sheet

        #max row
        max_row = active_sheet.max_row

        #take rows and append date
        tip_izdeliya = [] # J
        nomenklatura = [] # B
        harakteristika = [] # D
        shtrihkod = [] # K
        edinitca = [] # G
        kolichestvo = [] # M
        chena_postypleniya = [] # O
        nds = [20] #const
        chena_roznichnaya = [] # L
        spisok_all = [tip_izdeliya,tip_izdeliya,nomenklatura,harakteristika,shtrihkod,edinitca,kolichestvo,chena_postypleniya,nds,chena_roznichnaya]

        for i in range(5,(max_row - 1)):
            j = active_sheet["J" + str(i)].value
            tip_izdeliya.append(j)
            b = active_sheet["B" + str(i)].value
            nomenklatura.append(b)
            d = active_sheet["D" + str(i)].value
            harakteristika.append(d)
            k = active_sheet["K" + str(i)].value
            shtrihkod.append(int(k))
            g = active_sheet["G" + str(i)].value
            edinitca.append(g)
            m = active_sheet["M" + str(i)].value
            kolichestvo.append(m)
            o = active_sheet["O" + str(i)].value
            chena_postypleniya.append(o)
            l = active_sheet["L" + str(i)].value
            chena_roznichnaya.append(l)
            const = 20
            nds.append(const)

        #write lists
        filename = "ШаблонЗагрузки.xlsx"
        active_excel_1 = load_workbook(filename=filename, data_only=True)  # getting data
        active_sheet_1 = active_excel_1.active  # active sheet

        #max row in active_excel_1
        max_row_1 = active_sheet_1.max_row

        for g in range(len(tip_izdeliya)):
            for h in range(len(spisok_all)):
                _=active_sheet_1.cell(column=h+1, row=max_row_1+g+1, value=spisok_all[h][g])

        #max row before write for format
        max_row_2 = active_sheet_1.max_row
        for s in range(1, max_row_2 + 1):
            active_sheet_1["E" + str(s)].number_format = '#,##0'
            active_sheet_1["H" + str(s)].number_format='#,##0.00'
            active_sheet_1["J" + str(s)].number_format='#,##0.00'

        try:
            active_excel_1.save("ШаблонЗагрузки.xlsx")
        except PermissionError:
            print(Fore.RED)
            print("ЗАКРОЙТЕ ШАБЛОН ЗАГРУЗКИ!")

        print(Fore.BLUE)
        print("ДАННЫЕ ЗАПИСАНЫ!")

    else:
        print(Fore.RED)
        print("ФАЙЛ НЕ СООТВЕТСТВУЕТ ФОРМАТУ!")

rezult_check()


