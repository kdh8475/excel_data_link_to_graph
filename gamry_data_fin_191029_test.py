from openpyxl import load_workbook
import os
import gamry_parser as parser
import xlwings as xw

path_data = "./files/191016/"
path_format = "./files/PCFC template (2019)_DHK_v1_.xlsx"
path_file = "./files/111.xlsx"

file_list = os.listdir(path_data)
print(file_list)

book = xw.Book(path_format)
app = xw.apps.active

ws_iv = book.sheets['j-V-P']
ws_eis = book.sheets['EIS']

for files in file_list:
#j-V-P(iv)
    if files.find('650') is not -1 and files.find('LSV') is not -1 and files.find('OCP') is -1:
                data = 'I4'
                print(files)
                file_name = path_data + files
                ca = parser.ChronoAmperometry(to_timestamp=True)
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['T']
                ws_iv.range(data).options(index=False).value = dta
    if files.find('600') is not -1 and files.find('LSV') is not -1 and files.find('OCP') is -1:
                data = 'P4'
                print(files)
                file_name = path_data + files
                ca = parser.ChronoAmperometry(to_timestamp=True)
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['T']
                ws_iv.range(data).options(index=False).value = dta
    if files.find('550') is not -1 and files.find('LSV') is not -1 and files.find('OCP') is -1:
                data = 'W4'
                print(files)
                file_name = path_data + files
                ca = parser.ChronoAmperometry(to_timestamp=True)
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['T']
                ws_iv.range(data).options(index=False).value = dta
    if files.find('500') is not -1 and files.find('LSV') is not -1 and files.find('OCP') is -1:
                data = 'AD4'
                print(files)
                file_name = path_data + files
                ca = parser.ChronoAmperometry(to_timestamp=True)
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['T']
                ws_iv.range(data).options(index=False).value = dta
# EIS(eis)
    if files.find('650') is not -1 and files.find('EIS') is not -1:
            if files.find('OCV') is not -1:
                data = 'D5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('750mV') is not -1:
                data = 'N5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('550mV') is not -1:
                data = 'X5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta

    if files.find('600') is not -1 and files.find('EIS') is not -1:
            if files.find('OCV') is not -1:
                data = 'AH5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('750mV') is not -1:
                data = 'AR5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('550mV') is not -1:
                data = 'BB5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta

    if files.find('550') is not -1 and files.find('EIS') is not -1:
            if files.find('OCV') is not -1:
                data = 'BL5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('750mV') is not -1:
                data = 'BV5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('550mV') is not -1:
                data = 'CF5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta

    if files.find('500') is not -1 and files.find('EIS') is not -1:
            if files.find('OCV') is not -1:
                data = 'CP5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('750mV') is not -1:
                data = 'CZ5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta
            if files.find('550mV') is not -1:
                data = 'DJ5'
                print(files)
                file_name = path_data + files
                ca = parser.Impedance()
                ca.load(filename=file_name)
                dta = ca.get_curve_data()
                del dta['Freq'], dta['Zmod'], dta['Zphz']
                ws_eis.range(data).options(index=False).value = dta

book.sheets('j-V-P graph').activate()
book.save(path_file)
app.quit()