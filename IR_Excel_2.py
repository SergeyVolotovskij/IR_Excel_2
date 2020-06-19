#импортируем необходимые библиотеки
from colorama import init #для стиля ветового
from colorama import Fore, Back, Style #для стиля ветового
init()#для стиля цветового
from openpyxl import workbook
from openpyxl.styles import Font, Color, colors
from openpyxl import load_workbook
import openpyxl

#для удобства вводим название файла
filename = "Список.xlsx"

#вытянули данные с документа
active_excel = load_workbook(filename=filename,data_only=True)#data_only=True

#делаем или смотрим активный лист
active_sheet = active_excel.active

#нужно понять максимальный размер данных на листе
max_row = active_sheet.max_row
max_column = active_sheet.max_column

#выведем информацию
print(Fore.YELLOW)
print("КОЛИЧЕСТВО СТРОК: " + str(max_row))
print("КОЛИЧЕСТВО КОЛОНОК: " + str(max_column))

#делаем формы ввода (сделать в виде форм)
print(Fore.GREEN)
a = input("ВВЕДИТЕ ГРУППУ: ")
b = input("ВВЕДИТЕ ЧАСТЬ ИМЕНИ: ")

#забираем наименование номенклатуры
name = []
for i in range(2,(max_row + 1)):
    e = active_sheet["A" + str(i)].value
    name.append(e)

#поиск по группе (CD, MD...)
# name_2 = []
e = 0
for q in name:
    f = name[e][:]
    if a and b in f:
        print(str(e) + ' ' + f)
    # name_2.append(f)
    e += 1
# print(name_2)


# print(name)

#забирае характеристику номенклатуры
characteristic = []
for i in range(2,(max_row + 1)):
    e = active_sheet["B" + str(i)].value
    characteristic.append(e)
# print(characteristic)

#забираем ШК
barcode = []
for i in range(2,(max_row + 1)):
    e = active_sheet["E" + str(i)].value
    barcode.append(e)
# print(barcode)

#забираем единицы измерения
unit = []
for i in range(2,(max_row + 1)):
    e = active_sheet["C" + str(i)].value
    unit.append(e)
# print(unit)

# #осуществляем поиск
# if a in name_2:
#     print(Fore.WHITE)
#     print("КОЛИЧЕСТВО ВХОЖДЕНИЙ = " + str(name_2.count(a)))
#     # for a in name_2:


