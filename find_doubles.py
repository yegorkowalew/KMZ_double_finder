import openpyxl
from copy import copy

from datetime import datetime
import time
import os
import inspect

# Условия правильной работы скрипта:
    # Обрабатываемые поля:
    # Уз. - узел состящий из деталей
    #     Столбец: "п/п" - Значение означающее что это узел "Уз."
    #     Столбец: "Наименование" тип узла (указывать не обязательно)
    #     Столбец: "Позиция" название узла (указывать обязательно)
    #     Столбец: "Кол-во узел" (указывать не обязательно)
    #     Столбец: "Кол-во изделие" (указывать не обязательно)
    #     Столбец: "Кол-во заказ" (указывать не обязательно)
    #     Столбец: "Кол-во НЗП" (указывать не обязательно)
    #     Столбец: "Массив" (указывать не обязательно)
    #     Столбец: "Размеры чист." Описание узла (указывать не обязательно, но желательно)
    #     Столбец: "Размеры заготовки" (указывать не обязательно)
    #     Столбец: "Операции" (указывать не обязательно)
    #     Столбец: "Розцеховка" (указывать не обязательно)

    # Деталь - родитель Уз.
    #     Столбец: "п/п" - Номер по порядку детали (Указывать не обязательно)
    #     Столбец: "Наименование" тип узла (указывать обязательно)
    #     Столбец: "Позиция" название детали (указывать обязательно)
    #     Столбец: "Кол-во узел" (указывать не обязательно)
    #     Столбец: "Кол-во изделие" (указывать не обязательно)
    #     Столбец: "Кол-во заказ" по этому полю будут подсчитываться суммы дубликатов (указывать обязательно)
    #     Столбец: "Кол-во НЗП" (указывать не обязательно)
    #     Столбец: "Массив" (указывать не обязательно)
    #     Столбец: "Размеры чист." (указывать не обязательно)
    #     Столбец: "Размеры заготовки" (указывать не обязательно)
    #     Столбец: "Операции" (указывать не обязательно)
    #     Столбец: "Розцеховка" таблица будет разбита на листы в которых каждый цех по отдельности (указывать обязательно)

# === SETTINGS ===
# parse_file = 'c:\\Users\\i.kovalenko\\parser\\find_doubles\\140.xlsx' # путь к файлу выгруженному с одноэса
first_row = 6
end_coll = 0
second_row = 0
after_column = 8

end_names = {
        'npp': [1, "п/п",], 
        'types': [2, "Наименование",], 
        'name': [3, "Позиция",], 
        'knot': [4, "узел",], 
        'product': [5, "изделие",], 
        'order': [6, "заказ",], 
        'nzp': [7, "НЗП",], 
        'massiv': [8, "Массив",], 
        'size_clean': [9, "Размеры чист.",], 
        'size_workpiece': [10, "Размеры заготовки",], 
        'operations': [11, "Операции",], 
        'shop': [12, "Розцеховка",], 
        }



def if_num(cell):
    try:
        eval(cell)
        return True
    except:
        return False

class Unit:
    def __init__(self, types, name, knot, product, order, nzp, massiv, size_clean, size_workpiece, operations, shop, row):
        """Constructor for Unit"""
        self.types = types # Столбец "Наименование" (не обязательный)
        self.name = name # Столбец "Позиция" (обязательный)
        self.knot = knot # Столбец "Узел" (не обязательный)
        self.product = product # Столбец "Изделие" (не обязательный)
        self.order = order # Столбец "Заказ" (обязательный)
        self.nzp = nzp # Столбец: "Кол-во НЗП" (указывать не обязательно)
        self.massiv = massiv # Столбец: "Массив" (указывать не обязательно)
        self.size_clean = size_clean # Столбец: "Размеры чист." Описание узла (указывать не обязательно, но желательно)
        self.size_workpiece = size_workpiece # Столбец: "Размеры заготовки" (указывать не обязательно)
        self.operations = operations # Столбец: "Операции" (указывать не обязательно)
        self.shop = shop # Столбец: "Розцеховка" (указывать не обязательно)
        self.row = row # поле для хранения номера строки

class Detail:
    def __init__(self, npp, types, name, knot, product, order, nzp, massiv, size_clean, size_workpiece, operations, shop, doubles, row):
        """Constructor for Detail"""
        self.npp = npp # "п/п" - Номер по порядку детали (Указывать не обязательно, желательно самому пронумеровать)
        self.types = types # Столбец: "Наименование" тип узла (указывать обязательно)
        self.name = name # Столбец "Позиция" (обязательный)
        self.knot = knot # Столбец "Узел" (не обязательный)
        self.product = product # Столбец "Изделие" (не обязательный)
        # self.order = order # Столбец: "Кол-во заказ" по этому полю будут подсчитываться суммы дубликатов (указывать обязательно)
        if if_num(order):
            self.order = eval(order)
        else:
            self.order = int(order)
        self.nzp = nzp # Столбец: "Кол-во НЗП" (указывать не обязательно)
        self.massiv = massiv # Столбец: "Массив" (указывать не обязательно)
        self.size_clean = size_clean # Столбец: "Размеры чист." (указывать не обязательно)
        self.size_workpiece = size_workpiece # Столбец: "Размеры заготовки" (указывать не обязательно)
        self.operations = operations # Столбец: "Операции" (указывать не обязательно)
        # self.shop = shop # Столбец: "Розцеховка" таблица будет разбита на листы в которых каждый цех по отдельности (указывать обязательно)
        self.shop = shop.split("-") # разбиваем строку вида 104-101, на список [104, 101]
        self.doubles = doubles # список дубликатов 
        self.row = row # поле для хранения номера строки

def work_with_dir():
    folder_for_find_xlsx = os.path.dirname(os.path.abspath(inspect.stack()[0][1])) # полный путь к папке из которой выполняется файл
    list_files = [] # список файлов с полными путями
    for file in os.listdir(folder_for_find_xlsx): # во всех файлах папки
            if file.endswith(".xlsx"): # находим файл совпадающий с шаблоном
                list_files.append(os.path.join(folder_for_find_xlsx, file)) # добавляем файл, с полным путем, подходящий шаблону, в список
    return list_files

def file_to_wb(file_name):
    return openpyxl.load_workbook(filename = file_name)

if __name__ == '__main__':
    start_time = time.process_time() # Засекаем время
    xls_file = work_with_dir()[0] # первый файл в найденых файлах
    print('Работаю с файлом: ' + xls_file)

    work_wb = file_to_wb(xls_file) # из файла открываем книгу
    sheet = work_wb[work_wb.sheetnames[0]] # берем первый лист в файле
    new_sheet = work_wb.create_sheet("Посчитано") # создаем новый лист в файле

    units = []
    details = {}

    for row in range(first_row, sheet.max_row): # нашел последнюю строку в таблице
        if sheet.cell(row=row, column=end_names.get('name')[0]).value:
            second_row = row

    for col in range(1, sheet.max_column): # нашел последний столбец в таблице
        end_coll = col

    for row in range(first_row, second_row):
        if sheet.cell(row=row, column=1).value == "Уз.":
            npp = 0
            unit = Unit(
            sheet.cell(row=row, column=end_names.get('types')[0]).value,
            sheet.cell(row=row, column=end_names.get('name')[0]).value,
            sheet.cell(row=row, column=end_names.get('knot')[0]).value,
            sheet.cell(row=row, column=end_names.get('product')[0]).value,
            sheet.cell(row=row, column=end_names.get('order')[0]).value,
            sheet.cell(row=row, column=end_names.get('nzp')[0]).value,
            sheet.cell(row=row, column=end_names.get('massiv')[0]).value,
            sheet.cell(row=row, column=end_names.get('size_clean')[0]).value,
            sheet.cell(row=row, column=end_names.get('size_workpiece')[0]).value,
            sheet.cell(row=row, column=end_names.get('operations')[0]).value,
            sheet.cell(row=row, column=end_names.get('shop')[0]).value,
            row,)
            units.append(unit)
        else:
            npp +=1
            unit = Detail(
            npp,
            sheet.cell(row=row, column=end_names.get('types')[0]).value,
            sheet.cell(row=row, column=end_names.get('name')[0]).value,
            sheet.cell(row=row, column=end_names.get('knot')[0]).value,
            sheet.cell(row=row, column=end_names.get('product')[0]).value,
            sheet.cell(row=row, column=end_names.get('order')[0]).value,
            sheet.cell(row=row, column=end_names.get('nzp')[0]).value,
            sheet.cell(row=row, column=end_names.get('massiv')[0]).value,
            sheet.cell(row=row, column=end_names.get('size_clean')[0]).value,
            sheet.cell(row=row, column=end_names.get('size_workpiece')[0]).value,
            sheet.cell(row=row, column=end_names.get('operations')[0]).value,
            sheet.cell(row=row, column=end_names.get('shop')[0]).value,
            [],
            row,)
            details.update({row:unit})

    for detail_for_add in details.values(): # нашел дубликаты и дописал их в поле .doubles
        for detail in details.values(): 
            if detail_for_add.name == detail.name:
                detail_for_add.doubles.append(detail.row)

    after_end_names = {}
    colum = 0
    for key, value in end_names.items():
        if value[0] < after_column+1:
            after_end_names.update({key:value})
        else: 
            if value[0] < after_column+len(units):
                colum = value[0]
                for i in units:
                    # print(i)
                    colum += 1
                    after_end_names.update({i.row:[colum, i.name]})
        
    # print(len(units))
    # print(after_end_names)
    # for key, value in after_end_names.items():
    #     print(value[0])
    # for detail_for_add in details:
        # print(detail_for_add.shop)

    # new_sheet
    # print(end_names['types'])
    # for key, value in end_names.items():
    #     strr = 'name'
    #     try:
    #         print(getattr(details[3], key))
    #     except:
    #         print('no')

    added_rows = []
    row_step = first_row
    for row, detail in details.items():
        if row in added_rows:
            pass
        else:
            added_rows += detail.doubles
            sum_details = 0
            for i in detail.doubles:
                sum_details += sum_details + details.get(i).order
            new_sheet.cell(row=row_step, column=end_names.get('types')[0]).value = detail.types
            new_sheet.cell(row=row_step, column=end_names.get('name')[0]).value = detail.name
            new_sheet.cell(row=row_step, column=end_names.get('knot')[0]).value = detail.knot
            new_sheet.cell(row=row_step, column=end_names.get('product')[0]).value = detail.product
            new_sheet.cell(row=row_step, column=end_names.get('order')[0]).value = sum_details
            new_sheet.cell(row=row_step, column=end_names.get('nzp')[0]).value = detail.nzp
            new_sheet.cell(row=row_step, column=end_names.get('massiv')[0]).value = detail.massiv
            new_sheet.cell(row=row_step, column=end_names.get('size_clean')[0]).value = detail.size_clean
            new_sheet.cell(row=row_step, column=end_names.get('size_workpiece')[0]).value = detail.size_workpiece
            new_sheet.cell(row=row_step, column=end_names.get('operations')[0]).value = detail.operations
            new_sheet.cell(row=row_step, column=end_names.get('shop')[0]).value = "-".join(detail.shop)
            row_step += 1
            # print(sum_details)
            # print('----------')

    # for row in range(1, len(details)):
        # print(len(details))
        # detail = details[row-1]
        # print(detail.doubles)
        # for key, value in end_names.items():
            # new_sheet.cell(row=row, column=value[0]).value = str(getattr(detail, key))

    work_wb.save(xls_file)
    
    # from myfolder.myfile import myfunc

    print("Завершил работу с файлом: %s. За %s секунд. Время: %s"
        %(
            xls_file,
            round(time.process_time() - start_time, 3),
            datetime.strftime(datetime.now(), "%H:%M:%S"),
            )
        )