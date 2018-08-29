import os
import inspect

select = {0: "Все"}

def work_with_dir():
    folder_for_find_xlsx = os.path.dirname(os.path.abspath(inspect.stack()[0][1])) # полный путь к папке из которой выполняется файл
    list_files = [] # список файлов с полными путями
    for file in os.listdir(folder_for_find_xlsx): # во всех файлах папки
            if file.endswith(".xlsx"): # находим файл совпадающий с шаблоном
                list_files.append(os.path.join(folder_for_find_xlsx, file)) # добавляем файл, с полным путем, подходящий шаблону, в список
    return list_files

def list_dir():
    folder_for_find_xlsx = os.path.dirname(os.path.abspath(inspect.stack()[0][1])) # полный путь к папке из которой выполняется файл
    ii = 1
    for file in os.listdir(folder_for_find_xlsx): # во всех файлах папки
            if file.endswith(".xlsx"): # находим файл совпадающий с шаблоном
                select.update({ii:file})
                ii += 1
    return select


if __name__ == '__main__':
    print('Какие файлы обрабатывать?')
    for key, val in list_dir().items():
        print(key, ' ', val)
    print('Введите ответ. Можно выбрать несколько вариантов через пробел.')
    user_input = input()
    user_input = user_input.split(' ')
    false_int_user_input = []
    max_input = 0
    for int_user_input in user_input:
        try:
            int_user_input = int(int_user_input)
        except:
            false_int_user_input.append(int_user_input)

    user_input = list(set(user_input) - set(false_int_user_input))

    # print(user_input)
    print(max(user_input))
    # dd = 
    # print()
    if (max(user_input)) > (max(select.keys())):
        print('yo')
        

    # for int_user_input in user_input:
    #     try:
    #         if int_user_input > max_input:
    #             false_int_user_input.append(user_input.pop(user_input.index(int_user_input)))
    #     except:
    #         false_int_user_input.append(int_user_input)
    # print(false_int_user_input)
    # print(list(set(user_input) & set(false_int_user_input)))
    
    # print()