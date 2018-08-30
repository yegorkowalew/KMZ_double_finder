import os,sys
import inspect
import time

# select = {0: "Все"}
# 
def work_with_dir():
    folder_for_find_xlsx = os.path.dirname(os.path.abspath(inspect.stack()[0][1])) # полный путь к папке из которой выполняется файл
    list_files = [] # список файлов с полными путями
    for file in os.listdir(folder_for_find_xlsx): # во всех файлах папки
            if file.endswith(".xlsx"): # находим файл совпадающий с шаблоном
                list_files.append(os.path.join(folder_for_find_xlsx, file)) # добавляем файл, с полным путем, подходящий шаблону, в список
    return list_files

def list_dir():
    # print(os.path.dirname(os.path.realpath(sys.argv[0])))
    # folder_for_find_xlsx = os.path.dirname(os.path.abspath(inspect.stack()[0][1])) # полный путь к папке из которой выполняется файл
    folder_for_find_xlsx = os.path.dirname(os.path.realpath(sys.argv[0])) # полный путь к папке из которой выполняется файл
    list_dir = [] # список файлов с полными путями
    for file in os.listdir(folder_for_find_xlsx): # во всех файлах папки
            if file.endswith(".xlsx"): # находим файл совпадающий с шаблоном
                list_dir.append(file) # добавляем файл, с полным путем, подходящий шаблону, в список
    return list_dir

def check_input(files_lst, user_input):
    files_list = set()
    failed = set()
    user_input = user_input.split(' ')
    for val in user_input:
        if val == '':
            return files_lst, []
        try:
            num = int(val)
            if num == 0:
                for f in files_lst:
                    files_list.add(f)
            elif num < 0 or num > len(files_lst):
                failed.add(val)
            else:
                files_list.add(files_lst[num - 1])
        except:
            failed.add(val)
    return list(files_list), list(failed)

def user_input_work():
    # print(os.path.realpath(__file__))
    # print(os.path.basename(os.path.realpath(sys.argv[0])))
    
    # print(os.path.dirname(os.path.abspath(__file__)))
    print('Какие файлы обрабатывать?')
    i = 1
    for val in list_dir():
        print(i, ' ', val)
        i += 1
    print('Введите ответ. Можно выбрать несколько вариантов через пробел.')
    user_input = input()
    for_work = check_input(list_dir(), user_input)
    if for_work[1] == []:
        print('Обрабатываю файлы:\n','\n '.join(for_work[0]))
        # print(os.path.dirname(os.path.realpath(sys.argv[0])))
        # new_list = 
        # print(new_list)
        return [os.path.dirname(os.path.realpath(sys.argv[0]))+'\\'+x for x in for_work[0]]
    else:
        if for_work[0] == []:
            print('Обрабатывать нечего.\n Неверный ввод: ', ', '.join(for_work[1]))
            print('Можно закрыть окно.')
            time.sleep(5)
            raise SystemExit
        else:
            print('Обрабатываю файлы:\n','\n '.join(for_work[0]))
            print('Неверный ввод: ', ', '.join(for_work[1]))
            return for_work[0]

if __name__ == '__main__':
    user_input_work()
    print('-----------')
    time.sleep(50)