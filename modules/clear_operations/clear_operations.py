from settings import operations_names, file_name

def say_hallo(cell, operations_names):
    print('hallo', ' ', cell)

test_cell = "Тестовая, ячейка служ, зап"

def clear_name(cell_for_clear):
    """
    Разбиваю ячейку на ',' эти ячейки разбиваю на ' ' в них ищу похожее слово из file_name
    """
    all_words = []
    all_words = cell_for_clear.split(',')
    for words in all_words:
        for word in words.split(' '):
            for key, value in file_name.items():
                for ss in value:
                    if ss == word.lower():
                        return key
    return False


if __name__ == '__main__':
    """
    Обработка файла состоит из таких частей:
    - Обработка Названия (первые несколько строк таблицы с полями:
        - Служебная №:
        - Заказчик:
        - Заказ №:
        - Отгрузка:
    ), Надо их правильно обозвать, правильно расположить.
    """

    print(clear_name(test_cell))