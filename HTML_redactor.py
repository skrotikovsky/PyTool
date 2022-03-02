# файл исключительно для редактирования всех html файлов
import os
import Exel_Otchet


def get_html_dirs(path):
    html_files = []
    for root, dirs, files in os.walk(path):  # поиск всех html файлов в указанной директории получаем массив директорий
        for file in files:
            if file.endswith(".html"):
                html_files.append(os.path.join(root, file))
    return html_files


def file_reader(path):  # чтение файла
    with open(path, 'r', encoding="utf-8") as html_file:
        a = html_file.read()
        return a


def refactor_of_html(path):  # заменяем старую таблицу на новую исправленную
    sub_table = Exel_Otchet.sub_table(file_reader(path))
    correct_sub_table = Exel_Otchet.sub_table_changer(sub_table)
    html_file = file_reader(path)
    return html_file.replace(sub_table, correct_sub_table)


# ------------------------------------------------------------------------------------------
def write_html_file(path):  # изменяем таблицы во всех html файлах которые нашел цикл
    file_html = refactor_of_html(path)
    with open(path, 'w', encoding="utf-8") as file_html_to_write:
        file_html_to_write.write(file_html)
# ------------------------------------------------------------------------------------------


def html_redactor(path):  # пробегает по ВСЕМ html в директории
    for i in get_html_dirs(path):
        write_html_file(i)

# "C:/Users/skrut/OneDrive/Рабочий стол/текстовики1" test dir
