import os
import Exel_Otchet
html_files = []
otchet_data = [0, 0, 0, 0]


for root, dirs, files in os.walk("C:/Users/skrut/OneDrive/Рабочий стол/текстовики1"):  # поиск всех html файлов в указанной директории
    for file in files:
        if file.endswith(".html"):
            html_files.append(os.path.join(root, file))


def file_reader(path):  # чтение файла
    with open(path, 'r', encoding="utf-8") as html_file:
        a = html_file.read()
        return a


def refactor_of_html(path):  # заменяем старую таблицу на новую исправленную
    sub_table = Exel_Otchet.sub_table(file_reader(path))
    correct_sub_table = Exel_Otchet.sub_table_changer(sub_table)
    html_file = file_reader(path)
    return html_file.replace(sub_table, correct_sub_table)


def write_html_file(path):  # изменяем таблицы во всех html файлах которые нашел цикл
    with open(path, 'r', encoding="utf-8") as file_html:
        file_html = refactor_of_html(path)
        with open(path, 'w', encoding="utf-8") as file_html_to_write:
            file_html_to_write.write(file_html)


def html_redactor():  # забиндить кнопке в приложении запуск этой функции
    for i in html_files:
        data_for_otchet = Exel_Otchet.table_data(Exel_Otchet.td_strings(Exel_Otchet.sub_table_changer(
                            Exel_Otchet.sub_table(file_reader(i)))))
        global otchet_data
        otchet_data[0] += int(data_for_otchet[8])
        otchet_data[1] += int(data_for_otchet[9])
        otchet_data[2] += int(data_for_otchet[10])
        otchet_data[3] += int(data_for_otchet[11])
        write_html_file(i)


print(otchet_data)
html_redactor()
print(otchet_data)


