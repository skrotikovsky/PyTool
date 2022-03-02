# написать файл читающий таблицу и запоминающий ее данные и отдающий в HTML redactor
import re


def file_reader(path):  # чтение файла
    with open(path, 'r', encoding="utf-8") as html_file:
        return html_file.read()


def sub_table(file_html):  # поиск таблицы по шаблону
    match = re.findall("<table class=\"testSummaryTable\">.*<table class=\"mainTable\" id = \"GetTable\">", file_html,
                       flags=re.S)
    return match[0]


def sub_table_changer(table):  # Новая таблица по шаблону без лишних столбцов
    new_table = ''
    strings = table.split('\n')
    for i in range(len(strings)):
        if i == 7 or i == 9 or i == 18 or i == 20:
            pass
        else:
            if i != len(table.split('\n')) - 1:
                strings[i] += '\n'
            new_table += strings[i]
    return new_table


def td_strings(new_table):  # Извлечение самих строк для последущего извеления данных для формирования отчетов
    match = re.findall("<td class=.*</td>", new_table)
    new_match = []
    for i in range(len(match)):
        new_match.append(match[i])
    return new_match


def table_data(data_strings):  # Числа из таблицы
    data_of_table = []
    for i in data_strings:
        i = re.search("\">.*<", i)[0]
        i = i[2:len(i) - 1]
        data_of_table.append(i)
    return data_of_table


def name_of_model(path):  # Имя модели для генерации CSV
    html_file = file_reader(path)
    model_name = re.findall("testName\">.*<", html_file)[0][10: -1]
    return model_name


def exel_table_field_format(name, data):  # Перевод в формат строки CSV файла
    return name + "," + ",".join(data[7:])


exel_table_field_format(name_of_model(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики1/tester.html"),
                        table_data(td_strings(sub_table_changer(
                            sub_table(file_reader(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики1/tester.html"))))))



# sub_table_changer(sub_table(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики/TESTER.txt"))
# алгоритм: достаем все строки начинающиеся с <td> и извлекаем из них данные
# что бы заменить надо будет исзвлечь через subtable изначальную таблицу, удалить из нее строки ненужные и с помощью sub поставить обратно
# sub_table(file_reader(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики/TESTER.txt"))
#  file_reader(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики")
