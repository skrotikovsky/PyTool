# написать файл читающий таблицу и запоминающий ее данные и отдающий в HTML redactor
import re


def file_reader(path):  # чтение файла
    with open(path, 'r', encoding="utf-8") as html_file:
        return html_file.read()


def sub_table(path):  # поиск таблицы по шаблону
    file_html = file_reader(path)
    match = re.findall("<table class=\"testSummaryTable\">.*<table class=\"mainTable\" id = \"GetTable\">", file_html,
                       flags=re.S)
    return match[0]


def sub_table_changer(path):  # Новая таблица по шаблону без лишних столбцов
    table = sub_table(path)
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


def td_strings(path):  # Извлечение самих строк для последущего извеления данных для формирования отчетов
    new_table = sub_table(path)
    match = re.findall("<td class=.*</td>", new_table)
    new_match = []
    for i in range(len(match)):
        new_match.append(match[i])
    return new_match


def table_data(path):  # Числа из таблицы
    data_strings = td_strings(path)
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


def get_exel_format(path, is_edited_table):  # Перевод в формат строки CSV файла
    name = name_of_model(path)
    data = table_data(path)
    if not is_edited_table:
        data = data[0:4] + data[5:6] + data[7:13] + data[14:15] + data[16:]
        # небольшой костыль тк шаблон всегда один то попросту склеиваем части массива без ненужных элементов
    return name


print(get_exel_format("C:/Users/skrut/OneDrive/Рабочий стол/текстовики1/"
                      "KT301P.00ZK.000TW01-1004.AS.TD04_KG.html", True))
