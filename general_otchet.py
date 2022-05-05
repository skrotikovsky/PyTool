import os
from datetime import date
import pyexcel
import openpyxl
import xlwt
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
import xlsxwriter

path_exel = r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/Просмотр/Коллизии/2022.4.24"
path_of_main = r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples"


def get_xlsx_in_dir(path):
    names = os.listdir(path)
    file_list = []
    for i in range(len(names)):
        if names[i].endswith(".xlsx"):
            names[i] = path + '/' + names[i]
            file_list.append(names[i])
    #    print(file_list)
    return file_list


def find_xlsx(path):
    html_dirs = []
    for root, dirs, files in os.walk(path):  # поиск всех html файлов в указанной директории получаем массив директорий
        for file in files:
            html_dirs.append(root)
    #    print(list(set(html_dirs)))
    return list(set(html_dirs))


def get_exel_array(path):  # представляет отчет в виде двумерного массива
    return pyexcel.get_array(file_name=get_xlsx_in_dir(path)[0])


def get_otchet_rows_array(path):  # из названия файла достает марку и делает приписку а так же достает новые коллизии
    # и их общее количество
    rows_array = []
    data = get_exel_array(path)
    for i in data[1:len(data) - 1]:
        marks = i[0].split('.')
        i[0] = marks[len(marks) - 1]
        #        print(i[0], i[3], i[5])
        rows_array.append([f'Кол-во конфликтов между {i[0]}', i[1], i[4]])
    for i in range(len(rows_array)):
        for j in range(3):
            if rows_array[i][j] == '' or rows_array[i][j] == 'Коллизий не обнаружено':
                rows_array[i][j] = 0
    return rows_array


def get_otchet_marks_array(path):  # достает только марки и делает приписку
    marks_array = []
    data = get_exel_array(path)
    for i in data[1:len(data) - 1]:
        marks = i[0].split('.')
        i[0] = marks[len(marks) - 1]
        #        print(i[0], i[3], i[5])
        marks_array.append(f'Кол-во конфликтов между {i[0]}')
    return marks_array


def get_set_of_marks(path):  # оставляет только уникальные марки и создает словарь "марка": массив(пустой пока что)
    list_of_dicts = list(map(lambda x: {x: []}, list(set(get_otchet_marks_array(path)))))
    new_dict = {}

    for i in list_of_dicts:
        new_dict.update(i)
    #   print(new_dict)
    return new_dict


def get_marks_and_row_dict(path):  # заполняет массивы предыдущего словаря и если в строках есть одинаковые марки -
    # схлопывает их
    rows_array = get_otchet_rows_array(path)
    marks_dict = get_set_of_marks(path)
    for i in rows_array:
        mark = marks_dict[i[0]]
        if not mark:
            marks_dict[i[0]] = [int(i[1]), int(i[2])]
        else:
            marks_dict[i[0]] = [int(mark[0]) + int(i[1]), int(mark[1]) + int(i[2])]
    return marks_dict


def get_main_otchet_array(main_path):  # читает главный отчет и представляет его в виде двумерного массива
    return pyexcel.get_array(file_name=get_xlsx_in_dir(main_path)[0])


def write_if_main_is_empty(path,
                           main_path):  # если главный отчет пустой - заполняет его марками в первом столбце(теми
    # которые
    # были в начальных данных) и заполняет два ближних столбца
    marks_to_write = (list(map(lambda x: [x[0], x[1][0], x[1][1]], get_marks_and_row_dict(path).items())))
    marks_to_write.insert(0, ['Конфликты', 'Новые', 'Общее кол-во'])
    marks_to_write.insert(0, ['', f'{date.today()}', ''])
    #    marks_to_write = {'pyexel_sheet1': marks_to_write}
    #    print(marks_to_write)
    #    pyexcel.save_as(bookdict=marks_to_write,
    #                    dest_file_name=r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx")
    df = pd.DataFrame(marks_to_write)
    print(df)
    writer = pd.ExcelWriter(
        r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx", engine='xlsxwriter')
    df.to_excel(writer, 'openpyxl', index=False, index_label=False, header=False)
    writer.save()


def write_if_data_exists(path, main_path):  # если главный отчет уже заполнен - добавляет данные по новой проверке
    # в начало списка проверок
    main_exel_array = get_main_otchet_array(main_path)
    marks_to_write = (list(map(lambda x: [x[0], x[1][0], x[1][1]], get_marks_and_row_dict(path).items())))
    #    print(marks_to_write)
    #    print(main_exel_array)
    for i in range(len(main_exel_array)):
        if i == 0:
            main_exel_array[i].insert(1, f'{date.today()}')
            main_exel_array[i].insert(2, '')
        elif i == 1:
            main_exel_array[i].insert(1, 'Новые')
            main_exel_array[i].insert(2, 'Общее кол-во')
        else:
            main_exel_array[i].insert(1, marks_to_write[i - 2][1])
            main_exel_array[i].insert(2, marks_to_write[i - 2][2])
    #    main_exel_array = {'pyexel_sheet1': main_exel_array}
    #    print(main_exel_array)
    #    pyexcel.save_book_as(bookdict=main_exel_array,
    #    dest_file_name=r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx")
    #    print(df)
    #    print(ws)
    # Append the rows of the DataFrame to your worksheet
    #    for r in dataframe_to_rows(df, index=False, header=False):
    #        print(r)
    #        ws.append(r[0])
    df = pd.DataFrame(main_exel_array)
    print(df)
    writer = pd.ExcelWriter(
        r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx", engine='xlsxwriter')
    df.to_excel(writer, 'openpyxl', index=False, index_label=False, header=False)
    writer.save()


def write_data_in_main_otchet(path, main_path):  # заполняет данными в зависимости от того заполнен ли главный отчет
    # или он пуст
    main_array = get_main_otchet_array(main_path)
    print(main_array)
    print(main_array)
    for i in main_array:
        if not i:
            continue
        else:
            return write_if_data_exists
    return write_if_main_is_empty


'''    if not main_array:
        write_if_main_is_empty(path)
    else:
        write_if_data_exists(path, main_path)'''

write_data_in_main_otchet(path_exel, path_of_main)(path_exel, path_of_main)

'''  
exel_array = pyexcel.get_array(file_name=get_xlsx_in_dir(path_exel)[0])
marks_array = []
rows_array = []

for i in exel_array[1:len(exel_array) - 1]:
    marks = i[0].split('.')
    i[0] = marks[len(marks) - 1]
    print(i[0], i[3], i[5])
    rows_array.append([f'Кол-во конфликтов между {i[0]}', i[3], i[4]])
    marks_array.append(f'Кол-во конфликтов между {i[0]}')

print(marks_array)
marks_dict = {}

for i in range(len(rows_array)):
    for j in range(3):
        if rows_array[i][j] == '':
            rows_array[i][j] = 0

for i in set(marks_array):
    marks_dict[i] = []

for i in rows_array:
    mark = marks_dict[i[0]]
    if not mark:
        marks_dict[i[0]] = [int(i[1]), int(i[2])]
    else:
        marks_dict[i[0]] = [int(mark[0]) + int(i[1]), int(mark[1]) + int(i[2])]
print(marks_dict)

main_exel_array = pyexcel.get_array(file_name=get_xlsx_in_dir(path_of_main)[0])

if not main_exel_array:
    marks_to_write = (list(map(lambda x: [x[0], x[1][0], x[1][1]], marks_dict.items())))
    marks_to_write.insert(0, ['Конфликты', 'Новые', 'Общее кол-во'])
    marks_to_write.insert(0, ['', f'{date.today()}', ''])
    marks_to_write = {'pyexel_sheet1': marks_to_write}
    print(marks_to_write)
    pyexcel.save_as(bookdict=marks_to_write,
                    dest_file_name=r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx")
# top_players = pandas.read_excel(get_html_in_dir(path)[0])
# top_players.head()
else:
    marks_to_write = (list(map(lambda x: [x[0], x[1][0], x[1][1]], marks_dict.items())))
    print(marks_to_write)
    print(main_exel_array)
    for i in range(len(main_exel_array)):
        if i == 0:
            main_exel_array[i] = main_exel_array[i] + [f'{date.today()}', '']

        elif i == 1:
            main_exel_array[i] = main_exel_array[i] + ['Новые', 'Общее кол-во']
        else:
            main_exel_array[i] = main_exel_array[i] + [marks_to_write[i - 2][1], marks_to_write[i - 2][2]]
    main_exel_array = {'pyexel_sheet1': main_exel_array}
    print(main_exel_array)
    pyexcel.save_book_as(bookdict=main_exel_array,
                         dest_file_name=r"C:/Users/skrut/OneDrive/Рабочий стол/exelExamples/KT101R_Главный отчет.xlsx")'''''
