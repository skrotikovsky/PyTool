import csv
import os
import Exel_Otchet


def get_dirs(path):
    html_dirs = []
    for root, dirs, files in os.walk(path):  # поиск всех html файлов в указанной директории получаем массив директорий
        for file in files:
            html_dirs.append(root)
    print(list(set(html_dirs)))
    return list(set(html_dirs))


def get_html_in_dir(path):
    names = os.listdir(path)
    file_list = []
    for i in range(len(names)):
        if names[i].endswith(".html"):
            names[i] = path + '/' + names[i]
            file_list.append(names[i])
    print(file_list)
    return file_list


def create_csv_rows(path):
    files = get_html_in_dir(path)
    for i in range(len(files)):
        files[i] = Exel_Otchet.get_exel_format(files[i], False)
    print(files)
    return files


def create_csv_file(path):
    rows = create_csv_rows(path)
    dirs = get_dirs(path)
    for i in dirs:
        with open(i+'/exel.csv', 'w', newline='', encoding="utf-8-sig") as csv_file:
            file_writer = csv.writer(csv_file, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            file_writer.writerow(['name', 'Допуск', 'Конфликты', 'Новых', 'Активных', 'Подтверждено', 'Тип', 'Статус'])
            for row in rows:
                print(row)
                file_writer.writerow(row)


'''print(get_html_in_dir("C:/Users/skrut/OneDrive/Рабочий стол/текстовики1"))
print(os.listdir("C:/Users/skrut/OneDrive/Рабочий стол/csvtest"))
for i in os.listdir("C:/Users/skrut/OneDrive/Рабочий стол/csvtest"):
    if i.endswith(".txt"):
        print(i)
print()
print(get_html_in_dir('C:/Users/skrut/OneDrive/Рабочий стол/csvtest'))
'''
create_csv_file("C:/Users/skrut/OneDrive/Рабочий стол/csvtest")