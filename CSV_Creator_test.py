import csv

with open(r"C:/Users/skrut/OneDrive/Рабочий стол/текстовики/cartridge_accounting2.csv", 'w', newline='', encoding="utf-8-sig") as csv_file:
    file_writer = csv.writer(csv_file, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
    file_writer.writerow(['name', 'Допуск', 'Конфликты', 'Новых', 'Активных', 'Подтверждено', 'Тип', 'Статус'])
    file_writer.writerow(['2019-02-26', 'CE255X', 'nv print'])
    file_writer.writerow(['2019-02-26', 'CE255X', 'nv print'])
    file_writer.writerow(['2019-02-26', 'CE255X', 'хайхай'])
    file_writer.writerow(['2019-02-26', 'CE255X', 'ивангай'])
