import os
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils.exceptions import InvalidFileException
from csv import writer


EXTENSIONS = ['csv', 'xls', 'xlsx']
OPERATORS = {'>': (lambda x, y: x > y, 'Больше, чем'),
             '<': (lambda x, y: x < y, 'Меньше, чем'),
             '=': (lambda x, y: x == y, 'Равно'),
             'contains': (lambda x, y: x in y, 'Содержит'),
             '~contains': (lambda x, y: x not in y, 'Не содержит')}


def from_excel_to_csv(excel_filename: str, csv_filename=None, has_headers=True, delimiter=';'):
    csv_filename = (f"{'.'.join(excel_filename.split('.')[:-1])}.csv"
                    if csv_filename is None else csv_filename)
    try:
        workbook = openpyxl.open(excel_filename)
    except InvalidFileException:
        raise NameError(f'Неподдерживаемое расширение файла: {excel_filename.split(".")[-1]}')
    sheet: Worksheet = workbook[workbook.sheetnames[0]]
    if has_headers:
        print(sheet.min_row)
        headers = [sheet.cell(sheet.min_row, col).value.strip() for col in range(1, sheet.max_column + 1)
                   if sheet.cell(1, col).value.strip()]
    else:
        headers = []
    with open(csv_filename, 'w', newline='', encoding='utf-8') as csv_file:
        csv_writer = writer(csv_file, delimiter=';')
        if headers:
            csv_writer.writerow(headers)
        for row in range(sheet.min_row + 1, sheet.max_row + 1):
            values = [sheet.cell(row, col).value
                      for col in range(sheet.min_column, sheet.max_column + 1)]
            csv_writer.writerow(values)


def parse_query(query):
    statements = []
    idx = -1
    statement = ''
    while idx < len(query) - 1:
        idx += 1
        if query[idx] == '"':
            idx += 1
            if idx == len(query):
                break
            statement += query[idx]
            while idx < len(query) - 1 and query[idx + 1] != '"':
                idx += 1
                statement += query[idx]
            statements.append(statement)
            statement = ''
            idx = idx + 1 if idx < len(query) - 1 else idx
        elif query[idx] == ' ':
            if statement:
                statements.append(statement)
                statement = ''
        else:
            statement += query[idx]
    return query


def show_help():
    print('\n'.join([f'{key}: {val}' for key, val in OPERATORS.items()]))


def split_files():
    add_headers = input('Добавлять ли в каждый выходной файл шапку с заголовками? (y\\n) '
                        ).lower() == 'y'
    while True:
        try:
            rows_count = int(input('Введите количество строк в каждом файле: '))
            break
        except ValueError:
            print('Неверный формат, введите число ещё раз')
    filters = {}
    set_filters = input('Будете ли Вы устанавливать фильтры? (y\\n) ')
    while set_filters:
        print(f"{'*' * 20} Меню фильтрации {'*' * 20}")
        correct_filenames = [f_name for f_name in os.listdir() if f_name.split('.')[-1] in EXTENSIONS]
        print(' '.join(correct_filenames))
        filename = None
        while True:
            filename = input('Введите название файла: ')
            if filename.strip() in correct_filenames:
                break
            else:
                print('Неверное название файла')
        filters[filename] = []
        print('Введите запросы для фильтрации в виде: НАЗВАНИЕ_СТОЛБЦА ОПЕРАТОР ЗНАЧЕНИЕ '
              '(and/or ОПЕРАТОР ЗНАЧЕНЕИЕ)n',
              f'Для выхода из меню фильтрации для файла {filename} введите /exit',
              'Для получения операторов и их функций введите /help', sep='\n', end='\n\n')
        while True:
            filter_query = input()
            if filter_query == '/help':
                show_help()
                continue
            elif filter_query == '/exit':
                break
            try:
                statements = parse_query(filter_query)
                if not statements:
                    print('Значение фильтра не может быть пустым')
                    continue
            except ValueError:
                print('Неверно введены значения')


def unite_files():
    pass


def main():
    actions = {'разбить': split_files, '1': split_files,
               'объединить': unite_files, '2': unite_files}
    while True:
        action = input('Выберите действие: разбить(1) / объединить(2): ')
        if actions.get(action.lower()):
            action = actions[action]
            break
        else:
            print('Неверное значение', end='\n\n')
    action()


if __name__ == '__main__':
    main()