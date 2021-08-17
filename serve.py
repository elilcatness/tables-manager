import os
import csv
import openpyxl

EXTENSIONS = ['csv', 'xls', 'xlsx']

OPERATORS = {'>': (lambda x, y: x > y, 'Больше, чем'),
             '<': (lambda x, y: x < y, 'Меньше, чем'),
             '=': (lambda x, y: x == y, 'Равно'),
             'contains': (lambda x, y: x in y, 'Содержит'),
             '~contains': (lambda x, y: x not in y, 'Не содержит')}


class ExcelIterator:
    def __init__(self, excel):
        self.workbook = excel.get_workbook()
        self.current_row = excel.get_start()

    def __next__(self):
        sheet = self.workbook.active
        if self.current_row <= sheet.max_row:
            row = [sheet.cell(self.current_row, col).value for col in range(1, sheet.max_column + 1)]
            self.current_row += 1
            return row
        raise StopIteration


class ExcelReader:
    def __init__(self, filename, start_from=1):
        self.workbook = openpyxl.open(filename)
        self.start_from = start_from

    def get_workbook(self):
        return self.workbook

    def get_start(self):
        return self.start_from

    def __iter__(self):
        return ExcelIterator(self)


def get_headers(filename, delimiter=';'):
    extension = filename.split('.')[-1]
    if extension not in EXTENSIONS:
        return None
    if extension == 'csv':
        f = open(filename, newline='', encoding='utf-8')
        data_reader = csv.reader(f, delimiter=delimiter)
    else:
        f = None
        data_reader = iter(ExcelReader(filename))
    headers = next(data_reader)
    if f:
        f.close()
    return headers


def process_query(statements):
    logical_indexes = [i for i in range(len(statements)) if statements[i] in ('and', 'or')]
    processed_statements = []
    for idx in logical_indexes:
        if idx < len(statements) - 1:
            statements.insert(idx + 1, statements[0])
        else:
            statements.append(statements[0])
        processed_statements.extend([tuple(statements[idx - 3:idx]), statements[idx],
                                     tuple(statements[idx + 1:idx + 4])])
    return statements


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
    if statement:
        statements.append(statement)
    return statements


def show_help():
    print('\n'.join([f'{key}: {val}' for key, val in OPERATORS.items()]))


def get_extension(filename):
    return filename.split('.')[-1]


def get_reader(filename, delimiter=';'):
    extension = get_extension(filename)
    if get_extension(filename) not in EXTENSIONS:
        return None
    if extension == 'csv':
        file = open(filename, newline='', encoding='utf-8')
        reader = csv.reader(file, delimiter=delimiter)
        return reader, file
    return ExcelReader(filename), None


def split_files(filenames, rows_count, filters, headers=None, delimiter=';'):
    for filename in filenames:
        count, file_count = -1, 0
        reader, main_file = get_reader(filename, delimiter, skip_headers=True)
        csv_file, writer = None, None
        for row in reader:
            count += 1
            if count % rows_count == 0 or count == 0:
                file_count += 1
                if csv_file:
                    csv_file.close()
                csv_file = open(f'{".".join(filename.split(".")[:-1])}_{file_count}.csv',
                                'w', newline='', encoding='utf-8')
                writer = csv.writer(csv_file, delimiter=delimiter)
                if headers:
                    writer.writerow(headers[filename])
            writer.writerow(row)
        if main_file:
            main_file.close()


def manage_split_files():
    add_headers = input('Добавлять ли в каждый выходной файл шапку с заголовками? (y\\n) '
                        ).lower() == 'y'
    while True:
        try:
            rows_count = int(input('Введите количество строк в каждом файле: '))
            break
        except ValueError:
            print('Неверный формат, введите число ещё раз')
    filters = {}
    set_filters = input('Будете ли Вы устанавливать фильтры? (y\\n) ').lower() == 'y'
    correct_filenames = [f_name for f_name in os.listdir() if f_name.split('.')[-1] in EXTENSIONS]
    headers = {}
    for filename in correct_filenames:
        headers[filename] = get_headers(filename, delimiter=';')
    while set_filters:
        print(f"{'*' * 20} Меню фильтрации {'*' * 20}")
        print(' '.join(correct_filenames))
        while True:
            filename = input('Введите название файла: ')
            if filename.strip() in correct_filenames:
                break
            else:
                print('Неверное название файла')
        filters[filename] = []
        headers = get_headers(filename)
        print(f"Столбцы: {';'.join(headers)}")
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
                operators_statements = statements[1::3]
                if not statements:
                    print('Значение фильтра не может быть пустым')
                elif statements[0].lower() not in map(str.lower, headers):
                    print('Указан несуществующий столбец')
                elif len(statements) < 3:
                    print('Недостаточно аргументов')
                elif len(statements) > 3 and ('and' not in statements and 'or' not in statements):
                    print('Упущены логические И/ИЛИ')
                elif any(map(lambda statement: statement not in OPERATORS.keys(), operators_statements)):
                    print(', '.join([f"{statement}" for statement in operators_statements]),
                          'Are not operators', sep=' - ')
                else:
                    filters[filename] = process_query(statements)
                    break
            except ValueError:
                print('Неверно введены значения')
    for filename in correct_filenames:
        split_file(filename, rows_count, filters, )
    split_files(correct_filenames, rows_count, filters, headers=headers if add_headers else None)


def unite_files():
    pass


def manage_unite_files():
    pass


def main():
    actions = {'разбить': manage_split_files, '1': manage_split_files,
               'объединить': manage_unite_files, '2': manage_unite_files}
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