import re
import json
import csv
import sqlite3
import os.path
import sys
import subprocess
import datetime
from itertools import cycle
from downloader import Downloader


def install(package):
    """
    Устанавливает пакет
    :param package: название пакета (str)
    :return: код завершения процесса (int) или текст ошибки (str)
    """
    try:
        result = subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    except subprocess.CalledProcessError as result:
        return result

    return result


try:
    import xlrd
except ImportError:
    exit_code = install("xlrd")
    if exit_code == 0:
        import xlrd
    else:
        print("При установке пакета возникла ошибка! {}".format(exit_code))
        exit(0)


class Reader:
    """Класс для парсинга расписания MIREA из xlsx файлов"""

    def __init__(self, src, path_to_db=None):
        """Инициализация клсса
            src(str): Абсолютный путь к XLS файлу
        """
        self.book = xlrd.open_workbook(src)
        self.sheets = self.book.sheet_by_index(0)
        self.json_file = "table.json"  # Имя файла, в который записывается расписание групп в JSON формате
        self.csv_file = "table.csv"  # Имя файла, в который записывается расписание групп в CSV формате
        self.path_to_db = path_to_db  # Путь к файлу базы данных

    def read(self, write_to_json_file=False, write_to_csv_file=False, write_to_db=False):
        """Объединяет расписания отдельных групп и записывает в файлы
            write_to_json_file(bol): Записывать ли в JSON файл
            write_to_csv_file(bol): Записывать ли в CSV файл
        """

        def get_day_num(day_name):
            days = {
                'ПОНЕДЕЛЬНИК': 1,
                'ВТОРНИК': 2,
                'СРЕДА': 3,
                'ЧЕТВЕРГ': 4,
                'ПЯТНИЦА': 5,
                'СУББОТА': 6
            }
            return days[day_name]

        column_range = []
        timetable = {}
        group_list = []
        # строка с названиями групп
        row = self.sheets.row(1)
        for is_group in row:  # Поиск названий групп
            group = str(is_group.value)
            group = re.search(r"([А-Я]+-\w+-\w+)", group)
            if group:  # Если название найдено, то получение расписания этой группы

                if not group_list:  # инициализация списка диапазонов пар
                    column_range = {
                        1: [],
                        2: [],
                        3: [],
                        4: [],
                        5: [],
                        6: []
                    }

                    inicial_row_num = 3  # Номер строки, с которой начинается отсчет пар

                    para_count = 0  # Счетчик количества пар
                    # Перебор столбца с номерами пар и вычисление на основании количества пар в день диапазона выбора ячеек

                    day_num_val, para_num_val, para_time_val, para_week_num_val = 0, 0, 0, 0
                    for para_num in range(inicial_row_num, len(self.sheets.col(row.index(is_group) - 4))):

                        day_num_col = self.sheets.cell(para_num, row.index(is_group) - 5)
                        if day_num_col.value != '':
                            day_num_val = get_day_num(day_num_col.value)

                        para_num_col = self.sheets.cell(para_num, row.index(is_group) - 4)
                        if para_num_col.value != '':
                            para_num_val = para_num_col.value
                            if isinstance(para_num_val, float):
                                para_num_val = int(para_num_val)
                                if para_num_val > para_count:
                                    para_count = para_num_val

                        para_time_col = self.sheets.cell(para_num, row.index(is_group) - 3)
                        if para_time_col.value != '':
                            para_time_val = para_time_col.value.replace('-', ':')
                        para_week_num = self.sheets.cell(para_num, row.index(is_group) - 1)
                        if para_week_num.value != '':
                            if para_week_num.value == 'I':
                                para_week_num_val = 1
                            else:
                                para_week_num_val = 2

                        para_string_index = para_num

                        if re.findall(r'\d+:\d+', para_time_val, flags=re.A):
                            para_range = (para_num_val, para_time_val, para_week_num_val, para_string_index)
                            column_range[day_num_val].append(para_range)

                print(group.group(0))
                group_list.append(group.group(0))
                one_time_table = self.read_one_group(row.index(is_group), column_range)  # По номеру столбца
                for key in one_time_table.keys():
                    timetable[key] = one_time_table[key]  # Добавление в общий словарь

        if write_to_json_file:  # Запись в JSON файл
            self.write_to_json(timetable)
        if write_to_csv_file or write_to_db:  # Запись в CSV файл, Запись в базу данных
            self.write(timetable, write_to_csv_file, write_to_db)
        return group_list

    @staticmethod
    def format_other_cells(cell):
        cell = cell.split("\n")
        return cell

    @staticmethod
    def format_prepod_name(cell):
        return re.split(r' {2,}|\n', cell)

    @staticmethod
    def format_room_name(cell):
        notes_dict = {
            'МП-1': "ул. Малая Пироговская, д.1",
            'В-78': "Проспект Вернадского, д.78",
            'В-86': "Проспект Вернадского, д.86",
            'С-20': "ул. Стромынка, 20",
            'СГ-22': "5-я ул. Соколиной горы, д.22"
        }

        if isinstance(cell, float):
            cell = int(cell)
        string = str(cell)
        for pattern in notes_dict:
            regex_result = re.findall(pattern, string, flags=re.A)
            if regex_result:
                string = string.replace(' ', '').replace('*', '').replace('\n', '')

                string = re.sub(regex_result[0], notes_dict[regex_result[0]] + " ", string, flags=re.A)

        return re.split(r' {2,}|\n', string)

    @staticmethod
    def format_name(temp_name):
        """Разбор строки 'Предмет' на название дисциплины и номера
            недель включения и исключения
            temp_name(str)
        """

        def if_diapason_week(para_string):
            start_week = re.findall(r"\d+-", para_string)
            start_week = re.sub("-", "", start_week[0])
            end_week = re.findall(r"-\d+", para_string)
            end_week = re.sub("-", "", end_week[0])
            weeks = []
            for week in range(int(start_week), int(end_week) + 1):
                weeks.append(week)
            return weeks

        result = []
        temp_name = temp_name.replace(" ", "  ")

        temp_name = re.sub(r"(кр\. {2,})", "кр.", temp_name, flags=re.A)
        temp_name = re.sub(r"(н[\d,. ]*\+)", "", temp_name, flags=re.A)

        temp_name = re.findall(r"((?:\s*[\W\s]*)(?:|кр[ .]\s*|\d+-\d+|[\d,. ]*)\s*\s*(?:|[\W\s]*|\D*)*)(?:\s\s|\Z|\n)",
                               temp_name, flags=re.A)
        if isinstance(temp_name, list):
            for item in temp_name:
                if len(item) > 0:
                    if_except = re.search(r"(кр[. \w])", item, flags=re.A)
                    if_include = re.search(r"( н[. ])|(н[. ])|(\d\s\W)|(\d+\s+\D)", item, flags=re.A)
                    _except = ""
                    _include = ""
                    item = re.sub(r"\(", "", item, flags=re.A)
                    item = re.sub(r"\)", "", item, flags=re.A)
                    if if_except:
                        if re.search(r"\d+-\d+", item, flags=re.A):
                            _except = if_diapason_week(item)
                            item = re.sub(r"\d+-\d+", "", item, flags=re.A)
                        else:
                            _except = re.findall(r"(\d+)", item, flags=re.A)
                        item = re.sub(r"(кр[. \w])", "", item, flags=re.A)
                        item = re.sub(r"(\d+[,. ]+)", "", item, flags=re.A)
                        name = re.sub(r"( н[. ])", "", item, flags=re.A)
                    elif if_include:
                        if re.search(r"\d+-\d+", item):
                            _include = if_diapason_week(item)
                            item = re.sub(r"\d+-\d+", "", item, flags=re.A)
                        else:
                            _include = re.findall(r"(\d+)", item, flags=re.A)
                        item = re.sub(r"(\d+[,. н]+)", "", item, flags=re.A)
                        name = re.sub(r"(н[. ])", "", item, flags=re.A)
                    else:
                        name = item
                    # name = re.sub(r"  ", " ", name)
                    name = name.replace("  ", " ")
                    name = name.strip()
                    one_str = [name, _include, _except]
                    result.append(one_str)
        return result

    def write(self, timetable, write_to_csv=False, write_to_db=False):
        """Запись словаря 'timetable' в CSV файл или в базу данных
            - путь к файлу базы данных
        """

        def create_table(group):
            """Если Таблица не создана, создать таблицу
                table_name - Название таблицы
            """
            c.execute("DROP TABLE IF EXISTS {}".format(group))
            c.execute("""CREATE TABLE {} (day TEXT, para TEXT, time DATE,
                      week TEXT, name TEXT, type TEXT, room TEXT, prepod TEXT,
                      include TEXT, exception TEXT)""".format(group))

        def data_append(group, day_num, para_num, para_time, parity, discipline, para_type, room, teacher, include_week,
                        exception_week):
            """Добавление данных в базу данных"""
            c.execute("""INSERT INTO {} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""".format(group),
                      (day_num, para_num, para_time, parity, discipline, para_type, room, teacher, include_week,
                       exception_week))

        conn = sqlite3.connect(self.path_to_db)
        c = conn.cursor()

        with open(self.csv_file, "a", encoding="utf-8", newline="") as fh:
            if write_to_csv is not False:
                writer = csv.DictWriter(fh, fieldnames=["group_name", "day", "para", "week", "name",
                                                        "type", "room", "prepod", "include",
                                                        "exception"], quoting=csv.QUOTE_ALL)
                writer.writeheader()
            for group_name, value in sorted(timetable.items()):
                if write_to_db is not False:
                    table_name = group_name.replace('-', '_').replace(' ', '_').replace('(', '_') \
                                     .replace(')', '_').replace('.', '_').replace(',', '_')[0:10]
                    create_table(table_name)
                for n_day, day_item in sorted(value.items()):
                    for n_para, para_item in sorted(day_item.items()):
                        for n_week, item in sorted(para_item.items()):
                            day = n_day.split("_")[1]
                            para = n_para.split("_")[1]
                            week = n_week.split("_")[1]
                            for dist in item:
                                time = dist['time']
                                if "include" in dist:
                                    include = str(dist["include"])[1:-1]
                                else:
                                    include = ""
                                if "exception" in dist:
                                    exception = str(dist["exception"])[1:-1]
                                else:
                                    exception = ""
                                if write_to_csv is not False:
                                    writer.writerow(dict(group_name=group_name, day=day, para=para, week=week,
                                                         name=dist["name"], type=dist["type"], room=dist["room"],
                                                         prepod=dist["prepod"], include=include, exception=exception))
                                if write_to_db is not False:
                                    data_append(table_name, day, para, time, week,
                                                dist["name"], dist["type"],
                                                dist["room"], dist["prepod"],
                                                include, exception)
            conn.commit()
            c.close()

    def write_to_json(self, timetable):
        """Запись словаря 'timetable' в JSON файл
            timetable(dict)
        """
        with open(self.json_file, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(timetable, ensure_ascii=False, indent=4))

    def read_one_group(self, discipline_col_num, cell_range):
        """Получение расписания одной группы
            discipline_col_num(int): Номер столбца колонки 'Предмет'
            range(dict): Диапазон выбора ячеек
        """
        one_group = {}
        group_name = self.sheets.cell(1, discipline_col_num).value  # Название группы
        one_group[group_name] = {}  # Инициализация словаря

        # перебор по дням недели (понедельник-суббота)
        # номер дня недели (1-6)
        for day_num in cell_range:
            one_day = {}

            for para_range in cell_range[day_num]:
                para_num = para_range[0]
                time = para_range[1]
                week_num = para_range[2]
                string_index = para_range[3]

                # Перебор одного дня (1-6 пара)
                if "para_{}".format(para_num) not in one_day:
                    one_day["para_{}".format(para_num)] = {}

                # Получение данных об одной паре
                tmp_name = str(self.sheets.cell(string_index, discipline_col_num).value)
                tmp_name = self.format_name(tmp_name)

                if isinstance(tmp_name, list) and tmp_name != []:

                    para_type = self.sheets.cell(string_index, discipline_col_num + 1).value
                    teacher = self.format_prepod_name(self.sheets.cell(string_index, discipline_col_num + 2).value)
                    room = self.format_room_name(self.sheets.cell(string_index, discipline_col_num + 3).value)

                    max_len = max(len(tmp_name), len(teacher), len(room))
                    if len(tmp_name) < max_len:
                        tmp_name = cycle(tmp_name)
                    if len(teacher) < max_len:
                        teacher = cycle(teacher)
                    if len(room) < max_len:
                        room = cycle(room)

                    para_tuple = list(zip(tmp_name, teacher, room))
                    for tuple_item in para_tuple:
                        name = tuple_item[0][0]
                        include = tuple_item[0][1]
                        exception = tuple_item[0][2]
                        teacher = tuple_item[1]
                        room = tuple_item[2]

                        if isinstance(room, float):
                            room = int(room)

                        one_para = {"time": time, "name": name, "type": para_type, "prepod": teacher, "room": room}
                        if include:
                            one_para["include"] = include
                        if exception:
                            one_para["exception"] = exception

                        if name and room:
                            if "week_{}".format(week_num) not in one_day["para_{}".format(para_num)]:
                                one_day["para_{}".format(para_num)][
                                    "week_{}".format(week_num)] = []  # Инициализация списка
                            one_day["para_{}".format(para_num)]["week_{}".format(week_num)].append(one_para)

                    # Объединение расписания
                    one_group[group_name]["day_{}".format(day_num)] = one_day

        return one_group


if __name__ == "__main__":

    Downloader = Downloader(path_to_error_log='logs/downloadErrorLog.csv', file_dir='xls/',
                            file_type='xlsx')
    Downloader.download()

    for i in os.scandir("xls"):
        xlsx_path = os.path.join("xls", i.name)
        print(xlsx_path)

        try:
            reader = Reader(xlsx_path, "table.db")
            res = reader.read(write_to_json_file=False, write_to_csv_file=False, write_to_db=True)
        except Exception as err:
            print(err)
            with open('logs/ErrorLog.txt', 'a+') as file:
                file.write(str(datetime.datetime.now()) + ':' + str(i) + 'message:' + str(err) + '\n')
            continue

    # reader = Reader
    # result = reader.format_name("2,6,10,14 н Экология\n4,8,12,16 Правоведение")
    # print(result)
    # result = reader.format_name("Деньги, кредит, банки кр. 2,8,10 н.")
    # print(result)
    # result = reader.format_name("Орг. Химия (1-8 н.)")
    # print(result)
    # result = reader.format_name("Web-технологии в деятельности экономических субьектов")
    # print(result)
    # result = reader.format_name("1,5,9,13 н Оперционные системы\n3,7,11,15 н  Оперционные системы")
    # print(result)
    # result = reader.format_name("3,7,11,15  н Физика                                                     кр. 5,8,13.17 н Организация ЭВМ и Систем")
    # print(result)
    # result = reader.format_name("2,6,10,14 н Кроссплатформенная среда исполнения программного обеспечения 4,8,12,16 н Кроссплатформенная среда исполнения программного обеспечения")
    # print(result)
    # result = reader.format_name("кр 1,13 н Интеллектуальные информационные системы")
    # print(result)
    # result = reader.format_name("кр1 н Модели информационных процессов и систем")
    # print(result)
    # result = reader.format_name("Разработка ПАОИиАС 2,6,10,14 н.+4,8,12,16н.")
    # print(result)
    #
    # result = reader.format_prepod_name("Козлова Г.Г.\nИсаев Р.А.")
    # print(result)
    #
    # result = reader.format_room_name("В-78*\nБ-105")
    # print(result)
    #
    # result = reader.format_room_name("23452     Б-105")
    # print(result)
