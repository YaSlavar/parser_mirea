import re
import json
import csv
import sqlite3
import os.path
import sys
import subprocess
import datetime
from itertools import cycle
from sqlite3.dbapi2 import Connection
import traceback
import xlrd
from downloader import Downloader



class Reader:
    """Класс для парсинга расписания MIREA из xlsx файлов"""
    connect_to_db: Connection

    def __init__(self, path_to_json=None, path_to_csv=None, path_to_db=None):
        """Инициализация клсса
            src(str): Абсолютный путь к XLS файлу
        """

        try:
            import xlrd
        except ImportError:
            exit_code = self.install("xlrd")
            if exit_code == 0:
                import xlrd
            else:
                print("При установке пакета возникла ошибка! {}".format(exit_code))
                exit(0)

        # Имя файла, в который записывается расписание групп в JSON формате
        if path_to_json is not None:
            self.json_file = path_to_json
        else:
            self.json_file = "table.json"

        # Имя файла, в который записывается расписание групп в CSV формате
        if path_to_csv is not None:
            self.csv_file = path_to_csv
        else:
            self.csv_file = "table.csv"

        # Путь к файлу базы данных
        if path_to_db is not None:
            self.db_file = path_to_db
        else:
            self.db_file = 'table.db'

        self.log_file_path = 'logs/ErrorLog.txt'



    @staticmethod
    def install(package):
        """
        Устанавливает пакет
        :param package: название пакета (str)
        :return: код завершения процесса (int) или текст ошибки (str)
        """
        try:
            result = subprocess.check_call(['pip', 'install', package])
        except subprocess.CalledProcessError as result:
            return result

        return result

    def run(self, xlsx_dir, write_to_json_file=False, write_to_csv_file=False, write_to_db=False):
        """
        Выполнение парсинга данных
        :param write_to_db:
        :param write_to_csv_file:
        :param write_to_json_file:
        :type xlsx_dir: str
        """

        def remove_old_file(path_to_file):
            if os.path.isfile(path_to_file):
                os.remove(path_to_file)

        def get_doc_type_code(doc_type_str):
            doc_type_list = {
                'semester': 0,
                'zach': 1,
                'exam': 2
            }
            return doc_type_list[doc_type_str]

        remove_old_file(self.db_file)
        remove_old_file(self.json_file)
        remove_old_file(self.csv_file)

        self.connect_to_db = sqlite3.connect(self.db_file)

        for path, dirs, files in os.walk(xlsx_dir):
            for file_name in files:
                path_to_xlsx_file = os.path.join(path, file_name)
                xlsx_doc_type = get_doc_type_code(os.path.dirname(os.path.relpath(path_to_xlsx_file, start='xls')))

                try:
                    self.read(path_to_xlsx_file, xlsx_doc_type, write_to_json_file=write_to_json_file,
                              write_to_csv_file=write_to_csv_file, write_to_db=write_to_db)
                except Exception as err:
                    print(err, traceback.format_exc())
                    with open(self.log_file_path, 'a+') as log_file:
                        log_file.write(
                            str(datetime.datetime.now()) + ': ' + str(path_to_xlsx_file) + ' message: ' + str(
                                traceback.format_exc()) + '\n')
                    continue

    def read(self, xlsx_parh, doc_type, write_to_json_file=False, write_to_csv_file=False, write_to_db=False):
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
            return days[day_name.upper()]

        def get_month_num(month_name):
            months = {
                'ЯНВАРЬ': 1,
                'ФЕВРАЛЬ': 2,
                'МАРТ': 3,
                'АПРЕЛЬ': 4,
                'МАЙ': 5,
                'ИЮНЬ': 6,
                'ИЮЛЬ': 7,
                'АВГУСТ': 8,
                'СЕНТЯБРЬ': 9,
                'ОКТЯБРЬ': 10,
                'НОЯБРЬ': 11,
                'ДЕКАБРЬ': 12
            }
            return months[month_name.upper().replace(' ', '')]

        def get_column_range_semester(xlsx_sheet, group_name_cell, group_name_row_index):
            """

            :param group_name_row_index: 
            :param xlsx_sheet:
            :param group_name_cell:
            :return:
            """
            # инициализация списка диапазонов пар
            week_range = {
                1: [],
                2: [],
                3: [],
                4: [],
                5: [],
                6: []
            }

            initial_row_num = group_name_row_index + 1  # Номер строки, с которой начинается отсчет пар
            para_count = 0  # Счетчик количества пар
            # Перебор столбца с номерами пар и вычисление на основании количества пар в день диапазона выбора ячеек

            day_num_val, para_num_val, para_time_val, para_week_num_val = 0, 0, 0, 0
            for para_num in range(initial_row_num, len(xlsx_sheet.col(group_name_row.index(group_name_cell)))):

                day_num_col = xlsx_sheet.cell(para_num, group_name_row.index(group_name_cell) - 5)
                if day_num_col.value != '':

                    day_num_val = get_day_num(day_num_col.value)

                para_num_col = xlsx_sheet.cell(para_num, group_name_row.index(group_name_cell) - 4)
                if para_num_col.value != '':
                    para_num_val = para_num_col.value
                    if isinstance(para_num_val, float):
                        para_num_val = int(para_num_val)
                        if para_num_val > para_count:
                            para_count = para_num_val

                para_time_col = xlsx_sheet.cell(para_num, group_name_row.index(group_name_cell) - 3)
                if para_time_col.value != '':
                    para_time_val = str(para_time_col.value).replace('-', ':')

                para_week_num = xlsx_sheet.cell(para_num, group_name_row.index(group_name_cell) - 1)
                if para_week_num.value != '':
                    if para_week_num.value == 'I':
                        para_week_num_val = 1
                    elif para_week_num.value == 'II':
                        para_week_num_val = 2
                else:
                    if para_week_num_val == 1:
                        para_week_num_val = 2
                    else:
                        para_week_num_val = 1

                para_string_index = para_num

                if re.findall(r'\d+:\d+', str(para_time_val), flags=re.A):
                    para_range = (para_num_val, para_time_val, para_week_num_val, para_string_index)
                    week_range[day_num_val].append(para_range)

            return week_range

        def get_column_range_exam(xlsx_sheet, group_name_cell, group_name_row_index):
            """

            :param group_name_row_index: 
            :param xlsx_sheet:
            :param group_name_cell:
            :return:
            """

            def fix_date_range(date_range_list):
                is_fuck = False
                this_index, this_date = 0, 0
                for row in date_range_list:
                    if None in row:
                        is_fuck = True
                    elif is_fuck is True:
                        this_date = datetime.date(datetime.datetime.now().year, row[1], row[0])
                        this_index = date_range.index(row)
                        break

                for range_index in range(this_index, 0, -1):
                    if date_range_list[range_index - 1][0] != date_range_list[range_index][0]:
                        this_date = this_date - datetime.timedelta(1)
                    date_range_list[range_index - 1][0] = this_date.day
                    date_range_list[range_index - 1][1] = this_date.month

                return date_range_list

            initial_row_nam = group_name_row_index + 1  # Номер строки, с которой начинается отсчет пар

            date_range = []
            # Перебор столбца с номерами пар и вычисление на основании количества пар в день диапазона выбора ячеек
            date_num_val, month_num_val = None, None

            for day_num in range(initial_row_nam, len(xlsx_sheet.col(group_name_row.index(group_name_cell)))):

                month_num_col = xlsx_sheet.cell(day_num, group_name_row.index(group_name_cell) - 2)
                if month_num_col.value != '':
                    month_num_val = get_month_num(month_num_col.value)

                date_num_col = xlsx_sheet.cell(day_num, group_name_row.index(group_name_cell) - 1)
                if date_num_col.value != '':
                    if isinstance(date_num_col.value, float):
                        temp_date_num_val = str(round(date_num_col.value))
                    else:
                        temp_date_num_val = str(date_num_col.value)

                    date_num_val = re.findall(r'\d+', temp_date_num_val)

                    if date_num_val:
                        date_num_val = int(date_num_val[0])
                    else:
                        break

                date_range.append([date_num_val, month_num_val, day_num])

            date_range = fix_date_range(date_range)

            date_range_dict = {}
            for row in date_range:
                this_row_date = datetime.date(datetime.datetime.now().year, row[1], row[0])
                if this_row_date.strftime("%d.%m") not in date_range_dict:
                    date_range_dict[this_row_date.strftime("%d.%m")] = []

                date_range_dict[this_row_date.strftime("%d.%m")].append(row[2])

            return date_range_dict

        print(xlsx_parh)
        book = xlrd.open_workbook(xlsx_parh)
        sheet = book.sheet_by_index(0)

        DOC_TYPE_EXAM = 2
        column_range = []
        timetable = {}
        group_list = []

        # Индекс строки с названиями групп
        group_name_row_num = 1
        # Поиск строки, содержащей названия групп
        for row_index in range(len(sheet.col(1))):
            group_name_row = sheet.row_values(row_index)
            if len(group_name_row) > 0:
                group_row_str = " ".join(str(x) for x in group_name_row)
                gr = re.findall(r"([А-Я]+-\w+-\w+)", group_row_str, re.I)
                if gr:
                    group_name_row_num = row_index
                    break

        group_name_row = sheet.row(group_name_row_num)
        for group_cell in group_name_row:  # Поиск названий групп
            group = str(group_cell.value)
            group = re.search(r"([А-Я]+-\w+-\w+)", group)
            if group:  # Если название найдено, то получение расписания этой группы
                print(group.group(0))
                # обновляем column_range, если левее группы нет разметки с неделями, используем старый
                if not group_list and doc_type != DOC_TYPE_EXAM:
                    column_range = get_column_range_semester(sheet, group_cell, group_name_row_num)
                elif not group_list and doc_type == DOC_TYPE_EXAM:
                    column_range = get_column_range_exam(sheet, group_cell, group_name_row_num)

                group_list.append(group.group(0))

                if doc_type != DOC_TYPE_EXAM:
                    one_time_table = self.read_one_group_for_semester(
                        sheet, group_name_row.index(group_cell), group_name_row_num, column_range)  # По номеру столбца
                else:
                    one_time_table = self.read_one_group_for_exam(
                        sheet, group_name_row.index(group_cell), group_name_row_num, column_range)  # По номеру столбца

                # print(one_time_table)

                for key in one_time_table.keys():
                    timetable[key] = one_time_table[key]  # Добавление в общий словарь

        if write_to_json_file is not False:  # Запись в JSON файл
            self.write_to_json(timetable, doc_type)
        if write_to_csv_file is not False or write_to_db is not False:  # Запись в CSV файл, Запись в базу данных
            self.write(doc_type, timetable, write_to_csv_file, write_to_db)
        return group_list

    @staticmethod
    def format_other_cells(cell):
        cell = cell.split("\n")
        return cell

    @staticmethod
    def format_teacher_name(cell):
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
        temp_name = re.sub(r"((, *|)кроме {1,})", " кр.", temp_name, flags=re.A)
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
                        item = re.sub(r"(\d+[,. н]+)", "", item, flags=re.A)
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

    @staticmethod
    def get_para_num_from_time(time_str):
        time_dict = {
            "9:00": 1,
            "10:40": 2,
            "13:10": 3,
            "14:50": 4,
            "16:30": 5,
            "18:10": 6,
            "19:50": 7,
            "20:10": 8
        }
        if time_str in time_dict:
            return time_dict[time_str]
        else:
            return 0

    def write(self, doc_type, timetable, write_to_csv=False, write_to_db=False):
        """Запись словаря 'timetable' в CSV файл или в базу данных
            - путь к файлу базы данных
        """

        def drop_table(group):
            db_cursor.execute("DROP TABLE IF EXISTS {}".format(group))

        def create_table(group):
            """Если Таблица не создана, создать таблицу
                table_name - Название таблицы
            """
            db_cursor.execute("""SELECT count(*) FROM sqlite_master WHERE type='table' AND name='{}';""".format(group))

            is_create = db_cursor.fetchall()[0][0]

            if is_create != 1:
                db_cursor.execute("""CREATE TABLE {} (doc_type NUMERIC, date TEXT, day NUMERIC, para NUMERIC, time DATE,
                          week NUMERIC, name TEXT, type TEXT, room TEXT, prepod TEXT, include TEXT, exception TEXT)""".format(
                    group))

        def data_append(group, doc_type, date, day_num, para_num, para_time, parity, discipline, para_type, room,
                        teacher,
                        include_week,
                        exception_week):
            """Добавление данных в базу данных"""
            db_cursor.execute("""INSERT INTO {} VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""".format(group),
                              (doc_type, date, day_num, para_num, para_time, parity, discipline,
                               para_type, room, teacher, include_week, exception_week))

        db_cursor = self.connect_to_db.cursor()

        with open(self.csv_file, "a", encoding="utf-8", newline="") as fh:
            if write_to_csv is not False:
                writer = csv.DictWriter(fh, fieldnames=["doc_type", "group_name", "day", "para", "time", "week", "name",
                                                        "type", "room", "prepod", "include",
                                                        "exception"], quoting=csv.QUOTE_ALL)
                writer.writeheader()
            for group_name, value in sorted(timetable.items()):
                # print(group_name, value)
                if write_to_db is not False:
                    table_name = re.findall(r"([А-Я]+-\w+-\w+)", group_name, re.I)
                    if len(table_name) > 0:
                        table_name = table_name[0]

                        table_name = table_name.replace('-', '_').replace(' ', '_').replace('(', '_') \
                                         .replace(')', '_').replace('.', '_').replace(',', '_')[0:10]

                    create_table(table_name)
                    for n_day, day_item in sorted(value.items()):
                        for n_para, para_item in sorted(day_item.items()):
                            for n_week, item in sorted(para_item.items()):
                                day = n_day.split("_")[1]
                                para = n_para.split("_")[1]
                                week = n_week.split("_")[1]
                                for dist in item:
                                    date = dist['date']
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
                                        writer.writerow(dict(group_name=group_name, day=day, para=para,
                                                             time=time, week=week, name=dist["name"], type=dist["type"],
                                                             room=dist["room"], prepod=dist["prepod"], include=include,
                                                             exception=exception))
                                    if write_to_db is not False:
                                        data_append(table_name, doc_type, date, day, para, time, week,
                                                    dist["name"], dist["type"],
                                                    dist["room"], dist["prepod"],
                                                    include, exception)
            self.connect_to_db.commit()
            db_cursor.close()

    def write_to_json(self, timetable, doc_type):
        """Запись словаря 'timetable' в JSON файл
            timetable(dict)
        """
        with open(self.json_file, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(timetable, ensure_ascii=False, indent=4))

    def read_one_group_for_semester(self, sheet, discipline_col_num, group_name_row_num, cell_range):
        """
            Получение расписания одной группы
            discipline_col_num(int): Номер столбца колонки 'Предмет'
            range(dict): Диапазон выбора ячеек
        """
        one_group = {}
        group_name = sheet.cell(group_name_row_num, discipline_col_num).value  # Название группы
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
                tmp_name = str(sheet.cell(string_index, discipline_col_num).value)
                tmp_name = self.format_name(tmp_name)

                if isinstance(tmp_name, list) and tmp_name != []:

                    para_type = sheet.cell(string_index, discipline_col_num + 1).value
                    teacher = self.format_teacher_name(sheet.cell(string_index, discipline_col_num + 2).value)
                    room = self.format_room_name(sheet.cell(string_index, discipline_col_num + 3).value)

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

                        one_para = {"date": None, "time": time, "name": name, "type": para_type, "prepod": teacher,
                                    "room": room}
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

    def read_one_group_for_exam(self, sheet, discipline_col_num, group_name_row_num, cell_range):
        """
            Получение расписания одной группы для формы экзаменационной сессии
            discipline_col_num(int): Номер столбца колонки 'Предмет'
            range(dict): Диапазон выбора ячеек
        """
        EXAM_PASS_INDEX = 0
        one_group = {}
        group_name = sheet.cell(group_name_row_num, discipline_col_num).value  # Название группы
        one_group[group_name] = {}

        for date in cell_range:
            string_index = cell_range[date][0]
            para_type, dist_name, teacher = None, None, None
            if len(cell_range[date]) > 1:
                para_type_index = cell_range[date][0]
                dist_name_index = cell_range[date][1]
                prepod_name_index = cell_range[date][2]

                para_type = sheet.cell(para_type_index, discipline_col_num).value
                dist_name = sheet.cell(dist_name_index, discipline_col_num).value
                teacher = sheet.cell(prepod_name_index, discipline_col_num).value

            time = sheet.cell(string_index, discipline_col_num + 1).value
            if isinstance(time, str):
                time = time.replace("-", ":")
            if isinstance(time, float):
                time = xlrd.xldate.xldate_as_datetime(time, 0).strftime("%H:%M")

            para_num = Reader.get_para_num_from_time(time)

            room = sheet.cell(string_index, discipline_col_num + 2).value
            if isinstance(room, float):
                room = int(room)

            one_day = {"para_{}".format(para_num): {}}

            if dist_name is not None:
                one_para = {"date": date, "time": time, "name": dist_name, "type": para_type, "prepod": teacher,
                            "room": room,
                            "include": '', "exception": ''}

                if dist_name and room:
                    if "week_{}".format(EXAM_PASS_INDEX) not in one_day["para_{}".format(para_num)]:
                        one_day["para_{}".format(para_num)][
                            "week_{}".format(EXAM_PASS_INDEX)] = []  # Инициализация списка
                    one_day["para_{}".format(para_num)]["week_{}".format(EXAM_PASS_INDEX)].append(one_para)

            # Объединение расписания
            one_group[group_name]["day_{}".format(date)] = one_day

        return one_group


if __name__ == "__main__":
    Downloader = Downloader(path_to_error_log='logs/downloadErrorLog.csv', base_file_dir='xls/')
    Downloader.download()

    reader = Reader(path_to_db="table.db")
    reader.run('xls', write_to_db=True, write_to_json_file=True)
