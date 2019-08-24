import re
import json
import csv
import sqlite3
import os.path
import sys
import subprocess


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
    """Класс для парсинга расписания MIREA из xls файлов"""

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
        timetable = {}
        group_list = []
        # строка с названиями групп
        row = self.sheets.row(1)
        for is_group in row:  # Поиск названий групп
            group = str(is_group.value)
            group = re.search(r"([А-Я]+-\w+-\w+)", group)
            if group:  # Если название найдено, то получение расписания этой группы
                print(group.group(0))
                group_list.append(group.group(0))
                one_time_table = self.read_one_group(row.index(is_group))  # По номеру столбца
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
        cell = re.sub(r'\n', '  ', cell, flags=re.A)
        cell = cell.split('~')
        return cell

    # TODO: хз что делать с этими гребанными регулярками гребанный re.sub использует FullMach

    @staticmethod
    def format_name(temp_name):
        """Разбор строки 'Предмет' на название дисциплины и номера
            недель включения и исключения
            temp_name(str)
        """

        def if_diapason_week(item):
            start_week = re.findall(r"\d+-", item)
            start_week = re.sub("-", "", start_week[0])
            end_week = re.findall(r"-\d+", item)
            end_week = re.sub("-", "", end_week[0])
            weeks = []
            for i in range(int(start_week), int(end_week) + 1):
                weeks.append(i)
            return weeks

        result = []

        # temp_name = re.findall(r"(\s*[\W\s]*(?:|кр[ .]\s*|\d+-\d+|[\d,. ]*)\s*[\D\W.]\s*(?:|[\W\s]|\D)*)(?:\s\s|\Z|\n|\b)", temp_name, flags=re.A)
        # temp_name = re.findall(r"(\s*[\W\s]*(?:|кр[ .]\s*|\d+-\d+|[\d,. ]*)\s*[\D\W.]\s*(?:|[\W\s]|\D)*)(?:\s\s|\Z|\n)", temp_name, flags=re.A)

        temp_name = temp_name.replace(" ", "  ")

        temp_name = re.findall(r"((?:\s*[\W\s]*)(?:|кр[ .]\s*|\d+-\d+|[\d,. ]*)\s*\s*(?:|[\W\s]*|\D*)*)(?:\s\s|\Z|\n)",
                               temp_name, flags=re.A)
        if isinstance(temp_name, list):
            for item in temp_name:
                if len(item) > 0:
                    if_except = re.search(r"(кр[. ])", item)
                    if_include = re.search(r"( н[. ])|(н[. ])|(\d\s\W)|(\d+\s+\D)", item)
                    _except = ""
                    _include = ""
                    item = re.sub(r"\(", "", item)
                    item = re.sub(r"\)", "", item)
                    if if_except:
                        if re.search(r"\d+-\d+", item):
                            _except = if_diapason_week(item)
                            item = re.sub(r"\d+-\d+", "", item)
                        else:
                            _except = re.findall(r"(\d+)", item)
                        item = re.sub(r"(кр[. ])", "", item)
                        item = re.sub(r"(\d+[,. ]+)", "", item)
                        name = re.sub(r"( н[. ])", "", item)
                    elif if_include:
                        if re.search(r"\d+-\d+", item):
                            _include = if_diapason_week(item)
                            item = re.sub(r"\d+-\d+", "", item)
                        else:
                            _include = re.findall(r"(\d+)", item)
                        item = re.sub(r"(\d+[,. н]+)", "", item)
                        name = re.sub(r"(н[. ])", "", item)
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

        global writer, c, conn, table_name

        def create_table(table_name):
            """Если Таблица не создана, создать таблицу
                arg1 - Название таблицы
            """
            c.execute("DROP TABLE IF EXISTS {}".format(table_name))
            c.execute("""CREATE TABLE {} (day TEXT, para TEXT,
                      week TEXT, name TEXT, type TEXT, room TEXT, prepod TEXT,
                      include TEXT, exception TEXT)""".format(table_name))

        def data_append(table_name, day, para, week, name, type, room, prepod, include, exception):
            """Добавление данных в базу данных"""
            c.execute("""INSERT INTO {} ('day', 'para', 'week', 'name', 'type',
                      'room','prepod','include','exception') VALUES (?,?,?,?,?,?,?,?,?)""".format(table_name),
                      (day, para, week, name, type, room, prepod, include, exception))

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
                                    data_append(table_name, day, para, week,
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

    def read_one_group(self, col):
        """Получение расписания одной группы
            col(int): Номер столбца колонки 'Предмет'
        """
        one_group = {}
        group_name = self.sheets.cell(1, col).value  # Название группы
        one_group[group_name] = {}  # Инициализация словаря
        colimn_range = [(3, 15), (15, 27), (27, 39), (39, 51), (51, 63), (63, 75)]
        # перебор по дням недели (понедельник-пятница)
        for column in colimn_range:
            one_day = {}
            # номер дня недели (1-6)
            day_num = colimn_range.index(column) + 1
            week_num = 1
            para_num = 1
            # Перебор одного дня (1-6 пара)
            for string in range(column[0], column[1]):
                if week_num == 1:
                    one_day["para_{}".format(para_num)] = {}
                # Получение данных об одной паре
                tmp_name = str(self.sheets.cell(string, col).value)

                tmp_name = self.format_name(tmp_name)
                if isinstance(tmp_name, list):

                    para_type = self.sheets.cell(string, col + 1).value
                    prepod = self.sheets.cell(string, col + 2).value
                    room = self.sheets.cell(string, col + 3).value

                    for item in tmp_name:
                        name = item[0]
                        include = item[1]
                        exception = item[2]

                        if isinstance(room, float):
                            room = int(room)

                        one_para = {"name": name, "type": para_type, "prepod": prepod, "room": room}
                        if include:
                            one_para["include"] = include
                        if exception:
                            one_para["exception"] = exception

                        if name:
                            if "week_{}".format(week_num) not in one_day["para_{}".format(para_num)]:
                                one_day["para_{}".format(para_num)][
                                    "week_{}".format(week_num)] = []  # Инициализация списка
                            one_day["para_{}".format(para_num)]["week_{}".format(week_num)].append(one_para)

                # Изменение четности недели и инкремент пары
                if week_num == 1:
                    week_num = 2
                elif week_num == 2:
                    week_num = 1
                    para_num += 1
            # Объединение расписания
            one_group[group_name]["day_{}".format(day_num)] = one_day
        return one_group


if __name__ == "__main__":

    for i in os.scandir("xls"):
        xlsx_path = os.path.join("xls", i.name)
        print(xlsx_path)
        reader = Reader(xlsx_path, "table.db")
        res = reader.read(write_to_json_file=False, write_to_csv_file=False, write_to_db=True)

    # reader = Reader("xls/КБиСП 1 курс 2 сем .xlsx", "table.db")
    # res = reader.read(write_to_json_file=False, write_to_csv_file=False, write_to_db=True)

    # print(res)

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

    # result = reader.format_prepod_name("Козлова Г.Г.\nИсаев Р.А.")
    # print(result)
