from reader import Reader
from downloader import Downloader
from writer import New_to_old_table
import sys
import os.path
import json


def run(start_semester=None, input_file=None, output_file=None, groups=None):
    base_name = os.path.join(os.path.dirname(sys.argv[0]), "table.db")
    template = os.path.join(os.path.dirname(sys.argv[0]), "template.xlsx")

    if start_semester is None:
        source_file = input("Введите путь к файлу расписания в новой форме и нажмите Enter ...  ")
        out_file = input("Введите имя создаваемого файла таблицы по граппам(без расширения)...  ")
        out_file = os.path.join(os.path.dirname(sys.argv[0]), out_file + ".xlsx")
        print(out_file)
        reader = Reader(source_file, base_name)
        group_name_list = reader.read(write_to_json_file=True, write_to_csv_file=True, write_to_db=True)
        start_date = input("Введите дату начала семестра. Например: 09.02.2018 ...  ")
        print("В текущем файле имеюися следующие группы:")
        for name in group_name_list:
            print(name)
        group_list = input("Введите названия групп, которые Вы хотите видеть в таблице по группам (через запятую). "
                           "Например: БНБО-01-16, БНБО-02-16 ...   ")
        group_list = group_list.replace(" ", "").split(",")

        writer = New_to_old_table(template, base_name, out_file, start_date, group_list)
        writer.run()

    else:
        source_file = os.path.join(os.path.dirname(sys.argv[0]), input_file)
        out_file = os.path.join(os.path.dirname(sys.argv[0]), output_file)

        reader = Reader(source_file, base_name)
        group_name_list = reader.read(write_to_json_file=True, write_to_csv_file=True, write_to_db=True)
        for name in group_name_list:
            print(name)
        writer = New_to_old_table(template, base_name, out_file, start_semester, groups)
        writer.run()


if __name__ == "__main__":

    while True:
        mode = int(input("Выберите режим работы: \n"
                         "1) Конвертация на основе настроек файла config.json\n"
                         "2) Конвертация на основе пользовательского ввода (ручной ввод данных о конвертируемых файлах и т.д.)\n"
                         "3) Загрузить расписание с сайта МИРЭА и сформировать файл БД SQLite3\n"
                         "0) Завершить выполнение скрипта!\n"))

        try:
            if mode == 1:

                with open(os.path.join(os.path.dirname(sys.argv[0]), "config.json"), encoding="utf-8") as fh:
                    conf = json.loads(fh.read())
                    start_semester = conf["start_semester"]
                    del conf['start_semester']
                    for kurs in conf:
                        input_file = conf[kurs]["input_files"]
                        output_file = conf[kurs]["output_files"]
                        groups = conf[kurs]["groups"]
                        run(start_semester, input_file, output_file, groups)

                print("Конвертация успешно выполнена!\n")

            elif mode == 2:
                run()
                print("Конвертация успешно выполнена!\n")
                continue

            elif mode == 3:
                Downloader = Downloader(path_to_error_log='logs/downloadErrorLog.csv', base_file_dir='xls/')
                Downloader.download()

                reader = Reader(path_to_db="table.db")
                reader.run('xls', write_to_db=True, write_to_json_file=True)
                print("\nКонвертация успешно выполнена!\n\n")
                continue

            elif mode == 0:
                exit(0)

        except FileNotFoundError as err:
            print("Ошибка! Не найден файл шаблона 'template.xlsx' или файлы исходного расписания")
            continue
        except Exception as err:
            print(err, "\n")
            continue
