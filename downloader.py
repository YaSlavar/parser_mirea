from urllib import request
import sys
import subprocess
from bs4 import BeautifulSoup
import os
import os.path
import datetime
import re


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
    exit_code = install("beautifulsoup4")
    if exit_code == 0:
        import xlrd
    else:
        print("При установке пакета возникла ошибка! {}".format(exit_code))
        exit(0)


class Downloader:
    def __init__(self, path_to_error_log='errors/downloadErrorLog.csv', base_file_dir='xls/', except_types=None):
        """

        :type file_type: list
        """
        self.url = 'http://www.mirea.ru/education/schedule-main/schedule/'
        self.path_to_error_log = path_to_error_log
        self.base_file_dir = base_file_dir
        self.file_type = ['xls', 'xlsx']
        self.except_types = except_types
        self.download_dir = {
            "zach": [r'zach', r'zachety'],
            "exam": [r'zima', r'ekz', r'ekzam', r'ekzameny', r'sessiya'],
            "semester": [r'']
        }

    @staticmethod
    def save_file(url, path):
        """
        :param url: Путь до web страницы
        :param path: Путь с именем для сохраняемого файла
        """
        resp = request.urlopen(url).read()
        with open(path, 'wb') as file:
            file.write(resp)

    def get_dir(self, file_name):
        for dir_name in self.download_dir:
            for pattern in self.download_dir[dir_name]:
                if re.search(pattern, file_name, flags=re.IGNORECASE):
                    return dir_name

    def download(self):

        response = request.urlopen(self.url)  # Запрос страницы
        site = str(response.read())  # Чтение страницы в переменную
        response.close()

        parse = BeautifulSoup(site, "html.parser")  # Объект BS с параметром парсера
        xls_list = parse.findAll('a', {"class": "xls"})  # поиск в HTML Всех классов с разметой Html

        # Списки адресов на файлы
        url_files = [x['href'].replace('https', 'http') for x in xls_list]  # Сохранение списка адресов сайтов
        progress_all = len(url_files)

        count_file = 0
        # Сохранение файлов
        for url_file in url_files:  # цикл по списку
            file_name = url_file.split('/')[-1]
            # print(file_name)
            try:
                if file_name.split('.')[1] in self.file_type:
                    subdir = self.get_dir(file_name)
                    path_to_file = os.path.join(self.base_file_dir, subdir, file_name)
                    if subdir not in self.except_types:
                        if not os.path.isdir(os.path.join(self.base_file_dir, subdir)):
                            os.makedirs(os.path.join(self.base_file_dir, subdir), exist_ok=False)

                        self.save_file(url_file, path_to_file)
                    else:
                        continue

                    count_file += 1  # Счетчик для отображения скаченных файлов в %
                    print('{} -- {}'.format(path_to_file, count_file / progress_all * 100))
                else:
                    count_file += 1  # Счетчик для отображения скаченных файлов в %

            except Exception as err:
                with open(self.path_to_error_log, 'a+') as file:
                    file.write(
                        str(datetime.datetime.now()) + ': ' + url_file + ' message:' + str(err) + '\n', )
                pass


if __name__ == "__main__":
    Downloader = Downloader(path_to_error_log='logs/downloadErrorLog.csv', base_file_dir='xls/', except_types=None)
    Downloader.download()
