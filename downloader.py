from urllib import request
import sys
import subprocess
from bs4 import BeautifulSoup
import os.path
import datetime


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
    def __init__(self, path_to_error_log='errors/downloadErrorLog.csv', file_dir='xls/', file_type='xls'):
        self.url = 'http://www.mirea.ru/education/schedule-main/schedule/'
        self.path_to_error_log = path_to_error_log
        self.file_dir = file_dir
        self.file_type = file_type

    @staticmethod
    def save_file(url, path):
        """
        :param url: Путь до web страницы
        :param path: Путь с именем для сохраняемого файла
        """
        resp = request.urlopen(url).read()
        with open(path, 'wb') as file:
            file.write(resp)

    def download(self):

        response = request.urlopen(self.url)  # Запрос страницы
        site = str(response.read())  # Чтение страницы в переменную
        response.close()

        parse = BeautifulSoup(site, "html.parser")  # Объект BS с параметром парсера
        xls_list = parse.findAll('a', {"class": "xls"})  # поиск в HTML Всех классов с разметой Html

        # Списки адресов на файлы
        url_files = [x['href'].replace('https', 'http') for x in xls_list]  # Сохранение списка адресов сайтов
        progress_all = len(url_files)

        i = 0
        # Сохранение файлов
        for url_file in url_files:  # цикл по списку
            path_file = url_file.split('/')[-1]
            try:
                if path_file.split('.')[1] == self.file_type:
                    path_file = os.path.join(self.file_dir, path_file)
                    self.save_file(url_file, path_file)
                    i += 1  # Счетчик для отображения скаченных файлов в %
                    print('{} -- {}'.format(path_file, i / progress_all * 100))
                else:
                    i += 1  # Счетчик для отображения скаченных файлов в %

            except Exception as err:
                with open(self.path_to_error_log, 'a+') as file:
                    file.write(
                        str(datetime.datetime.now()) + ': ' + url_file + ' message:' + str(err) + '\n', )
                pass


if __name__ == "__main__":
    Downloader = Downloader(path_to_error_log='logs/downloadErrorLog.csv', file_dir='xls/', file_type='xlsx')
    Downloader.download()
