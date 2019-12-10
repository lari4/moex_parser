import argparse
import datetime
import requests
import sys
import time
import threading
import xlsxwriter

from functools import wraps
from loguru import logger


HEADERS = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1)'
                      ' AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}


def test_internet_connection():
    try:
        response = requests.get(
            "http://google.com",
            headers=HEADERS,
        )
        if response.ok:
            return True
    except Exception as err:
        logger.debug(f"Internet connection error: {err}")
        return False


def test_moex_connection():
    try:
        response = requests.get(
            "https://www.moex.com/",
            headers=HEADERS,
        )
        if response.ok:
            return True
    except Exception as err:
        logger.debug(f"Moex connection error: {err}")
        return False


def rate_limited(max_per_second):
    """
    Decorator that make functions not be called faster than
    """
    lock = threading.Lock()
    min_interval = 1.0 / float(max_per_second)

    def decorate(func):
        last_time_called = [0.0]

        @wraps(func)
        def rate_limited_function(*args, **kwargs):
            lock.acquire()
            elapsed = time.perf_counter() - last_time_called[0]
            left_to_wait = min_interval - elapsed

            if left_to_wait > 0:
                time.sleep(left_to_wait)

            lock.release()

            ret = func(*args, **kwargs)
            last_time_called[0] = time.perf_counter()
            return ret

        return rate_limited_function

    return decorate


def valid_date(s):
    try:
        return datetime.datetime.strptime(s, "%d.%m.%Y")
    except ValueError:
        msg = "Неправильный формат даты: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)


@rate_limited(1)
def parse(security, date):
    logger.info(f"Скачивание данных для: {security} {date}")
    base_url = "https://www.moex.com/api/contract/OpenOptionService/"
    response = requests.get(
        f"{base_url}{date}/F/{security}/json",
        headers=HEADERS,
    )
    if response.ok:
        if response.json():
            return date, response.json()
    else:
        logger.error(f"Ошибка получения данных для: {security} {date}")


def save_to_excel(path, data):
    logger.info(f"Сохранение данных в файл {path}")
    try:
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet()
        for line, date_data in enumerate(data[::-1]):
            column = 0
            requested_date = date_data[0]
            worksheet.write(line, column, requested_date)
            column += 1
            for _ in date_data[1]:
                if column == 1:
                    worksheet.write(line, column, _['Date'])
                    column += 1
                worksheet.write(line, column, _['PhysicalLong'])
                column += 1
                worksheet.write(line, column, _['PhysicalShort'])
                column += 1
                worksheet.write(line, column, _['JuridicalLong'])
                column += 1
                worksheet.write(line, column, _['JuridicalShort'])
                column += 1
                worksheet.write(line, column, _['Summary'])
                column += 1
        workbook.close()
        return True
    except Exception as err:
        logger.error(f"Ошибка сохранения в файл: {err}")
        return False


def main(security, date_from, date_to, path):
    logger.info(f"Запуск скачивания с параметрами инструмент={security} с={date_from} "
                f"по={date_to} директория={path}")
    if not test_internet_connection():
        logger.error(f"Нет подключения к интернету!")
        sys.exit(0)
    if not test_moex_connection():
        logger.error(f"Нет подключения к сайту москвоской биржи!")
        sys.exit(0)
    if not date_to:
        date_to = date_from
    date_delta = date_to - date_from
    data = []
    for date in range(date_delta.days + 1):
        day = date_from + datetime.timedelta(days=date)
        parsed_data = parse(security, day.strftime("%d.%m.%Y"))
        if parsed_data:
            data.append(parsed_data)
        else:
            logger.info(f"Для {security} нет данных за {day.strftime('%d.%m.%Y')}")
    if data:
        save_to_excel(
            path=f'{security}_{date_from.strftime("%d.%m.%Y")}'
                 f'_{date_to.strftime("%d.%m.%Y")}.xlsx',
            data=data,
        )
    else:
        logger.info(f"Нет данных для сохранения")
    logger.info(f"Завершение работы")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Парсер сайта москвоской биржи.')
    parser.add_argument(
        '-s',
        '--security',
        help='Код базового актива.',
    )
    parser.add_argument(
        "-df",
        "--datefrom",
        help="Дата от (включительно)- формат DD.MM.YYYY",
        type=valid_date,
        default=datetime.date.today()
    )
    parser.add_argument(
        "-dt",
        "--dateto",
        help="Дата до (включительно)- формат DD.MM.YYYY",
        type=valid_date,
    )
    parser.add_argument(
        "-p",
        "--path",
        help="Директория для сохранения - ",
    )
    parser.add_argument(
        "-l",
        "--loglevel",
        help="Уровень логирования",
        default='INFO'
    )
    args = parser.parse_args()
    if args.loglevel == 'INFO':
        logger.remove()
        logger.add(sys.stdout, level="INFO")
    main(args.security, args.datefrom, args.dateto, args.path)
