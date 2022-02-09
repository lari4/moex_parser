import argparse
import datetime
import requests
import sys
import time
import threading
import xlsxwriter

from functools import wraps
from loguru import logger
from pathlib import Path, PurePath


HEADERS = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1)'
                      ' AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}


def parse_number(number):
    number = number.replace('\xa0', '')
    number = number.replace(',', '.')
    return number


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
            "https://moex.com/",
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
        def rate_limited_function(*a, **kwargs):
            lock.acquire()
            elapsed = time.perf_counter() - last_time_called[0]
            left_to_wait = min_interval - elapsed

            if left_to_wait > 0:
                time.sleep(left_to_wait)

            lock.release()

            ret = func(*a, **kwargs)
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

        title_right_border_bold = workbook.add_format()
        title_right_border_bold.set_align('center')
        title_right_border_bold.set_align('vcenter')
        title_right_border_bold.set_bold(True)
        title_right_border_bold.set_text_wrap(True)
        title_right_border_bold.set_right(5)

        title_bottom_border_bold = workbook.add_format()
        title_bottom_border_bold.set_align('center')
        title_bottom_border_bold.set_align('vcenter')
        title_bottom_border_bold.set_bold(True)
        title_bottom_border_bold.set_text_wrap(True)
        title_bottom_border_bold.set_bottom(5)

        title_right_bottom_border_bold = workbook.add_format()
        title_right_bottom_border_bold.set_align('center')
        title_right_bottom_border_bold.set_align('vcenter')
        title_right_bottom_border_bold.set_bold(True)
        title_right_bottom_border_bold.set_text_wrap(True)
        title_right_bottom_border_bold.set_right(5)
        title_right_bottom_border_bold.set_bottom(5)

        title_format = workbook.add_format()
        title_format.set_align('center')
        title_format.set_align('vcenter')
        title_format.set_bold(True)
        title_format.set_text_wrap(True)

        right_bold_border = workbook.add_format()
        right_bold_border.set_right(5)

        right_bottom_bold_border = workbook.add_format()
        right_bottom_bold_border.set_right(5)
        right_bottom_bold_border.set_bottom(5)

        bottom_bold_border = workbook.add_format()
        bottom_bold_border.set_bottom(5)

        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, '', right_bold_border)
        worksheet.write(1, 0, '', right_bold_border)
        worksheet.write(1, 1, '', right_bold_border)
        worksheet.set_column(0, 1, 12)
        worksheet.set_column(2, 20, 8)

        worksheet.merge_range(0, 1, 0, 6, "Открытые позиции", cell_format=title_right_border_bold)
        worksheet.merge_range(0, 7, 0, 11, "Изменение", cell_format=title_right_border_bold)
        worksheet.merge_range(0, 12, 0, 16, "Относительное изменение в %", cell_format=title_right_border_bold)
        worksheet.merge_range(0, 17, 0, 21, "Количество лиц", cell_format=title_right_border_bold)

        worksheet.merge_range(1, 2, 1, 3, "Физ лица", cell_format=title_format)
        worksheet.merge_range(1, 4, 1, 5, "Ю лица", cell_format=title_format)
        worksheet.write(1, 6, '', right_bold_border)
        worksheet.merge_range(1, 7, 1, 8, "Физ лица", cell_format=title_format)
        worksheet.merge_range(1, 9, 1, 10, "Ю лица", cell_format=title_format)
        worksheet.write(1, 11, '', right_bold_border)
        worksheet.merge_range(1, 12, 1, 13, "Физ лица", cell_format=title_format)
        worksheet.merge_range(1, 14, 1, 15, "Ю лица", cell_format=title_format)
        worksheet.write(1, 16, '', right_bold_border)
        worksheet.merge_range(1, 17, 1, 18, "Физ лица", cell_format=title_format)
        worksheet.merge_range(1, 19, 1, 20, "Ю лица", cell_format=title_format)
        worksheet.write(1, 21, '', right_bold_border)
        worksheet.set_row(2, 35, cell_format=title_format)
        worksheet.write(2, 0, 'Запрошенная дата', title_right_bottom_border_bold)
        worksheet.write(2, 1, 'Предоставленная дата', title_right_bottom_border_bold)

        column = 2
        for _ in range(4):
            worksheet.write(2, column, 'Длинные', title_bottom_border_bold)
            column += 1
            worksheet.write(2, column, 'Короткие', title_bottom_border_bold)
            column += 1
            worksheet.write(2, column, 'Длинные', title_bottom_border_bold)
            column += 1
            worksheet.write(2, column, 'Короткие', title_bottom_border_bold)
            column += 1
            worksheet.write(2, column, 'Итого', title_right_bottom_border_bold)
            column += 1
        for line, date_data in enumerate(data[::-1], start=3):
            if line == len(data) + 2:
                style = right_bottom_bold_border
                bottom_style = bottom_bold_border
            else:
                style = right_bold_border
                bottom_style = None
            column = 0
            requested_date = date_data[0]

            worksheet.write(line, column, requested_date, style)
            column += 1
            for _ in date_data[1]:
                if column == 1:
                    worksheet.write(line, column, _['Date'], style)
                    column += 1
                worksheet.write_number(line, column, float(parse_number(_['PhysicalLong'])), bottom_style)
                column += 1
                worksheet.write_number(line, column, float(parse_number(_['PhysicalShort'])), bottom_style)
                column += 1
                worksheet.write_number(line, column, float(parse_number(_['JuridicalLong'])), bottom_style)
                column += 1
                worksheet.write_number(line, column, float(parse_number(_['JuridicalShort'])), bottom_style)
                column += 1
                worksheet.write_number(line, column, float(parse_number(_['Summary'])), style)
                column += 1
        workbook.close()
        return True
    except Exception as err:
        logger.error(f"Ошибка сохранения в файл: {err}")
        return False


def main(security, date_from, date_to, path):
    if not date_to:
        file_name = f'Данные по {security} за {date_from.strftime("%d.%m.%Y")}.xlsx'
        date_to = date_from
    else:
        file_name = f'Данные по {security} за {date_from.strftime("%d.%m.%Y")}' \
                    f'-{date_to.strftime("%d.%m.%Y")}.xlsx'
    logger.info(f"Запуск скачивания с параметрами инструмент={security}"
                f" с={date_from.strftime('%d.%m.%Y')}"
                f" по={date_to.strftime('%d.%m.%Y')} директория={path}")
    if not test_internet_connection():
        logger.error(f"Нет подключения к интернету!")
        sys.exit(0)
    if not test_moex_connection():
        logger.error(f"Нет подключения к сайту москвоской биржи!")
        sys.exit(0)
    if not path:
        path = PurePath()
    save_dir = Path(path)
    if not save_dir.exists() and not save_dir.is_dir():
        logger.error(f"Папка {path} не существует!")
        sys.exit(0)
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
        if not save_to_excel(
            path=save_dir / file_name,
            data=data,
        ):
            attempts = 0
            while attempts < 5:
                new_path = Path(input("Введите новую директорию для сохранения: "))
                if save_to_excel(
                    path=new_path / file_name,
                    data=data,
                ):
                    break
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
        required=True,
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
    print(1/0)
