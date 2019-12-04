import argparse
import datetime
import time
import threading

from functools import wraps


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
        return datetime.datetime.strptime(s, "%d-%m-%Y")
    except ValueError:
        msg = "Неправильный формат даты: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)


@rate_limited(1)
def parse(security, date):
    base_url = "https://www.moex.com/api/contract/OpenOptionService/"


def main(security, date_from, date_to, path):
    if not date_to:
        date_to = date_from
    date_delta = date_to - date_from
    data = []
    for date in range(date_delta.days + 1):
        day = date_from + datetime.timedelta(days=date)
        parse(security, day.strftime("%d-%m-%Y"))


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
        help="Дата от (включительно)- формат DD-MM-YYYY",
        type=valid_date,
        default=datetime.date.today()
    )
    parser.add_argument(
        "-dt",
        "--dateto",
        help="Дата до (включительно)- формат DD-MM-YYYY",
        type=valid_date,
    )
    parser.add_argument(
        "-p",
        "--path",
        help="Директория для сохранения - ",
    )
    args = parser.parse_args()
    main(args.security, args.datefrom, args.dateto, args.path)
