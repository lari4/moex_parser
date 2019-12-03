import argparse
import datetime


def valid_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        msg = "Неправильный формат даты: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)


def parse(security, date):
    base_url = "https://www.moex.com/api/contract/OpenOptionService/"



def main(security):
    data = parse(security)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Scrape a moex.')
    parser.add_argument(
        '-s',
        '--security',
        help='Код базового актива.')
    parser.add_argument(
        "-df",
        "--datefrom",
        help="Дата от - формат YYYY-MM-DD",
        required=True,
        type=valid_date)
    parser.add_argument(
        "-dt",
        "--dateto",
        help="Дата до - формат YYYY-MM-DD",
        type=valid_date)
    args = parser.parse_args()
    main(args.security)
