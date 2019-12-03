import argparse


def parse(security):
    pass


def main(security):
    data = parse(security)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Scrape a moex.')
    parser.add_argument(
        '-s',
        '--security',
        help='Код базового актива.')
    args = parser.parse_args()
    main(args.security)