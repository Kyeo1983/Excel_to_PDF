import os
import logging
from logging.handlers import TimedRotatingFileHandler
import xlwings as xw
from pathlib import Path
import argparse
from api import config

# Set logging
cfg = config.cfg
log_folder = cfg["log_folder"]
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler = TimedRotatingFileHandler(filename=f'{log_folder}/conversion.log', when="midnight", interval=1, backupCount=30)
handler.setFormatter(formatter)
logger = logging.getLogger('app')
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)
logger.info('Starting subprocess conversion job...')


# Parser for arguments
parser = argparse.ArgumentParser(description='Conversion Arguments.')
parser.add_argument('-f', metavar='f', type=str, nargs='?', default="",
                    help='input excel file path')
parser.add_argument('-d', metavar='d', type=str, nargs='?', default="",
                    help='destination pdf file path')
parser.add_argument('-s', metavar='s', type=str, nargs='?', default="Sheet1",
                    help='sheet to convert')
args = parser.parse_args()
print("Input args", args)
FILEPATH = args.f
DESTPATH = args.d
SHEET = args.s


if __name__ == '__main__':
    logger.info('> cpdfy > input file={}, sheet={}, dest={}'.format(FILEPATH, SHEET, DESTPATH))

    logger.info('> cpdfy > in try block')
    try:
        book = xw.Book(FILEPATH)
        sheet = book.sheets[SHEET]
        logger.info('> cpdfy > opened excel')
    except:
        logger.exception('> cpdfy > error')

    book.api.ExportAsFixedFormat(0, DESTPATH)
    logger.info('> transform_to_pdf > succeed output file = {}'.format(DESTPATH))
