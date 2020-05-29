import os
import logging
from datetime import datetime
import xlwings as xw
from api import config
from pathlib import Path

# Set logging
logger = logging.getLogger('app')
cfg= config.cfg

def get_file_name(file_type):
     return datetime.now().strftime('%Y%m%d%H%M%S%f') + ('.xlsx' if file_type =='exl' else '.pdf')

# save a file in temporary location.
def save_file(file_type,request):

    file = request.files['file']
    dist_path =cfg['working_file_path']
    file_name = get_file_name(file_type)
    dest_path = os.path.join(dist_path,file_name)
    logger.info('> transform_to_pdf > save_file > saving file to {}'.format(dest_path))
    file.save(dest_path)
    return dest_path #str(Path(dist_path) / Path(file_name))


#transform Excel file to pdf.
def transform_to_pdf(file_path, sheet):
    logger.info('> transform_to_pdf > input var file_path={}'.format(file_path))
    dist_path = None

    try:
        logger.info('> transform_to_pdf > in try block')
        book = xw.Book(file_path)
        sheet = book.sheets[sheet]
        logger.info('> transform_to_pdf > opened excel')

        dist_path = cfg['working_file_path']
        dist_path = Path(dist_path) / Path(file_path).name
        dist_path = dist_path.with_suffix('.pdf')
        dist_path = str(dist_path.resolve())
        logger.info('> transform_to_pdf > target output file = {}'.format(dist_path))

        book.api.ExportAsFixedFormat(0, dist_path)
        logger.info('> transform_to_pdf > succeed output file = {}'.format(dist_path))

        book.close()
        logger.info('> transform_to_pdf > closed book {}'.format(file_path))
    except:
        logger.exception(">transform_to_pdf > error encountered")

    return dist_path
