import pythoncom
from flask import Flask, request,send_from_directory
from api import pdfy
from api import config
from pathlib import Path
import logging
from flask_cors import CORS

import subprocess
from subprocess import Popen
devnull = subprocess.DEVNULL


# App
logger = logging.getLogger('app')
app = Flask(__name__)
CORS(app, supports_credentials=True)

# from api import config
cfg = config.cfg

# parameter defaults
defaults = { 'sheet': 'Sheet1', 'test': False }


def getQueryParams(get_param, request):
    args = request.args
    if get_param == 'test':
        return 'test' in args
    elif get_param in args:
        return args[get_param]

    return defaults[get_param]


@app.route('/api/v1/excel_to_pdf',methods=['POST'])
def excel_to_pdf():
    sheet = getQueryParams('sheet', request)
    is_test = getQueryParams('test', request)

    logger.info('> excel_to_pdf > running test job: {}'.format(is_test))
    if is_test:
        sheet = defaults['sheet'] # use default
        test_path = Path(cfg['working_file_path']) / 'test.xlsx'
        file_path = Path(cfg['working_file_path']) / pdfy.get_file_name('exl')
        logger.info('> excel_to_pdf > test {}, file {}'.format(test_path, file_path))
        file_path.write_bytes(test_path.read_bytes())
        file_path = str(file_path.resolve())
    else:
        file_path = pdfy.save_file('exl',request)

    #transforming a file into pdf
    result = pdfy.transform_to_pdf(file_path, sheet=sheet)
    logger.info('> excel_to_pdf > returning file {}'.format(result))

    if result is not None:
        result = Path(result).name
        return send_from_directory('import', result, as_attachment=True)
    return None
