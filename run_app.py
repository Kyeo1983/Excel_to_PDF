from api import config
import logging
from logging.handlers import TimedRotatingFileHandler
cfg = config.cfg
loglvl = logging.DEBUG if cfg["log_level"] == "DEBUG" else logging.INFO
log_folder = cfg["log_folder"]
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler = TimedRotatingFileHandler(filename=f'{log_folder}/app.log', when="midnight", interval=1, backupCount=30)
handler.setFormatter(formatter)
logger = logging.getLogger('app')
logger.addHandler(handler)
logger.setLevel(loglvl)


import os
from api.excel_to_pdf import app
logger.info('Starting Excel-to-Pdf Engine...')

if __name__ == "__main__":
    app.debug = True
    host = os.environ.get('IP', '0.0.0.0')
    port = int(os.environ.get('PORT', 5001))
    app.run(host=host, port=port)
