import time
import re
import os
import stat
import logging
import sys
from logging.handlers import RotatingFileHandler

def setup_debug_log(nomeloja):
    log_formatter = logging.Formatter('%(asctime)s %(levelname)s - %(funcName)s : Linha %(lineno)d - %(message)s')
    debug_log_file = '{}.log'.format(nomeloja)
    logger = logging.getLogger()
    # 20MB allocated for the main debug log file, one for the current, and one backup.
    fileHandler = RotatingFileHandler(debug_log_file, mode='a', maxBytes=10*1024*1024, backupCount=1, encoding=None, delay=0)
    fileHandler.setFormatter(log_formatter)
    logger.addHandler(fileHandler)

    consoleHandler = logging.StreamHandler()
    consoleHandler.setFormatter(log_formatter)
    logger.addHandler(consoleHandler)

    logger.setLevel(logging.INFO)

    return logger