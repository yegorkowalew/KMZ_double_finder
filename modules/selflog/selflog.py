import logging

logger = logging.getLogger('double_finder')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('find_doubles.log')
fh.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('[%(asctime)s] [LINE:%(lineno)d] %(levelname)-8s %(message)s',"%Y-%m-%d %H:%M:%S")
formatter_console = logging.Formatter('[%(asctime)s] %(levelname)-8s %(message)s',"%Y-%m-%d %H:%M:%S")
ch.setFormatter(formatter_console)
fh.setFormatter(formatter)
logger.addHandler(ch)
logger.addHandler(fh)

if __name__ == '__main__':
# 'application' example code
    logger.debug('debug message')
    logger.info('info message')
    logger.warning('warn message')
    logger.error('error message')
    logger.critical('critical message')