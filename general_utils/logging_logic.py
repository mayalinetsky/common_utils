"""
Module that encapsulates the package's logger
"""

import logging
import sys
from io import StringIO

API_LOGGER = logging.getLogger(__package__)
API_LOGGER.setLevel(logging.DEBUG)

__formatter__ = logging.Formatter('[%(asctime)s %(filename)s->%(funcName)s():%(lineno)s] %(levelname)s: %(message)s')
__handler__ = logging.StreamHandler(stream=sys.stdout)
__handler__.setFormatter(__formatter__)

API_LOGGER.addHandler(__handler__)  # prints to stdout
API_LOGGER.propagate = False


class TqdmLogFormatter(object):
    def __init__(self, logger):
        self._logger = logger

    def __enter__(self):
        self.__original_formatters = list()

        for handler in self._logger.handlers:
            self.__original_formatters.append(handler.formatter)

            handler.terminator = ''
            # todo figure out a way to print with the logger's formatter
            formatter = logging.Formatter('%(message)s')
            handler.setFormatter(formatter)

        return self._logger

    def __exit__(self, exc_type, exc_value, exc_traceback):
        for handler, formatter in zip(self._logger.handlers, self.__original_formatters):
            handler.terminator = '\n'
            handler.setFormatter(formatter)


class TqdmLogger(StringIO):
    def __init__(self, logger):
        super().__init__()

        self._logger = logger

    def write(self, buffer):
        with TqdmLogFormatter(self._logger) as logger:
            logger.info(buffer)

    def flush(self):
        pass


TQDM_API_LOGGER = TqdmLogger(API_LOGGER)
