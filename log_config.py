"""
Shared logging configuration for leads-bot scripts.
Logs to console (INFO) and rotating file leads_bot.log (DEBUG).
"""

import logging
import os
from logging.handlers import TimedRotatingFileHandler

_LOG_FILE = os.path.join(os.path.dirname(__file__), "leads_bot.log")


def setup_logging(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    fmt_console = logging.Formatter("%(message)s")
    fmt_file    = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(fmt_console)

    file_handler = TimedRotatingFileHandler(
        _LOG_FILE, when="midnight", backupCount=7, encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(fmt_file)

    logger.addHandler(console)
    logger.addHandler(file_handler)
    return logger
