import os
import threading
import logging
import logging.config
formatter = logging.Formatter(fmt='%(asctime)s  %(levelname)-8s [%(threadName)s-%(name)s] %(message)s',
                              datefmt="%Y-%m-%d %H:%M:%S")


def clean_up_logs(file_path):
    files = os.listdir(file_path)
    for file in files:
        if 'lookup' in file and file.endswith('.log'):
            os.remove(os.path.join(file_path, file))
        if file == "MainThread.log":
            os.remove(os.path.join(file_path, file))


class ThreadLogFilter(logging.Filter):
    """
    This filter only show log entries for specified thread name
    """

    def __init__(self, thread_name, *args, **kwargs):
        logging.Filter.__init__(self, *args, **kwargs)
        self.thread_name = thread_name

    def filter(self, record):
        return record.threadName == self.thread_name


def start_thread_logging(file_path):
    """
    Add a log handler to separate file for current thread
    """

    log_handler = None

    if logging.root.level == logging.DEBUG:
        thread_name = threading.Thread.getName(threading.current_thread())
        log_file = file_path + thread_name + '.log'
        log_handler = logging.FileHandler(log_file)

        log_handler.setLevel(logging.DEBUG)

        my_formatter = logging.Formatter(fmt='%(asctime)s  %(levelname)-8s [%(threadName)s-%(name)s] %(message)s',
                                         datefmt="%Y-%m-%d %H:%M:%S")
        log_handler.setFormatter(my_formatter)

        log_filter = ThreadLogFilter(thread_name)
        log_handler.addFilter(log_filter)

        logger = logging.getLogger()
        logger.addHandler(log_handler)

    return log_handler


def stop_thread_logging(log_handler):

    if log_handler is not None:
        # Remove thread log handler from root logger
        logging.getLogger().removeHandler(log_handler)

        # Close the thread log handler so that the lock on log file can be released
        log_handler.close()


def config_root_logger(file_path, log_level):
    log_file = file_path+'DB.log'

    root_formatter = '%(asctime)s  %(levelname)-8s [%(threadName)s-%(name)s] %(message)s'
    date_format = "%Y-%m-%d %H:%M:%S"

    logging.config.dictConfig({
        'version': 1,
        'formatters': {
            'root_formatter': {
                'format': root_formatter,
                'datefmt': date_format
            }
        },
        'handlers': {
            'log_file': {
                'class': 'logging.FileHandler',
                'level': log_level,
                'filename': log_file,
                'formatter': 'root_formatter',
            }
        },
        'loggers': {
            '': {
                'handlers': [
                    'log_file',
                ],
                'level': log_level,
                'propagate': True
            }
        }
    })
