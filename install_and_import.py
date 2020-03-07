import logging
import subprocess
logger = logging.getLogger("Install and Import")


def install_and_import(package):
    import importlib
    try:
        importlib.import_module(package)
    except ImportError:
        logger.info("Installing %s",package)
        import pip
        subprocess.call(['pip3', 'install', package])
    finally:
        globals()[package] = importlib.import_module(package)
