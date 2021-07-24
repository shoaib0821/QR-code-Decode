import logging
import os


def get_logger():
    """
    function to configure the logger
    :return:
    """
    debug_enabled = os.getenv("DEBUG", False)
    logging.basicConfig(level=logging.DEBUG if debug_enabled else logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    if debug_enabled:
        logging.debug("DEBUG logging enabled")

    return logging
