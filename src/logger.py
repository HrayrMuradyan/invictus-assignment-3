import logging
import sys

class LevelFormatter(logging.Formatter):
    def format(self, record):
        record.level = f"[{record.levelname}]"
        return super().format(record)

def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """
    Configures the global logging settings for the application.
    """
    
    # Create the format
    date_format = '%H:%M:%S'
    log_format = '%(asctime)s - %(level)-9s - %(message)s'
    formatter = LevelFormatter(log_format, datefmt=date_format)

    # Get the root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # Clear existing handlers (prevents duplicate logs if function is called twice)
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    # Console Handler (StreamHandler)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    return logging.getLogger("CLI")