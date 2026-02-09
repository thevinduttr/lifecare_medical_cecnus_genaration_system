
# src/utils/logger.py
import logging
import os
from src.utils.load_yaml import LOGS_PATH

# Ensure the directory exists
os.makedirs(LOGS_PATH, exist_ok=True)

class CustomLoggerAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        extra = kwargs.get("extra", {})
        kwargs["extra"] = extra  # Update kwargs with the extra dictionary
        return msg, kwargs

def _create_logger(portal_name):
    """
    Creates and returns a logger specific to the given portal
    
    Args:
        portal_name (str): Name of the portal (e.g., 'ngi', 'takaful')
        
    Returns:
        CustomLoggerAdapter: Logger adapter configured for the specified portal
    """
    # Create a log file path for this portal
    log_file = os.path.join(LOGS_PATH, f"{portal_name}.log")
    
    # Create a logger with the portal name
    portal_logger = logging.getLogger(portal_name)
    
    # Clear any existing handlers to avoid duplicate logs
    if portal_logger.handlers:
        portal_logger.handlers.clear()
    
    # Set log level
    portal_logger.setLevel(logging.DEBUG)
    
    # Create file handler for this portal
    file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
    
    # Set formatter
    formatter = logging.Formatter(
        "{asctime} - {levelname} - {filename}:{lineno} - {message}",
        style="{",
        datefmt="%Y-%m-%d %H:%M"
    )
    file_handler.setFormatter(formatter)
    
    # Add handler to logger
    portal_logger.addHandler(file_handler)
    
    # Set propagation to False to prevent duplicate logs
    portal_logger.propagate = False
    
    # Return a CustomLoggerAdapter for this portal logger
    return CustomLoggerAdapter(portal_logger, {})

logger = _create_logger("app")

# Pre-create loggers for each portal so they can be imported directly

nlg_logger = _create_logger("nlg")
dubaiinsurance_logger = _create_logger("dubaiinsurance")
sukoon_logger = _create_logger("sukoon")
takaful_logger = _create_logger("takaful")
union_logger = _create_logger("union")
# alsagr_logger = _create_logger("alsagr")
adnic_logger = _create_logger("adnic")
gig_logger = _create_logger("gig")
daman_logger = _create_logger("daman")
medgulf_logger = _create_logger("medgulf")
ngi_logger = _create_logger("ngi")
alittihad_logger = _create_logger("alittihad")
ison_logger = _create_logger("ison")
qatar_logger = _create_logger('qatar')
wataniatakaful_logger = _create_logger('wataniatakaful')
orient_logger = _create_logger('orient')
rak_logger = _create_logger('rak')
dni_logger = _create_logger('dni')
daman_logger = _create_logger('daman')
maxhealth_logger = _create_logger('maxhealth')
fidelity_logger = _create_logger('Fidelity')