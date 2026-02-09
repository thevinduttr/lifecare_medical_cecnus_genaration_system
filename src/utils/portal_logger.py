import logging
import os
from datetime import datetime

def setup_portal_logger(portal_name, request_id):
    """Setup a logger specific to a portal within a request ID folder"""
    
    # Create base logs directory if it doesn't exist
    base_logs_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "logs")
    os.makedirs(base_logs_dir, exist_ok=True)
    
    # Create request ID specific directory
    request_logs_dir = os.path.join(base_logs_dir, str(request_id))
    os.makedirs(request_logs_dir, exist_ok=True)
    
    # Create log file with timestamp and portal name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(request_logs_dir, f"{portal_name}_{timestamp}.log")
    
    # Create and configure logger
    logger = logging.getLogger(f"{portal_name}_{request_id}")
    logger.setLevel(logging.DEBUG)
    
    # Remove existing handlers if any
    logger.handlers = []
    
    # Create file handler
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger
