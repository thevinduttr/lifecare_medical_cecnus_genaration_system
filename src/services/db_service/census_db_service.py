import base64
import os
import json
from src.services.db_config.config import DB_HOST, DB_NAME, DB_USER, DB_PASSWORD
from src.services.db_config.db_connect import MySQLDatabase
from src.utils.logger import logger

def fetch_pending_census_uploads():
    """
    Fetch the first pending census upload request from Census_Excel_Uploads.
    Returns:
        dict: The request record or None.
    """
    try:
        db = MySQLDatabase(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD)
        if not db.connect():
            logger.error("Failed to connect to database")
            return None
        
        query = "SELECT * FROM Census_Excel_Uploads WHERE status = 'Pending' ORDER BY created_at ASC LIMIT 1"
        logger.info(f"Executing query: {query}")
        data = db.fetch_all(query)
        logger.info(f"Query returned {len(data) if data else 0} rows")
        
        if data and len(data) > 0:
            logger.info(f"Found pending request: ID={data[0].get('id')}")
        else:
            logger.info("No pending requests found")
            
        db.disconnect()
        
        if data and len(data) > 0:
            return data[0]
        return None
    except Exception as e:
        logger.error(f"Error fetching pending uploads: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return None

def update_upload_status(id, status):
    """
    Update the status of a Census_Excel_Uploads record.
    """
    try:
        db = MySQLDatabase(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD)
        if not db.connect():
            return False
            
        result = db.update_record("Census_Excel_Uploads", {"status": status}, f"id = {id}")
        db.disconnect()
        return result
    except Exception as e:
        logger.error(f"Error updating upload status for id {id}: {e}")
        return False

def save_base64_to_file(base64_str, output_path):
    """
    Decodes a base64 string and saves it to a file.
    """
    try:
        file_data = base64.b64decode(base64_str)
        with open(output_path, 'wb') as f:
            f.write(file_data)
        return True
    except Exception as e:
        logger.error(f"Error saving base64 to file {output_path}: {e}")
        return False

def insert_generated_census(upload_id, portal, file_path):
    """
    Reads a file, encodes it to base64, and inserts a record into Census_Portal_Excels.
    """
    try:
        if not os.path.exists(file_path):
            logger.error(f"Generated file not found: {file_path}")
            return False
            
        with open(file_path, "rb") as f:
            encoded_string = base64.b64encode(f.read()).decode('utf-8')
            
        db = MySQLDatabase(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD)
        if not db.connect():
            return False
            
        data = {
            "upload_id": upload_id,
            "portal": portal,
            "census": encoded_string,
            "status": "Completed"
        }
        
        db.insert_record("Census_Portal_Excels", data)
        db.disconnect()
        logger.info(f"Successfully inserted census for {portal} (Upload ID: {upload_id})")
        return True
    except Exception as e:
        logger.error(f"Error inserting generated census for {portal}: {e}")
        return False
