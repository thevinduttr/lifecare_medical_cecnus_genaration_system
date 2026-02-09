import base64
import os
import json
import time
import gc
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

def wait_for_file_unlock(file_path, max_attempts=10, delay=1.0):
    """
    Wait for a file to be unlocked and readable.
    Returns True if file is accessible, False if still locked after max attempts.
    """
    for attempt in range(max_attempts):
        try:
            # Try to open the file in read+write mode to check if it's locked
            with open(file_path, "rb") as f:
                f.read(1024)  # Try to read a small portion
            logger.info(f"File {os.path.basename(file_path)} is now accessible (attempt {attempt + 1})")
            return True
        except (PermissionError, IOError) as e:
            if "Permission denied" in str(e) or "being used by another process" in str(e):
                logger.warning(f"File {os.path.basename(file_path)} is locked, waiting... (attempt {attempt + 1}/{max_attempts})")
                time.sleep(delay)
                # Force garbage collection to help release any lingering file handles
                gc.collect()
            else:
                logger.error(f"Unexpected file error: {e}")
                return False
        except Exception as e:
            logger.error(f"Error checking file accessibility: {e}")
            return False
    
    logger.error(f"File {os.path.basename(file_path)} remains locked after {max_attempts} attempts")
    return False

def insert_generated_census(upload_id, portal, file_path):
    """
    Reads a file, encodes it to base64, and inserts a record into Census_Portal_Excels.
    Includes retry logic for file locking issues.
    """
    try:
        if not os.path.exists(file_path):
            logger.error(f"Generated file not found: {file_path}")
            return False
        
        # Check file size to ensure it's been fully written
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            logger.error(f"Generated file is empty: {file_path}")
            return False
        
        logger.info(f"Attempting to read census file: {os.path.basename(file_path)} ({file_size:,} bytes)")
        
        # Wait for file to be unlocked (important for Excel COM-generated files)
        if not wait_for_file_unlock(file_path):
            logger.error(f"Unable to access file due to locking: {file_path}")
            return False
        
        # Read the file with retry mechanism
        encoded_string = None
        for attempt in range(3):
            try:
                with open(file_path, "rb") as f:
                    file_content = f.read()
                    encoded_string = base64.b64encode(file_content).decode('utf-8')
                    logger.info(f"Successfully read and encoded file (attempt {attempt + 1})")
                    break
            except (PermissionError, IOError) as e:
                logger.warning(f"File read attempt {attempt + 1} failed: {e}")
                if attempt < 2:
                    time.sleep(2.0)  # Wait 2 seconds before retry
                    gc.collect()
                else:
                    raise e
        
        if not encoded_string:
            logger.error(f"Failed to read file content after retries: {file_path}")
            return False
            
        # Insert into database
        db = MySQLDatabase(DB_HOST, DB_NAME, DB_USER, DB_PASSWORD)
        if not db.connect():
            logger.error(f"Failed to connect to database for {portal}")
            return False
            
        data = {
            "upload_id": upload_id,
            "portal": portal,
            "census": encoded_string,
            "status": "Completed"
        }
        
        db.insert_record("Census_Portal_Excels", data)
        db.disconnect()
        logger.info(f"Successfully inserted census for {portal} (Upload ID: {upload_id}) - {len(encoded_string):,} chars encoded")
        return True
        
    except Exception as e:
        logger.error(f"Error inserting generated census for {portal}: {e}")
        logger.error(f"Full error details: {type(e).__name__}: {str(e)}")
        
        # Additional debugging info for permission errors
        if "Permission denied" in str(e):
            logger.error(f"File permissions issue detected. File: {file_path}")
            if os.path.exists(file_path):
                try:
                    stat_info = os.stat(file_path)
                    logger.error(f"File size: {stat_info.st_size} bytes, Modified: {time.ctime(stat_info.st_mtime)}")
                except:
                    logger.error("Unable to get file statistics")
        
        return False
