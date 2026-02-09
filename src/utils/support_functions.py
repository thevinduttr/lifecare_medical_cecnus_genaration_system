def convert_to_int(dropdownValue , excelValue):
    try:
        if excelValue is None:
            return dropdownValue
        else:
            return str(int(excelValue))
          
    except ValueError:
        return excelValue
    
def get_replaced_referral_id(value):
    try:
        return value.replace("/", "-")
    except Exception as e:
        return Exception(f"Error replacing value: {e}")
    
def get_original_referral_id(value):
    try:
        return value.replace("-", "/")
    except Exception as e:
        return Exception(f"Error replacing value: {e}")
    
import sqlite3
import mysql.connector
from datetime import datetime
import uuid

def update_abc_table(req_id, portal_name, user_email, error, db, db_type='mysql'):
    """
    Update ABC table with provided values and auto-generated timestamp and retry_id
    
    Parameters:
    - req_id: Request ID
    - portal_name: Name of the portal
    - user_email: User's email address
    - error: Error message/description
    - db: Database configuration dictionary
    - db_type: Database type (default: 'mysql')
    
    Returns:
    - bool: True if successful, False otherwise
    """
    
    # Generate current timestamp and retry_id
    current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    retry_id = str(uuid.uuid4())
    
    try:
        # MySQL connection (default)
        if db_type.lower() == 'mysql':
            conn = mysql.connector.connect(**db)
            cursor = conn.cursor()
            
            query = """
            INSERT INTO abc (req_id, portal_name, user_email, error, datetime, retry_id)
            VALUES (%s, %s, %s, %s, %s, %s)
            """
            cursor.execute(query, (req_id, portal_name, user_email, error, current_datetime, retry_id))
        
        # SQLite connection
        elif db_type.lower() == 'sqlite':
            conn = sqlite3.connect(db.get('database', 'ab.db'))
            cursor = conn.cursor()
            
            query = """
            INSERT INTO abc (req_id, portal_name, user_email, error, datetime, retry_id)
            VALUES (?, ?, ?, ?, ?, ?)
            """
            cursor.execute(query, (req_id, portal_name, user_email, error, current_datetime, retry_id))
        
        # PostgreSQL connection (using psycopg2)
        elif db_type.lower() == 'postgresql':
            import psycopg2
            conn = psycopg2.connect(**db)
            cursor = conn.cursor()
            
            query = """
            INSERT INTO abc (req_id, portal_name, user_email, error, datetime, retry_id)
            VALUES (%s, %s, %s, %s, %s, %s)
            """
            cursor.execute(query, (req_id, portal_name, user_email, error, current_datetime, retry_id))
        
        # Commit the transaction
        conn.commit()
        print(f"Record inserted successfully with retry_id: {retry_id}")
        return True
        
    except Exception as e:
        print(f"Error updating ABC table: {str(e)}")
        if 'conn' in locals():
            conn.rollback()
        return False
        
    finally:
        if 'cursor' in locals():
            cursor.close()
        if 'conn' in locals():
            conn.close()




def extract_dropdown_values(page, tpa, network, region, dropdown_selector: str, field_name: str, output_file: str, portal_name: str) -> list:
    print(f"Extracting values from: {dropdown_selector}")
    # config = BenefitConfig(tpa=tpa, network=network, region=region)
    
    try:
        dropdown = page.locator(dropdown_selector)
        
        if not dropdown.is_visible():
            raise Exception(f"Dropdown {dropdown_selector} not visible")

        options = dropdown.locator("option").all_inner_texts()
        values = [opt for opt in options if opt]

        # Format the result line
        result_line = (
            f"'Portal': {portal_name}, "
            f"'tpa': '{tpa}', "
            f"'Region': '{region}', "
            f"'Network': '{network}', "
            f"'field name': '{field_name}', "
            f"values: {values}\n"
        )

        # Write the result to the output file
        with open(output_file, "a", encoding="utf-8") as f:
            f.write(result_line)

        print(f"Extracted {len(values)} values written to {output_file}")
        print("Extracted values:", values)
        return values
        
    except Exception as e:
        print(f"Error: {e}")
        return []

