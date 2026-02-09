import os
import shutil
import pandas as pd
import sys

# Ensure src can be imported
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.utils.load_yaml import ATTACHMENTS_SAVE_DIR
from src.services.excel_service.adnic_census_map import adnic_map_census_data
from src.services.excel_service.daman_census_map import daman_map_census_data
# Import other mappers as needed...

def create_dummy_files():
    """Create dummy Excel files in ATTACHMENTS_SAVE_DIR for testing."""
    os.makedirs(ATTACHMENTS_SAVE_DIR, exist_ok=True)
    
    # Dummy data
    data = {
        'Beneficiary First Name': ['John Doe', 'Jane Doe'],
        'Relation': ['Principal', 'Spouse'],
        'Gender': ['Male', 'Female'],
        'DOB': ['1990-01-01', '1992-05-20'],
        'Category': ['A', 'B'],
        'Marital status': ['Married', 'Married'],
        'Nationality': ['United Kingdom', 'India'],
        'Visa Issued Emirates': ['Dubai', 'Abu Dhabi'],
        'Salary Type': ['HSB', 'LSB'],
        'Monthly salary': [5000, 3000],
        'NLGIC Code': [123, 456]
    }
    df = pd.DataFrame(data)
    
    # Needs Nationality_Updated sheet for most mappers
    nat_data = {'AL SAGR': ['United Kingdom', 'India'], 'Updated': ['UK', 'IND']}
    nat_df = pd.DataFrame(nat_data)
    
    # Save standard census file
    file_path = os.path.join(ATTACHMENTS_SAVE_DIR, "Census_Input.xlsx")
    with pd.ExcelWriter(file_path) as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        nat_df.to_excel(writer, sheet_name='Nationality_Updated', index=False)
        
    # Create copies required by mappers
    shutil.copy(file_path, os.path.join(ATTACHMENTS_SAVE_DIR, "Medical__Census_Input.xlsx"))
    shutil.copy(file_path, os.path.join(ATTACHMENTS_SAVE_DIR, "CensusData-TEMPLATE_Common with Nationality.xlsx"))
    
    print("Dummy files created.")

def verify_mapper(name, func):
    print(f"Testing {name}...")
    try:
        func('default')
        print(f"✅ {name} executed successfully.")
    except Exception as e:
        print(f"❌ {name} failed: {e}")

if __name__ == "__main__":
    create_dummy_files()
    
    # Test a few key mappers
    verify_mapper("ADNIC", adnic_map_census_data)
    verify_mapper("DAMAN", daman_map_census_data)
    print("Verification Setup Complete. Ready to run specific tests.")
