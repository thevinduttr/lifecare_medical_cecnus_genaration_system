import pandas as pd
from datetime import datetime
import os
from src.utils.load_yaml import ATTACHMENTS_SAVE_DIR 
from src.utils.load_yaml import ATTACHMENTS_SAVE_DIR, AURA_GENERATED_CENSUS_DIR, REFERRAL_FILE_STORE_DIR
from src.utils.support_functions import get_replaced_referral_id

# Define a function to calculate age based on DOB
def calculate_age(dob):
    today = datetime.today()
    return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))

def aura_map_census_data(id):

    census_filename = ""
    if id == 'default':
        for file_name in os.listdir(ATTACHMENTS_SAVE_DIR):
            if not file_name.startswith("Medical__") and file_name.endswith(".xlsx"):
                census_filename = file_name
        
        census_filepath = os.path.join(ATTACHMENTS_SAVE_DIR, census_filename)

    else:
        for file_name in os.listdir(os.path.join(REFERRAL_FILE_STORE_DIR, get_replaced_referral_id(id))):
            if not file_name.startswith("Medical__") and file_name.endswith(".xlsx"):
                census_filename = file_name

        census_filepath = os.path.join(REFERRAL_FILE_STORE_DIR,get_replaced_referral_id(id), census_filename)

    
    # Load the Excel file
    census_filepath = os.path.join(ATTACHMENTS_SAVE_DIR, census_filename)

    # Load the data from 'Sheet1'
    sheet1_data = pd.read_excel(census_filepath, sheet_name='Sheet1')
    # Load the data from 'Nationality_Updated'
    nationality_df = pd.read_excel(
            census_filepath, sheet_name="Nationality_Updated")

    # Mappings
    nationality_mapping = dict(
        zip(nationality_df['AL SAGR'], nationality_df['TAKAFUL EMARAT']))
    # Create the auto-incrementing 'S. No.' and 'Employee No.' columns
    sheet1_data['S. No.'] = range(1, len(sheet1_data) + 1)
    sheet1_data['Employee No'] = range(1, 1 + len(sheet1_data))  

    # Replace 'Principal' with 'Employee' in the 'Relation' column
    sheet1_data['Relation'] = sheet1_data['Relation'].replace('Principal', 'Employee')

    # Modify the 'Category' column to add "Category" before each value
    sheet1_data['Category'] = sheet1_data['Category'].apply(lambda x: f"Category {x}")

    # Map the original data to the new structure
    new_structure = pd.DataFrame({
        'S. No.': sheet1_data['S. No.'],
        'Employee No': sheet1_data['Employee No'],
        'Employee Name': sheet1_data['Beneficiary First Name'],	
        'Relationship': sheet1_data['Relation'],
        # 'Date of Birth (MM/DD/YY)': pd.to_datetime(sheet1_data['DOB']).dt.strftime('%m/%d/%y'),
        'Date of Birth (DD/MM/YY)': pd.to_datetime(sheet1_data['DOB']).dt.date,
        'Gender': sheet1_data['Gender'],
        'Marital Status': sheet1_data['Marital status'],
        'Nationality': sheet1_data['Nationality'].map(nationality_mapping),
        'Visa Issuance Emirates': sheet1_data['Visa Issued Emirates'],
        'Category': sheet1_data['Category'],
        'Member Type': sheet1_data['Salary Type'],
    })

    # Define the file path for saving
    output_file_path = os.path.join(AURA_GENERATED_CENSUS_DIR, "aura_map.xlsx")

    # Create the directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

    # Save the new structure to an Excel file
    new_structure.to_excel(output_file_path, index=False)

    print(f"File saved at: {output_file_path}")
