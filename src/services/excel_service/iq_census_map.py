import xlwings as xw
import pandas as pd
import os
from src.utils.load_yaml import ATTACHMENTS_SAVE_DIR, IQ2HEALTH_TEMPLATES_DIR, IQ2HEALTH_GENERATED_CENSUS_DIR, REFERRAL_FILE_STORE_DIR
from src.utils.support_functions import get_replaced_referral_id
 
def iq_map_census_data(id):
    try:

        census_filename = ""
        if id == 'default':
            for file_name in os.listdir(ATTACHMENTS_SAVE_DIR):
                if not file_name.startswith("Medical__") and file_name.endswith(".xlsx"):
                    census_filename = file_name
            excel_file_path = os.path.join(ATTACHMENTS_SAVE_DIR, census_filename)

        else:
            for file_name in os.listdir(os.path.join(REFERRAL_FILE_STORE_DIR, get_replaced_referral_id(id))):
                if not file_name.startswith("Medical__") and file_name.endswith(".xlsx"):
                    census_filename = file_name
            excel_file_path = os.path.join(REFERRAL_FILE_STORE_DIR, get_replaced_referral_id(id), census_filename)

        print(f"Reading Excel file from {excel_file_path}")
        excel_data_df = pd.read_excel(excel_file_path, sheet_name='Sheet1')
        nationality_df = pd.read_excel(excel_file_path, sheet_name='Nationality_Updated')
 
     
        if 'Nationality' not in excel_data_df.columns or 'IQ Portal' not in nationality_df.columns:
            raise ValueError("Required Nationlaity columns are missing from the dataframes.")
 
        # Merge the dataframes
        merged_df = pd.merge(
        excel_data_df,
        nationality_df[['AL SAGR', 'IQ Portal']],
        left_on='Nationality',
        right_on='AL SAGR',
        how='left',
        suffixes=('', '_new')
        )
 
        # Replace the Nationality values with the corresponding IQ Portal values
        merged_df['Nationality'] = merged_df['IQ Portal'].combine_first(merged_df['Nationality'])
 
        # Drop the additional columns used for merging
        merged_df = merged_df.drop(columns=['AL SAGR', 'IQ Portal'])
 
        merged_df['Salary Type'] = merged_df['Salary Type'].apply(lambda x: 'NLSB' if x == 'HSB' else 'LSB')
 
        print("Excel DataFrames loaded successfully.")
 
        # Load the existing .xlsm file using xlwings
        template_file_path = os.path.join(IQ2HEALTH_TEMPLATES_DIR, "Census_Template_AE.xlsm")
 
        with xw.App(visible=False) as app:
            wb = app.books.open(template_file_path)
            ws = wb.sheets['INPUT - Census']
            print("Inserting data to macro XLSM template...")
 
            for index, row in merged_df.iterrows():
                ws.range(f"B{index + 2}").value = row['DOB']
                ws.range(f"C{index + 2}").value = row['Relation']
                ws.range(f"D{index + 2}").value = 'Category ' + str(row['Category'])
                ws.range(f"E{index + 2}").value = row['Gender']
                ws.range(f"F{index + 2}").value = row['Marital status']
                ws.range(f"G{index + 2}").value = row['Nationality']
                ws.range(f"H{index + 2}").value = row.get('Visa Issued Emirates', 'N/A')  # Handle missing values
                ws.range(f"I{index + 2}").value = row['Salary Type']
 
            # Save the workbook with macros preserved
            output_file_path = os.path.join(IQ2HEALTH_GENERATED_CENSUS_DIR, "Census_Template_AE.xlsm")
            wb.save(output_file_path)
 
        print("Cencus Data mapping macro sheet and saving successful.")
 
    except Exception as e:
        print(f"Error: {e}")
