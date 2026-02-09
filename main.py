import asyncio
import base64
import json
import os
import shutil
import time
import traceback

from src.services.db_service.census_db_service import (
    fetch_pending_census_uploads,
    update_upload_status,
    insert_generated_census,
    save_base64_to_file
)
from src.utils.load_yaml import (
    ATTACHMENTS_SAVE_DIR,
    ADNIC_GENERATED_CENSUS_DIR,
    DAMAN_GENERATED_CENSUS_DIR,
    GIG_GENERATED_CENSUS_DIR,
    IQ2HEALTH_GENERATED_CENSUS_DIR,
    SUKOON_GENERATED_CENSUS_DIR,
    NLG_GENERATED_CENSUS_DIR,
    AURA_GENERATED_CENSUS_DIR,
    MAXHEALTH_GENERATED_CENSUS_DIR,
    DUBAIINSURANCE_GENERATED_CENSUS_DIR,
    ISON_GENERATED_CENSUS_DIR,
    EMAIL_GENERATED_CENSUS_DIR
)
from src.utils.clear_folder import clear_files
from src.utils.logger import logger

# Import Census Mapping Functions
from src.services.excel_service.adnic_census_map import adnic_map_census_data
from src.services.excel_service.daman_census_map import daman_map_census_data
from src.services.excel_service.gig_census_map import gig_map_census_data
from src.services.excel_service.iq_census_map import iq_map_census_data
from src.services.excel_service.sukoon_census_map import sukoon_map_census_data
from src.services.excel_service.nlg_census_map import nlg_map_census_data
from src.services.excel_service.aura_census_map import aura_map_census_data
from src.services.excel_service.maxHealth_census_map import maxHealth_map_census_data
from src.services.excel_service.dubai_census_map import dubai_map_census_data
from src.services.excel_service.ison_census_map import ison_map_census_data
from src.services.excel_service.emails_cencus_map import email_map_census_data

# Mapping of Portal Names to (Function, OutputDir, OutputFilename)
CENSUS_MAPPING = {
    "ADNIC": (adnic_map_census_data, ADNIC_GENERATED_CENSUS_DIR, "MemberUpload.xlsx"),
    "DAMAN": (daman_map_census_data, DAMAN_GENERATED_CENSUS_DIR, "SME_Member_Details_Template.xlsx"),
    "GIG": (gig_map_census_data, GIG_GENERATED_CENSUS_DIR, "gig_map.xlsx"),
    "IQ": (iq_map_census_data, IQ2HEALTH_GENERATED_CENSUS_DIR, "Census_Template_AE.xlsm"),
    "SUKOON": (sukoon_map_census_data, SUKOON_GENERATED_CENSUS_DIR, "MemberCensusData.xlsx"),
    "NLG": (nlg_map_census_data, NLG_GENERATED_CENSUS_DIR, "MemberUpload.xlsx"),
    "AURA": (aura_map_census_data, AURA_GENERATED_CENSUS_DIR, "aura_map.xlsx"),
    "MAXHEALTH": (maxHealth_map_census_data, MAXHEALTH_GENERATED_CENSUS_DIR, "MaxHealth.xlsx"),
    "DUBAIINSURANCE": (dubai_map_census_data, DUBAIINSURANCE_GENERATED_CENSUS_DIR, "Dubaiinsurance_map.xlsx"),
    "ISON": (ison_map_census_data, ISON_GENERATED_CENSUS_DIR, "ison_map.xlsx"),
}

# Email Portals Mapping
EMAIL_PORTALS = [
    "ALLIANZ", "BUPA", "CIGNA", "HANSE_MERKUR", "NOW_HEALTH", 
    "APRIL_INTERNATIONAL", "QATAR_INSURANCE"
]

def get_mapper_for_portal(portal_name):
    """
    Returns the mapper tuple for a given portal name.
    Handles standard mappings and email portal group.
    """
    if portal_name in CENSUS_MAPPING:
        return CENSUS_MAPPING[portal_name]
    
    if portal_name in EMAIL_PORTALS:
        return (email_map_census_data, EMAIL_GENERATED_CENSUS_DIR, "Lifecare_Census Template.xlsx")
    
    return None

async def run_census_loop():
    logger.info("Starting Census Processing Loop...")
    while True:
        try:
            # 1. Clear directories
            await clear_files()
            
            # 2. Fetch pending request
            req = fetch_pending_census_uploads()
            if not req:
                time.sleep(10)
                continue
                
            upload_id = req['id']
            logger.info(f"Processing Request ID: {upload_id}")
            
            # 3. Update status to Processing
            update_upload_status(upload_id, "Processing")
            
            # 4. Save Input File (Multiple copies to satisfy different mappers)
            # Standard input required by most mappers
            input_path = os.path.join(ATTACHMENTS_SAVE_DIR, "Census_Input.xlsx")
            save_base64_to_file(req['census_file'], input_path)
            
            # Copy for mappers expecting "Medical_" prefix (DAMAN)
            medical_copy_path = os.path.join(ATTACHMENTS_SAVE_DIR, "Medical_Census_Input.xlsx")
            shutil.copy(input_path, medical_copy_path)
            
            # Copy for Email mapper expecting specific name
            email_copy_path = os.path.join(ATTACHMENTS_SAVE_DIR, "CensusData-TEMPLATE_Common with Nationality.xlsx")
            shutil.copy(input_path, email_copy_path)
            
            # 5. Parse Portals
            portals_json = req['portals']
            try:
                if isinstance(portals_json, str):
                    portals = json.loads(portals_json)
                else:
                    portals = portals_json 
            except Exception as e:
                logger.error(f"Failed to parse portals JSON: {e}")
                update_upload_status(upload_id, "Failed")
                continue
                
            logger.info(f"Requested Portals: {portals}")
            
            # 6. Execute Mappers
            has_error = False
            processed_mappers = set() # Track executed mappers to avoid re-running for same group (e.g. Email)
            
            for portal in portals:
                mapper_info = get_mapper_for_portal(portal)
                if not mapper_info:
                    logger.warning(f"No mapper found for portal: {portal}")
                    continue
                
                func, output_dir, filename = mapper_info
                
                # Run mapper only if not already run (important for grouped email portals)
                if func not in processed_mappers:
                    logger.info(f"Running mapper for {portal}...")
                    try:
                        func('default')
                        processed_mappers.add(func)
                    except Exception as e:
                        logger.error(f"Error running mapper for {portal}: {e}")
                        logger.error(traceback.format_exc())
                        has_error = True
                        continue
                
                # 7. Check output and Insert
                output_path = os.path.join(output_dir, filename)
                if os.path.exists(output_path):
                    if insert_generated_census(upload_id, portal, output_path):
                        logger.info(f"Generated census for {portal} saved to DB.")
                    else:
                        logger.error(f"Failed to insert census for {portal} to DB.")
                        has_error = True
                else:
                    logger.error(f"Expected output file not found for {portal}: {output_path}")
                    has_error = True
            
            # 8. Update Final Status
            final_status = "Failed" if has_error else "Completed"
            update_upload_status(upload_id, final_status)
            logger.info(f"Request {upload_id} finished with status: {final_status}")
            
            # Sleep briefly before next poll
            time.sleep(5)
            
        except Exception as e:
            logger.error(f"Critical error in main loop: {e}")
            logger.error(traceback.format_exc())
            time.sleep(10)

if __name__ == "__main__":
    try:
        asyncio.run(run_census_loop())
    except KeyboardInterrupt:
        print("Shutting down...")