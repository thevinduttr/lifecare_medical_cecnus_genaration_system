import asyncio
import base64
import json
import os
import shutil
import time
import traceback
import gc

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
    logger.info("Available mappers: %s", list(CENSUS_MAPPING.keys()) + EMAIL_PORTALS)
    
    while True:
        # Processing tracking variables
        requested_portals = []
        completed_portals = []
        failed_portals = []
        processing_errors = []
        
        try:
            # 1. Clear directories
            await clear_files()
            
            # 2. Fetch pending request
            req = fetch_pending_census_uploads()
            if not req:
                time.sleep(10)
                continue
                
            upload_id = req['id']
            logger.info(f"="*60)
            logger.info(f"Processing Request ID: {upload_id}")
            logger.info(f"="*60)
            
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
            
            # 5. Parse Portals and Other Data
            portals_json = req['portals']
            other_data_json = req.get('other_data', '{}')
            
            try:
                if isinstance(portals_json, str):
                    portals = json.loads(portals_json)
                else:
                    portals = portals_json 
                    
                requested_portals = portals.copy() if portals else []
                logger.info(f"Requested Portals ({len(requested_portals)}): {requested_portals}")
                
                # Parse other_data JSON
                try:
                    if isinstance(other_data_json, str):
                        other_data = json.loads(other_data_json) if other_data_json else {}
                    else:
                        other_data = other_data_json if other_data_json else {}
                    logger.info(f"Other Data: {other_data}")
                except Exception as parse_error:
                    logger.warning(f"Failed to parse other_data JSON: {parse_error}, using empty dict")
                    other_data = {}
                
            except Exception as e:
                error_msg = f"Failed to parse portals JSON: {e}"
                logger.error(error_msg)
                processing_errors.append(error_msg)
                update_upload_status(upload_id, "Failed")
                continue
            
            # 6. Execute Mappers with detailed tracking
            processed_mappers = set()  # Track executed mappers to avoid re-running for same group (e.g. Email)
            
            for portal in portals:
                portal_start_time = time.time()
                logger.info(f"Processing portal: {portal}")
                
                try:
                    mapper_info = get_mapper_for_portal(portal)
                    if not mapper_info:
                        error_msg = f"No mapper found for portal: {portal}"
                        logger.warning(error_msg)
                        failed_portals.append({"portal": portal, "reason": "No mapper available"})
                        processing_errors.append(error_msg)
                        continue
                    
                    func, output_dir, filename = mapper_info
                    
                    # Run mapper only if not already run (important for grouped email portals)
                    if func not in processed_mappers:
                        logger.info(f"Running mapper function for {portal}...")
                        try:
                            # Pass other_data to mappers that support it (like GIG)
                            if portal == 'GIG':
                                func('default', other_data)
                            else:
                                func('default')
                            processed_mappers.add(func)
                            logger.info(f"Mapper function completed for {portal}")
                        except Exception as e:
                            error_msg = f"Error running mapper for {portal}: {str(e)}"
                            logger.error(error_msg)
                            logger.error(f"Full traceback for {portal}: {traceback.format_exc()}")
                            failed_portals.append({"portal": portal, "reason": f"Mapper execution error: {str(e)}"})
                            processing_errors.append(error_msg)
                            continue
                    else:
                        logger.info(f"Mapper function already executed for {portal} (shared mapper)")
                    
                    # 7. Check output and Insert
                    output_path = os.path.join(output_dir, filename)
                    if os.path.exists(output_path):
                        file_size = os.path.getsize(output_path)
                        logger.info(f"Output file found for {portal}: {filename} ({file_size:,} bytes)")
                        
                        # Add a brief delay for Excel COM files to ensure they're fully closed
                        if portal in ['DAMAN', 'IQ'] and func.__name__ in ['daman_map_data', 'iq_map_data']:
                            logger.info(f"Waiting for Excel file to be fully released for {portal}...")
                            time.sleep(2.0)
                            import gc
                            gc.collect()  # Force garbage collection to help release Excel handles
                        
                        if insert_generated_census(upload_id, portal, output_path):
                            portal_duration = time.time() - portal_start_time
                            logger.info(f"✅ SUCCESS - {portal} completed in {portal_duration:.2f}s")
                            completed_portals.append(portal)
                        else:
                            error_msg = f"Failed to insert census for {portal} to database"
                            logger.error(error_msg)
                            failed_portals.append({"portal": portal, "reason": "Database insertion failed"})
                            processing_errors.append(error_msg)
                    else:
                        error_msg = f"Expected output file not found for {portal}: {output_path}"
                        logger.error(error_msg)
                        # List what files are actually in the directory
                        if os.path.exists(output_dir):
                            actual_files = os.listdir(output_dir)
                            logger.error(f"Files found in {output_dir}: {actual_files}")
                        else:
                            logger.error(f"Output directory does not exist: {output_dir}")
                        
                        failed_portals.append({"portal": portal, "reason": "Output file not generated"})
                        processing_errors.append(error_msg)
                        
                except Exception as e:
                    error_msg = f"Unexpected error processing {portal}: {str(e)}"
                    logger.error(error_msg)
                    logger.error(f"Full traceback for {portal}: {traceback.format_exc()}")
                    failed_portals.append({"portal": portal, "reason": f"Unexpected error: {str(e)}"})
                    processing_errors.append(error_msg)
            
            # 8. Generate Processing Summary
            logger.info(f"\n" + "="*60)
            logger.info(f"PROCESSING SUMMARY - Request ID: {upload_id}")
            logger.info(f"="*60)
            
            # Portal statistics
            total_requested = len(requested_portals)
            total_completed = len(completed_portals)
            total_failed = len(failed_portals)
            success_rate = (total_completed / total_requested * 100) if total_requested > 0 else 0
            
            logger.info(f"Total Requested Portals: {total_requested}")
            logger.info(f"Successfully Completed: {total_completed}")
            logger.info(f"Failed: {total_failed}")
            logger.info(f"Success Rate: {success_rate:.1f}%")
            
            # Detailed results
            if completed_portals:
                logger.info(f"\n✅ COMPLETED PORTALS ({len(completed_portals)}):")
                for portal in completed_portals:
                    logger.info(f"   - {portal}")
            
            if failed_portals:
                logger.error(f"\n❌ FAILED PORTALS ({len(failed_portals)}):")
                for failed in failed_portals:
                    logger.error(f"   - {failed['portal']}: {failed['reason']}")
            
            # Check for any requested portals that weren't processed
            not_processed = [p for p in requested_portals if p not in completed_portals and p not in [f['portal'] for f in failed_portals]]
            if not_processed:
                logger.warning(f"\n⚠️  PORTALS NOT PROCESSED ({len(not_processed)}):")
                for portal in not_processed:
                    logger.warning(f"   - {portal}: Not attempted")
            
            # Log all errors encountered
            if processing_errors:
                logger.error(f"\n🚨 PROCESSING ERRORS ({len(processing_errors)}):")
                for idx, error in enumerate(processing_errors, 1):
                    logger.error(f"   {idx}. {error}")
            
            # Final status determination
            if total_completed == total_requested and not processing_errors:
                final_status = "Completed"
                logger.info(f"\n🎉 ALL PORTALS COMPLETED SUCCESSFULLY!")
            elif total_completed > 0:
                final_status = "Partial"
                logger.warning(f"\n⚠️  PARTIAL SUCCESS: {total_completed}/{total_requested} portals completed")
            else:
                final_status = "Failed"
                logger.error(f"\n💥 ALL PORTALS FAILED")
            
            # 9. Update Final Status
            update_upload_status(upload_id, final_status)
            logger.info(f"\nRequest {upload_id} finished with status: {final_status}")
            logger.info(f"="*60 + "\n")
            
            # Sleep briefly before next poll
            time.sleep(5)
            
        except Exception as e:
            error_msg = f"Critical error in main loop: {e}"
            logger.error(error_msg)
            logger.error(f"Full traceback: {traceback.format_exc()}")
            
            # Try to update status if we have upload_id
            try:
                if 'upload_id' in locals():
                    update_upload_status(upload_id, "Failed")
                    logger.error(f"Request {upload_id} marked as failed due to critical error")
            except Exception as status_error:
                logger.error(f"Failed to update status after critical error: {status_error}")
            
            time.sleep(10)

if __name__ == "__main__":
    try:
        asyncio.run(run_census_loop())
    except KeyboardInterrupt:
        print("Shutting down...")