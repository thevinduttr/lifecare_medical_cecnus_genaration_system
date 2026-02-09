"""
Fixed comprehensive test for Medical RPA Census System
"""
import os
import shutil
import sys
import time

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import mappers
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

from src.utils.load_yaml import *

# Test configuration
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

def setup_directories():
    """Create all necessary directories"""
    print("=== Creating Directories ===")
    directories = [
        ADNIC_GENERATED_CENSUS_DIR, DAMAN_GENERATED_CENSUS_DIR, GIG_GENERATED_CENSUS_DIR,
        IQ2HEALTH_GENERATED_CENSUS_DIR, SUKOON_GENERATED_CENSUS_DIR, NLG_GENERATED_CENSUS_DIR,
        AURA_GENERATED_CENSUS_DIR, MAXHEALTH_GENERATED_CENSUS_DIR, 
        DUBAIINSURANCE_GENERATED_CENSUS_DIR, ISON_GENERATED_CENSUS_DIR, EMAIL_GENERATED_CENSUS_DIR
    ]
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"[OK] Created: {os.path.basename(directory)}")

def clear_output_directories():
    """Clear all output directories before testing with better file lock handling"""
    print("=== Clearing Output Directories ===")
    directories_to_clear = [
        ADNIC_GENERATED_CENSUS_DIR, DAMAN_GENERATED_CENSUS_DIR, GIG_GENERATED_CENSUS_DIR,
        IQ2HEALTH_GENERATED_CENSUS_DIR, SUKOON_GENERATED_CENSUS_DIR, NLG_GENERATED_CENSUS_DIR,
        AURA_GENERATED_CENSUS_DIR, MAXHEALTH_GENERATED_CENSUS_DIR, 
        DUBAIINSURANCE_GENERATED_CENSUS_DIR, ISON_GENERATED_CENSUS_DIR, EMAIL_GENERATED_CENSUS_DIR
    ]
    
    for directory in directories_to_clear:
        if os.path.exists(directory):
            try:
                files_to_remove = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
                for file in files_to_remove:
                    file_path = os.path.join(directory, file)
                    try:
                        os.remove(file_path)
                    except PermissionError as e:
                        print(f"[WARNING] Cannot remove {file} (file in use): {e}")
                        # Try waiting and retry once
                        time.sleep(2)
                        try:
                            os.remove(file_path)
                            print(f"[OK] Removed {file} on retry")
                        except:
                            print(f"[SKIP] Keeping {file} (Excel may still have it open)")
                    except Exception as e:
                        print(f"[WARNING] Error removing {file}: {e}")
                
                remaining_files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
                if not remaining_files:
                    print(f"[OK] Cleared: {os.path.basename(directory)}")
                else:
                    print(f"[PARTIAL] {os.path.basename(directory)} - {len(remaining_files)} files remain")
                    
            except Exception as e:
                print(f"[WARNING] Error clearing {directory}: {e}")
    
    return True

def setup_test_files():
    """Setup required test files"""
    print("=== Setting Up Test Files ===")
    
    # Check for source files in order of preference
    source_candidates = [
        os.path.join(ATTACHMENTS_SAVE_DIR, "Census_Input.xlsx"),
        os.path.join(os.getcwd(), "for lifecare.xlsx"),
        os.path.join(os.getcwd(), "Census_Input.xlsx"),
    ]
    
    source_file = None
    for candidate in source_candidates:
        if os.path.exists(candidate):
            source_file = candidate
            print(f"[INFO] Using source file: {os.path.basename(candidate)}")
            break
    
    if not source_file:
        print(f"[ERROR] No source Excel file found! Checked:")
        for candidate in source_candidates:
            print(f"         - {candidate}")
        return False
    
    # Copy source file to attachments as Census_Input.xlsx (only if different)
    dest_census = os.path.join(ATTACHMENTS_SAVE_DIR, "Census_Input.xlsx")
    if source_file != dest_census:
        shutil.copy(source_file, dest_census)
        print(f"[OK] Created: Census_Input.xlsx")
    else:
        print(f"[OK] Using existing: Census_Input.xlsx")
    
    # Create files required by different mappers
    files_to_create = [
        "Medical_Census_Input.xlsx",  # For DAMAN
        "CensusData-TEMPLATE_Common with Nationality.xlsx"  # For email mapper
    ]
    
    for filename in files_to_create:
        dest_path = os.path.join(ATTACHMENTS_SAVE_DIR, filename)
        if os.path.exists(dest_path):
            os.remove(dest_path)
        shutil.copy(dest_census, dest_path)
        print(f"[OK] Created: {filename}")
    
    # Show all files in attachments
    print(f"[INFO] Files in attachments:")
    for file in os.listdir(ATTACHMENTS_SAVE_DIR):
        if file.endswith('.xlsx'):
            size = os.path.getsize(os.path.join(ATTACHMENTS_SAVE_DIR, file))
            print(f"         - {file} ({size:,} bytes)")
    
    return True

def test_mapper(portal_name, mapper_func, output_dir, expected_filename):
    """Test a single census mapper"""
    print(f"\n{'='*50}")
    print(f"Testing {portal_name}")
    print(f"{'='*50}")
    
    try:
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"[INFO] Running {portal_name} mapper...")
        start_time = time.time()
        mapper_func('default')
        duration = time.time() - start_time
        
        output_path = os.path.join(output_dir, expected_filename)
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"[PASS] SUCCESS - {expected_filename}")
            print(f"       Size: {file_size:,} bytes, Time: {duration:.2f}s")
            return True
        else:
            print(f"[FAIL] Output file not found: {expected_filename}")
            files = os.listdir(output_dir) if os.path.exists(output_dir) else []
            if files:
                print(f"       Found files: {files}")
            return False
            
    except Exception as e:
        print(f"[FAIL] Error: {str(e)}")
        return False

def test_email_portals():
    """Test email portals mapper"""
    print(f"\n{'='*50}")
    print("Testing Email Portals")
    print(f"{'='*50}")
    
    try:
        os.makedirs(EMAIL_GENERATED_CENSUS_DIR, exist_ok=True)
        
        print(f"[INFO] Running email mapper...")
        start_time = time.time()
        email_map_census_data('default')
        duration = time.time() - start_time
        
        expected_filename = "Lifecare_Census Template.xlsx"
        output_path = os.path.join(EMAIL_GENERATED_CENSUS_DIR, expected_filename)
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"[PASS] SUCCESS - {expected_filename}")
            print(f"       Size: {file_size:,} bytes, Time: {duration:.2f}s")
            return True
        else:
            print(f"[FAIL] Output file not found: {expected_filename}")
            files = os.listdir(EMAIL_GENERATED_CENSUS_DIR) if os.path.exists(EMAIL_GENERATED_CENSUS_DIR) else []
            if files:
                print(f"       Found files: {files}")
            return False
            
    except Exception as e:
        print(f"[FAIL] Error: {str(e)}")
        return False

def main():
    """Main test function"""
    print("="*50)
    print("MEDICAL RPA CENSUS SYSTEM TEST")
    print("="*50)
    print(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Setup
    setup_directories()
    clear_output_directories()  # Clear any existing files with proper handling
    if not setup_test_files():
        return
    
    # Test all mappers
    print(f"\n{'='*50}")
    print("TESTING MAPPERS")
    print(f"{'='*50}")
    
    results = {}
    
    # Test individual mappers
    for portal_name, (mapper_func, output_dir, expected_filename) in CENSUS_MAPPING.items():
        results[portal_name] = test_mapper(portal_name, mapper_func, output_dir, expected_filename)
    
    # Test email portals
    results['EMAIL_PORTALS'] = test_email_portals()
    
    # Summary
    print(f"\n{'='*50}")
    print("SUMMARY")
    print(f"{'='*50}")
    
    passed = sum(results.values())
    total = len(results)
    
    for portal, success in results.items():
        status = "[PASS]" if success else "[FAIL]"
        print(f"{portal:15} {status}")
    
    print(f"\nResults: {passed}/{total} passed ({(passed/total)*100:.1f}%)")
    
    if passed == total:
        print("\n[SUCCESS] All mappers working!")
    else:
        failed = [p for p, s in results.items() if not s]
        print(f"\n[WARNING] Failed: {', '.join(failed)}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()