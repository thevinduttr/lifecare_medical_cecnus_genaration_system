"""
Quick test for GIG mapper only
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.services.excel_service.gig_census_map import gig_map_census_data

try:
    print("Testing GIG mapper...")
    gig_map_census_data('default')
    print("GIG mapper completed successfully")
except Exception as e:
    print(f"GIG mapper failed: {e}")
    import traceback
    traceback.print_exc()