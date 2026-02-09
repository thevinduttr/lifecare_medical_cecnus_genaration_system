import os
import datetime
from src.utils.load_yaml import TRACE_BASE_DIR

def get_trace_path(portal_name, referral_id):
    portal_lower = portal_name.lower()
    date_str = datetime.datetime.now().strftime("%d%m%y_%H%M%S")
    trace_path = os.path.join(TRACE_BASE_DIR, portal_name, f"Trace_{portal_lower}_{date_str}_{referral_id}.zip")
    os.makedirs(os.path.dirname(trace_path), exist_ok=True)
    return trace_path