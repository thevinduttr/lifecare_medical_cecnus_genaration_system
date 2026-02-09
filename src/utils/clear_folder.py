import os
from src.utils.load_yaml import (
    ATTACHMENTS_SAVE_DIR,
    NLG_GENERATED_CENSUS_DIR,
    AURA_GENERATED_CENSUS_DIR,
    IQ2HEALTH_GENERATED_CENSUS_DIR,
    SUKOON_GENERATED_CENSUS_DIR,
    MAXHEALTH_GENERATED_CENSUS_DIR,
    ADNIC_GENERATED_CENSUS_DIR,
    GIG_GENERATED_CENSUS_DIR,
    DAMAN_GENERATED_CENSUS_DIR,
    DUBAIINSURANCE_GENERATED_CENSUS_DIR,
    ISON_GENERATED_CENSUS_DIR,
    EMAIL_GENERATED_CENSUS_DIR
)

async def remove_files_from_subfolders(directory):
    """Removes all files within subfolders of the specified directory.

    Args:
      directory: The directory path.
    """
    if not os.path.exists(directory):
        return
        
    for root, dirs, files in os.walk(directory):
        for file in files:
            try:
                os.remove(os.path.join(root, file))
            except Exception as e:
                print(f"Error removing {file}: {e}")

async def clear_files():
    """Clear all generated census files from output directories."""
    
    await remove_files_from_subfolders(ATTACHMENTS_SAVE_DIR)
    await remove_files_from_subfolders(AURA_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(NLG_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(IQ2HEALTH_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(SUKOON_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(MAXHEALTH_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(ADNIC_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(GIG_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(DAMAN_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(DUBAIINSURANCE_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(ISON_GENERATED_CENSUS_DIR)
    await remove_files_from_subfolders(EMAIL_GENERATED_CENSUS_DIR)
