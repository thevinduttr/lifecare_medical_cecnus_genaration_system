import yaml
import os

# Define the path to the configuration yaml file
CONFIG_PATH = os.path.join(os.path.dirname(__file__), '../../config.yaml')

# Load the configuration from the YAML file once
with open(CONFIG_PATH, "r") as file:
    config = yaml.safe_load(file)

# Paths
ATTACHMENTS_SAVE_DIR = config['paths']['attachments_save_dir']
REFERRAL_FILE_STORE_DIR = config['paths']['referral_file_store_dir'] # Still used by some mappers even if empty
LOGS_PATH = os.path.join(os.path.dirname(__file__), '../../logs')  # Default logs path

# *** Company specific directories ***
IQ2HEALTH_GENERATED_CENSUS_DIR = config['iq2health']['generated_census_dir']
IQ2HEALTH_TEMPLATES_DIR = config['iq2health']['template_dir']

AURA_GENERATED_CENSUS_DIR = config['aura']['generated_census_dir']

NLG_GENERATED_CENSUS_DIR = config['NLG']['generated_census_dir']
NLG_TEMPLATES_DIR = config['NLG']['templates_dir']

DUBAIINSURANCE_GENERATED_CENSUS_DIR = config['dubaiinsurance']['generated_census_dir']
DUBAIINSURANCE_TEMPLATES_DIR = config['dubaiinsurance']['templates_dir']

ISON_GENERATED_CENSUS_DIR = config['ison']['generated_census_dir']
ISON_TEMPLATES_DIR = config['ison']['templates_dir']

SUKOON_GENERATED_CENSUS_DIR = config['sukoon']['generated_census_dir']
SUKOON_TEMPLATES_DIR = config['sukoon']['templates_dir']

MAXHEALTH_GENERATED_CENSUS_DIR = config['maxHealth']['generated_census_dir']
MAXHEALTH_TEMPLATES_DIR = config['maxHealth']['templates_dir']

ADNIC_GENERATED_CENSUS_DIR = config['adnic']['generated_census_dir']
ADNIC_TEMPLATES_DIR = config['adnic']['templates_dir']

GIG_GENERATED_CENSUS_DIR = config['GIG']['generated_census_dir']
GIG_TEMPLATES_DIR = config['GIG']['templates_dir']

DAMAN_GENERATED_CENSUS_DIR = config['daman']['generated_census_dir']
DAMAN_TEMPLATES_DIR = config['daman']['templates_dir']

EMAIL_GENERATED_CENSUS_DIR = config['email_portals']['generated_census_dir']
EMAIL_CENCUS_TEMPLATE_DIR = config['email_portals']['templates_dir']
