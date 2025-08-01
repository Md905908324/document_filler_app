# In utils/formatter.py
import re
import pandas as pd
import logging
import os

logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "Documents", "document_filler.log"),
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

datetime_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$")

def sanitize_key(k):
    return k.replace(" ", "_").replace("(", "").replace(")", "").replace(",", "").replace("'", "")

def format_value(val):
    try:
        dt = pd.to_datetime(val, errors='coerce')
        if pd.notna(dt):
            return dt.strftime("%m/%d/%Y")
    except:
        pass
    return str(val)