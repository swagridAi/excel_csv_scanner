# config.py
"""
Configuration settings for Excel and CSV Inspector.
"""
import os
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List

# File extensions
EXCEL_EXTENSIONS = ['.xlsx', '.xlsm', '.xls', '.xlsb']
CSV_EXTENSIONS = ['.csv']
ALL_SUPPORTED_EXTENSIONS = EXCEL_EXTENSIONS + CSV_EXTENSIONS

# Default values for Excel file metadata
EXCEL_DEFAULTS = {
    "is_excel": True,
    "sheet_count": 0,
    "row_count": 0,
    "column_count": 0,
    "has_vba": False,
    "has_pivot_table": False,
    "has_custom_xml": False,
    "has_embedded_objects": False,
    "sheets_info": [],
    "sheet_names": [],
    "excel_format": "unknown"
}

# Default values for CSV file metadata
CSV_DEFAULTS = {
    "is_csv": True,
    "row_count": 0,
    "column_count": 0,
    "encoding": "utf-8",
    "delimiter": ",",
    "has_header": True,
    "has_vba": False,
    "has_pivot_table": False,
    "has_empty_values": False,
    "empty_values_count": 0,
    "has_quotes": False
}

# Encoding detection settings
ENCODING_SAMPLE_SIZE = 4096  # Number of bytes to read for encoding detection

# CSV parsing settings
CSV_SNIFF_LINES = 5  # Number of lines to examine when detecting CSV properties
CSV_DELIMITERS_TO_GUESS = [',', ';', '\t', '|', ':']  # Possible delimiters

# Excel parsing settings
EXCEL_MAX_ROWS = 5000  # Maximum number of rows to scan in Excel sheets

# Report settings
REPORT_COLUMN_ORDER = [
    "filename", 
    "file_path", 
    "extension", 
    "size_bytes", 
    "size_kb", 
    "modified_date", 
    "created_date", 
    "is_excel", 
    "is_csv", 
    "excel_format", 
    "encoding", 
    "delimiter", 
    "sheet_count", 
    "row_count", 
    "column_count", 
    "has_vba", 
    "has_pivot_table",
    "has_custom_xml",
    "has_embedded_objects"
]

# Excel report styling
EXCEL_HEADER_STYLE = {
    "fill_color": "E9ECEF",
    "font_bold": True,
    "border": True
}

# Logging settings
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_DIR = os.path.join(os.path.expanduser('~'), '.excel_csv_inspector', 'logs')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f'inspector_{datetime.now().strftime("%Y%m%d")}.log')
DEFAULT_LOG_LEVEL = 'INFO'

def get_config() -> Dict[str, Any]:
    """
    Get the default configuration for the application.
    
    Returns:
        Dictionary with configuration settings
    """
    # Default configuration
    config = {
        # Input/output settings
        'input_folder': '.',  # Current directory
        'output_file': 'report.xlsx',  # Default output file
        
        # Processing settings
        'recursion_depth': 0,  # Default to no recursion
        'file_extensions': ALL_SUPPORTED_EXTENSIONS,
        
        # Performance settings
        'max_memory': 1024 * 1024 * 100,  # 100MB max memory usage
        'chunk_size': 1024 * 1024,  # 1MB chunk size for reading
        
        # Logging settings
        'log_level': DEFAULT_LOG_LEVEL,
        'log_file': LOG_FILE,
        
        # Feature flags
        'detect_vba': True,
        'detect_pivot_tables': True,
        'analyze_csv_structure': True,
    }
    
    return config