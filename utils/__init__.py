# utils/__init__.py
"""
Utility modules for Excel and CSV Inspector Tool.
"""

from .error_utils import safe_parser, safe_operation, with_fallback, ErrorCollection
from .file_utils import (
    find_files, 
    get_file_metadata, 
    read_binary_sample, 
    read_text_sample,
    detect_encoding,
    detect_csv_delimiter,
    get_line_count,
    file_scanner
)
from .metadata_utils import (
    get_basic_metadata,
    get_excel_file_metadata,
    get_csv_file_metadata,
    get_extended_metadata
)

__all__ = [
    'safe_parser', 'safe_operation', 'with_fallback', 'ErrorCollection',
    'find_files', 'get_file_metadata', 'read_binary_sample', 'read_text_sample',
    'detect_encoding', 'detect_csv_delimiter', 'get_line_count', 'file_scanner',
    'get_basic_metadata', 'get_excel_file_metadata', 'get_csv_file_metadata',
    'get_extended_metadata'
]