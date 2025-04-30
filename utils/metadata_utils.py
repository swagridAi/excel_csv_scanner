# utils/metadata_utils.py
"""
Utilities for extracting metadata from files.

This module provides specialized functions for extracting metadata from 
different file types, building on the basic file operations in file_utils.py.
"""
import os
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, Union, List
import json

from config import ENCODING_SAMPLE_SIZE
from utils.error_utils import safe_operation
from utils.file_utils import (
    read_binary_sample, read_text_sample, detect_encoding,
    detect_csv_delimiter, get_file_metadata
)

# Configure logger
logger = logging.getLogger(__name__)


def get_basic_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get basic metadata for a file.
    
    A wrapper around get_file_metadata from file_utils.py for backward compatibility.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with basic metadata (size, modified date, etc.)
    """
    return get_file_metadata(file_path)


def get_excel_file_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get specialized metadata for Excel files.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Dictionary with Excel-specific metadata
    """
    # Start with basic metadata
    metadata = get_file_metadata(file_path)
    
    # Add Excel-specific flags
    metadata["is_excel"] = True
    
    # Distinguish between Excel formats
    extension = metadata["extension"].lower()
    if extension == '.xls':
        metadata["excel_format"] = "legacy"
        metadata["excel_format_description"] = "Legacy Excel Binary Format (.xls)"
    elif extension == '.xlsx':
        metadata["excel_format"] = "xml"
        metadata["excel_format_description"] = "Excel Open XML Format (.xlsx)"
    elif extension == '.xlsm':
        metadata["excel_format"] = "macro"
        metadata["excel_format_description"] = "Excel Open XML Format with Macros (.xlsm)"
    elif extension == '.xlsb':
        metadata["excel_format"] = "binary"
        metadata["excel_format_description"] = "Excel Binary Format (.xlsb)"
    else:
        metadata["excel_format"] = "unknown"
        metadata["excel_format_description"] = f"Unknown Excel Format ({extension})"
    
    # Format dates for consistent display
    for date_field in ['created_date', 'modified_date', 'accessed_date']:
        if date_field in metadata and metadata[date_field]:
            if isinstance(metadata[date_field], datetime):
                metadata[f"{date_field}_iso"] = metadata[date_field].isoformat()
                metadata[f"{date_field}_formatted"] = metadata[date_field].strftime("%Y-%m-%d %H:%M:%S")
    
    return metadata


def get_csv_file_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get specialized metadata for CSV files.
    
    Args:
        file_path: Path to the CSV file
        
    Returns:
        Dictionary with CSV-specific metadata
    """
    # Start with basic metadata
    metadata = get_file_metadata(file_path)
    
    # Add CSV-specific flags
    metadata["is_csv"] = True
    
    # Detect encoding and delimiter
    encoding = detect_encoding(file_path)
    delimiter = detect_csv_delimiter(file_path, encoding)
    
    metadata["encoding"] = encoding
    metadata["delimiter"] = delimiter
    
    # Format dates for consistent display
    for date_field in ['created_date', 'modified_date', 'accessed_date']:
        if date_field in metadata and metadata[date_field]:
            if isinstance(metadata[date_field], datetime):
                metadata[f"{date_field}_iso"] = metadata[date_field].isoformat()
                metadata[f"{date_field}_formatted"] = metadata[date_field].strftime("%Y-%m-%d %H:%M:%S")
    
    # Try to detect if file has headers
    try:
        sample = read_text_sample(file_path, 1024, encoding)
        lines = sample.splitlines()
        if len(lines) >= 2:
            # Heuristic: if first row has different data types than second row,
            # it's likely a header
            first_row = lines[0].split(delimiter)
            second_row = lines[1].split(delimiter)
            
            # Check if first row is all strings while second row has numbers
            first_row_has_numbers = any(cell.replace('.', '', 1).isdigit() for cell in first_row)
            second_row_has_numbers = any(cell.replace('.', '', 1).isdigit() for cell in second_row)
            
            metadata["likely_has_header"] = (not first_row_has_numbers and second_row_has_numbers)
        else:
            metadata["likely_has_header"] = False
    except Exception as e:
        logger.debug(f"Error detecting headers for {file_path}: {str(e)}")
        metadata["likely_has_header"] = False
    
    return metadata


def get_file_content_hash(file_path: Union[str, Path], algorithm: str = 'sha256') -> str:
    """
    Calculate a hash of the file contents.
    
    Args:
        file_path: Path to the file
        algorithm: Hash algorithm to use ('md5', 'sha1', 'sha256', etc.)
        
    Returns:
        Hash value as a hexadecimal string
    """
    import hashlib
    
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    hash_alg = getattr(hashlib, algorithm)()
    chunk_size = 8192  # 8KB chunks
    
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(chunk_size), b''):
            hash_alg.update(chunk)
    
    return hash_alg.hexdigest()


def get_mime_type(file_path: Union[str, Path]) -> str:
    """
    Detect the MIME type of a file.
    
    Args:
        file_path: Path to the file
        
    Returns:
        MIME type string
    """
    import mimetypes
    
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # Initialize MIME types database if needed
    if not mimetypes.inited:
        mimetypes.init()
    
    # Get MIME type from extension
    mime_type, _ = mimetypes.guess_type(path)
    
    # If MIME type couldn't be determined from extension, try magic
    if not mime_type:
        try:
            import magic
            mime_type = magic.from_file(path, mime=True)
        except (ImportError, Exception):
            # If python-magic isn't available or fails, use a simple mapping
            extension_map = {
                '.csv': 'text/csv',
                '.xls': 'application/vnd.ms-excel',
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
                '.xlsb': 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
            }
            mime_type = extension_map.get(path.suffix.lower(), 'application/octet-stream')
    
    return mime_type or 'application/octet-stream'


def detect_file_encoding_with_bom(file_path: Union[str, Path]) -> str:
    """
    Detect file encoding, handling BOM (Byte Order Mark) markers.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Detected encoding
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # BOM signatures
    boms = {
        b'\xef\xbb\xbf': 'utf-8-sig',       # UTF-8 with BOM
        b'\xfe\xff': 'utf-16-be',           # UTF-16 (Big Endian)
        b'\xff\xfe': 'utf-16-le',           # UTF-16 (Little Endian)
        b'\xff\xfe\x00\x00': 'utf-32-le',   # UTF-32 (Little Endian)
        b'\x00\x00\xfe\xff': 'utf-32-be',   # UTF-32 (Big Endian)
    }
    
    # Read the first few bytes to check for BOM
    try:
        with open(path, 'rb') as f:
            raw = f.read(4)  # Read 4 bytes for BOM
            
        # Check for BOM markers
        for bom, encoding in boms.items():
            if raw.startswith(bom):
                return encoding
    except Exception as e:
        logger.debug(f"Error checking for BOM in {path}: {str(e)}")
    
    # If no BOM found, use chardet for detection
    return detect_encoding(path)


def get_extended_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get comprehensive metadata for any supported file type.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with extended metadata
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    extension = path.suffix.lower()
    
    # Start with basic metadata
    metadata = get_file_metadata(path)
    
    # Add MIME type
    metadata["mime_type"] = get_mime_type(path)
    
    # Add content hash for data integrity
    try:
        metadata["content_hash"] = get_file_content_hash(path, 'sha256')
    except Exception as e:
        logger.debug(f"Error calculating hash for {path}: {str(e)}")
    
    # Add file encoding with BOM detection for text files
    if not metadata.get("is_binary", False):
        metadata["encoding"] = detect_file_encoding_with_bom(path)
    
    # Format date values for better readability
    for date_field in ['created_date', 'modified_date', 'accessed_date']:
        if date_field in metadata and metadata[date_field]:
            if isinstance(metadata[date_field], datetime):
                metadata[f"{date_field}_iso"] = metadata[date_field].isoformat()
                metadata[f"{date_field}_formatted"] = metadata[date_field].strftime("%Y-%m-%d %H:%M:%S")
    
    # Prepare structure for user information (will be populated by specific parsers)
    metadata["user_info"] = {}
    
    # Add file type specific metadata
    if extension in ['.xlsx', '.xlsm', '.xls', '.xlsb']:
        excel_metadata = get_excel_file_metadata(path)
        metadata.update(excel_metadata)
    elif extension == '.csv':
        csv_metadata = get_csv_file_metadata(path)
        metadata.update(csv_metadata)
    
    # If user_info is empty, remove it
    if not metadata["user_info"]:
        del metadata["user_info"]
    
    return metadata


def get_metadata_batch(file_paths: List[Union[str, Path]]) -> List[Dict[str, Any]]:
    """
    Get metadata for a batch of files.
    
    Args:
        file_paths: List of paths to the files
        
    Returns:
        List of dictionaries with metadata
    """
    results = []
    
    for path in file_paths:
        try:
            metadata = get_extended_metadata(path)
            results.append(metadata)
        except Exception as e:
            logger.warning(f"Error getting metadata for {path}: {str(e)}")
            # Add basic error record
            results.append({
                "file_path": str(path),
                "filename": Path(path).name if isinstance(path, (str, Path)) else str(path),
                "error": str(e)
            })
    
    return results


def extract_owner_info_from_path(file_path: Union[str, Path]) -> Optional[Dict[str, Any]]:
    """
    Attempt to extract ownership information from file path and system attributes.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with ownership information or None if not available
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    try:
        owner_info = {}
        
        # Try to get file owner on Unix-like systems
        try:
            import pwd
            import grp
            stat_info = path.stat()
            owner_info["owner_user"] = pwd.getpwuid(stat_info.st_uid).pw_name
            owner_info["owner_group"] = grp.getgrgid(stat_info.st_gid).gr_name
        except (ImportError, AttributeError, KeyError):
            # Not on Unix or user/group not found
            pass
        
        # Try to get file owner on Windows
        try:
            import win32security
            import win32con
            
            security_descriptor = win32security.GetFileSecurity(
                str(path),
                win32security.OWNER_SECURITY_INFORMATION
            )
            
            owner_sid = security_descriptor.GetSecurityDescriptorOwner()
            name, domain, _ = win32security.LookupAccountSid(None, owner_sid)
            
            owner_info["owner_user"] = name
            owner_info["owner_domain"] = domain
        except (ImportError, Exception):
            # Not on Windows or failed to get owner
            pass
        
        return owner_info if owner_info else None
        
    except Exception as e:
        logger.debug(f"Error extracting ownership info: {str(e)}")
        return None