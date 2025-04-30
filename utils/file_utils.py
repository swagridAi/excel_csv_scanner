# utils/file_utils.py
"""
Comprehensive utilities for file operations.

This module centralizes all file operations including finding files,
reading file content with encoding detection, and extracting metadata.
"""
import os
import sys
import mmap
import chardet
import logging
import io
import platform
from datetime import datetime
from pathlib import Path
from typing import Generator, List, Dict, Any, Optional, Union, BinaryIO, TextIO, Tuple, Callable
import shutil

from config import (
    ENCODING_SAMPLE_SIZE, 
    ALL_SUPPORTED_EXTENSIONS,
    CSV_DELIMITERS_TO_GUESS
)
from utils.error_utils import safe_operation, ErrorCollection

# Configure logger
logger = logging.getLogger(__name__)


def find_files(
    folder_path: Union[str, Path], 
    extensions: Optional[List[str]] = None, 
    recursion_depth: int = -1,
    min_size: int = 0,
    max_size: Optional[int] = None,
    include_pattern: Optional[str] = None,
    exclude_pattern: Optional[str] = None
) -> Generator[Path, None, None]:
    """
    Find all files with the given extensions in the specified folder.
    
    Args:
        folder_path: Path to the folder to scan
        extensions: List of file extensions to look for (e.g., ['.csv', '.xlsx'])
                    If None, uses all supported extensions from config
        recursion_depth: How deep to recurse (-1 for unlimited)
        min_size: Minimum file size in bytes to include (0 to disable)
        max_size: Maximum file size in bytes to include (None to disable)
        include_pattern: Optional regex pattern for filenames to include
        exclude_pattern: Optional regex pattern for filenames to exclude
        
    Yields:
        Path objects for each matching file
    """
    import re
    
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Invalid folder path: {folder_path}")
    
    # Use default extensions if none provided
    if extensions is None:
        extensions = ALL_SUPPORTED_EXTENSIONS
    
    # Ensure all extensions start with a dot and are lowercase
    extensions = [
        ext.lower() if ext.startswith('.') else f'.{ext}'.lower() 
        for ext in extensions
    ]
    
    # Compile regex patterns if provided
    include_regex = re.compile(include_pattern, re.IGNORECASE) if include_pattern else None
    exclude_regex = re.compile(exclude_pattern, re.IGNORECASE) if exclude_pattern else None
    
    def _is_target_file(path: Path) -> bool:
        """Check if a file matches all criteria."""
        # Extension check
        if not path.is_file() or path.suffix.lower() not in extensions:
            return False
        
        # Size check
        if min_size > 0 or max_size is not None:
            try:
                size = path.stat().st_size
                if size < min_size:
                    return False
                if max_size is not None and size > max_size:
                    return False
            except (OSError, IOError):
                logger.warning(f"Could not determine size for {path}")
                return False
        
        # Pattern checks
        if include_regex and not include_regex.search(path.name):
            return False
        if exclude_regex and exclude_regex.search(path.name):
            return False
        
        return True
    
    # Handle the non-recursive case
    if recursion_depth == 0:
        for item in folder.iterdir():
            if _is_target_file(item):
                yield item
        return
                
    # Handle recursive case
    for root, _, files in os.walk(folder):
        current_depth = len(Path(root).relative_to(folder).parts)
        
        # Skip if we've gone too deep
        if recursion_depth > 0 and current_depth > recursion_depth:
            continue
            
        for file in files:
            file_path = Path(root) / file
            if _is_target_file(file_path):
                yield file_path


def get_file_metadata(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get comprehensive metadata for a file.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with metadata (size, dates, type, etc.)
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    try:
        stat_info = path.stat()
        
        # Basic metadata
        metadata = {
            "file_path": str(path),
            "filename": path.name,
            "extension": path.suffix.lower(),
            "size_bytes": stat_info.st_size,
            "size_kb": round(stat_info.st_size / 1024, 2),
            "size_mb": round(stat_info.st_size / (1024 * 1024), 2),
            "size_human": get_human_size(stat_info.st_size),
            "modified_date": datetime.fromtimestamp(stat_info.st_mtime),
        }
        
        # Add creation date if available
        if hasattr(stat_info, 'st_birthtime'):  # macOS, BSD
            metadata["created_date"] = datetime.fromtimestamp(stat_info.st_birthtime)
        else:  # Linux and others use st_ctime
            metadata["created_date"] = datetime.fromtimestamp(stat_info.st_ctime)
        
        # Add access date
        metadata["accessed_date"] = datetime.fromtimestamp(stat_info.st_atime)
        
        # Get and add platform-specific file attributes
        try:
            platform_attrs = get_platform_file_attributes(path)
            if platform_attrs:
                metadata.update(platform_attrs)
        except Exception as e:
            logger.debug(f"Error getting platform-specific attributes: {str(e)}")
        
        # Get file owner information
        try:
            owner_info = get_file_owner(path)
            if owner_info:
                metadata.update(owner_info)
        except Exception as e:
            logger.debug(f"Error getting file owner: {str(e)}")
        
        # File type flags
        metadata["is_excel"] = path.suffix.lower() in ['.xls', '.xlsx', '.xlsm', '.xlsb']
        metadata["is_csv"] = path.suffix.lower() == '.csv'
        
        # Try to detect if file is empty
        metadata["is_empty"] = stat_info.st_size == 0
        
        # Try to detect if file is binary
        if not metadata["is_empty"]:
            try:
                sample = read_binary_sample(path, 1024)
                metadata["is_binary"] = is_binary_data(sample)
            except Exception:
                metadata["is_binary"] = False
        
        # Add formatted date strings for better display
        for date_field in ['created_date', 'modified_date', 'accessed_date']:
            if date_field in metadata and metadata[date_field]:
                metadata[f"{date_field}_formatted"] = metadata[date_field].strftime("%Y-%m-%d %H:%M:%S")
                metadata[f"{date_field}_iso"] = metadata[date_field].isoformat()

        return metadata
    
    except Exception as e:
        logger.error(f"Error getting metadata for {file_path}: {str(e)}")
        return {
            "file_path": str(path),
            "filename": path.name,
            "extension": path.suffix.lower() if path.suffix else "",
            "error": f"Failed to get file metadata: {str(e)}"
        }


def get_platform_file_attributes(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get platform-specific file attributes.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with platform-specific attributes
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    attrs = {}
    
    # Windows-specific attributes
    if platform.system() == 'Windows':
        try:
            import win32api
            import win32con
            
            file_attr = win32api.GetFileAttributes(str(path))
            
            attrs["is_readonly"] = bool(file_attr & win32con.FILE_ATTRIBUTE_READONLY)
            attrs["is_hidden"] = bool(file_attr & win32con.FILE_ATTRIBUTE_HIDDEN)
            attrs["is_system"] = bool(file_attr & win32con.FILE_ATTRIBUTE_SYSTEM)
            attrs["is_archive"] = bool(file_attr & win32con.FILE_ATTRIBUTE_ARCHIVE)
            attrs["is_temporary"] = bool(file_attr & win32con.FILE_ATTRIBUTE_TEMPORARY)
            attrs["is_offline"] = bool(file_attr & win32con.FILE_ATTRIBUTE_OFFLINE)
            attrs["is_encrypted"] = bool(file_attr & win32con.FILE_ATTRIBUTE_ENCRYPTED)
            
            # Get file version information if available
            try:
                info = win32api.GetFileVersionInfo(str(path), '\\')
                if info:
                    # Try to extract file version
                    ms = info['FileVersionMS']
                    ls = info['FileVersionLS']
                    attrs["file_version"] = f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
            except Exception:
                pass
        except ImportError:
            # win32api not available
            pass
    
    # Unix-specific attributes
    elif platform.system() in ['Linux', 'Darwin']:
        try:
            import stat
            
            file_stat = path.stat()
            
            # File permissions
            attrs["permissions"] = oct(file_stat.st_mode)[-3:]  # Last 3 digits of octal
            attrs["is_readonly"] = not bool(file_stat.st_mode & stat.S_IWUSR)
            attrs["is_executable"] = bool(file_stat.st_mode & stat.S_IXUSR)
            
            # Special flags (macOS only)
            if platform.system() == 'Darwin':
                try:
                    import xattr
                    
                    # Check for macOS quarantine flag
                    try:
                        quarantine = xattr.getxattr(str(path), 'com.apple.quarantine')
                        attrs["is_quarantined"] = bool(quarantine)
                    except:
                        attrs["is_quarantined"] = False
                    
                    # Check for file tags
                    try:
                        tags_attr = xattr.getxattr(str(path), 'com.apple.metadata:_kMDItemUserTags')
                        if tags_attr:
                            import plistlib
                            tags = plistlib.loads(tags_attr)
                            attrs["tags"] = tags
                    except:
                        pass
                except ImportError:
                    # xattr not available
                    pass
        except Exception as e:
            logger.debug(f"Error getting Unix file attributes: {str(e)}")
    
    return attrs


def get_file_owner(file_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get file owner information.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with owner information
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    owner_info = {}
    
    # Try to get file owner on Unix-like systems
    if platform.system() in ['Linux', 'Darwin']:
        try:
            import pwd
            import grp
            
            stat_info = path.stat()
            owner_info["owner_id"] = stat_info.st_uid
            owner_info["group_id"] = stat_info.st_gid
            
            try:
                owner = pwd.getpwuid(stat_info.st_uid)
                owner_info["owner_user"] = owner.pw_name
                owner_info["owner_gecos"] = owner.pw_gecos  # Full name and other info
            except KeyError:
                owner_info["owner_user"] = str(stat_info.st_uid)
            
            try:
                group = grp.getgrgid(stat_info.st_gid)
                owner_info["owner_group"] = group.gr_name
            except KeyError:
                owner_info["owner_group"] = str(stat_info.st_gid)
        except ImportError:
            # pwd/grp modules not available
            pass
    
    # Try to get file owner on Windows
    elif platform.system() == 'Windows':
        try:
            import win32security
            import win32con
            
            # Get security descriptor
            security_descriptor = win32security.GetFileSecurity(
                str(path),
                win32security.OWNER_SECURITY_INFORMATION | win32security.GROUP_SECURITY_INFORMATION
            )
            
            # Get owner
            owner_sid = security_descriptor.GetSecurityDescriptorOwner()
            owner_name, owner_domain, _ = win32security.LookupAccountSid(None, owner_sid)
            owner_info["owner_user"] = owner_name
            owner_info["owner_domain"] = owner_domain
            owner_info["owner_sid"] = str(owner_sid)
            
            # Try to get group
            try:
                group_sid = security_descriptor.GetSecurityDescriptorGroup()
                if group_sid:
                    group_name, group_domain, _ = win32security.LookupAccountSid(None, group_sid)
                    owner_info["owner_group"] = group_name
                    owner_info["group_domain"] = group_domain
            except:
                pass
            
            # Get additional ACL information
            try:
                dacl = security_descriptor.GetSecurityDescriptorDacl()
                if dacl:
                    # Check if the file is shared with others
                    ace_count = dacl.GetAceCount()
                    if ace_count > 1:  # More than just the owner
                        owner_info["is_shared"] = True
                        
                        # Get the first few ACEs for reference
                        users_with_access = []
                        for i in range(min(ace_count, 5)):  # Limit to first 5
                            ace = dacl.GetAce(i)
                            if ace:
                                ace_type, ace_flags, ace_mask, ace_sid = ace
                                try:
                                    user, domain, _ = win32security.LookupAccountSid(None, ace_sid)
                                    users_with_access.append(f"{domain}\\{user}")
                                except:
                                    users_with_access.append(str(ace_sid))
                        
                        if users_with_access:
                            owner_info["users_with_access"] = users_with_access
                    else:
                        owner_info["is_shared"] = False
            except Exception as e:
                logger.debug(f"Error getting ACL information: {str(e)}")
        except (ImportError, Exception) as e:
            logger.debug(f"Error getting Windows file owner: {str(e)}")
    
    return owner_info


def get_human_size(size_bytes: int) -> str:
    """
    Convert a file size in bytes to a human-readable string.
    
    Args:
        size_bytes: File size in bytes
        
    Returns:
        Human-readable file size string (e.g., "1.23 MB")
    """
    if size_bytes < 1024:
        return f"{size_bytes} bytes"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


@safe_operation(operation_name="reading binary sample")
def read_binary_sample(file_path: Union[str, Path], sample_size: int = 4096) -> bytes:
    """
    Read a binary sample from the beginning of a file.
    
    Args:
        file_path: Path to the file
        sample_size: Number of bytes to read
        
    Returns:
        Bytes containing the file sample
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    with open(path, 'rb') as f:
        return f.read(sample_size)


@safe_operation(operation_name="reading text sample")
def read_text_sample(
    file_path: Union[str, Path], 
    sample_size: int = 4096,
    encoding: Optional[str] = None
) -> str:
    """
    Read a text sample from the beginning of a file.
    
    Args:
        file_path: Path to the file
        sample_size: Number of bytes to read
        encoding: File encoding or None to auto-detect
        
    Returns:
        String containing the file sample
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # Auto-detect encoding if not specified
    if encoding is None:
        encoding = detect_encoding(path)
    
    with open(path, 'r', encoding=encoding, errors='replace') as f:
        return f.read(sample_size)


def detect_encoding(file_path: Union[str, Path]) -> str:
    """
    Detect the encoding of a file.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Detected encoding or 'utf-8' as fallback
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    try:
        # Read a binary sample for encoding detection
        sample = read_binary_sample(path, ENCODING_SAMPLE_SIZE)
        
        # Use chardet to detect encoding
        result = chardet.detect(sample)
        detected_encoding = result.get('encoding')
        confidence = result.get('confidence', 0)
        
        # Log detection results
        logger.debug(f"Detected encoding for {path.name}: {detected_encoding} (confidence: {confidence:.2f})")
        
        # Return detected encoding or utf-8 as fallback
        return detected_encoding or 'utf-8'
        
    except Exception as e:
        logger.warning(f"Error detecting encoding for {path}: {str(e)}")
        return 'utf-8'  # Default to UTF-8


def detect_csv_delimiter(file_path: Union[str, Path], encoding: Optional[str] = None) -> str:
    """
    Detect the delimiter used in a CSV file.
    
    Args:
        file_path: Path to the CSV file
        encoding: File encoding or None to auto-detect
        
    Returns:
        Detected delimiter or ',' as fallback
    """
    import csv
    
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # Auto-detect encoding if not specified
    if encoding is None:
        encoding = detect_encoding(path)
    
    try:
        # Read a sample of the file
        with open(path, 'r', encoding=encoding, errors='replace') as f:
            sample_lines = [f.readline() for _ in range(5)]
            sample = ''.join(line for line in sample_lines if line)
        
        if not sample:
            return ','
        
        # Try to use csv.Sniffer to detect the delimiter
        try:
            dialect = csv.Sniffer().sniff(sample)
            return dialect.delimiter
        except Exception:
            # If sniffer fails, try to guess based on common delimiters
            first_line = sample_lines[0]
            best_delimiter = ','
            max_count = 0
            
            for delimiter in CSV_DELIMITERS_TO_GUESS:
                count = first_line.count(delimiter)
                if count > max_count:
                    max_count = count
                    best_delimiter = delimiter
            
            return best_delimiter
            
    except Exception as e:
        logger.warning(f"Error detecting delimiter for {path}: {str(e)}")
        return ','  # Default to comma


def is_binary_data(data: bytes) -> bool:
    """
    Check if data is likely binary (non-text).
    
    Args:
        data: Bytes to check
        
    Returns:
        True if the data is likely binary, False otherwise
    """
    # Check for NULL bytes, which are unlikely in text files
    if b'\x00' in data:
        return True
    
    # Count printable vs non-printable characters
    printable = sum(1 for byte in data if 32 <= byte <= 126 or byte in (9, 10, 13))
    return printable / len(data) < 0.75 if data else False


def memory_mapped_reader(file_path: Union[str, Path]) -> mmap.mmap:
    """
    Create a memory-mapped reader for a file.
    This allows efficient random access to large files.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Memory-mapped file object
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    file_size = path.stat().st_size
    if file_size == 0:
        raise ValueError(f"Cannot memory-map empty file: {path}")
    
    fd = os.open(path, os.O_RDONLY)
    return mmap.mmap(fd, 0, access=mmap.ACCESS_READ)


def get_line_count(
    file_path: Union[str, Path], 
    encoding: Optional[str] = None,
    chunk_size: int = 1024 * 1024  # 1MB chunks
) -> int:
    """
    Efficiently count lines in a file.
    
    Args:
        file_path: Path to the file
        encoding: File encoding for text files, or None for binary
        chunk_size: Size of chunks to read at once
        
    Returns:
        Number of lines in the file
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # For very small files, just read the whole thing
    file_size = path.stat().st_size
    if file_size < 1024 * 10:  # 10KB
        if encoding:
            with open(path, 'r', encoding=encoding, errors='replace') as f:
                return sum(1 for _ in f)
        else:
            with open(path, 'rb') as f:
                return sum(1 for _ in f)
    
    # For larger files, use more efficient methods
    try:
        # First try memory mapping for efficiency
        try:
            with memory_mapped_reader(path) as mmapped:
                return mmapped.read().count(b'\n') + 1
        except (ValueError, OSError, IOError) as e:
            logger.debug(f"Memory mapping failed for {path}: {str(e)}")
            
        # Fall back to chunk reading if memory mapping fails
        newline_count = 0
        if encoding:
            with open(path, 'r', encoding=encoding, errors='replace') as f:
                for chunk in iter(lambda: f.read(chunk_size), ''):
                    newline_count += chunk.count('\n')
        else:
            with open(path, 'rb') as f:
                for chunk in iter(lambda: f.read(chunk_size), b''):
                    newline_count += chunk.count(b'\n')
                    
        # Add 1 if file doesn't end with newline
        return newline_count + 1
        
    except Exception as e:
        logger.warning(f"Error counting lines in {path}: {str(e)}")
        return 0


def create_file_handle(
    file_path: Union[str, Path],
    mode: str = 'r',
    encoding: Optional[str] = None,
    **kwargs
) -> Union[TextIO, BinaryIO]:
    """
    Create an appropriate file handle based on mode and path.
    
    Args:
        file_path: Path to the file
        mode: File open mode ('r', 'rb', 'w', etc.)
        encoding: Encoding for text modes, or None to auto-detect for reading
        **kwargs: Additional arguments to pass to open()
        
    Returns:
        File handle object
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # Auto-detect encoding for text read modes
    if 'b' not in mode and 'r' in mode and encoding is None:
        encoding = detect_encoding(path)
    
    # Handle text modes
    if 'b' not in mode:
        return open(path, mode, encoding=encoding or 'utf-8', **kwargs)
    # Handle binary modes
    else:
        return open(path, mode, **kwargs)


def iter_chunked_file(
    file_path: Union[str, Path],
    chunk_size: int = 1024 * 1024,  # 1MB chunks
    mode: str = 'r',
    encoding: Optional[str] = None,
    **kwargs
) -> Generator[Union[str, bytes], None, None]:
    """
    Iterate through a file in chunks efficiently.
    
    Args:
        file_path: Path to the file
        chunk_size: Size of chunks to read
        mode: File open mode ('r', 'rb')
        encoding: Encoding for text modes, or None to auto-detect
        **kwargs: Additional arguments to pass to open()
        
    Yields:
        Chunks of the file (str for text mode, bytes for binary)
    """
    with create_file_handle(file_path, mode, encoding, **kwargs) as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            yield chunk


def process_file_chunks(
    file_path: Union[str, Path],
    processor: Callable[[Union[str, bytes]], Any],
    mode: str = 'r',
    chunk_size: int = 1024 * 1024,  # 1MB chunks
    encoding: Optional[str] = None,
    **kwargs
) -> List[Any]:
    """
    Process a file in chunks using a given processor function.
    
    Args:
        file_path: Path to the file
        processor: Function to process each chunk
        mode: File open mode ('r', 'rb')
        chunk_size: Size of chunks to read
        encoding: Encoding for text modes, or None to auto-detect
        **kwargs: Additional arguments to pass to open()
        
    Returns:
        List of results from the processor function
    """
    results = []
    for chunk in iter_chunked_file(file_path, chunk_size, mode, encoding, **kwargs):
        result = processor(chunk)
        if result is not None:
            results.append(result)
    return results


def file_scanner(
    folder_path: Union[str, Path],
    extensions: List[str],
    processor: Callable[[Path], Dict[str, Any]],
    recursion_depth: int = -1,
    error_handler: Optional[Callable[[Path, Exception], None]] = None
) -> Tuple[List[Dict[str, Any]], ErrorCollection]:
    """
    Scan a folder for files and process each file.
    
    Args:
        folder_path: Path to the folder to scan
        extensions: List of file extensions to process
        processor: Function to process each file
        recursion_depth: How deep to recurse (-1 for unlimited)
        error_handler: Optional function to handle errors
        
    Returns:
        Tuple of (results_list, errors_collection)
    """
    results = []
    errors = ErrorCollection()
    
    for file_path in find_files(folder_path, extensions, recursion_depth):
        try:
            result = processor(file_path)
            results.append(result)
        except Exception as e:
            errors.add(f"Error processing {file_path}", e)
            if error_handler:
                error_handler(file_path, e)
    
    return results, errors


def get_temp_file_path(prefix: str = "temp", suffix: Optional[str] = None) -> Path:
    """
    Generate a temporary file path.
    
    Args:
        prefix: Prefix for the temporary file
        suffix: Optional suffix/extension
        
    Returns:
        Path object for the temporary file
    """
    import tempfile
    
    temp_dir = tempfile.gettempdir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{prefix}_{timestamp}{suffix or ''}"
    return Path(temp_dir) / filename


def make_path_safe(filename: str) -> str:
    """
    Make a filename safe for all operating systems.
    
    Args:
        filename: Original filename
        
    Returns:
        Safe filename
    """
    # Replace characters that are problematic on various operating systems
    unsafe_chars = '<>:"/\\|?*'
    safe_name = ''.join(c if c not in unsafe_chars else '_' for c in filename)
    
    # Handle special names on Windows
    if sys.platform == 'win32':
        # Handle reserved names
        reserved_names = [
            'CON', 'PRN', 'AUX', 'NUL', 
            'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
            'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
        ]
        name_without_ext = safe_name.split('.')[0].upper()
        if name_without_ext in reserved_names:
            safe_name = f"_{safe_name}"
    
    # Trim leading/trailing whitespace and periods
    safe_name = safe_name.strip().strip('.')
    
    # If we end up with an empty name, use a default
    if not safe_name:
        safe_name = "unnamed_file"
    
    return safe_name


def ensure_directory_exists(directory_path: Union[str, Path]) -> Path:
    """
    Ensure a directory exists, creating it if necessary.
    
    Args:
        directory_path: Path to the directory
        
    Returns:
        Path object for the directory
    """
    path = Path(directory_path) if isinstance(directory_path, str) else directory_path
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    return path


def copy_file_safely(
    source_path: Union[str, Path],
    dest_path: Union[str, Path],
    overwrite: bool = False
) -> Path:
    """
    Copy a file with additional safety checks.
    
    Args:
        source_path: Path to the source file
        dest_path: Path to the destination file
        overwrite: Whether to overwrite the destination if it exists
        
    Returns:
        Path object for the copied file
    """
    src_path = Path(source_path) if isinstance(source_path, str) else source_path
    dst_path = Path(dest_path) if isinstance(dest_path, str) else dest_path
    
    # Check source
    if not src_path.exists():
        raise FileNotFoundError(f"Source file not found: {src_path}")
    if not src_path.is_file():
        raise ValueError(f"Source is not a file: {src_path}")
    
    # Ensure destination directory exists
    ensure_directory_exists(dst_path.parent)
    
    # Handle existing destination
    if dst_path.exists() and not overwrite:
        # Generate a unique name
        base = dst_path.stem
        ext = dst_path.suffix
        counter = 1
        while dst_path.exists():
            dst_path = dst_path.with_name(f"{base}_{counter}{ext}")
            counter += 1
    
    # Copy the file
    return Path(shutil.copy2(src_path, dst_path))


def get_file_access_history(file_path: Union[str, Path]) -> Optional[Dict[str, Any]]:
    """
    Attempt to get file access history information.
    
    This is only available on certain systems and requires additional modules.
    
    Args:
        file_path: Path to the file
        
    Returns:
        Dictionary with access history or None if not available
    """
    path = Path(file_path) if isinstance(file_path, str) else file_path
    
    # Windows file access history
    if platform.system() == 'Windows':
        try:
            # Requires the pywin32 module
            import win32com.client
            import win32con
            import win32file
            import win32security
            from datetime import datetime
            
            # Try to get last access info using FileSystem object
            history = {}
            
            # First try to get USN journal entries (requires admin rights)
            try:
                # Get the volume where the file is located
                drive_letter = os.path.splitdrive(path)[0]
                if drive_letter:
                    # Create a handle to the volume
                    hVolume = win32file.CreateFile(
                        f"\\\\.\\{drive_letter}",
                        win32con.GENERIC_READ,
                        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                        None,
                        win32con.OPEN_EXISTING,
                        0,
                        None
                    )
                    
                    # Get USN journal information
                    journal_info = win32file.GetUsnJournalInformation(hVolume)
                    if journal_info:
                        history["has_journal"] = True
                        # Note: Actual USN record access requires more complex implementation
                        # and admin privileges, so we just note that the journal exists
            except:
                pass
            
            # Get basic file access information
            try:
                # Convert the path to a shell item
                shell = win32com.client.Dispatch("Shell.Application")
                folder = shell.NameSpace(os.path.dirname(path))
                file_item = folder.ParseName(os.path.basename(path))
                
                # Try to get access times
                for i in range(0, 300):  # Try all possible property indices
                    prop_name = folder.GetDetailsOf(None, i)
                    if prop_name.lower() in ['date accessed', 'last accessed', 'accessed']:
                        value = folder.GetDetailsOf(file_item, i)
                        if value:
                            history["access_time_shell"] = value
                    elif prop_name.lower() in ['date modified', 'last modified', 'modified']:
                        value = folder.GetDetailsOf(file_item, i)
                        if value:
                            history["modified_time_shell"] = value
                    elif prop_name.lower() in ['date created', 'created']:
                        value = folder.GetDetailsOf(file_item, i)
                        if value:
                            history["created_time_shell"] = value
                    elif prop_name.lower() in ['owner', 'author']:
                        value = folder.GetDetailsOf(file_item, i)
                        if value:
                            history["owner_shell"] = value
            except:
                pass
                
            # Try to get security audit information
            try:
                # Check if the file has auditing enabled
                security_info = win32security.GetFileSecurity(
                    str(path),
                    win32security.SACL_SECURITY_INFORMATION
                )
                sacl = security_info.GetSecurityDescriptorSacl()
                if sacl and sacl.GetAceCount() > 0:
                    history["has_audit_trail"] = True
                    # Note: Reading the actual audit trail requires admin privileges
                    # and accessing the Windows Event Log
            except:
                history["has_audit_trail"] = False
                
            return history if history else None
            
        except ImportError:
            # Required modules not available
            return None
        except Exception as e:
            logger.debug(f"Error getting Windows file access history: {str(e)}")
            return None
    
    # macOS file access history
    elif platform.system() == 'Darwin':
        try:
            # Try to get Spotlight metadata
            import subprocess
            
            try:
                # Use mdls to get metadata attributes
                result = subprocess.run(
                    ['mdls', str(path)],
                    capture_output=True,
                    text=True,
                    check=True
                )
                
                if result.stdout:
                    history = {}
                    lines = result.stdout.splitlines()
                    
                    for line in lines:
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip()
                            
                            # Extract relevant metadata
                            if 'kMDItemLastUsedDate' in key:
                                history["last_used_date"] = value
                            elif 'kMDItemUsedDates' in key:
                                history["used_dates"] = value
                            elif 'kMDItemLastUsedApplication' in key:
                                history["last_used_application"] = value
                            elif 'kMDItemUsedApplications' in key:
                                history["used_applications"] = value
                            elif 'kMDItemContentCreationDate' in key:
                                history["creation_date"] = value
                            elif 'kMDItemContentModificationDate' in key:
                                history["modification_date"] = value
                            elif 'kMDItemAuthors' in key:
                                history["authors"] = value
                    
                    return history if history else None
            except subprocess.SubprocessError:
                pass
                
            return None
            
        except Exception as e:
            logger.debug(f"Error getting macOS file access history: {str(e)}")
            return None
    
    # Not supported on other platforms
    return None