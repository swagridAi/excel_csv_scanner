# main.py
"""
Main entry point for the Excel and CSV Inspector tool.
"""
import os
import sys
import logging
import argparse
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Union
import time
import platform

from config import (
    get_config, LOG_FORMAT, LOG_FILE, 
    DEFAULT_LOG_LEVEL, ALL_SUPPORTED_EXTENSIONS
)
from utils.file_utils import (
    find_files, 
    get_file_metadata,
    file_scanner,
    ensure_directory_exists,
    get_human_size
)
from utils.error_utils import ErrorCollection, safe_operation
from parsers.parser_registry import create_parser_registry, analyze_file
from report.reporter import Reporter


def get_folder_metadata(folder_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get comprehensive metadata for a folder.
    
    Args:
        folder_path: Path to the folder
        
    Returns:
        Dictionary with folder metadata (creation date, owner, etc.)
    """
    logger = logging.getLogger(__name__)
    path = Path(folder_path) if isinstance(folder_path, str) else folder_path
    
    try:
        stat_info = path.stat()
        
        # Basic metadata
        metadata = {
            "folder_path": str(path),
            "folder_name": path.name,
            "modified_date": datetime.fromtimestamp(stat_info.st_mtime),
        }
        
        # Add creation date if available
        if hasattr(stat_info, 'st_birthtime'):  # macOS, BSD
            metadata["folder_created_date"] = datetime.fromtimestamp(stat_info.st_birthtime)
        else:  # Linux and others use st_ctime
            metadata["folder_created_date"] = datetime.fromtimestamp(stat_info.st_ctime)
        
        # Add access date
        metadata["folder_accessed_date"] = datetime.fromtimestamp(stat_info.st_atime)
        
        # Get folder owner information
        try:
            owner_info = get_folder_owner(path)
            if owner_info:
                metadata.update(owner_info)
        except Exception as e:
            logger.debug(f"Error getting folder owner: {str(e)}")
            
        # Format dates for better display
        for date_field in ['folder_created_date', 'folder_modified_date', 'folder_accessed_date']:
            if date_field in metadata and metadata[date_field]:
                metadata[f"{date_field}_formatted"] = metadata[date_field].strftime("%Y-%m-%d %H:%M:%S")
                metadata[f"{date_field}_iso"] = metadata[date_field].isoformat()
                
        # Add folder size information
        try:
            folder_size = get_folder_size(path)
            metadata["folder_size_bytes"] = folder_size
            metadata["folder_size_kb"] = round(folder_size / 1024, 2)
            metadata["folder_size_mb"] = round(folder_size / (1024 * 1024), 2)
            metadata["folder_size_human"] = get_human_size(folder_size)
        except Exception as e:
            logger.debug(f"Error calculating folder size: {str(e)}")
            
        return metadata
        
    except Exception as e:
        logger.error(f"Error getting metadata for folder {folder_path}: {str(e)}")
        return {
            "folder_path": str(path),
            "folder_name": path.name,
            "error": f"Failed to get folder metadata: {str(e)}"
        }


def get_folder_owner(folder_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Get owner information for a folder.
    
    Args:
        folder_path: Path to the folder
        
    Returns:
        Dictionary with owner information
    """
    logger = logging.getLogger(__name__)
    path = Path(folder_path) if isinstance(folder_path, str) else folder_path
    owner_info = {}
    
    # Try to get folder owner on Unix-like systems
    if platform.system() in ['Linux', 'Darwin']:
        try:
            import pwd
            import grp
            
            stat_info = path.stat()
            owner_info["folder_owner_id"] = stat_info.st_uid
            owner_info["folder_group_id"] = stat_info.st_gid
            
            try:
                owner = pwd.getpwuid(stat_info.st_uid)
                owner_info["folder_owner_user"] = owner.pw_name
                owner_info["folder_owner_gecos"] = owner.pw_gecos  # Full name and other info
            except KeyError:
                owner_info["folder_owner_user"] = str(stat_info.st_uid)
            
            try:
                group = grp.getgrgid(stat_info.st_gid)
                owner_info["folder_owner_group"] = group.gr_name
            except KeyError:
                owner_info["folder_owner_group"] = str(stat_info.st_gid)
        except ImportError:
            # pwd/grp modules not available
            logger.debug("pwd/grp modules not available for Unix folder ownership detection")
            pass
    
    # Try to get folder owner on Windows
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
            owner_info["folder_owner_user"] = owner_name
            owner_info["folder_owner_domain"] = owner_domain
            owner_info["folder_owner_sid"] = str(owner_sid)
            
            # Try to get group
            try:
                group_sid = security_descriptor.GetSecurityDescriptorGroup()
                if group_sid:
                    group_name, group_domain, _ = win32security.LookupAccountSid(None, group_sid)
                    owner_info["folder_owner_group"] = group_name
                    owner_info["folder_group_domain"] = group_domain
            except:
                pass
                
            # Check folder permissions and inheritance
            try:
                dacl = security_descriptor.GetSecurityDescriptorDacl()
                if dacl:
                    # Get information about folder permissions
                    ace_count = dacl.GetAceCount()
                    folder_permissions = []
                    
                    for i in range(min(ace_count, 5)):  # Limit to first 5 ACEs
                        ace = dacl.GetAce(i)
                        if ace:
                            ace_type, ace_flags, ace_mask, ace_sid = ace
                            try:
                                user, domain, _ = win32security.LookupAccountSid(None, ace_sid)
                                ace_info = {
                                    "user": f"{domain}\\{user}",
                                    "type": ace_type,
                                    "inherited": bool(ace_flags & win32security.INHERITED_ACE)
                                }
                                folder_permissions.append(ace_info)
                            except:
                                pass
                    
                    if folder_permissions:
                        owner_info["folder_permissions"] = folder_permissions
            except Exception as e:
                logger.debug(f"Error getting folder permissions: {str(e)}")
                
        except ImportError:
            logger.debug("win32security module not available for Windows folder ownership detection")
        except Exception as e:
            logger.debug(f"Error getting Windows folder owner: {str(e)}")
    
    return owner_info


def get_folder_size(folder_path: Union[str, Path]) -> int:
    """
    Calculate the total size of a folder and its contents.
    
    Args:
        folder_path: Path to the folder
        
    Returns:
        Size in bytes
    """
    logger = logging.getLogger(__name__)
    path = Path(folder_path) if isinstance(folder_path, str) else folder_path
    total_size = 0
    
    try:
        for item in path.glob('**/*'):
            if item.is_file():
                total_size += item.stat().st_size
    except Exception as e:
        logger.debug(f"Error calculating folder size: {str(e)}")
    
    return total_size


def setup_logging(log_level: str) -> logging.Logger:
    """Set up logging."""
    numeric_level = getattr(logging, log_level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError(f"Invalid log level: {log_level}")
    
    # Create log directory if it doesn't exist
    log_dir = os.path.dirname(LOG_FILE)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
    
    logging.basicConfig(
        level=numeric_level,
        format=LOG_FORMAT,
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(LOG_FILE)
        ]
    )
    
    return logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description='Inspect Excel and CSV files for metadata, VBA, and PivotTables.'
    )
    parser.add_argument(
        '-i', '--input',
        help='Input folder path to scan'
    )
    parser.add_argument(
        '-o', '--output',
        help='Output file path for the report'
    )
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Scan subfolders recursively'
    )
    parser.add_argument(
        '-l', '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
        default=DEFAULT_LOG_LEVEL,
        help='Set the logging level'
    )
    parser.add_argument(
        '-f', '--filter',
        help='Filter files by extension (comma separated, e.g., .xlsx,.csv)'
    )
    parser.add_argument(
        '--min-size',
        type=int,
        default=0,
        help='Minimum file size in bytes to include'
    )
    parser.add_argument(
        '--max-size',
        type=int,
        help='Maximum file size in bytes to include'
    )
    parser.add_argument(
        '--include',
        help='Regex pattern for filenames to include'
    )
    parser.add_argument(
        '--exclude',
        help='Regex pattern for filenames to exclude'
    )
    parser.add_argument(
        '--no-folder-info',
        action='store_true',
        help='Disable extraction of folder metadata'
    )
    
    return parser.parse_args()


@safe_operation(operation_name="processing file")
def process_file(file_path: Path, registry) -> Dict[str, Any]:
    """
    Process a single file using the parser registry.
    
    Args:
        file_path: Path to the file to process
        registry: Parser registry
        
    Returns:
        Dictionary with analysis results
    """
    return analyze_file(file_path, registry)


def error_handler(file_path: Path, exception: Exception) -> None:
    """
    Handle errors during file processing.
    
    Args:
        file_path: Path to the file that caused the error
        exception: The exception that occurred
    """
    logger = logging.getLogger(__name__)
    logger.error(f"Error processing file {file_path}: {str(exception)}")


def main() -> None:
    """Main entry point."""
    start_time = time.time()
    
    # Parse arguments and get configuration
    args = parse_args()
    config = get_config()
    
    # Update config with command-line arguments
    if args.input:
        config['input_folder'] = args.input
    if args.output:
        config['output_file'] = args.output
    if args.log_level:
        config['log_level'] = args.log_level
    
    # Set recursion depth based on recursive flag
    if args.recursive:
        config['recursion_depth'] = -1
    else:
        config['recursion_depth'] = 0
    
    # Filter extensions if specified
    if args.filter:
        extensions = [ext.strip() for ext in args.filter.split(',')]
        config['file_extensions'] = [ext if ext.startswith('.') else f'.{ext}' for ext in extensions]
    
    # Set up logging
    logger = setup_logging(config['log_level'])
    
    logger.info(f"Starting Excel and CSV file inspection...")
    logger.info(f"Input folder: {config['input_folder']}")
    logger.info(f"Output file: {config['output_file']}")
    logger.info(f"File extensions: {', '.join(config['file_extensions'])}")
    
    # Create output directory if it doesn't exist
    ensure_directory_exists(os.path.dirname(config['output_file']))
    
    # Get input folder metadata and ownership information
    folder_metadata = {"folder_path": config['input_folder']}
    if not args.no_folder_info:
        try:
            input_folder = Path(config['input_folder'])
            folder_metadata = get_folder_metadata(input_folder)
            logger.info(f"Analyzing folder: {input_folder.name}")
            
            # Log folder owner information if available
            if 'folder_owner_user' in folder_metadata:
                logger.info(f"Folder owner: {folder_metadata['folder_owner_user']}")
                
                # Add domain info on Windows
                if 'folder_owner_domain' in folder_metadata:
                    logger.info(f"Folder owner domain: {folder_metadata['folder_owner_domain']}")
                
            if 'folder_owner_group' in folder_metadata:
                logger.info(f"Folder group: {folder_metadata['folder_owner_group']}")
                
            if 'folder_created_date_formatted' in folder_metadata:
                logger.info(f"Folder created: {folder_metadata['folder_created_date_formatted']}")
                
            if 'folder_size_human' in folder_metadata:
                logger.info(f"Folder size: {folder_metadata['folder_size_human']}")
                
        except Exception as e:
            logger.warning(f"Could not get folder metadata: {str(e)}")
    
    # Create parser registry
    registry = create_parser_registry()
    
    # Additional find_files arguments from command line
    find_files_args = {
        'min_size': args.min_size,
        'max_size': args.max_size,
        'include_pattern': args.include,
        'exclude_pattern': args.exclude
    }
    
    # Track counters
    file_count = 0
    excel_count = 0
    csv_count = 0
    vba_count = 0
    pivot_count = 0
    
    logger.info(f"Scanning for files...")
    
    try:
        # Use the file_scanner utility to scan and process files
        results, errors = file_scanner(
            config['input_folder'],
            config['file_extensions'],
            lambda file_path: process_file(file_path, registry),
            config['recursion_depth'],
            error_handler
        )
        
        # Add folder metadata to each result
        if not args.no_folder_info:
            for result in results:
                result['folder_metadata'] = folder_metadata
        
        # Count file types
        file_count = len(results)
        for result in results:
            if result.get('extension') == '.csv':
                csv_count += 1
            elif result.get('is_excel', False):
                excel_count += 1
                
            if result.get('has_vba', False):
                vba_count += 1
            if result.get('has_pivot_table', False):
                pivot_count += 1
        
        # Generate report
        reporter = Reporter()
        reporter.generate_report(results, config['output_file'])
        
        # Calculate and display stats
        end_time = time.time()
        duration = end_time - start_time
        
        logger.info(f"Finished scanning {file_count} files in {duration:.2f} seconds.")
        logger.info(f"Results: {excel_count} Excel files, {csv_count} CSV files")
        logger.info(f"Found {vba_count} files with VBA macros and {pivot_count} files with PivotTables")
        
        if 'folder_owner_user' in folder_metadata:
            logger.info(f"Folder '{Path(config['input_folder']).name}' is owned by: {folder_metadata['folder_owner_user']}")
        
        if errors:
            logger.warning(f"Encountered {len(errors)} errors during processing")
            logger.debug(errors.summary())
        
        logger.info(f"Report generated: {config['output_file']}")
        
    except Exception as e:
        logger.error(f"Error during file scanning: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()