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
from typing import Dict, Any, List, Optional
import time

from config import (
    get_config, LOG_FORMAT, LOG_FILE, 
    DEFAULT_LOG_LEVEL, ALL_SUPPORTED_EXTENSIONS
)
from utils.file_utils import (
    find_files, 
    get_file_metadata,
    file_scanner,
    ensure_directory_exists
)
from utils.error_utils import ErrorCollection, safe_operation
from parsers.parser_registry import create_parser_registry, analyze_file
from report.reporter import Reporter


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
        
        if errors:
            logger.warning(f"Encountered {len(errors)} errors during processing")
            logger.debug(errors.summary())
        
        logger.info(f"Report generated: {config['output_file']}")
        
    except Exception as e:
        logger.error(f"Error during file scanning: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()