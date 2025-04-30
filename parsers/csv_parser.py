# parsers/csv_parser.py
"""
CSV Parser implementation for extracting metadata from CSV files.
"""
import csv
import pandas as pd
import re
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional, Union

from config import CSV_DEFAULTS, ENCODING_SAMPLE_SIZE, CSV_SNIFF_LINES, CSV_DELIMITERS_TO_GUESS, CSV_EXTENSIONS
from utils.error_utils import safe_parser, safe_operation, with_fallback
from utils.file_utils import (
    read_text_sample, 
    read_binary_sample, 
    detect_encoding, 
    detect_csv_delimiter,
    get_line_count,
    create_file_handle,
    iter_chunked_file
)
from utils.metadata_utils import get_csv_file_metadata
from .base_parser import BaseParser


class CSVParser(BaseParser):
    """
    Parser for CSV files.
    
    This parser extracts metadata from CSV files including row and column counts,
    delimiter, encoding, and other relevant information.
    """
    
    def __init__(self):
        """Initialize the CSV parser with default values and supported extensions."""
        super().__init__(
            default_values=CSV_DEFAULTS,
            file_extensions=CSV_EXTENSIONS
        )
        # Common patterns for user information in CSV comments or headers
        self.user_patterns = {
            'creator': [
                r'(?i)created\s+by[:\s]+([^,;\r\n]+)',
                r'(?i)author[:\s]+([^,;\r\n]+)',
                r'(?i)generated\s+by[:\s]+([^,;\r\n]+)',
                r'(?i)prepared\s+by[:\s]+([^,;\r\n]+)'
            ],
            'modified_by': [
                r'(?i)modified\s+by[:\s]+([^,;\r\n]+)',
                r'(?i)edited\s+by[:\s]+([^,;\r\n]+)',
                r'(?i)updated\s+by[:\s]+([^,;\r\n]+)'
            ],
            'company': [
                r'(?i)company[:\s]+([^,;\r\n]+)',
                r'(?i)organization[:\s]+([^,;\r\n]+)',
                r'(?i)dept\.?[:\s]+([^,;\r\n]+)',
                r'(?i)department[:\s]+([^,;\r\n]+)'
            ],
            'contact': [
                r'(?i)contact[:\s]+([^,;\r\n]+)',
                r'(?i)email[:\s]+([^\s,;\r\n]+)',
                r'(?i)phone[:\s]+([^,;\r\n]+)'
            ]
        }
    
    @safe_parser(default_return=CSV_DEFAULTS)
    def parse(self, file_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Parse a CSV file and extract relevant information.
        
        Args:
            file_path: Path to the CSV file
            
        Returns:
            Dictionary with extracted information
        """
        path = Path(file_path) if isinstance(file_path, str) else file_path
        
        # Start with comprehensive CSV metadata
        result = get_csv_file_metadata(path)
        
        # Get encoding and delimiter from metadata
        encoding = result.get("encoding", "utf-8")
        delimiter = result.get("delimiter", ",")
        
        # Count rows and columns using the most efficient method available
        counting_func = with_fallback(
            lambda: self._count_with_pandas(path, encoding, delimiter),
            lambda: self._manual_count(path, encoding, delimiter),
            "Pandas CSV parsing failed, using manual counting"
        )
        
        row_count, column_count = counting_func()
        result.update({
            "row_count": row_count,
            "column_count": column_count,
            "has_vba": False,  # CSV files don't have VBA
            "has_pivot_table": False,  # CSV files don't have pivot tables
            "sheet_count": 1,  # CSV files are single "sheet"
        })
        
        # Try to detect CSV structure (header presence, data types)
        structure_info = self._analyze_structure(path, encoding, delimiter, sample_rows=10)
        if structure_info:
            result.update(structure_info)
        
        # Try to extract user information from CSV
        user_info = self._extract_user_info(path, encoding, delimiter)
        if user_info:
            result["user_info"] = user_info
        
        return result
    
    @safe_operation(operation_name="counting rows and columns with pandas")
    def _count_with_pandas(self, file_path: Path, encoding: str, delimiter: str) -> Tuple[int, int]:
        """
        Count rows and columns using pandas.
        
        Args:
            file_path: Path to the CSV file
            encoding: File encoding
            delimiter: CSV delimiter
            
        Returns:
            Tuple of (row_count, column_count)
        """
        df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, 
                         low_memory=True, memory_map=True)
        return len(df), len(df.columns)
    
    @safe_operation(operation_name="manually counting rows and columns")
    def _manual_count(self, file_path: Path, encoding: str, delimiter: str) -> Tuple[int, int]:
        """
        Manually count rows and columns in a CSV file.
        
        Args:
            file_path: Path to the CSV file
            encoding: File encoding
            delimiter: CSV delimiter
            
        Returns:
            Tuple of (row_count, column_count)
        """
        # Count lines efficiently
        row_count = get_line_count(file_path, encoding)
        
        # Get column count from first line
        try:
            with create_file_handle(file_path, 'r', encoding=encoding, errors='replace') as f:
                first_line = f.readline().strip()
                if delimiter in first_line:
                    column_count = len(first_line.split(delimiter))
                else:
                    # If delimiter not found, try CSV reader
                    f.seek(0)
                    csv_reader = csv.reader([first_line], delimiter=delimiter)
                    for row in csv_reader:
                        column_count = len(row)
                        break
                    else:
                        column_count = 1
        except Exception as e:
            self.logger.warning(f"Error counting columns: {str(e)}")
            column_count = 0
            
        return row_count, column_count
    
    @safe_operation(operation_name="analyzing CSV structure")
    def _analyze_structure(self, file_path: Path, encoding: str, 
                          delimiter: str, sample_rows: int = 10) -> Optional[Dict[str, Any]]:
        """
        Analyze the structure of a CSV file.
        
        Args:
            file_path: Path to the CSV file
            encoding: File encoding
            delimiter: CSV delimiter
            sample_rows: Number of rows to sample for analysis
            
        Returns:
            Dictionary with structure information or None if analysis fails
        """
        try:
            # Try to read the CSV with pandas for structure analysis
            df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, 
                             nrows=sample_rows, low_memory=True)
            
            # Check for header presence
            has_header = True  # Assume header is present by default
            
            # Check if first row is likely a header (string types)
            numeric_columns = sum(1 for col in df.columns if pd.api.types.is_numeric_dtype(df[col]))
            if numeric_columns == len(df.columns):
                has_header = False  # All columns are numeric, likely no header
                
            # Get column types
            column_types = {col: str(df[col].dtype) for col in df.columns}
            
            # Check for empty values
            empty_values_count = df.isna().sum().sum()
            has_empty_values = empty_values_count > 0
            
            # Check for quotation characters
            sample_text = '\n'.join(df.astype(str).values.flatten()[:100])
            has_quotes = '"' in sample_text or "'" in sample_text
            
            return {
                "has_header": has_header,
                "column_types": column_types,
                "has_empty_values": has_empty_values,
                "empty_values_count": int(empty_values_count),
                "has_quotes": has_quotes,
            }
            
        except Exception as e:
            self.logger.debug(f"Structure analysis failed: {str(e)}")
            return None
    
    @safe_operation(operation_name="extracting user information from CSV")
    def _extract_user_info(self, file_path: Path, encoding: str, delimiter: str) -> Optional[Dict[str, Any]]:
        """
        Extract potential user information from CSV file.
        
        This method looks for common patterns that might indicate user information
        in CSV comments, headers, or metadata rows.
        
        Args:
            file_path: Path to the CSV file
            encoding: File encoding
            delimiter: CSV delimiter
            
        Returns:
            Dictionary with user information or None if no user info found
        """
        try:
            # Read the first several lines for analysis
            with create_file_handle(file_path, 'r', encoding=encoding, errors='replace') as f:
                # Read first 20 lines which might contain metadata or comments
                header_lines = [f.readline() for _ in range(20) if f.readline()]
                header_text = '\n'.join(header_lines)
            
            user_info = {}
            
            # Look for patterns in the header text
            for info_type, patterns in self.user_patterns.items():
                for pattern in patterns:
                    matches = re.search(pattern, header_text)
                    if matches and matches.group(1).strip():
                        user_info[info_type] = matches.group(1).strip()
                        break
            
            # If nothing found in header comments, check if there might be metadata rows
            if not user_info:
                try:
                    # Read first 10 rows
                    df = pd.read_csv(file_path, encoding=encoding, delimiter=delimiter, nrows=10)
                    
                    # Check if any column names match user metadata
                    columns = df.columns.tolist()
                    for col in columns:
                        col_lower = str(col).lower()
                        if any(keyword in col_lower for keyword in ['created by', 'author', 'user', 'creator']):
                            # Check the first non-null value
                            values = df[col].dropna()
                            if not values.empty:
                                user_info['creator'] = str(values.iloc[0])
                        
                        elif any(keyword in col_lower for keyword in ['modified by', 'updated by', 'editor']):
                            values = df[col].dropna()
                            if not values.empty:
                                user_info['modified_by'] = str(values.iloc[0])
                        
                        elif any(keyword in col_lower for keyword in ['company', 'organization', 'dept']):
                            values = df[col].dropna()
                            if not values.empty:
                                user_info['company'] = str(values.iloc[0])
                except Exception as e:
                    self.logger.debug(f"Failed to check for user info in data rows: {str(e)}")
            
            # Check for email addresses which might indicate user information
            email_pattern = r'[\w\.-]+@[\w\.-]+\.\w+'
            email_matches = re.findall(email_pattern, header_text)
            if email_matches and 'contact' not in user_info:
                user_info['contact'] = email_matches[0]
            
            return user_info if user_info else None
            
        except Exception as e:
            self.logger.debug(f"User info extraction failed: {str(e)}")
            return None