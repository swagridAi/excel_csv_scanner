# parsers/excel_parser.py
"""
Excel parser implementation for extracting metadata from Excel files.
"""
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional, Union
import pandas as pd
import openpyxl
import xlrd
import os
import zipfile

from config import EXCEL_DEFAULTS, EXCEL_MAX_ROWS, EXCEL_EXTENSIONS
from utils.error_utils import safe_parser, safe_operation, with_fallback
from utils.file_utils import read_binary_sample
from utils.metadata_utils import get_excel_file_metadata
from .base_parser import BaseParser


class ExcelParser(BaseParser):
    """
    Parser for Excel files.
    
    This parser extracts metadata from Excel files including row and column counts,
    sheet counts, and other relevant information.
    """
    
    def __init__(self):
        """Initialize the Excel parser with default values and supported extensions."""
        super().__init__(
            default_values=EXCEL_DEFAULTS,
            file_extensions=EXCEL_EXTENSIONS
        )
        # Map file extensions to their parser methods
        self.format_parsers = {
            '.xls': self._parse_xls,
            '.xlsx': self._parse_xlsx,
            '.xlsm': self._parse_xlsx,
            '.xlsb': self._parse_xlsx,
        }
    
    @safe_parser(default_return=EXCEL_DEFAULTS)
    def parse(self, file_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Parse an Excel file and extract metadata.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with extracted metadata
        """
        path = Path(file_path) if isinstance(file_path, str) else file_path
        
        # Start with comprehensive Excel metadata
        result = get_excel_file_metadata(path)
        
        # Get the appropriate parser method for this file type
        extension = path.suffix.lower()
        parser_method = self.format_parsers.get(extension)
        
        if not parser_method:
            result["error"] = f"Unsupported Excel format: {extension}"
            return result
        
        # Parse the file using the appropriate method
        excel_result = parser_method(path)
        result.update(excel_result)
        
        return result
    
    @safe_parser(default_return=EXCEL_DEFAULTS)
    def _parse_xls(self, file_path: Path) -> Dict[str, Any]:
        """
        Parse old Excel format (.xls).
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with extracted metadata
        """
        self.logger.debug(f"Parsing .xls file: {file_path}")
        
        # Open the workbook in read-only mode for better performance
        workbook = xlrd.open_workbook(file_path, on_demand=True)
        sheet_names = workbook.sheet_names()
        sheet_count = len(sheet_names)
        
        result = {
            "sheet_count": sheet_count,
            "sheet_names": sheet_names,
        }
        
        # Get information about each sheet
        sheets_info = []
        total_rows = 0
        max_cols = 0
        
        for sheet_name in sheet_names:
            try:
                sheet = workbook.sheet_by_name(sheet_name)
                row_count = sheet.nrows
                col_count = sheet.ncols
                total_rows += row_count
                max_cols = max(max_cols, col_count)
                
                sheets_info.append({
                    "name": sheet_name,
                    "row_count": row_count,
                    "column_count": col_count,
                    "visible": True,  # xlrd doesn't easily expose sheet visibility
                })
            except Exception as e:
                self.logger.warning(f"Error analyzing sheet '{sheet_name}': {str(e)}")
        
        # Add overall statistics
        result.update({
            "sheets_info": sheets_info,
            "row_count": total_rows,  # Total rows across all sheets
            "column_count": max_cols,  # Maximum columns in any sheet
            "avg_row_count": total_rows // sheet_count if sheet_count else 0,
        })
        
        # Check for embedded objects
        try:
            ole_dirs = workbook.xls_workbook.dir_list
            ole_names = [item for item, _ in ole_dirs]
            result["has_embedded_objects"] = any("Embed" in name for name in ole_names)
        except Exception:
            result["has_embedded_objects"] = False
        
        return result
    
    @safe_parser(default_return=EXCEL_DEFAULTS)
    def _parse_xlsx(self, file_path: Path) -> Dict[str, Any]:
        """
        Parse modern Excel formats (.xlsx, .xlsm, .xlsb).
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with extracted metadata
        """
        self.logger.debug(f"Parsing modern Excel file: {file_path}")
        
        # Use read_only mode for better performance
        workbook = openpyxl.load_workbook(
            file_path, 
            read_only=True, 
            data_only=True,
            keep_vba=False
        )
        
        sheet_names = workbook.sheetnames
        sheet_count = len(sheet_names)
        
        result = {
            "sheet_count": sheet_count,
            "sheet_names": sheet_names,
        }
        
        # Get information about each sheet
        sheets_info = []
        total_rows = 0
        max_cols = 0
        
        for sheet_name in sheet_names:
            try:
                sheet = workbook[sheet_name]
                row_count, col_count = self._get_sheet_dimensions(sheet)
                total_rows += row_count
                max_cols = max(max_cols, col_count)
                
                # Check if sheet is hidden
                visible = True
                if hasattr(sheet, 'sheet_state'):
                    visible = sheet.sheet_state == 'visible'
                
                sheets_info.append({
                    "name": sheet_name,
                    "row_count": row_count,
                    "column_count": col_count,
                    "visible": visible,
                })
            except Exception as e:
                self.logger.warning(f"Error analyzing sheet '{sheet_name}': {str(e)}")
                sheets_info.append({
                    "name": sheet_name,
                    "error": str(e),
                })
        
        # Add overall statistics
        result.update({
            "sheets_info": sheets_info,
            "row_count": total_rows,  # Total rows across all sheets
            "column_count": max_cols,  # Maximum columns in any sheet
            "avg_row_count": total_rows // sheet_count if sheet_count else 0,
        })
        
        # Try to extract workbook properties
        try:
            if hasattr(workbook, 'properties'):
                props = workbook.properties
                if props:
                    properties = {}
                    for attr in ['creator', 'lastModifiedBy', 'created', 'modified', 'title', 'subject', 'keywords', 'category']:
                        if hasattr(props, attr):
                            value = getattr(props, attr)
                            if value is not None:
                                properties[attr] = str(value)
                    
                    if properties:
                        result["properties"] = properties
        except Exception as e:
            self.logger.debug(f"Error extracting workbook properties: {str(e)}")
        
        # Check for custom XML parts
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                # Look for custom XML
                custom_xml_files = [name for name in z.namelist() if 'customXml/' in name]
                result["has_custom_xml"] = len(custom_xml_files) > 0
                
                # Check for embedded objects (OLE objects)
                has_embedded = any('embeddings/' in name for name in z.namelist())
                result["has_embedded_objects"] = has_embedded
        except Exception as e:
            self.logger.debug(f"Error checking for custom XML: {str(e)}")
        
        return result
    
    @safe_operation(operation_name="getting sheet dimensions")
    def _get_sheet_dimensions(self, sheet) -> Tuple[int, int]:
        """
        Get row and column counts from an openpyxl sheet.
        
        Args:
            sheet: openpyxl worksheet object
            
        Returns:
            Tuple of (row_count, column_count)
        """
        # Strategy 1: Use sheet dimensions if available
        if hasattr(sheet, 'calculate_dimension') and callable(getattr(sheet, 'calculate_dimension')):
            try:
                dimension = sheet.calculate_dimension()
                if dimension and dimension != 'A1:A1':
                    try:
                        min_row, min_col, max_row, max_col = openpyxl.utils.range_boundaries(dimension)
                        return max_row - min_row + 1, max_col - min_col + 1
                    except Exception as e:
                        self.logger.debug(f"Error parsing sheet dimensions: {str(e)}")
            except Exception as e:
                self.logger.debug(f"Error calculating sheet dimensions: {str(e)}")
        
        # Strategy 2: Try to get max_row and max_column directly
        if hasattr(sheet, 'max_row') and hasattr(sheet, 'max_column'):
            try:
                if sheet.max_row > 0 and sheet.max_column > 0:
                    return sheet.max_row, sheet.max_column
            except Exception as e:
                self.logger.debug(f"Error getting max_row/max_column: {str(e)}")
        
        # Strategy 3: Scan for non-empty cells (limited to EXCEL_MAX_ROWS for performance)
        try:
            max_row = 0
            max_col = 0
            for i, row in enumerate(sheet.iter_rows(max_row=EXCEL_MAX_ROWS)):
                if i > max_row:
                    max_row = i
                for j, cell in enumerate(row):
                    if cell.value is not None and j > max_col:
                        max_col = j
            
            if max_row > 0 or max_col > 0:
                return max_row + 1, max_col + 1  # +1 because indices are 0-based
        except Exception as e:
            self.logger.debug(f"Error scanning sheet for dimensions: {str(e)}")
        
        # If we reach here, we couldn't determine dimensions reliably
        self.logger.warning("Could not determine sheet dimensions accurately")
        return 0, 0