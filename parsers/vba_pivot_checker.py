# parsers/vba_pivot_checker.py
"""
Module for detecting VBA macros and PivotTables in Excel files.
"""
from pathlib import Path
from typing import Dict, Any, Tuple, Callable, Optional, Union
import zipfile
import re
import openpyxl
from oletools.olevba import VBA_Parser
import logging

from config import EXCEL_EXTENSIONS
from utils.error_utils import safe_parser, safe_operation, with_fallback
from utils.file_utils import read_binary_sample
from utils.metadata_utils import get_excel_file_metadata
from .base_parser import BaseParser


class VBAPivotChecker(BaseParser):
    """
    Checker for VBA macros and PivotTables in Excel files.
    
    This class provides functionality to detect if an Excel file contains
    VBA macros and/or PivotTables, using different detection strategies
    based on the Excel file format.
    """
    
    def __init__(self):
        """Initialize the VBA and PivotTable checker."""
        super().__init__(
            default_values={"has_vba": False, "has_pivot_table": False},
            file_extensions=EXCEL_EXTENSIONS
        )
        # Map file extensions to their checker methods
        self.check_strategies = {
            '.xlsx': (self._check_vba_new_excel, self._check_pivot_table_new_excel),
            '.xlsm': (self._check_vba_new_excel, self._check_pivot_table_new_excel),
            '.xlsb': (self._check_vba_new_excel, self._check_pivot_table_new_excel),
            '.xls': (self._check_vba_old_excel, self._check_pivot_table_old_excel),
        }
    
    @safe_parser(default_return={"has_vba": False, "has_pivot_table": False})
    def parse(self, file_path: Union[str, Path]) -> Dict[str, bool]:
        """
        Check if an Excel file contains VBA macros and/or PivotTables.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with has_vba and has_pivot_table flags
        """
        path = Path(file_path) if isinstance(file_path, str) else file_path
        extension = path.suffix.lower()
        
        # Get the appropriate checker functions for this file type
        vba_checker, pivot_checker = self.check_strategies.get(
            extension, 
            (lambda x: False, lambda x: False)  # Default for unknown extensions
        )
        
        # Execute the checks
        has_vba = vba_checker(path)
        has_pivot_table = pivot_checker(path)
        
        # Build detailed result with additional information
        result = {
            "has_vba": has_vba,
            "has_pivot_table": has_pivot_table
        }
        
        # If we found VBA, try to get more details
        if has_vba:
            vba_details = self._get_vba_details(path)
            if vba_details:
                result.update(vba_details)
        
        return result
    
    @safe_operation(operation_name="checking for VBA in modern Excel")
    def _check_vba_new_excel(self, file_path: Path) -> bool:
        """
        Check if a modern Excel file contains VBA macros.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the file contains VBA macros, False otherwise
        """
        # Primary method using oletools
        try:
            vba_parser = VBA_Parser(file_path)
            has_vba = vba_parser.detect_vba_macros()
            vba_parser.close()
            return has_vba
        except Exception as e:
            self.logger.debug(f"VBA detection with oletools failed: {str(e)}")
            
        # Fallback method: check for vbaProject.bin in the zip
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                return any(name == 'xl/vbaProject.bin' for name in z.namelist())
        except Exception as e:
            self.logger.debug(f"VBA detection with zipfile failed: {str(e)}")
            
        # Filename-based heuristic (xlsm usually has macros)
        if file_path.suffix.lower() == '.xlsm':
            return True
            
        return False
    
    @safe_operation(operation_name="checking for VBA in legacy Excel")
    def _check_vba_old_excel(self, file_path: Path) -> bool:
        """
        Check if a legacy Excel file contains VBA macros.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the file contains VBA macros, False otherwise
        """
        try:
            vba_parser = VBA_Parser(file_path)
            has_vba = vba_parser.detect_vba_macros()
            vba_parser.close()
            return has_vba
        except Exception as e:
            self.logger.debug(f"VBA detection for legacy Excel failed: {str(e)}")
            
            # Try an alternative method - check for VBA module names
            try:
                workbook = xlrd.open_workbook(file_path, on_demand=True)
                names = workbook.name_obj_list
                # Check for names that look like VBA modules
                vba_name_patterns = ['Module', 'Class', 'ThisWorkbook', 'Sheet']
                for name in names:
                    if any(pattern in name.name for pattern in vba_name_patterns):
                        return True
            except Exception:
                pass
                
            return False
    
    @safe_operation(operation_name="getting VBA details")
    def _get_vba_details(self, file_path: Path) -> Optional[Dict[str, Any]]:
        """
        Get details about VBA macros in an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with VBA details or None if extraction fails
        """
        try:
            vba_parser = VBA_Parser(file_path)
            
            if not vba_parser.detect_vba_macros():
                vba_parser.close()
                return None
            
            # Extract VBA information
            details = {
                "vba_project_name": "Unknown",
                "vba_module_count": 0,
                "vba_modules": [],
                "vba_has_autoexec": False,
                "vba_suspicious": False,
            }
            
            # Try to get module names
            try:
                modules = []
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
                    if vba_filename:
                        modules.append(vba_filename)
                        
                        # Check for potentially suspicious patterns
                        suspicious_patterns = [
                            "Shell", "WScript.Shell", "ExecuteExcel4Macro",
                            "ActiveX", "CreateObject", "Environ", 
                            "Open.*For.*Output", "Write.*#", "Binary"
                        ]
                        
                        for pattern in suspicious_patterns:
                            if re.search(pattern, vba_code, re.IGNORECASE):
                                details["vba_suspicious"] = True
                                break
                
                details["vba_module_count"] = len(modules)
                details["vba_modules"] = modules[:10]  # Limit to first 10 modules
                
                # Check for auto-executable macros
                auto_exec_names = ["autoopen", "autoexec", "auto_open", "document_open", "workbook_open"]
                details["vba_has_autoexec"] = any(
                    any(auto_name in module.lower() for auto_name in auto_exec_names)
                    for module in modules
                )
                
            except Exception as e:
                self.logger.debug(f"Error extracting VBA module details: {str(e)}")
            
            vba_parser.close()
            return details
            
        except Exception as e:
            self.logger.debug(f"Error getting VBA details: {str(e)}")
            return None
    
    @safe_operation(operation_name="checking for PivotTables in modern Excel")
    def _check_pivot_table_new_excel(self, file_path: Path) -> bool:
        """
        Check if a modern Excel file contains PivotTables.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the file contains PivotTables, False otherwise
        """
        # Strategy 1: Check for pivot-related files in the ZIP structure (fast)
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                file_list = z.namelist()
                
                # Check for pivot cache definition parts
                if any('pivotCache' in name for name in file_list):
                    return True
                
                # Check for pivot table definition parts
                if any('pivotTable' in name for name in file_list):
                    return True
                    
                # Check for pivot mentions in worksheet xml
                for name in file_list:
                    if re.match(r'xl/worksheets/sheet\d+.xml', name):
                        try:
                            content = z.read(name).decode('utf-8', errors='ignore')
                            if '<pivotTable' in content or '<pivotCache' in content:
                                return True
                        except Exception:
                            pass
        except Exception as e:
            self.logger.debug(f"Error checking for PivotTables in ZIP: {str(e)}")
            
        # Strategy 2: Use openpyxl (more thorough but slower)
        try:
            # First check if the file is large - if so, avoid loading the whole workbook
            if file_path.stat().st_size > 10 * 1024 * 1024:  # Skip detailed check for files > 10MB
                self.logger.debug(f"Skipping detailed PivotTable check for large file: {file_path}")
                return False
                
            # Load the workbook in read-only mode
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_vba=False)
            
            # Check each sheet
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                
                # Check pivot tables attribute
                if hasattr(sheet, '_pivot_tables') and sheet._pivot_tables:
                    return True
                    
                # Alternative check for pivot tables via relationships
                if hasattr(sheet, '_rels') and sheet._rels:
                    for rel in sheet._rels.values():
                        if 'pivotTable' in rel.target:
                            return True
        except Exception as e:
            self.logger.debug(f"Error checking for PivotTables with openpyxl: {str(e)}")
            
        return False
    
    @safe_operation(operation_name="checking for PivotTables in legacy Excel")
    def _check_pivot_table_old_excel(self, file_path: Path) -> bool:
        """
        Check if a legacy Excel file contains PivotTables.
        This is a best-effort approach, as .xls format doesn't expose this easily.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if the file contains PivotTables, False otherwise
        """
        import xlrd
        
        try:
            workbook = xlrd.open_workbook(file_path, on_demand=True)
            
            # Heuristic approach: check for named ranges that are likely related to pivot tables
            names = workbook.name_obj_list
            for name in names:
                if 'pivot' in name.name.lower():
                    return True
                    
            # Check for sheets with names suggesting pivot tables
            for sheet_name in workbook.sheet_names():
                if 'pivot' in sheet_name.lower() or 'pt' == sheet_name.lower():
                    return True
                    
            return False
        except Exception as e:
            self.logger.debug(f"Error checking for PivotTables in XLS: {str(e)}")
            return False