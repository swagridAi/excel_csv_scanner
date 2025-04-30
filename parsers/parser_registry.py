# parsers/parser_registry.py
"""
Module providing a parser registry to simplify file analysis.
"""
from pathlib import Path
from typing import Dict, Any, List, Optional

from config import ALL_SUPPORTED_EXTENSIONS
from utils.error_utils import safe_operation
from .base_parser import FileParserRegistry
from .csv_parser import CSVParser
from .excel_parser import ExcelParser
from .vba_pivot_checker import VBAPivotChecker


def create_parser_registry() -> FileParserRegistry:
    """
    Create and configure a file parser registry with all available parsers.
    
    Returns:
        Configured FileParserRegistry instance
    """
    registry = FileParserRegistry()
    
    # Register CSV parser
    csv_parser = CSVParser()
    registry.register(csv_parser)
    
    # Register Excel parser
    excel_parser = ExcelParser()
    registry.register(excel_parser)
    
    # Register VBA and PivotTable checker
    vba_pivot_checker = VBAPivotChecker()
    registry.register(vba_pivot_checker)
    
    return registry


@safe_operation(operation_name="analyzing file")
def analyze_file(file_path: Path, registry: Optional[FileParserRegistry] = None) -> Dict[str, Any]:
    """
    Analyze a file using the appropriate parsers from the registry.
    This function combines results from all applicable parsers.
    
    Args:
        file_path: Path to the file to analyze
        registry: Optional parser registry (will create one if None)
        
    Returns:
        Combined analysis results from all applicable parsers
    """
    # Create registry if not provided
    if registry is None:
        registry = create_parser_registry()
    
    # Get all parsers that support this file type
    extension = file_path.suffix.lower()
    
    # Start with an empty result
    result = {}
    
    # Get Excel parser for Excel files
    if extension in ['.xlsx', '.xlsm', '.xls', '.xlsb']:
        # Get basic Excel information
        excel_parser = registry.parsers.get(extension)
        if excel_parser:
            excel_result = excel_parser.parse(file_path)
            result.update(excel_result)
        
        # Get VBA and PivotTable information
        vba_pivot_checker = registry.parsers.get(extension)
        if vba_pivot_checker and hasattr(vba_pivot_checker, '_check_vba_new_excel'):
            vba_pivot_result = vba_pivot_checker.parse(file_path)
            result.update(vba_pivot_result)
    
    # Get CSV parser for CSV files
    elif extension == '.csv':
        csv_parser = registry.parsers.get(extension)
        if csv_parser:
            csv_result = csv_parser.parse(file_path)
            result.update(csv_result)
    
    # If we didn't get any results, try to get basic file info
    if not result and extension in ALL_SUPPORTED_EXTENSIONS:
        # Find any parser that can give us basic info
        for ext, parser in registry.parsers.items():
            if hasattr(parser, 'get_file_info'):
                result = parser.get_file_info(file_path)
                result["error"] = f"No parser found for {extension}"
                break
    
    return result