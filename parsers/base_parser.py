# parsers/base_parser.py
"""
Base parser module for Excel and CSV Inspector Tool.

This module provides the base classes for file parsers and a registry
to manage multiple parsers based on file extensions.
"""
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

class BaseParser:
    """
    Base class for file parsers.
    
    This abstract class provides common functionality for file parsers,
    including logging and default values handling.
    """
    
    def __init__(self, default_values: Dict[str, Any], file_extensions: List[str]):
        """
        Initialize the parser with default values and supported extensions.
        
        Args:
            default_values: Dictionary of default values to use when a field can't be parsed
            file_extensions: List of file extensions supported by this parser
        """
        self.default_values = default_values
        self.file_extensions = file_extensions
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
    
    def parse(self, file_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Parse a file and extract relevant information.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Dictionary with extracted information
        """
        raise NotImplementedError("Subclasses must implement parse method")


class FileParserRegistry:
    """
    Registry for file parsers.
    
    This class maintains a mapping of file extensions to parsers
    to simplify file analysis.
    """
    
    def __init__(self):
        """Initialize the registry."""
        self.parsers = {}  # Maps file extensions to parsers
    
    def register(self, parser: BaseParser) -> None:
        """
        Register a parser for its supported file extensions.
        
        Args:
            parser: Parser to register
        """
        for extension in parser.file_extensions:
            self.parsers[extension.lower()] = parser
    
    def get_parser_for_file(self, file_path: Union[str, Path]) -> Optional[BaseParser]:
        """
        Get the appropriate parser for a file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Appropriate parser or None if no parser is found
        """
        path = Path(file_path) if isinstance(file_path, str) else file_path
        extension = path.suffix.lower()
        
        return self.parsers.get(extension)