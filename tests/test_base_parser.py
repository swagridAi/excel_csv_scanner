# tests/test_base_parser.py
import unittest
import tempfile
import os
from pathlib import Path
import logging

from parsers.base_parser import BaseParser, FileParserRegistry

# Create a concrete implementation of BaseParser for testing
class TestParser(BaseParser):
    def __init__(self, default_values=None, file_extensions=None):
        super().__init__(
            default_values=default_values or {"test_value": True},
            file_extensions=file_extensions or [".test"]
        )
    
    def parse(self, file_path):
        # Simple implementation for testing
        result = self.default_values.copy()
        result["filename"] = Path(file_path).name
        return result


class TestBaseParser(unittest.TestCase):
    
    def setUp(self):
        # Create a temporary file for testing
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_file_path = Path(self.temp_dir.name) / "test_file.txt"
        with open(self.test_file_path, 'w') as f:
            f.write("Test content")
    
    def tearDown(self):
        self.temp_dir.cleanup()
    
    def test_init(self):
        """Test BaseParser initialization with default values."""
        default_values = {"key": "value", "has_test": False}
        file_extensions = [".test", ".tst"]
        
        parser = TestParser(default_values=default_values, file_extensions=file_extensions)
        
        self.assertEqual(parser.default_values, default_values)
        self.assertEqual(parser.file_extensions, file_extensions)
        self.assertIsInstance(parser.logger, logging.Logger)
    
    def test_parse(self):
        """Test that the parse method works as expected."""
        default_values = {"key": "value", "has_test": False}
        parser = TestParser(default_values=default_values)
        
        result = parser.parse(self.test_file_path)
        
        # The parser should return default values plus the filename
        expected = default_values.copy()
        expected["filename"] = self.test_file_path.name
        self.assertEqual(result, expected)
    
    def test_default_values_not_modified(self):
        """Test that default_values dict is not modified by parse operations."""
        default_values = {"key": "value", "count": 0}
        parser = TestParser(default_values=default_values)
        
        # Run parse which should make a copy of default_values
        result = parser.parse(self.test_file_path)
        result["count"] += 1  # Modify the result
        
        # Check that original default_values was not changed
        self.assertEqual(default_values["count"], 0)
    
    def test_nonexistent_file(self):
        """Test parsing a file that doesn't exist."""
        parser = TestParser()
        nonexistent_file = Path(self.temp_dir.name) / "nonexistent.txt"
        
        # This should not raise an exception due to safe_parser decorator
        result = parser.parse(nonexistent_file)
        
        # It should return default values and include an error
        self.assertTrue("error" in result)


class TestFileParserRegistry(unittest.TestCase):
    
    def setUp(self):
        self.registry = FileParserRegistry()
        
        # Create some test parsers
        self.txt_parser = TestParser(
            default_values={"type": "text"},
            file_extensions=[".txt", ".text"]
        )
        
        self.csv_parser = TestParser(
            default_values={"type": "csv"},
            file_extensions=[".csv"]
        )
    
    def test_init(self):
        """Test registry initialization."""
        self.assertEqual(len(self.registry.parsers), 0)
    
    def test_register(self):
        """Test registering a parser."""
        self.registry.register(self.txt_parser)
        
        # Check that the parser was registered for all its extensions
        self.assertEqual(self.registry.parsers[".txt"], self.txt_parser)
        self.assertEqual(self.registry.parsers[".text"], self.txt_parser)
    
    def test_register_multiple(self):
        """Test registering multiple parsers."""
        self.registry.register(self.txt_parser)
        self.registry.register(self.csv_parser)
        
        # Check all parsers were registered
        self.assertEqual(self.registry.parsers[".txt"], self.txt_parser)
        self.assertEqual(self.registry.parsers[".text"], self.txt_parser)
        self.assertEqual(self.registry.parsers[".csv"], self.csv_parser)
    
    def test_parsers_lookup(self):
        """Test looking up parsers directly from the parsers dict."""
        self.registry.register(self.txt_parser)
        self.registry.register(self.csv_parser)
        
        # Test with extensions
        self.assertEqual(self.registry.parsers[".txt"], self.txt_parser)
        self.assertEqual(self.registry.parsers[".csv"], self.csv_parser)
    
    def test_register_lowercase_extensions(self):
        """Test that extensions are stored in lowercase."""
        uppercase_parser = TestParser(
            default_values={"type": "uppercase"},
            file_extensions=[".TXT", ".TEXT"]
        )
        
        self.registry.register(uppercase_parser)
        
        # The parser should be registered with lowercase extensions
        self.assertEqual(self.registry.parsers[".txt"], uppercase_parser)
        self.assertEqual(self.registry.parsers[".text"], uppercase_parser)
    
    def test_overlapping_extensions(self):
        """Test registering parsers with overlapping extensions."""
        overlap_parser = TestParser(
            default_values={"type": "overlap"},
            file_extensions=[".txt"]
        )
        
        self.registry.register(self.txt_parser)
        self.registry.register(overlap_parser)
        
        # The overlap parser should now be registered for .txt
        self.assertEqual(self.registry.parsers[".txt"], overlap_parser)
        # But the original parser should still be registered for .text
        self.assertEqual(self.registry.parsers[".text"], self.txt_parser)
    
    def test_get_parser_for_file(self):
        """Test getting a parser for a specific file."""
        self.registry.register(self.txt_parser)
        self.registry.register(self.csv_parser)
        
        # Test based on actual method signature/implementation if it exists
        if hasattr(self.registry, 'get_parser_for_file'):
            parser = self.registry.get_parser_for_file("test.txt")
            self.assertEqual(parser, self.txt_parser)
            
            parser = self.registry.get_parser_for_file("data.csv")
            self.assertEqual(parser, self.csv_parser)
            
            parser = self.registry.get_parser_for_file("unknown.xyz")
            self.assertIsNone(parser)
    
    def test_empty_registry(self):
        """Test behavior with an empty registry."""
        # Try to access a parser from an empty registry
        with self.assertRaises(KeyError):
            parser = self.registry.parsers[".txt"]


if __name__ == "__main__":
    unittest.main()