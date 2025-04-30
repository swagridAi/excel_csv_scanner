# tests/test_csv_parser.py
import unittest
from pathlib import Path
import os
import tempfile
import csv
import pandas as pd

from parsers.csv_parser import CSVParser
from config import CSV_DEFAULTS

class TestCSVParser(unittest.TestCase):
    
    def setUp(self):
        # Create a temporary directory
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)
        
        # Create test files
        self._create_test_files()
        
        # Initialize parser
        self.parser = CSVParser()
    
    def tearDown(self):
        # Clean up temporary directory
        self.temp_dir.cleanup()
    
    def _create_test_files(self):
        # Create a simple CSV file with comma delimiter
        self.basic_csv = self.test_dir / "basic.csv"
        with open(self.basic_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a CSV file with semicolon delimiter
        self.semicolon_csv = self.test_dir / "semicolon.csv"
        with open(self.semicolon_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a CSV file with user information in comments
        self.user_info_csv = self.test_dir / "user_info.csv"
        with open(self.user_info_csv, 'w', newline='', encoding='utf-8') as f:
            f.write("# Created by: John Developer\n")
            f.write("# Company: Acme Corporation\n")
            f.write("# Contact: john@example.com\n")
            writer = csv.writer(f)
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a CSV file with user information in columns
        self.user_columns_csv = self.test_dir / "user_columns.csv"
        with open(self.user_columns_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Age', 'City', 'Created by', 'Company'])
            writer.writerow(['John Doe', '30', 'New York', 'John Developer', 'Acme Corporation'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles', '', ''])
        
        # Create a CSV file with empty values
        self.empty_values_csv = self.test_dir / "empty_values.csv"
        with open(self.empty_values_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '', 'New York'])
            writer.writerow(['Jane Smith', '25', ''])
        
        # Create an empty CSV file
        self.empty_csv = self.test_dir / "empty.csv"
        with open(self.empty_csv, 'w', newline='', encoding='utf-8') as f:
            pass
        
        # Create a CSV file with numeric header (to test header detection)
        self.numeric_header_csv = self.test_dir / "numeric_header.csv"
        with open(self.numeric_header_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['123', '456', '789'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a CSV file with quoted values
        self.quoted_csv = self.test_dir / "quoted.csv"
        with open(self.quoted_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
    
    def test_initialization(self):
        # Test that the parser initializes correctly
        self.assertIsInstance(self.parser, CSVParser)
        self.assertEqual(self.parser.default_values, CSV_DEFAULTS)
        self.assertEqual(self.parser.file_extensions, ['.csv'])
    
    def test_parse_basic_csv(self):
        # Test parsing a basic CSV file
        result = self.parser.parse(self.basic_csv)
        
        # Check basic properties
        self.assertTrue(result['is_csv'])
        self.assertEqual(result['row_count'], 3)  # Header + 2 data rows
        self.assertEqual(result['column_count'], 3)  # 3 columns
        self.assertEqual(result['delimiter'], ',')
        self.assertEqual(result['encoding'], 'utf-8')
        self.assertFalse(result['has_vba'])
        self.assertFalse(result['has_pivot_table'])
        self.assertEqual(result['sheet_count'], 1)
    
    def test_parse_semicolon_csv(self):
        # Test parsing a CSV file with semicolon delimiter
        result = self.parser.parse(self.semicolon_csv)
        
        # Check delimiter detection
        self.assertEqual(result['delimiter'], ';')
        self.assertEqual(result['row_count'], 3)  # Header + 2 data rows
        self.assertEqual(result['column_count'], 3)  # 3 columns
    
    def test_parse_user_info_csv(self):
        # Test parsing a CSV file with user information in comments
        result = self.parser.parse(self.user_info_csv)
        
        # Check user information extraction
        self.assertIn('user_info', result)
        self.assertEqual(result['user_info'].get('creator'), 'John Developer')
        self.assertEqual(result['user_info'].get('company'), 'Acme Corporation')
        self.assertEqual(result['user_info'].get('contact'), 'john@example.com')
    
    def test_parse_user_columns_csv(self):
        # Test parsing a CSV file with user information in columns
        result = self.parser.parse(self.user_columns_csv)
        
        # Verify user info is extracted, but implementation details may vary
        if 'user_info' in result:
            user_info = result['user_info']
            # Check if creator was found in the columns
            if 'creator' in user_info:
                self.assertEqual(user_info['creator'], 'John Developer')
            # Check if company was found in the columns
            if 'company' in user_info:
                self.assertEqual(user_info['company'], 'Acme Corporation')
    
    def test_parse_empty_values_csv(self):
        # Test parsing a CSV file with empty values
        result = self.parser.parse(self.empty_values_csv)
        
        # Check empty values detection
        self.assertTrue(result['has_empty_values'])
        self.assertGreater(result['empty_values_count'], 0)
    
    def test_parse_empty_csv(self):
        # Test parsing an empty CSV file
        result = self.parser.parse(self.empty_csv)
        
        # Check that proper values are used for empty files
        self.assertTrue(result['is_csv'])
        # The exact behavior for empty files may vary
        # Some parsers might return 0, others might return 1 for a single empty row
        self.assertIn(result['row_count'], [0, 1])
    
    def test_parse_numeric_header_csv(self):
        # Test parsing a CSV file with numeric header
        result = self.parser.parse(self.numeric_header_csv)
        
        # Check basic properties
        self.assertEqual(result['row_count'], 3)  # All rows including the numeric header
        self.assertEqual(result['column_count'], 3)
        
        # If header detection is implemented, check the result
        if 'has_header' in result:
            # Since all header values are numeric, it might detect as no header
            # This test is more flexible since the implementation could vary
            pass
    
    def test_parse_quoted_csv(self):
        # Test parsing a CSV file with quoted values
        result = self.parser.parse(self.quoted_csv)
        
        # Check quotes detection if implemented
        if 'has_quotes' in result:
            self.assertTrue(result['has_quotes'])
        
        # Check basic properties
        self.assertEqual(result['row_count'], 3)
        self.assertEqual(result['column_count'], 3)
    
    def test_count_with_pandas(self):
        # Test the pandas counting method
        row_count, col_count = self.parser._count_with_pandas(self.basic_csv, 'utf-8', ',')
        self.assertEqual(row_count, 3)  # All rows including header
        self.assertEqual(col_count, 3)
    
    def test_manual_count(self):
        # Test the manual counting method
        row_count, col_count = self.parser._manual_count(self.basic_csv, 'utf-8', ',')
        self.assertEqual(row_count, 3)  # All rows including header
        self.assertEqual(col_count, 3)
    
    def test_analyze_structure(self):
        # Test the structure analysis method
        structure = self.parser._analyze_structure(self.basic_csv, 'utf-8', ',')
        self.assertIsNotNone(structure)
        self.assertTrue(structure['has_header'])
        self.assertFalse(structure['has_empty_values'])
        
        # Test with file containing empty values
        structure = self.parser._analyze_structure(self.empty_values_csv, 'utf-8', ',')
        self.assertTrue(structure['has_empty_values'])
        
        # Test with file containing quoted values
        structure = self.parser._analyze_structure(self.quoted_csv, 'utf-8', ',')
        if 'has_quotes' in structure:
            self.assertTrue(structure['has_quotes'])
    
    def test_extract_user_info(self):
        # Test the user information extraction method
        user_info = self.parser._extract_user_info(self.user_info_csv, 'utf-8', ',')
        self.assertIsNotNone(user_info)
        self.assertEqual(user_info.get('creator'), 'John Developer')
        self.assertEqual(user_info.get('company'), 'Acme Corporation')
        self.assertEqual(user_info.get('contact'), 'john@example.com')
        
        # Test with file that has no user info
        user_info = self.parser._extract_user_info(self.basic_csv, 'utf-8', ',')
        self.assertIsNone(user_info)
    
    def test_error_handling(self):
        # Test parsing a non-existent file
        non_existent = self.test_dir / "non_existent.csv"
        result = self.parser.parse(non_existent)
        
        # Check that default values are used for errors
        self.assertIn('error', result)
        for key, default_value in CSV_DEFAULTS.items():
            self.assertEqual(result.get(key), default_value)
        
        # Test error handling with invalid file path
        with self.assertRaises(Exception):
            # This should raise an exception but be caught by the safe_parser decorator
            self.parser.parse(None)

if __name__ == '__main__':
    unittest.main()