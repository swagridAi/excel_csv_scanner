# tests/test_file_utils.py
import unittest
from pathlib import Path
import os
import tempfile
import csv

from utils.file_utils import (
    find_files,
    get_file_metadata,
    read_binary_sample,
    read_text_sample,
    detect_encoding,
    detect_csv_delimiter,
    get_line_count
)

class TestFileUtils(unittest.TestCase):
    
    def setUp(self):
        # Create a temporary directory
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)
        
        # Create test files
        self._create_test_files()
    
    def tearDown(self):
        # Clean up temporary directory
        self.temp_dir.cleanup()
    
    def _create_test_files(self):
        # Create a simple text file
        self.text_file = self.test_dir / "test.txt"
        with open(self.text_file, 'w', encoding='utf-8') as f:
            f.write("This is a test file.\nIt has multiple lines.\nThird line.")
        
        # Create a simple CSV file with comma delimiter
        self.csv_file = self.test_dir / "test.csv"
        with open(self.csv_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a simple CSV file with semicolon delimiter
        self.csv_file_semicolon = self.test_dir / "test_semicolon.csv"
        with open(self.csv_file_semicolon, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(['Name', 'Age', 'City'])
            writer.writerow(['John Doe', '30', 'New York'])
            writer.writerow(['Jane Smith', '25', 'Los Angeles'])
        
        # Create a subdirectory
        self.sub_dir = self.test_dir / "subdir"
        self.sub_dir.mkdir()
        
        # Create a file in the subdirectory
        self.sub_file = self.sub_dir / "subfile.txt"
        with open(self.sub_file, 'w', encoding='utf-8') as f:
            f.write("This is a file in a subdirectory.")
    
    def test_find_files(self):
        # Test finding all text files in the test directory
        files = list(find_files(self.test_dir, ['.txt']))
        self.assertEqual(len(files), 1)
        self.assertEqual(files[0], self.text_file)
        
        # Test finding all CSV files in the test directory
        files = list(find_files(self.test_dir, ['.csv']))
        self.assertEqual(len(files), 2)
        
        # Test finding all files in the test directory
        files = list(find_files(self.test_dir, ['.txt', '.csv']))
        self.assertEqual(len(files), 3)
        
        # Test recursive file finding
        files = list(find_files(self.test_dir, ['.txt'], recursion_depth=-1))
        self.assertEqual(len(files), 2)  # Should find both text files
        
        # Test non-recursive file finding
        files = list(find_files(self.test_dir, ['.txt'], recursion_depth=0))
        self.assertEqual(len(files), 1)  # Should only find the root text file
    
    def test_get_file_metadata(self):
        # Test getting metadata for a text file
        metadata = get_file_metadata(self.text_file)
        
        # Check basic metadata fields
        self.assertEqual(metadata['filename'], self.text_file.name)
        self.assertEqual(metadata['extension'], '.txt')
        self.assertGreater(metadata['size_bytes'], 0)
        self.assertFalse(metadata['is_excel'])
        self.assertFalse(metadata['is_csv'])
        self.assertFalse(metadata['is_empty'])
    
    def test_read_binary_sample(self):
        # Test reading a binary sample
        sample = read_binary_sample(self.text_file, 10)
        self.assertEqual(len(sample), 10)
        self.assertEqual(sample, b'This is a ')
    
    def test_read_text_sample(self):
        # Test reading a text sample
        sample = read_text_sample(self.text_file, 10)
        self.assertEqual(len(sample), 10)
        self.assertEqual(sample, 'This is a ')
    
    def test_detect_encoding(self):
        # Test encoding detection
        encoding = detect_encoding(self.text_file)
        self.assertIn(encoding.lower(), ['utf-8', 'ascii'])
    
    def test_detect_csv_delimiter(self):
        # Test delimiter detection for comma-separated CSV
        delimiter = detect_csv_delimiter(self.csv_file)
        self.assertEqual(delimiter, ',')
        
        # Test delimiter detection for semicolon-separated CSV
        delimiter = detect_csv_delimiter(self.csv_file_semicolon)
        self.assertEqual(delimiter, ';')
    
    def test_get_line_count(self):
        # Test line counting
        line_count = get_line_count(self.text_file, encoding='utf-8')
        self.assertEqual(line_count, 3)
        
        # Test line counting for CSV
        line_count = get_line_count(self.csv_file, encoding='utf-8')
        self.assertEqual(line_count, 3)  # Header + 2 data rows

if __name__ == '__main__':
    unittest.main()