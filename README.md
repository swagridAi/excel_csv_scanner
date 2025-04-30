# Excel and CSV Inspector

A powerful Python tool that scans folders for Excel and CSV files, extracts detailed metadata, and identifies special features like VBA macros and PivotTables.

## Features

- **File Discovery**: Recursively scan directories for Excel and CSV files with filtering options
- **Comprehensive Metadata Extraction**: 
  - Basic file properties (size, modified date, etc.)
  - Enhanced timestamp information (created, modified, accessed dates)
  - User metadata (creator, last modified by, company)
  - Excel-specific information (sheet count, row/column counts, etc.)
  - CSV structure analysis (delimiter detection, encoding, etc.)
- **Special Feature Detection**:
  - VBA macro detection with security analysis
  - PivotTable identification
  - Custom XML content detection
  - Sheet visibility analysis
- **User Information Tracking**:
  - Identify file creators and editors
  - Extract company and contact information
  - Track file ownership and sharing information
  - Access history and usage patterns (when available)
- **Multiple Output Formats**:
  - Excel reports with formatting and multiple sheets including:
    - Main data sheet with all file information
    - User Information sheet showing who created/modified files
    - Timeline sheet showing chronological file activity
  - CSV output for compatibility
  - Interactive HTML reports with tabbed interface
  - JSON export for programmatic use
- **Performance Optimizations**:
  - Memory-efficient processing for large files
  - Chunked reading operations
  - Multi-strategy parsing for robust results

## Installation

### From PyPI (Recommended)

```bash
pip install excel-csv-inspector
```

### From Source

```bash
git clone https://github.com/yourusername/excel-csv-inspector.git
cd excel-csv-inspector
pip install -e .
```

### Requirements

- Python 3.6+
- Required libraries: pandas, openpyxl, xlrd, oletools, pyxlsb, chardet

## Quick Start

### Basic Usage

Scan a folder for Excel and CSV files and generate a report:

```bash
excel-csv-inspector -i /path/to/folder -o report.xlsx
```

### Recursive Scanning

Scan a folder and all its subfolders:

```bash
excel-csv-inspector -i /path/to/folder -o report.xlsx -r
```

### Filtering Options

Filter files by extension:

```bash
excel-csv-inspector -i /path/to/folder -o report.xlsx -f .xlsx,.csv
```

Filter files by size:

```bash
excel-csv-inspector -i /path/to/folder -o report.xlsx --min-size 1024 --max-size 5242880
```

Filter files by pattern:

```bash
excel-csv-inspector -i /path/to/folder -o report.xlsx --exclude "~$.*" --include ".*2023.*"
```

### Output Formats

Generate different report formats:

```bash
# Excel report (default)
excel-csv-inspector -i /path/to/folder -o report.xlsx

# CSV report
excel-csv-inspector -i /path/to/folder -o report.csv

# HTML report
excel-csv-inspector -i /path/to/folder -o report.html

# JSON report
excel-csv-inspector -i /path/to/folder -o report.json
```

## Command Line Options

| Option | Description |
|--------|-------------|
| `-i, --input` | Input folder to scan (required) |
| `-o, --output` | Output file for the report (defaults to report.xlsx) |
| `-r, --recursive` | Scan subfolders recursively |
| `-l, --log-level` | Set logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `-f, --filter` | Filter files by extension (comma separated, e.g., .xlsx,.csv) |
| `--min-size` | Minimum file size in bytes to include |
| `--max-size` | Maximum file size in bytes to include |
| `--include` | Regex pattern for filenames to include |
| `--exclude` | Regex pattern for filenames to exclude |
| `--no-user-info` | Disable extraction of user metadata |
| `--no-timestamps` | Disable extraction of timestamp metadata |

## Report Contents

The generated report includes the following information for each file:

### Basic Information
- Filename
- File path
- Extension
- Size (bytes, KB, MB)

### Timestamp Information
- Created date and time
- Modified date and time
- Accessed date and time

### User Information
- Creator/Author
- Last modified by
- Company/Organization
- Owner (file system owner)
- Additional contacts

### Excel File Information
- Sheet count and names
- Row and column counts
- Sheet visibility status
- Document properties
- Custom XML content

### CSV File Information
- Delimiter used
- Encoding
- Header detection
- Row and column counts
- Data type analysis

### Special Features
- VBA macro presence and details
- PivotTable detection
- Embedded objects

## Excel Report Sheets

### Main Sheet
Contains all file data with enhanced formatting for user and timestamp fields

### User Information Sheet
Displays user metadata with the following information:
- File details (name, path, type)
- Creator information
- Last modified by information
- Company and department details

### Timeline Sheet
Provides a chronological view of file activity:
- Creation events
- Modification events
- Access events
- User details associated with each event

## Interactive HTML Reports

The HTML reports feature a tabbed interface with:

1. **File Analysis** tab - Complete file data in a sortable table
2. **User Information** tab - User cards showing who created or modified files
3. **Timeline** tab - Chronological visualization of file activity

## Project Structure

```
excel_csv_inspector/
├── main.py                      # Entry point for running the tool
├── config.py                    # Configuration settings
├── utils/
│   ├── file_utils.py            # File operations and metadata extraction
│   ├── metadata_utils.py        # Metadata extraction utilities
│   └── error_utils.py           # Error handling utilities
├── parsers/
│   ├── base_parser.py           # Base parser class with shared functionality
│   ├── csv_parser.py            # CSV file parser with user metadata detection
│   ├── excel_parser.py          # Excel file parser with enhanced metadata
│   ├── vba_pivot_checker.py     # VBA and PivotTable detection
│   └── parser_registry.py       # Parser registry and file analysis
├── report/
│   ├── reporter.py              # Report generation with multiple output formats
│   └── templates.py             # Report templates for different formats
├── tests/                       # Unit tests
├── requirements.txt
└── README.md
```

## Using as a Library

You can also use Excel and CSV Inspector as a library in your own Python code:

```python
from excel_csv_inspector.utils import find_files, get_extended_metadata
from excel_csv_inspector.parsers import analyze_file, create_parser_registry

# Find Excel and CSV files
files = list(find_files("/path/to/folder", [".xlsx", ".csv"], recursion_depth=-1))

# Create parser registry
registry = create_parser_registry()

# Analyze files
results = []
for file_path in files:
    result = analyze_file(file_path, registry)
    results.append(result)

# Process results
for result in results:
    if result.get("has_vba", False):
        print(f"VBA found in {result['filename']}")
    
    # Access user metadata
    if "user_info" in result:
        creator = result["user_info"].get("creator", "Unknown")
        print(f"File created by {creator}")
    
    # Access timestamp information
    created = result.get("created_date_formatted", "Unknown")
    modified = result.get("modified_date_formatted", "Unknown")
    print(f"Created: {created}, Last Modified: {modified}")
```

## Extracting User Metadata

The tool extracts user information from various sources:

### Excel Files
- Document properties (creator, lastModifiedBy)
- Custom document properties
- Core XML content (Office Open XML)
- VBA project properties

### CSV Files
- Header comments with user information
- Metadata rows at the beginning of files
- Standard patterns indicating creators or editors

### File System
- File ownership information
- Access control lists (Windows)
- File system attributes and permissions

## Error Handling

Excel and CSV Inspector is designed to be resilient when processing large numbers of files:

- Files that cannot be parsed will be included in the report with error information
- The tool will continue processing even if individual files fail
- Error summaries are provided in the logs and report

## Performance Considerations

- For large Excel files, the tool uses read-only mode to reduce memory usage
- Memory-mapped operations are used when possible for large file handling
- File types are detected efficiently without loading entire files
- Users with large collections can use filtering options to process files in batches

## Known Limitations

- CSV user metadata extraction depends on standardized formats
- Some timestamp information may not be available on all platforms
- File access history requires specific operating system support
- Some advanced features require additional dependencies

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Testing

Run the test suite:

```bash
python -m unittest discover -s tests
```