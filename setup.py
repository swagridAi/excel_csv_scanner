# setup.py
from setuptools import setup, find_packages

setup(
    name="excel_csv_inspector",
    version="1.0.0",
    description="Tool for analyzing Excel and CSV files to extract metadata and detect VBA and PivotTables",
    author="Excel CSV Inspector Team",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.3.0",
        "openpyxl>=3.0.7",
        "xlrd>=2.0.1",
        "oletools>=0.60",
        "pyxlsb>=1.0.9",
        "chardet>=4.0.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-csv-inspector=excel_csv_inspector.main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Information Technology",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Utilities",
    ],
    python_requires=">=3.6",
)