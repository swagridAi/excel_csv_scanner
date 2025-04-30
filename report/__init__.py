# report/__init__.py
"""
Report generation modules for Excel and CSV Inspector Tool.
"""

from .reporter import Reporter
from .templates import (
    EXCEL_TEMPLATE_HEADER, 
    HTML_REPORT_TEMPLATE, 
    TEXT_REPORT_TEMPLATE,
    JSON_SCHEMA
)

__all__ = [
    'Reporter',
    'EXCEL_TEMPLATE_HEADER',
    'HTML_REPORT_TEMPLATE',
    'TEXT_REPORT_TEMPLATE',
    'JSON_SCHEMA'
]