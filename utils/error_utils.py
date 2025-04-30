# utils/error_utils.py
"""
Error handling utilities for Excel and CSV Inspector.

This module provides functions and decorators for error handling,
making the application more robust when dealing with file parsing 
and other potentially error-prone operations.
"""
import logging
import functools
import traceback
from typing import Dict, Any, Callable, TypeVar, Optional, List, Tuple, Union

# Type variables for function signatures
T = TypeVar('T')
R = TypeVar('R')

# Configure module logger
logger = logging.getLogger(__name__)


class ErrorCollection:
    """
    Collection of errors that may occur during processing.
    
    This class allows multiple errors to be collected and summarized.
    """
    
    def __init__(self):
        """Initialize an empty error collection."""
        self.errors = []
    
    def add(self, description: str, exception: Exception) -> None:
        """
        Add an error to the collection.
        
        Args:
            description: Description of what operation was being performed
            exception: The exception that occurred
        """
        error_entry = {
            "description": description,
            "exception_type": type(exception).__name__,
            "exception_message": str(exception),
            "traceback": traceback.format_exc()
        }
        self.errors.append(error_entry)
    
    def __len__(self) -> int:
        """Return the number of errors in the collection."""
        return len(self.errors)
    
    def __bool__(self) -> bool:
        """Return True if there are errors in the collection."""
        return len(self.errors) > 0
    
    def summary(self) -> str:
        """
        Get a summary of all errors.
        
        Returns:
            String containing a summary of all errors
        """
        if not self.errors:
            return "No errors"
        
        summary_lines = [f"Error summary ({len(self.errors)} errors):"]
        for i, error in enumerate(self.errors, 1):
            summary_lines.append(f"{i}. {error['description']}: {error['exception_type']}: {error['exception_message']}")
        
        return "\n".join(summary_lines)
    
    def detailed_report(self) -> str:
        """
        Get a detailed report of all errors including stack traces.
        
        Returns:
            String containing a detailed report of all errors
        """
        if not self.errors:
            return "No errors"
        
        report_lines = [f"Detailed error report ({len(self.errors)} errors):"]
        for i, error in enumerate(self.errors, 1):
            report_lines.append(f"\n{i}. {error['description']}:")
            report_lines.append(f"   Exception: {error['exception_type']}: {error['exception_message']}")
            report_lines.append("   Traceback:")
            for line in error['traceback'].split('\n'):
                report_lines.append(f"   {line}")
        
        return "\n".join(report_lines)


def safe_parser(default_return: Dict[str, Any] = None) -> Callable:
    """
    Decorator for safely parsing files. If parsing fails, returns default values.
    
    Args:
        default_return: Default values to return if parsing fails
        
    Returns:
        Decorated function
    """
    def decorator(parse_func: Callable[..., Dict[str, Any]]) -> Callable[..., Dict[str, Any]]:
        @functools.wraps(parse_func)
        def wrapper(*args, **kwargs) -> Dict[str, Any]:
            try:
                return parse_func(*args, **kwargs)
            except Exception as e:
                # Get logger from class instance if available
                if len(args) > 0 and hasattr(args[0], 'logger'):
                    log = args[0].logger
                else:
                    log = logger
                    
                # Get file path for better error messages
                file_path = kwargs.get('file_path', None)
                if file_path is None and len(args) > 1:
                    file_path = args[1]
                
                # Log the error
                if file_path:
                    log.error(f"Error parsing {file_path}: {str(e)}")
                else:
                    log.error(f"Error in {parse_func.__name__}: {str(e)}")
                
                # Create error result
                result = default_return.copy() if default_return is not None else {}
                result["error"] = str(e)
                result["error_type"] = type(e).__name__
                
                # If we have a file path, include it in the result
                if file_path:
                    if isinstance(file_path, str):
                        filename = file_path.split('/')[-1].split('\\')[-1]
                    else:  # Assume Path object
                        filename = file_path.name
                        
                    result.update({
                        "filename": filename,
                        "file_path": str(file_path)
                    })
                
                return result
                
        return wrapper
    return decorator


def safe_operation(operation_name: str = None, log_level: int = logging.WARNING) -> Callable:
    """
    Decorator for safely performing operations that might fail.
    
    Args:
        operation_name: Name of the operation for logging
        log_level: Level at which to log errors
        
    Returns:
        Decorated function
    """
    def decorator(func: Callable[..., T]) -> Callable[..., Optional[T]]:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Optional[T]:
            op_name = operation_name or func.__name__
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # Get logger from class instance if available
                if len(args) > 0 and hasattr(args[0], 'logger'):
                    log = args[0].logger
                else:
                    log = logger
                
                log.log(log_level, f"Error {op_name}: {str(e)}")
                return None
        return wrapper
    return decorator


def with_fallback(primary_func: Callable[[], T], 
                 fallback_func: Callable[[], T], 
                 error_message: str = None) -> Callable[[], T]:
    """
    Create a function that tries primary_func and falls back to fallback_func if it fails.
    
    Args:
        primary_func: The function to try first
        fallback_func: The function to use if the first fails
        error_message: Optional message to log when falling back
        
    Returns:
        A function that implements the fallback behavior
    """
    def combined_func() -> T:
        try:
            return primary_func()
        except Exception as e:
            if error_message:
                logger.warning(f"{error_message}: {str(e)}")
            return fallback_func()
    return combined_func