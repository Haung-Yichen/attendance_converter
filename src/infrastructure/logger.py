"""
Logger Module

Provides a centralized logging system that outputs to both console and file.
"""

import logging
import sys
from pathlib import Path
from typing import Optional

# Application log file path (relative to project root)
_LOG_FILE_NAME = "app.log"


def _get_project_root() -> Path:
    """Get the project root directory."""
    return Path(__file__).parent.parent.parent


def get_logger(name: str, log_file: Optional[str] = None) -> logging.Logger:
    """
    Get or create a logger with console and file handlers.
    
    Args:
        name: Logger name (typically module name like "ExcelParser")
        log_file: Optional custom log file path. If None, uses default app.log
        
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    
    # Avoid adding handlers multiple times
    if logger.handlers:
        return logger
    
    logger.setLevel(logging.DEBUG)
    
    # Formatter for both handlers
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    
    # Console handler - INFO level and above
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # File handler - DEBUG level and above
    log_path = Path(log_file) if log_file else _get_project_root() / _LOG_FILE_NAME
    try:
        file_handler = logging.FileHandler(
            log_path, 
            mode="a", 
            encoding="utf-8"
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except (OSError, PermissionError) as e:
        # If file logging fails, just log to console
        logger.warning(f"無法建立日誌檔案 {log_path}: {e}")
    
    return logger
