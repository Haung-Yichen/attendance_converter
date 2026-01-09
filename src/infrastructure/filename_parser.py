"""
Filename Parser Module

Parses 701Client export filenames to extract year and month information.
"""

import re
from typing import Tuple, Optional


class FilenameParser:
    """
    Parses 701Client export filenames.
    
    Expected format: MonRepyymmdd (e.g., MonRep251201 = December 2025)
    """
    
    # Pattern: MonRep followed by 2-digit year, 2-digit month, 2-digit day
    PATTERN = re.compile(r'^MonRep(\d{2})(\d{2})(\d{2})')
    
    @classmethod
    def parse_report_date(cls, filename: str) -> Tuple[int, int]:
        """
        Parse year and month from 701Client export filename.
        
        Args:
            filename: The filename to parse (e.g., "MonRep251201.xlsx")
            
        Returns:
            Tuple of (year, month) where year is full 4-digit year
            
        Raises:
            ValueError: If filename doesn't match expected format
        """
        match = cls.PATTERN.match(filename)
        if not match:
            raise ValueError(f"Invalid filename format: {filename}. Expected format: MonRepyymmdd")
        
        yy = int(match.group(1))
        mm = int(match.group(2))
        # dd = int(match.group(3))  # Day is available but not needed
        
        # Convert 2-digit year to 4-digit (assuming 2000s)
        year = 2000 + yy
        
        # Validate month
        if not 1 <= mm <= 12:
            raise ValueError(f"Invalid month: {mm}")
        
        return year, mm
    
    @classmethod
    def try_parse_report_date(cls, filename: str) -> Optional[Tuple[int, int]]:
        """
        Try to parse year and month from filename, returning None on failure.
        
        Args:
            filename: The filename to parse
            
        Returns:
            Tuple of (year, month) or None if parsing fails
        """
        try:
            return cls.parse_report_date(filename)
        except ValueError:
            return None
