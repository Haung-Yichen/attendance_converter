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
    Extended format: MonRepyymmdd_nnnnn_nnnnn_YYYYMM.xlsx
      where the trailing _YYYYMM encodes the **data month**.
    """
    
    # Pattern: MonRep followed by 2-digit year, 2-digit month, 2-digit day
    PATTERN = re.compile(r'^MonRep(\d{2})(\d{2})(\d{2})')
    
    # Pattern: trailing _YYYYMM before the file extension
    # e.g. MonRep250101_00000_00200_202512.xlsx → year=2025, month=12
    DATA_MONTH_PATTERN = re.compile(r'_(\d{4})(\d{2})\.[^.]+$')
    
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

    @classmethod
    def try_parse_data_month(cls, filename: str) -> Optional[Tuple[int, int]]:
        """
        Extract the **data month** from the trailing ``_YYYYMM`` suffix.

        This is the month the report *covers*, as opposed to the export
        date encoded in the ``MonRepYYMMDD`` prefix.

        Args:
            filename: e.g. ``MonRep250101_00000_00200_202512.xlsx``

        Returns:
            ``(year, month)`` or ``None`` if the suffix is absent.
        """
        match = cls.DATA_MONTH_PATTERN.search(filename)
        if not match:
            return None
        year = int(match.group(1))
        month = int(match.group(2))
        if not 1 <= month <= 12:
            return None
        return year, month
