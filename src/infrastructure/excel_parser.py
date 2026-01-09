"""
Excel Parser Module

Handles parsing and cleaning raw Excel attendance exports.
Cleans messy data like "Name[]" patterns and "*" symbols.
"""

import re
from datetime import datetime, time, date
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from domain.entities import AttendanceRecord, AttendanceStatus


@dataclass
class RawAttendanceRow:
    """Raw attendance data from Excel."""
    name: str
    date: date
    check_in: Optional[time]
    check_out: Optional[time]


class ExcelParser:
    """
    Parses raw Excel attendance exports.
    
    Handles:
    - Cleaning "Name[]" garbage patterns
    - Removing "*" symbols from time values
    - Extracting attendance data into domain entities
    """
    
    # Pattern to clean name fields like "John[]" or "Mary[*]"
    NAME_CLEAN_PATTERN = re.compile(r'\[.*?\]')
    
    # Pattern to clean time values with asterisks
    TIME_CLEAN_PATTERN = re.compile(r'\*')
    
    def __init__(self):
        self._raw_data: List[RawAttendanceRow] = []
        self._unique_names: List[str] = []
        self._year: Optional[int] = None
    
    def parse_file(self, file_path: Path, year: Optional[int] = None) -> List[RawAttendanceRow]:
        """
        Parse an Excel file and extract attendance data.
        
        Args:
            file_path: Path to the Excel file
            year: Optional year to use for date parsing (for MM/DD formats)
            
        Returns:
            List of RawAttendanceRow objects
        """
        self._raw_data = []
        self._unique_names = []
        self._year = year
        
        if not file_path.exists():
            return []
        
        wb = load_workbook(file_path, data_only=True)
        
        # Parse ALL worksheets, not just the active one
        for ws in wb.worksheets:
            sheet_data = self._parse_worksheet(ws)
            self._raw_data.extend(sheet_data)
        
        # Extract unique names
        seen_names = set()
        for row in self._raw_data:
            if row.name not in seen_names:
                self._unique_names.append(row.name)
                seen_names.add(row.name)
        
        wb.close()
        return self._raw_data
    
    def _parse_worksheet(self, ws: Worksheet) -> List[RawAttendanceRow]:
        """Parse a worksheet and extract attendance rows.
        
        Expected format (MonRep export):
        - Column A (1): 部門 (Department)
        - Column B (2): 姓名 (Name)
        - Column C (3): 日期 (Date)
        - Column D (4): 遲到 (Late)
        - Column E (5): 早退 (Early leave)
        - Column F (6): 加班 (Overtime)
        - Column G (7): 工時 (Work hours)
        - Column H (8): 假別(1) / 上班 (Check-in)
        - Column I (9): 假別(2) / 下班 (Check-out)
        """
        rows = []
        
        # Detect header row by looking for column headers
        header_row = 1
        name_col = 2  # Default: B column
        date_col = 3  # Default: C column
        check_in_col = 8  # Default: H column (上班)
        check_out_col = 9  # Default: I column (下班)
        
        # Search for header row in first 10 rows, search up to 15 columns
        for row_idx in range(1, min(10, ws.max_row + 1)):
            for col_idx in range(1, min(15, ws.max_column + 1)):
                cell_value = str(ws.cell(row_idx, col_idx).value or '').strip()
                cell_lower = cell_value.lower()
                
                # Detect name column
                if any(keyword in cell_lower for keyword in ['name', '姓名', '員工']):
                    header_row = row_idx
                    name_col = col_idx
                
                # Detect date column
                if any(keyword in cell_lower for keyword in ['date', '日期']):
                    date_col = col_idx
                
                # Detect check-in column (上班)
                if '上班' in cell_value:
                    check_in_col = col_idx
                
                # Detect check-out column (下班)
                if '下班' in cell_value:
                    check_out_col = col_idx
        
        # Parse data rows - track name across rows since name only appears once per person
        current_name = None
        for row_idx in range(header_row + 1, ws.max_row + 1):
            name_cell = ws.cell(row_idx, name_col).value
            
            # Update current name if we find a new one
            if name_cell:
                current_name = self._clean_name(str(name_cell))
            
            # Skip if no name established yet
            if not current_name:
                continue
            
            # Extract date - skip non-date rows (summary rows at bottom)
            date_val = self._extract_date(ws.cell(row_idx, date_col).value, self._year)
            if not date_val:
                continue
            
            # Extract times from detected columns
            check_in = self._extract_time(ws.cell(row_idx, check_in_col).value)
            check_out = self._extract_time(ws.cell(row_idx, check_out_col).value)
            
            rows.append(RawAttendanceRow(
                name=current_name,
                date=date_val,
                check_in=check_in,
                check_out=check_out
            ))
        
        return rows
    
    def _clean_name(self, name: str) -> str:
        """Clean a name field by removing garbage patterns."""
        if not name:
            return ""
        # Remove [*] or [] patterns
        cleaned = self.NAME_CLEAN_PATTERN.sub('', name)
        return cleaned.strip()
    
    def _extract_date(self, value, year: Optional[int] = None) -> Optional[date]:
        """Extract date from a cell value.
        
        Handles formats like:
        - 12/01(一) - MM/DD with weekday in parentheses
        - 2024-12-01
        - 2024/12/01
        """
        if value is None:
            return None
        
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        
        # Try parsing string formats
        str_val = str(value).strip()
        
        # Handle MM/DD(weekday) format like "12/01(一)" or "12/05 (五 *)"
        # Remove the weekday part in parentheses including any spaces/asterisks
        weekday_pattern = re.compile(r'\s*\([一二三四五六日月火水木金土]\s*\*?\)')
        str_val_clean = weekday_pattern.sub('', str_val).strip()
        
        # If no year provided, use current year
        if year is None:
            year = datetime.now().year
        
        # Try MM/DD format (common in Taiwan/Japan)
        if '/' in str_val_clean and len(str_val_clean) <= 5:
            try:
                parts = str_val_clean.split('/')
                if len(parts) == 2:
                    month = int(parts[0])
                    day = int(parts[1])
                    return date(year, month, day)
            except (ValueError, IndexError):
                pass
        
        # Try standard formats
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y']:
            try:
                return datetime.strptime(str_val_clean, fmt).date()
            except ValueError:
                continue
        
        return None
    
    def _extract_time(self, value) -> Optional[time]:
        """Extract time from a cell value, cleaning asterisks."""
        if value is None:
            return None
        
        if isinstance(value, datetime):
            return value.time()
        if isinstance(value, time):
            return value
        
        # Clean and parse string
        str_val = self.TIME_CLEAN_PATTERN.sub('', str(value)).strip()
        
        for fmt in ['%H:%M:%S', '%H:%M', '%I:%M %p', '%I:%M:%S %p']:
            try:
                return datetime.strptime(str_val, fmt).time()
            except ValueError:
                continue
        
        return None
    
    def get_unique_names(self) -> List[str]:
        """Get list of unique staff names found in the file."""
        return self._unique_names
    
    def get_records_for_name(self, name: str) -> List[RawAttendanceRow]:
        """Get all records for a specific staff member."""
        return [row for row in self._raw_data if row.name == name]
    
    def get_records_by_month(
        self, 
        year: int, 
        month: int
    ) -> Dict[str, List[RawAttendanceRow]]:
        """
        Get records grouped by name for a specific month.
        
        Args:
            year: Year to filter
            month: Month to filter (1-12)
            
        Returns:
            Dictionary mapping names to their records
        """
        result: Dict[str, List[RawAttendanceRow]] = {}
        
        for row in self._raw_data:
            if row.date.year == year and row.date.month == month:
                if row.name not in result:
                    result[row.name] = []
                result[row.name].append(row)
        
        return result
    
    def convert_to_attendance_records(
        self,
        raw_rows: List[RawAttendanceRow]
    ) -> List[AttendanceRecord]:
        """
        Convert raw rows to domain AttendanceRecord objects.
        
        Args:
            raw_rows: List of raw attendance rows
            
        Returns:
            List of AttendanceRecord objects
        """
        records = []
        for raw in raw_rows:
            status = AttendanceStatus.ABSENT
            if raw.check_in or raw.check_out:
                status = AttendanceStatus.NORMAL  # Will be refined by logic layer
            
            records.append(AttendanceRecord(
                date=raw.date,
                check_in=raw.check_in,
                check_out=raw.check_out,
                status=status
            ))
        
        return records
