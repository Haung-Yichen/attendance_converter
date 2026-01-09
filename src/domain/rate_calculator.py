"""
Rate Calculator Module

Calculates attendance rates and determines color grading.
"""

from datetime import date
from calendar import monthrange
from typing import List, Set

from .entities import (
    Staff, StaffType, AttendanceRecord, AttendanceStatus,
    MonthlyAttendance, MonthlyStats, RateColorTier
)


class RateCalculator:
    """
    Calculates attendance rates and statistics.
    
    Provides:
    - Required work days calculation
    - Actual attendance days calculation
    - Attendance rate with 3-tier color grading
    """
    
    def __init__(self, holidays: Set[date] = None):
        """
        Initialize calculator.
        
        Args:
            holidays: Set of holiday dates to exclude from required days
        """
        self.holidays = holidays or set()
    
    def calculate_required_days(
        self, 
        staff: Staff, 
        year: int, 
        month: int
    ) -> int:
        """
        Calculate required work days for a staff member in a month.
        
        Args:
            staff: The staff member
            year: Year
            month: Month (1-12)
            
        Returns:
            Number of required work days
        """
        _, num_days = monthrange(year, month)
        required = 0
        
        for day in range(1, num_days + 1):
            d = date(year, month, day)
            
            # Skip holidays
            if d in self.holidays:
                continue
            
            # Check if staff should work on this day
            if staff.should_work_on(d):
                required += 1
        
        return required
    
    def calculate_actual_days(self, records: List[AttendanceRecord]) -> int:
        """
        Calculate actual attendance days from records.
        
        Args:
            records: List of attendance records
            
        Returns:
            Number of days with valid attendance
        """
        actual = 0
        for record in records:
            if record.status in (
                AttendanceStatus.NORMAL,
                AttendanceStatus.LATE,
                AttendanceStatus.EARLY_LEAVE,
                AttendanceStatus.ABNORMAL,
                AttendanceStatus.LEAVE,
            ):
                actual += 1
        return actual
    
    def calculate_rate(
        self,
        actual_days: int,
        required_days: int
    ) -> float:
        """
        Calculate attendance rate percentage.
        
        Args:
            actual_days: Number of days attended
            required_days: Number of days required
            
        Returns:
            Attendance rate as percentage (0-100)
        """
        if required_days == 0:
            return 100.0
        return (actual_days / required_days) * 100
    
    def get_rate_color(self, rate: float, threshold: int = 80) -> RateColorTier:
        """
        Get color tier for attendance rate.
        
        Args:
            rate: Attendance rate percentage
            threshold: Lower threshold (default 80%)
            
        Returns:
            RateColorTier enum value
        """
        if rate < threshold:
            return RateColorTier.RED
        elif rate < 90:
            return RateColorTier.YELLOW
        else:
            return RateColorTier.GREEN
    
    def calculate_monthly_attendance(
        self,
        staff: Staff,
        records: List[AttendanceRecord],
        year: int,
        month: int,
        threshold: int = 80
    ) -> MonthlyAttendance:
        """
        Calculate complete monthly attendance for a staff member.
        
        Args:
            staff: The staff member
            records: List of attendance records for the month
            year: Year
            month: Month
            threshold: Rate threshold for color grading
            
        Returns:
            MonthlyAttendance with all calculations
        """
        required_days = self.calculate_required_days(staff, year, month)
        actual_days = self.calculate_actual_days(records)
        rate = self.calculate_rate(actual_days, required_days)
        rate_color = self.get_rate_color(rate, threshold)
        
        return MonthlyAttendance(
            staff=staff,
            year=year,
            month=month,
            records=records,
            required_days=required_days,
            actual_days=actual_days,
            attendance_rate=rate,
            rate_color=rate_color
        )
    
    def calculate_monthly_stats(
        self,
        year: int,
        month: int,
        internal_count: int,
        external_count: int
    ) -> MonthlyStats:
        """
        Calculate monthly statistics.
        
        Args:
            year: Year
            month: Month
            internal_count: Number of internal staff
            external_count: Number of external staff
            
        Returns:
            MonthlyStats with all statistics
        """
        _, num_days = monthrange(year, month)
        
        # Calculate work days (Mon-Fri, excluding holidays)
        work_days = 0
        holiday_count = 0
        
        for day in range(1, num_days + 1):
            d = date(year, month, day)
            if d in self.holidays:
                holiday_count += 1
            elif d.weekday() < 5:  # Mon-Fri
                work_days += 1
        
        return MonthlyStats(
            year=year,
            month=month,
            required_work_days=work_days,
            holidays=holiday_count,
            total_staff_count=internal_count + external_count,
            internal_count=internal_count,
            external_count=external_count
        )
