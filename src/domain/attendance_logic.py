"""
Attendance Logic Module

Implements Strategy pattern for determining attendance status
based on staff type and time rules.
"""

from abc import ABC, abstractmethod
from datetime import time, datetime
from typing import Optional

from .entities import (
    Staff, StaffType, AttendanceRecord, AttendanceStatus,
    DailyAttendance, RateColorTier
)
from config.config_manager import TimeRule, ColorLogic


def parse_time(time_str: str) -> time:
    """Parse time string (HH:MM) to time object."""
    if not time_str:
        return time(0, 0)
    try:
        parts = time_str.split(':')
        return time(int(parts[0]), int(parts[1]))
    except (ValueError, IndexError):
        return time(0, 0)


class AttendanceStrategy(ABC):
    """Abstract base class for attendance determination strategies."""
    
    @abstractmethod
    def determine_status(
        self, 
        record: AttendanceRecord, 
        time_rule: TimeRule
    ) -> AttendanceStatus:
        """
        Determine the attendance status for a record.
        
        Args:
            record: The attendance record to evaluate
            time_rule: The time rule to apply
            
        Returns:
            The determined AttendanceStatus
        """
        pass
    
    @abstractmethod
    def get_colors(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule,
        color_logic: ColorLogic
    ) -> tuple[Optional[str], Optional[str]]:
        """
        Get colors for check-in and check-out cells.
        
        Args:
            record: The attendance record
            time_rule: The time rule to apply
            color_logic: The color logic settings
            
        Returns:
            Tuple of (in_color, out_color) - 'green', 'red', or None
        """
        pass
    
    @abstractmethod
    def get_remark(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule
    ) -> str:
        """
        Get remark for this record based on time violations.
        
        Returns:
            Remark string (e.g., '下班延遲打卡') or empty string
        """
        pass


class InternalAttendanceStrategy(AttendanceStrategy):
    """
    Attendance strategy for internal staff (內勤).
    
    Rules:
    - Late if check_in > in_end
    - Early leave if check_out < out_start
    """
    
    def determine_status(
        self, 
        record: AttendanceRecord, 
        time_rule: TimeRule
    ) -> AttendanceStatus:
        if record.check_in is None and record.check_out is None:
            return AttendanceStatus.ABSENT
        
        in_end = parse_time(time_rule.in_end)
        out_start = parse_time(time_rule.out_start)
        
        is_late = record.check_in is not None and record.check_in > in_end
        is_early = record.check_out is not None and record.check_out < out_start
        
        if is_late and is_early:
            return AttendanceStatus.ABNORMAL
        elif is_late:
            return AttendanceStatus.LATE
        elif is_early:
            return AttendanceStatus.EARLY_LEAVE
        else:
            return AttendanceStatus.NORMAL
    
    def get_colors(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule,
        color_logic: ColorLogic
    ) -> tuple[Optional[str], Optional[str]]:
        in_color = None
        out_color = None
        
        if record.check_in is None and record.check_out is None:
            return (None, None)
        
        in_end = parse_time(time_rule.in_end)
        out_start = parse_time(time_rule.out_start)
        
        # Check-in color
        if record.check_in is not None:
            if record.check_in <= in_end:
                # Normal check-in
                if color_logic.green_normal_in:
                    in_color = 'green'
            else:
                # Late check-in
                if color_logic.red_abnormal_in:
                    in_color = 'red'
        
        # Check-out color
        if record.check_out is not None:
            if record.check_out >= out_start:
                # Normal check-out
                if color_logic.green_normal_out:
                    out_color = 'green'
            else:
                # Early check-out
                if color_logic.red_abnormal_out:
                    out_color = 'red'
        
        return (in_color, out_color)
    
    def get_remark(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule
    ) -> str:
        """Get remark for delayed checkout."""
        if record.check_out is not None:
            out_end = parse_time(time_rule.out_end)
            if record.check_out > out_end:
                return "下班延遲打卡"
        return ""


class ExternalAttendanceStrategy(AttendanceStrategy):
    """
    Attendance strategy for external staff (外勤).
    
    Rules:
    - Late if check_in > in_end
    - CRITICAL: If check_out > 12:00 (post-noon), mark as Abnormal/Red
    """
    
    NOON = time(12, 0)
    
    def determine_status(
        self, 
        record: AttendanceRecord, 
        time_rule: TimeRule
    ) -> AttendanceStatus:
        if record.check_in is None and record.check_out is None:
            return AttendanceStatus.ABSENT
        
        in_end = parse_time(time_rule.in_end)
        
        is_late = record.check_in is not None and record.check_in > in_end
        
        # User requested modification: Post-noon is NOT abnormal, just delayed (remark)
        # is_post_noon = record.check_out is not None and record.check_out > self.NOON
        # if is_post_noon: return AttendanceStatus.ABNORMAL
        
        if is_late:
            return AttendanceStatus.LATE
        else:
            return AttendanceStatus.NORMAL
    
    def get_colors(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule,
        color_logic: ColorLogic
    ) -> tuple[Optional[str], Optional[str]]:
        in_color = None
        out_color = None
        
        if record.check_in is None and record.check_out is None:
            return (None, None)
        
        in_end = parse_time(time_rule.in_end)
        
        # Check-in color
        if record.check_in is not None:
            if record.check_in <= in_end:
                if color_logic.green_normal_in:
                    in_color = 'green'
            else:
                if color_logic.red_abnormal_in:
                    in_color = 'red'
        
        # Check-out color - modified logic
        # Post-noon is now considered NORMAL (Delayed), so use Green/Normal logic
        if record.check_out is not None:
            # Always apply normal color logic unless user wants special handling for overtime
            # Since "Delayed Check-out" for internal is usually Green, we do same here
            if color_logic.green_normal_out:
                out_color = 'green'
        
        return (in_color, out_color)
    
    def get_remark(
        self,
        record: AttendanceRecord,
        time_rule: TimeRule
    ) -> str:
        """Get remark for external staff - delayed checkout."""
        if record.check_out is not None:
            out_end = parse_time(time_rule.out_end)
            if record.check_out > out_end:
                return "下班延遲打卡"
        return ""


class AttendanceLogicFactory:
    """Factory for creating appropriate attendance strategy."""
    
    _strategies = {
        StaffType.INTERNAL: InternalAttendanceStrategy(),
        StaffType.EXTERNAL: ExternalAttendanceStrategy(),
    }
    
    @classmethod
    def get_strategy(cls, staff_type: StaffType) -> AttendanceStrategy:
        """Get the appropriate strategy for a staff type."""
        return cls._strategies.get(staff_type, InternalAttendanceStrategy())


def calculate_rate_color(rate: float, threshold: int = 80) -> RateColorTier:
    """
    Calculate the color tier based on attendance rate.
    
    Args:
        rate: Attendance rate as percentage (0-100)
        threshold: The lower threshold (default 80%)
        
    Returns:
        RateColorTier for display
    """
    if rate < threshold:
        return RateColorTier.RED
    elif rate < 90:
        return RateColorTier.YELLOW
    else:
        return RateColorTier.GREEN
