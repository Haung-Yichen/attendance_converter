"""
Domain Entities Module

Core domain entities using dataclasses for the attendance system.
These entities represent the core business concepts independent of infrastructure.
"""

from dataclasses import dataclass, field
from datetime import date, time
from enum import Enum, auto
from typing import List, Optional


class StaffType(Enum):
    """Type of staff member."""
    INTERNAL = auto()  # 內勤: Mon-Fri
    EXTERNAL = auto()  # 外勤: Mon, Wed, Fri only


class AttendanceStatus(Enum):
    """Status of attendance for a given day."""
    NORMAL = auto()        # 正常
    LATE = auto()          # 遲到
    EARLY_LEAVE = auto()   # 早退
    ABSENT = auto()        # 缺勤
    ABNORMAL = auto()      # 異常 (e.g., external staff post-noon checkout)
    LEAVE = auto()         # 請假
    HOLIDAY = auto()       # 假日
    NON_WORK_DAY = auto()  # 非工作日 (e.g., Tue/Thu for external staff)


class RateColorTier(Enum):
    """Color tier for attendance rate display."""
    RED = auto()     # < threshold (default 80%)
    YELLOW = auto()  # >= threshold and < 90%
    GREEN = auto()   # >= 90%


@dataclass
class Staff:
    """
    Represents a staff member.
    
    Attributes:
        name: Staff member's name
        staff_type: Internal or External classification
        work_days: List of weekday numbers (0=Mon, 4=Fri) this staff should work
    """
    name: str
    staff_type: StaffType
    work_days: List[int] = field(default_factory=list)
    
    def __post_init__(self):
        """Set default work days based on staff type if not provided."""
        if not self.work_days:
            if self.staff_type == StaffType.INTERNAL:
                self.work_days = [0, 1, 2, 3, 4]  # Mon-Fri
            else:
                self.work_days = [0, 2, 4]  # Mon, Wed, Fri
    
    def should_work_on(self, day: date) -> bool:
        """Check if staff should work on a given date."""
        return day.weekday() in self.work_days


@dataclass
class AttendanceRecord:
    """
    Represents a single day's attendance record.
    
    Attributes:
        date: The date of attendance
        check_in: Check-in time (None if not recorded)
        check_out: Check-out time (None if not recorded)
        status: Calculated attendance status
        remark: Optional remark/note for this record
    """
    date: date
    check_in: Optional[time] = None
    check_out: Optional[time] = None
    status: AttendanceStatus = AttendanceStatus.ABSENT
    remark: str = ""


@dataclass
class DailyAttendance:
    """
    Container for a staff member's attendance on a specific day.
    
    Attributes:
        staff: The staff member
        record: The attendance record
        in_color: Color to apply to check-in cell (None=no color)
        out_color: Color to apply to check-out cell (None=no color)
    """
    staff: Staff
    record: AttendanceRecord
    in_color: Optional[str] = None   # 'green', 'red', or None
    out_color: Optional[str] = None  # 'green', 'red', or None


@dataclass
class MonthlyAttendance:
    """
    Container for a staff member's monthly attendance summary.
    
    Attributes:
        staff: The staff member
        year: Year of the report
        month: Month of the report (1-12)
        records: List of daily attendance records
        required_days: Number of days staff was required to work
        actual_days: Number of days staff actually attended
        attendance_rate: Calculated attendance rate percentage
        rate_color: Color tier based on rate
    """
    staff: Staff
    year: int
    month: int
    records: List[AttendanceRecord] = field(default_factory=list)
    required_days: int = 0
    actual_days: int = 0
    attendance_rate: float = 0.0
    rate_color: RateColorTier = RateColorTier.RED


@dataclass 
class MonthlyStats:
    """
    Statistics for the monthly report.
    
    Attributes:
        year: Report year
        month: Report month
        required_work_days: Total work days in the month
        holidays: Number of holidays
        total_staff_count: Total staff in source
        internal_count: Number of internal staff
        external_count: Number of external staff
    """
    year: int
    month: int
    required_work_days: int = 0
    holidays: int = 0
    total_staff_count: int = 0
    internal_count: int = 0
    external_count: int = 0
