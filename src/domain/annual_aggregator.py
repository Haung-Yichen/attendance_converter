"""
Annual Aggregator Module

Domain layer service responsible for aggregating monthly attendance data
into annual summaries. Implements cross-month grouping with dynamic
denominator logic and employee status tracking (resignation / new-hire).
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Dict, List, Optional, Set, Tuple

from .entities import (
    MonthlyAttendance,
    RateColorTier,
    Staff,
    StaffType,
)


# =============================================================================
# Value Objects
# =============================================================================

class EmployeeStatus(Enum):
    """Employment status derived from appearance pattern across months."""
    ACTIVE = auto()       # 在職：持續出現至最後一個月
    RESIGNED = auto()     # 離職：前期出現但後期消失
    NEW_HIRE = auto()     # 新進：前期不在後期出現


@dataclass(frozen=True)
class MonthlySnapshot:
    """
    Immutable snapshot of a single month's attendance for one employee.

    Attributes:
        month: Calendar month (1-12).
        required_days: Days the employee was required to attend.
        actual_days: Days the employee actually attended.
        attendance_rate: Monthly attendance rate (0-100).
        rate_color: Color tier for the month.
    """
    month: int
    required_days: int
    actual_days: int
    attendance_rate: float
    rate_color: RateColorTier


@dataclass
class AnnualEmployeeSummary:
    """
    Annual attendance summary for a single employee.

    The denominator for the annual rate is based **only** on the months
    in which this employee appeared, not a fixed 12 months.

    Attributes:
        staff: The staff member entity.
        year: Report year.
        monthly_snapshots: Per-month data keyed by month number (1-12).
        total_required_days: Sum of required days across appeared months.
        total_actual_days: Sum of actual (valid for rate) days across appeared months.
        annual_attendance_rate: Weighted annual attendance rate.
        rate_color: Color tier for the annual rate.
        appeared_months: Sorted list of months in which the employee appeared.
        status: Derived employment status.
        first_month: First month the employee appeared.
        last_month: Last month the employee appeared.
    """
    staff: Staff
    year: int
    monthly_snapshots: Dict[int, MonthlySnapshot] = field(default_factory=dict)
    total_required_days: int = 0
    total_actual_days: int = 0
    annual_attendance_rate: float = 0.0
    rate_color: RateColorTier = RateColorTier.RED
    appeared_months: List[int] = field(default_factory=list)
    status: EmployeeStatus = EmployeeStatus.ACTIVE
    first_month: int = 0
    last_month: int = 0


# =============================================================================
# Aggregation Result
# =============================================================================

@dataclass
class AnnualAggregationResult:
    """
    Complete result of annual aggregation.

    Attributes:
        year: Report year.
        internal_summaries: Annual summaries for internal staff.
        external_summaries: Annual summaries for external staff.
        available_months: Sorted list of months that had source data.
        missing_months: Sorted list of months with no source data.
    """
    year: int
    internal_summaries: List[AnnualEmployeeSummary] = field(default_factory=list)
    external_summaries: List[AnnualEmployeeSummary] = field(default_factory=list)
    available_months: List[int] = field(default_factory=list)
    missing_months: List[int] = field(default_factory=list)


# =============================================================================
# Aggregator (Pure Domain Logic – No I/O)
# =============================================================================

class AnnualAggregator:
    """
    Aggregates multiple months of ``MonthlyAttendance`` data into
    ``AnnualEmployeeSummary`` objects.

    Design:
    - **Pure domain logic** – no file I/O, no framework dependency.
    - Grouping key: ``staff.name`` (employee identifier).
    - Denominator = sum of ``required_days`` only for months the employee appeared.
    - Status detection uses first/last appearance vs. the range of available months.
    """

    def __init__(self, rate_threshold: int = 80) -> None:
        """
        Args:
            rate_threshold: Percentage threshold for RED/YELLOW/GREEN tier.
        """
        self._rate_threshold = rate_threshold

    # --------------------------------------------------------------------- #
    # Public API
    # --------------------------------------------------------------------- #

    def aggregate(
        self,
        year: int,
        monthly_data: Dict[int, List[MonthlyAttendance]],
    ) -> AnnualAggregationResult:
        """
        Aggregate monthly attendance into annual summaries.

        Args:
            year: The target year.
            monthly_data: Mapping of ``month (1-12)`` → list of
                ``MonthlyAttendance`` for that month.

        Returns:
            ``AnnualAggregationResult`` with separated internal / external
            summaries, plus metadata about available / missing months.
        """
        available_months = sorted(monthly_data.keys())
        missing_months = [m for m in range(1, 13) if m not in monthly_data]

        # ---- Step 1: Group by employee name across all months ---- #
        employee_months: Dict[str, Dict[int, MonthlyAttendance]] = {}

        for month, attendances in monthly_data.items():
            for att in attendances:
                key = att.staff.name
                if key not in employee_months:
                    employee_months[key] = {}
                employee_months[key][month] = att

        # ---- Step 2: Build per-employee annual summaries ---- #
        internal_summaries: List[AnnualEmployeeSummary] = []
        external_summaries: List[AnnualEmployeeSummary] = []

        for _name, months_map in employee_months.items():
            summary = self._build_employee_summary(
                year, months_map, available_months,
            )
            if summary.staff.staff_type == StaffType.INTERNAL:
                internal_summaries.append(summary)
            else:
                external_summaries.append(summary)

        # ---- Step 3: Sort by annual attendance rate (ascending) ---- #
        internal_summaries.sort(key=lambda s: s.annual_attendance_rate)
        external_summaries.sort(key=lambda s: s.annual_attendance_rate)

        return AnnualAggregationResult(
            year=year,
            internal_summaries=internal_summaries,
            external_summaries=external_summaries,
            available_months=available_months,
            missing_months=missing_months,
        )

    # --------------------------------------------------------------------- #
    # Private helpers
    # --------------------------------------------------------------------- #

    def _build_employee_summary(
        self,
        year: int,
        months_map: Dict[int, MonthlyAttendance],
        available_months: List[int],
    ) -> AnnualEmployeeSummary:
        """Build an ``AnnualEmployeeSummary`` for one employee."""

        # Pick any MonthlyAttendance to get the Staff entity
        any_att = next(iter(months_map.values()))
        staff = any_att.staff

        appeared = sorted(months_map.keys())
        first_month = appeared[0]
        last_month = appeared[-1]

        # Build monthly snapshots and accumulate totals
        snapshots: Dict[int, MonthlySnapshot] = {}
        total_required = 0
        total_actual = 0

        for month in appeared:
            att = months_map[month]
            snapshot = MonthlySnapshot(
                month=month,
                required_days=att.required_days,
                actual_days=att.actual_days,
                attendance_rate=att.attendance_rate,
                rate_color=att.rate_color,
            )
            snapshots[month] = snapshot
            total_required += att.required_days
            total_actual += att.actual_days

        # Annual rate (weighted average across appeared months)
        annual_rate = (
            (total_actual / total_required * 100.0)
            if total_required > 0
            else 100.0
        )
        rate_color = self._get_rate_color(annual_rate)

        # Determine employment status
        status = self._determine_status(
            appeared, available_months,
        )

        return AnnualEmployeeSummary(
            staff=staff,
            year=year,
            monthly_snapshots=snapshots,
            total_required_days=total_required,
            total_actual_days=total_actual,
            annual_attendance_rate=annual_rate,
            rate_color=rate_color,
            appeared_months=appeared,
            status=status,
            first_month=first_month,
            last_month=last_month,
        )

    def _determine_status(
        self,
        appeared_months: List[int],
        available_months: List[int],
    ) -> EmployeeStatus:
        """
        Derive employment status based on appearance pattern.

        Rules:
        - If the employee's last appearance month equals the overall last
          available month, they are **ACTIVE**.
        - If the employee appeared but disappeared *before* the last
          available month, they are **RESIGNED**.
        - If the employee did *not* appear in the first available month
          but appeared later, they are a **NEW_HIRE**.
        - A person can be both new-hire and resigned if they appeared only
          in the middle; we prioritise RESIGNED in that case.
        """
        if not available_months or not appeared_months:
            return EmployeeStatus.ACTIVE

        overall_first = available_months[0]
        overall_last = available_months[-1]

        emp_first = appeared_months[0]
        emp_last = appeared_months[-1]

        is_new = emp_first > overall_first
        is_gone = emp_last < overall_last

        if is_gone:
            return EmployeeStatus.RESIGNED
        if is_new:
            return EmployeeStatus.NEW_HIRE
        return EmployeeStatus.ACTIVE

    def _get_rate_color(self, rate: float) -> RateColorTier:
        """Map a rate percentage to a colour tier."""
        if rate < self._rate_threshold:
            return RateColorTier.RED
        if rate < 90:
            return RateColorTier.YELLOW
        return RateColorTier.GREEN
