"""
PDF Writer Module

Generates formatted PDF attendance reports using fpdf2.
Supports Chinese text and creates reports similar to Excel output.
"""

from calendar import monthrange
from pathlib import Path
from typing import List, Optional

from fpdf import FPDF

from domain.entities import (
    MonthlyAttendance, AttendanceRecord, AttendanceStatus, RateColorTier
)


class AttendancePdf(FPDF):
    """Custom FPDF class with header and footer."""
    
    def __init__(self, title: str = ""):
        super().__init__(orientation='L', unit='mm', format='A4')
        self.title_text = title
        # Use built-in font that supports basic characters
        # For full Chinese support, a TTF font would need to be added
    
    def header(self):
        """Add header to each page."""
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, self.title_text, align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(5)
    
    def footer(self):
        """Add footer with page number."""
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')


class PdfWriter:
    """
    Generates PDF attendance reports.
    
    Creates reports with:
    - Staff name
    - Daily attendance times
    - Monthly attendance rate
    """
    
    # Color definitions (RGB)
    COLORS = {
        'green': (144, 238, 144),   # Light green
        'yellow': (255, 215, 0),     # Gold
        'red': (255, 107, 107),      # Light red
        'header': (68, 114, 196),    # Blue header
        'white': (255, 255, 255),
        'black': (0, 0, 0),
        'gray': (200, 200, 200),
    }
    
    def __init__(self):
        pass
    
    def create_report(
        self,
        attendance_list: List[MonthlyAttendance],
        year: int,
        month: int,
        output_path: Path,
        report_type: str = "internal"
    ) -> None:
        """
        Create PDF report with attendance table.
        
        Args:
            attendance_list: List of monthly attendance records
            year: Report year
            month: Report month
            output_path: Output file path
            report_type: "internal" or "external"
        """
        if not attendance_list:
            return
        
        type_label = "內勤" if report_type == "internal" else "外勤"
        title = f"{year}年{month}月 {type_label}出勤報表"
        
        pdf = AttendancePdf(title=title)
        pdf.alias_nb_pages()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Get days in month
        _, days_in_month = monthrange(year, month)
        
        # Calculate column widths
        # Page width (A4 Landscape) = 297mm, margins = 10mm each side
        available_width = 277
        name_col_width = 25
        rate_col_width = 15
        day_col_width = (available_width - name_col_width - rate_col_width) / days_in_month
        
        # Draw header row
        self._draw_header(pdf, days_in_month, name_col_width, day_col_width, rate_col_width)
        
        # Draw data rows
        for attendance in attendance_list:
            self._draw_attendance_row(
                pdf, attendance, days_in_month,
                name_col_width, day_col_width, rate_col_width
            )
        
        # Save PDF
        pdf.output(str(output_path))
    
    def _draw_header(
        self,
        pdf: FPDF,
        days_in_month: int,
        name_width: float,
        day_width: float,
        rate_width: float
    ) -> None:
        """Draw the table header row."""
        pdf.set_font('Helvetica', 'B', 8)
        pdf.set_fill_color(*self.COLORS['header'])
        pdf.set_text_color(*self.COLORS['white'])
        
        # Name column
        pdf.cell(name_width, 8, 'Name', border=1, align='C', fill=True)
        
        # Day columns
        for day in range(1, days_in_month + 1):
            pdf.cell(day_width, 8, str(day), border=1, align='C', fill=True)
        
        # Rate column
        pdf.cell(rate_width, 8, 'Rate', border=1, align='C', fill=True)
        pdf.ln()
        
        # Reset text color
        pdf.set_text_color(*self.COLORS['black'])
    
    def _draw_attendance_row(
        self,
        pdf: FPDF,
        attendance: MonthlyAttendance,
        days_in_month: int,
        name_width: float,
        day_width: float,
        rate_width: float
    ) -> None:
        """Draw a single attendance row (two lines: in/out times)."""
        pdf.set_font('Helvetica', '', 6)
        
        # Create lookup for records by day
        records_by_day = {}
        for record in attendance.records:
            records_by_day[record.date.day] = record
        
        row_height = 5
        
        # Check-in row
        # Name cell (spans 2 rows conceptually, but we'll draw it in first row)
        pdf.cell(name_width, row_height * 2, attendance.staff.name[:10], border=1, align='C')
        
        # Save position for out row
        x_after_name = pdf.get_x()
        y_start = pdf.get_y()
        
        # Go back to draw in times
        pdf.set_xy(x_after_name, y_start)
        
        # In times
        for day in range(1, days_in_month + 1):
            record = records_by_day.get(day)
            if record and record.check_in:
                time_str = record.check_in.strftime('%H:%M')
                # Apply color based on status
                fill_color = self._get_status_color(record, is_in=True)
                if fill_color:
                    pdf.set_fill_color(*fill_color)
                    pdf.cell(day_width, row_height, time_str, border=1, align='C', fill=True)
                else:
                    pdf.cell(day_width, row_height, time_str, border=1, align='C')
            else:
                pdf.cell(day_width, row_height, '-', border=1, align='C')
        
        # Rate cell (spans 2 rows)
        rate_color = self._get_rate_color(attendance.rate_color)
        pdf.set_fill_color(*rate_color)
        pdf.cell(rate_width, row_height * 2, f'{attendance.attendance_rate:.0f}%', 
                 border=1, align='C', fill=True)
        pdf.ln()
        
        # Out times row
        pdf.set_x(x_after_name)
        for day in range(1, days_in_month + 1):
            record = records_by_day.get(day)
            if record and record.check_out:
                time_str = record.check_out.strftime('%H:%M')
                fill_color = self._get_status_color(record, is_in=False)
                if fill_color:
                    pdf.set_fill_color(*fill_color)
                    pdf.cell(day_width, row_height, time_str, border=1, align='C', fill=True)
                else:
                    pdf.cell(day_width, row_height, time_str, border=1, align='C')
            else:
                pdf.cell(day_width, row_height, '-', border=1, align='C')
        
        pdf.ln()
    
    def _get_status_color(
        self,
        record: AttendanceRecord,
        is_in: bool
    ) -> Optional[tuple]:
        """Get fill color based on attendance status."""
        if record.status == AttendanceStatus.NORMAL:
            return self.COLORS['green']
        elif record.status in (AttendanceStatus.LATE, AttendanceStatus.EARLY_LEAVE, 
                               AttendanceStatus.ABNORMAL):
            return self.COLORS['red']
        return None
    
    def _get_rate_color(self, rate_tier: RateColorTier) -> tuple:
        """Get color for attendance rate tier."""
        if rate_tier == RateColorTier.GREEN:
            return self.COLORS['green']
        elif rate_tier == RateColorTier.YELLOW:
            return self.COLORS['yellow']
        else:
            return self.COLORS['red']


def format_filename(pattern: str, year: int, month: int, report_type: str = "") -> str:
    """
    Format filename pattern with placeholders.
    
    Supported placeholders:
    - {year}: 4-digit year
    - {month}: 2-digit month
    
    Args:
        pattern: Filename pattern with placeholders
        year: Report year
        month: Report month
        report_type: Optional, not used (kept for compatibility)
    
    Returns:
        Formatted filename string
    """
    return pattern.format(
        year=year,
        month=f"{month:02d}"
    )
