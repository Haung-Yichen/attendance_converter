"""
Report Service Module

Application layer service that orchestrates the attendance report generation.
Separates business logic from UI concerns (PyQt).
"""

from calendar import monthrange
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import List, Optional, Set, Tuple

from config.config_manager import AppConfig, TimeRule, ColorLogic
from domain.entities import (
    MonthlyAttendance, 
    StaffType, 
    AttendanceRecord ,
    Staff
)
from domain.staff_classifier import StaffClassifier
from domain.attendance_logic import AttendanceLogicFactory
from domain.rate_calculator import RateCalculator
from domain.sorting import sort_attendance_list
from infrastructure.logger import get_logger

logger = get_logger("ReportService")


@dataclass
class ReportGenerationParams:
    """
    Parameters for report generation.
    
    This dataclass encapsulates all parameters needed for report generation,
    decoupling the service from UI-specific types like AppConfig.
    """
    source_path: Path
    output_path: Path
    staff_csv_path: Path
    
    # Time rules
    internal_time_rule: TimeRule
    external_time_rule: TimeRule
    
    # Holidays
    holidays: Set[date]
    
    # Output settings
    rate_threshold: int = 80
    sort_by: str = "attendance_rate"
    
    # Color logic for Excel
    color_logic: Optional[ColorLogic] = None
    
    # PDF generation
    generate_pdf: bool = True
    separate_pdf: bool = True
    pdf_output_dir: Optional[str] = None
    pdf_filename_pattern: str = "出勤報表_{year}_{month}.pdf"
    internal_pdf_pattern: str = "內勤出勤報表_{year}_{month}.pdf"
    external_pdf_pattern: str = "外勤出勤報表_{year}_{month}.pdf"
    custom_font_path: Optional[str] = None  # Custom font path for PDF


@dataclass
class ReportResult:
    """Result of report generation."""
    success: bool
    output_path: Path
    skipped_names: List[str]
    internal_count: int = 0
    external_count: int = 0
    year: int = 0
    month: int = 0
    error_message: str = ""


class AttendanceReportService:
    """
    Application service for generating attendance reports.
    
    This service:
    - Orchestrates the report generation workflow
    - Depends only on domain entities and infrastructure, not on PyQt
    - Provides logging for key operations
    """
    
    def __init__(self):
        """Initialize the service."""
        pass
    
    def generate_report(self, params: ReportGenerationParams) -> ReportResult:
        """
        Generate the attendance report.
        
        Args:
            params: ReportGenerationParams containing all necessary configuration
            
        Returns:
            ReportResult with the outcome of report generation
            
        Raises:
            ValueError: If no data found or all staff were skipped
            PermissionError: If files cannot be written
        """
        from infrastructure.excel_parser import ExcelParser, UnclassifiedStaffError
        from infrastructure.excel_writer import ExcelWriter
        from infrastructure.filename_parser import FilenameParser
        
        logger.info(f"開始解析來源檔案: {params.source_path}")
        
        # Parse source file
        parser = ExcelParser()
        
        # Try to extract year from filename (MonRepyymmdd format)
        parsed_date = FilenameParser.try_parse_report_date(params.source_path.name)
        year_from_filename = parsed_date[0] if parsed_date else None
        
        raw_data = parser.parse_file(params.source_path, year=year_from_filename)
        
        if not raw_data:
            raise ValueError("來源檔案中沒有找到任何資料")
        
        # Determine year and month from data
        first_date = raw_data[0].date
        year = first_date.year
        month = first_date.month
        
        logger.info(f"解析完成: {year}年{month}月, 共 {len(raw_data)} 筆原始記錄")
        
        # Get records by month
        records_by_name = parser.get_records_by_month(year, month)
        
        if not records_by_name:
            raise ValueError(f"來源檔案中沒有 {year} 年 {month} 月的資料")
        
        # Load staff classification
        classifier = StaffClassifier()
        classifier.load_from_csv(params.staff_csv_path)
        
        # Get number of days in this month
        _, num_days = monthrange(year, month)
        
        # Calculate attendance with strict matching
        rate_calc = RateCalculator()
        
        internal_attendance: List[MonthlyAttendance] = []
        external_attendance: List[MonthlyAttendance] = []
        skipped_names: List[str] = []
        
        for name, raw_rows in records_by_name.items():
            staff = classifier.get_staff_by_name(name)
            
            # Strict Matching: If staff not in list, raise error to prompt user
            if not staff:
                # Previously we skipped, now we pause and ask
                # skipped_names.append(name)
                # continue
                raise UnclassifiedStaffError(name)
            
            # Convert to records
            records = parser.convert_to_attendance_records(raw_rows)
            
            # Apply logic
            strategy = AttendanceLogicFactory.get_strategy(staff.staff_type)
            time_rule = (
                params.internal_time_rule 
                if staff.staff_type == StaffType.INTERNAL 
                else params.external_time_rule
            )
            
            for record in records:
                record.status = strategy.determine_status(record, time_rule)
                record.remark = strategy.get_remark(record, time_rule)
            
            # Calculate work days for this staff member (based on staff type)
            work_days_set = set()
            for day in range(1, num_days + 1):
                d = date(year, month, day)
                if staff.should_work_on(d) and d not in params.holidays:
                    work_days_set.add(day)
            
            # Calculate monthly attendance with work_days filter
            monthly = rate_calc.calculate_monthly_attendance(
                staff, records, year, month,
                params.rate_threshold,
                work_days=work_days_set
            )
            
            if staff.staff_type == StaffType.INTERNAL:
                internal_attendance.append(monthly)
            else:
                external_attendance.append(monthly)
        
        # Scenario C: Complete failure - all staff were skipped
        if not internal_attendance and not external_attendance:
            if skipped_names:
                raise ValueError(
                    f"所有 {len(skipped_names)} 位人員都不在人員名單中，無法產生報表。\n"
                    f"請確認人員名單是否正確。"
                )
            else:
                raise ValueError("沒有任何人員資料可供處理")
        
        # Apply sorting based on config setting
        internal_attendance = sort_attendance_list(internal_attendance, params.sort_by)
        external_attendance = sort_attendance_list(external_attendance, params.sort_by)
        
        logger.info(
            f"資料處理完成: 內勤 {len(internal_attendance)} 人, "
            f"外勤 {len(external_attendance)} 人, 略過 {len(skipped_names)} 人"
        )
        
        # Generate Excel
        logger.info(f"開始寫入 Excel: {params.output_path}")
        writer = ExcelWriter(params.color_logic)
        writer.create_report(
            internal_attendance,
            external_attendance,
            year, month,
            params.output_path,
            holidays=params.holidays
        )
        logger.info("Excel 寫入完成")
        
        # Generate PDF if enabled
        if params.generate_pdf:
            try:
                self._generate_pdf_reports(
                    params, internal_attendance, external_attendance, year, month
                )
            except Exception as e:
                logger.error(f"PDF 生成失敗: {e}")
                # We don't fail the entire process if PDF fails, but we should inform
                # In a future update, we could add a warnings field to ReportResult
        
        return ReportResult(
            success=True,
            output_path=params.output_path,
            skipped_names=skipped_names,
            internal_count=len(internal_attendance),
            external_count=len(external_attendance),
            year=year,
            month=month
        )
    
    def _generate_pdf_reports(
        self,
        params: ReportGenerationParams,
        internal_attendance: List[MonthlyAttendance],
        external_attendance: List[MonthlyAttendance],
        year: int,
        month: int
    ) -> None:
        """Generate PDF report with Excel-like formatting.
        
        Always generates a combined PDF with internal staff first,
        then external staff on a new page.
        """
        from infrastructure.pdf_writer import PdfWriter, format_filename
        
        # Create PdfWriter with color_logic for proper color coding
        pdf_writer = PdfWriter(
            color_logic=params.color_logic,
            custom_font_path=params.custom_font_path
        )
        
        # Determine PDF output directory
        if params.pdf_output_dir:
            pdf_dir = Path(params.pdf_output_dir)
        else:
            pdf_dir = params.output_path.parent
        
        # Ensure PDF output dir exists
        pdf_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate combined PDF (internal + external in one file)
        combined_filename = format_filename(params.pdf_filename_pattern, year, month)
        combined_pdf_path = pdf_dir / combined_filename
        logger.info(f"開始寫入合併 PDF: {combined_pdf_path}")
        
        pdf_writer.create_combined_report(
            internal_attendance,
            external_attendance,
            year, month,
            combined_pdf_path,
            holidays=params.holidays
        )
        logger.info("合併 PDF 寫入完成")
    
    @staticmethod
    def build_params_from_config(
        config: AppConfig,
        source_path: Path,
        output_path: Path,
        generate_pdf: bool = True
    ) -> ReportGenerationParams:
        """
        Build ReportGenerationParams from AppConfig.
        
        This is a convenience method that bridges the config object to
        the params dataclass. UI code can use this to avoid manually
        constructing the params.
        
        Args:
            config: Application configuration (AppConfig from config_manager)
            source_path: Path to source Excel file
            output_path: Path for output Excel file
            generate_pdf: Whether to generate PDF reports
            
        Returns:
            ReportGenerationParams ready for generate_report()
        """
        # Build holidays set from config
        holidays_set: Set[date] = set()
        for date_str in config.holidays.custom_dates:
            try:
                # Expect format: YYYY-MM-DD
                parts = date_str.split('-')
                holidays_set.add(date(int(parts[0]), int(parts[1]), int(parts[2])))
            except (ValueError, IndexError):
                pass
        
        return ReportGenerationParams(
            source_path=source_path,
            output_path=output_path,
            staff_csv_path=Path(config.paths.staff_csv),
            internal_time_rule=config.time_rules.internal,
            external_time_rule=config.time_rules.external,
            holidays=holidays_set,
            rate_threshold=config.ui_prefs.rate_threshold,
            sort_by=config.output_settings.sort_by,
            color_logic=config.ui_prefs.color_logic,
            generate_pdf=generate_pdf,
            separate_pdf=config.output_settings.separate_pdf,
            pdf_output_dir=config.output_settings.pdf_output_dir,
            pdf_filename_pattern=config.output_settings.pdf_filename_pattern,
            internal_pdf_pattern=config.output_settings.internal_pdf_pattern,
            external_pdf_pattern=config.output_settings.external_pdf_pattern,
            custom_font_path=config.paths.custom_font_path or None,
        )
