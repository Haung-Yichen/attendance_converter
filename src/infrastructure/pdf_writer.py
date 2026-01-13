"""
PDF Writer Module

Generates formatted PDF attendance reports using fpdf2.
Replicates the Excel report format with color coding and table structure.
"""

import sys
from calendar import monthrange
from datetime import date
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Set

from fpdf import FPDF

from domain.entities import (
    MonthlyAttendance, AttendanceRecord, AttendanceStatus, RateColorTier
)
from config.config_manager import ColorLogic
from infrastructure.logger import get_logger

logger = get_logger("PdfWriter")


# ==============================================================================
# Font Configuration
# ==============================================================================
WINDOWS_FONT_PATHS: List[Path] = [
    Path("C:/Windows/Fonts/msjh.ttc"),       # 微軟正黑體 (Microsoft JhengHei)
    Path("C:/Windows/Fonts/msyh.ttc"),       # 微軟雅黑 (Microsoft YaHei)
    Path("C:/Windows/Fonts/simsun.ttc"),     # 宋體 (SimSun)
    Path("C:/Windows/Fonts/mingliu.ttc"),    # 細明體 (MingLiU)
]

MACOS_FONT_PATHS: List[Path] = [
    Path("/System/Library/Fonts/STHeiti Light.ttc"),
    Path("/System/Library/Fonts/STHeiti Medium.ttc"),
    Path("/System/Library/Fonts/PingFang.ttc"),
    Path("/Library/Fonts/Arial Unicode.ttf"),
    Path("/System/Library/Fonts/Supplemental/Songti.ttc"),
]

LINUX_FONT_PATHS: List[Path] = [
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
    Path("/usr/share/fonts/truetype/droid/DroidSansFallback.ttf"),
]

FALLBACK_FONT = "Helvetica"


def find_chinese_font(custom_font_path: Optional[str] = None) -> Optional[Path]:
    """
    Search for an available Chinese font with cross-platform support.
    """
    if custom_font_path:
        custom_path = Path(custom_font_path)
        if custom_path.exists():
            logger.info(f"使用自訂字型: {custom_path}")
            return custom_path
        else:
            logger.warning(f"自訂字型路徑不存在: {custom_path}")

    platform_fonts = _get_platform_fonts()
    for font_path in platform_fonts:
        if font_path.exists():
            logger.debug(f"找到系統字型: {font_path}")
            return font_path

    font_from_matplotlib = _try_matplotlib_font()
    if font_from_matplotlib:
        return font_from_matplotlib

    return None


def _get_platform_fonts() -> List[Path]:
    """Get the font search list for the current platform."""
    if sys.platform == 'win32':
        return WINDOWS_FONT_PATHS
    elif sys.platform == 'darwin':
        return MACOS_FONT_PATHS
    else:
        return LINUX_FONT_PATHS


def _try_matplotlib_font() -> Optional[Path]:
    """Try to find a Chinese font using matplotlib's font_manager."""
    try:
        from matplotlib import font_manager

        font_names = [
            'Microsoft JhengHei', 'Microsoft YaHei', 'SimHei',
            'PingFang TC', 'PingFang SC', 'Noto Sans CJK TC',
            'Noto Sans CJK SC', 'WenQuanYi Micro Hei',
        ]

        for font_name in font_names:
            font_path = font_manager.findfont(
                font_manager.FontProperties(family=font_name),
                fallback_to_default=False
            )
            if font_path and Path(font_path).exists():
                if 'DejaVu' not in font_path and 'Helvetica' not in font_path:
                    logger.info(f"透過 matplotlib 找到字型: {font_path}")
                    return Path(font_path)

    except ImportError:
        logger.debug("matplotlib 未安裝，跳過 font_manager 字型搜尋")
    except Exception as e:
        logger.debug(f"matplotlib font_manager 搜尋失敗: {e}")

    return None


# ==============================================================================
# AttendancePdf Class (A3 Landscape)
# ==============================================================================
class AttendancePdf(FPDF):
    """
    Custom FPDF class with Chinese font support for A3 landscape attendance reports.
    """

    _font_family: str = FALLBACK_FONT
    _font_loaded: bool = False

    def __init__(self, title: str = "", custom_font_path: Optional[str] = None):
        # A3 Landscape: 420mm x 297mm
        super().__init__(orientation='L', unit='mm', format='A3')
        self.title_text = title
        self._setup_chinese_font(custom_font_path)

    def _setup_chinese_font(self, custom_font_path: Optional[str] = None) -> None:
        """Load Chinese font if available."""
        font_path = find_chinese_font(custom_font_path)

        if font_path:
            try:
                self.add_font("ChineseFont", "", str(font_path), uni=True)
                self._font_family = "ChineseFont"
                self._font_loaded = True
                logger.info(f"成功載入中文字型: {font_path.name}")
            except Exception as e:
                logger.warning(f"無法載入中文字型 {font_path}: {e}")
                self._font_family = FALLBACK_FONT
                self._font_loaded = False
        else:
            logger.warning("無法找到中文字型，PDF 中文將無法正確顯示。")
            self._font_family = FALLBACK_FONT
            self._font_loaded = False

    @property
    def font_family_name(self) -> str:
        return self._font_family

    def header(self) -> None:
        """Draw page header with centered title."""
        self.set_font(self._font_family, '', 14)
        self.cell(0, 10, self.title_text, align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(3)

    def footer(self) -> None:
        """Draw page footer with page number."""
        self.set_y(-12)
        self.set_font(self._font_family, '', 8)
        self.cell(0, 10, f'第 {self.page_no()}/{{nb}} 頁', align='C')


# ==============================================================================
# PdfWriter Class
# ==============================================================================
class PdfWriter:
    """
    Generates PDF attendance reports that replicate the Excel format.

    Features:
    - A3 Landscape to accommodate 31+ columns
    - Two-row-per-person layout (check-in / check-out)
    - Color coding based on ColorLogic settings
    - Thick borders around each person's block
    """

    # RGB Color definitions (matching ExcelWriter)
    COLORS: Dict[str, Tuple[int, int, int]] = {
        'green': (144, 238, 144),
        'red': (255, 107, 107),
        'yellow': (255, 215, 0),
        'orange': (255, 165, 0),
        'blue': (107, 140, 255),
        'purple': (221, 160, 221),
        'pink': (255, 182, 193),
        'black': (51, 51, 51),
        'gray': (211, 211, 211),
        'header': (68, 114, 196),
        'white': (255, 255, 255),
    }

    # Hex to color name mapping
    HEX_TO_NAME: Dict[str, str] = {
        '#90EE90': 'green', '#90ee90': 'green',
        '#FF6B6B': 'red', '#ff6b6b': 'red',
        '#FFD700': 'yellow', '#ffd700': 'yellow',
        '#FFA500': 'orange', '#ffa500': 'orange',
        '#6B8CFF': 'blue', '#6b8cff': 'blue',
        '#DDA0DD': 'purple', '#dda0dd': 'purple',
        '#FFB6C1': 'pink', '#ffb6c1': 'pink',
        '#333333': 'black',
        '#D3D3D3': 'gray', '#d3d3d3': 'gray',
    }

    # Layout constants (mm) for A3 Landscape (420mm width)
    MARGIN = 8
    PAGE_WIDTH = 420
    PAGE_HEIGHT = 297
    
    NAME_COL_WIDTH = 22
    DAY_COL_WIDTH = 10.5
    REMARKS_COL_WIDTH = 30
    ACTUAL_COL_WIDTH = 16
    RATE_COL_WIDTH = 14
    
    HEADER_ROW_HEIGHT = 18
    DATA_ROW_HEIGHT = 6
    
    THIN_LINE = 0.2
    THICK_LINE = 0.6

    def __init__(
        self,
        color_logic: Optional[ColorLogic] = None,
        custom_font_path: Optional[str] = None
    ):
        self._color_logic = color_logic or ColorLogic()
        self._custom_font_path = custom_font_path

    def _get_rgb(self, color_value: str) -> Optional[Tuple[int, int, int]]:
        """Get RGB tuple from color name or hex code."""
        if not color_value or color_value in ('none', 'transparent'):
            return None

        if color_value in self.COLORS:
            return self.COLORS[color_value]

        if color_value in self.HEX_TO_NAME:
            return self.COLORS[self.HEX_TO_NAME[color_value]]

        return None

    def _is_dark_color(self, color_value: str) -> bool:
        """Check if color is dark (needs white text)."""
        dark_colors = {'black', 'blue', 'purple', 'header'}
        if color_value in dark_colors:
            return True
        if color_value in self.HEX_TO_NAME:
            return self.HEX_TO_NAME[color_value] in dark_colors
        return False

    def create_combined_report(
        self,
        internal_list: List[MonthlyAttendance],
        external_list: List[MonthlyAttendance],
        year: int,
        month: int,
        output_path: Path,
        holidays: Optional[Set[date]] = None
    ) -> None:
        """
        Create a combined PDF report with internal and external staff.

        Internal staff section is drawn first, followed by a page break,
        then the external staff section.
        """
        if not internal_list and not external_list:
            return

        title = f"{year}年{month}月 出勤報表"
        pdf = AttendancePdf(title=title, custom_font_path=self._custom_font_path)
        pdf.alias_nb_pages()
        pdf.set_auto_page_break(auto=False)

        holidays = holidays or set()

        # Internal Staff Section
        if internal_list:
            pdf.add_page()
            self._draw_section(
                pdf, internal_list, year, month,
                section_title="【內勤】",
                is_external=False,
                holidays=holidays
            )

        # External Staff Section (new page)
        if external_list:
            pdf.add_page()
            self._draw_section(
                pdf, external_list, year, month,
                section_title="【外勤】",
                is_external=True,
                holidays=holidays
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))
        logger.info(f"PDF 報表已儲存: {output_path}")

    def _draw_section(
        self,
        pdf: AttendancePdf,
        attendance_list: List[MonthlyAttendance],
        year: int,
        month: int,
        section_title: str,
        is_external: bool,
        holidays: Set[date]
    ) -> None:
        """Draw a complete section (internal or external) on the current page."""
        _, num_days = monthrange(year, month)

        # Determine work days
        if is_external:
            work_weekdays = {0, 2, 4}  # Mon, Wed, Fri
        else:
            work_weekdays = {0, 1, 2, 3, 4}  # Mon-Fri

        work_days = []
        for day in range(1, num_days + 1):
            d = date(year, month, day)
            if d.weekday() in work_weekdays and d not in holidays:
                work_days.append(day)

        weekday_names = ['一', '二', '三', '四', '五', '六', '日']
        num_work_days = len(work_days)

        # ─────────────────────────────────────────────────────────────────────
        # Dynamic layout calculation to fill the page width
        # ─────────────────────────────────────────────────────────────────────
        available_width = self.PAGE_WIDTH - 2 * self.MARGIN
        
        # Fixed column widths (scaled up for better readability)
        name_col_width = 28
        remarks_col_width = 38
        actual_col_width = 20
        rate_col_width = 18
        
        fixed_cols_width = name_col_width + remarks_col_width + actual_col_width + rate_col_width
        
        # Calculate day column width to fill remaining space
        remaining_width = available_width - fixed_cols_width
        day_col_width = remaining_width / num_work_days if num_work_days > 0 else 12
        
        # Ensure minimum day column width
        day_col_width = max(day_col_width, 10)
        
        # Calculate actual table width
        table_width = fixed_cols_width + day_col_width * num_work_days
        
        # Center the table horizontally
        start_x = (self.PAGE_WIDTH - table_width) / 2

        # Section title (centered)
        pdf.set_font(pdf.font_family_name, '', 14)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(0, 10, section_title, new_x='LMARGIN', new_y='NEXT', align='C')
        pdf.ln(3)

        # Store layout info for drawing methods
        layout = {
            'name_col_width': name_col_width,
            'day_col_width': day_col_width,
            'remarks_col_width': remarks_col_width,
            'actual_col_width': actual_col_width,
            'rate_col_width': rate_col_width,
        }

        # Draw header row
        self._draw_header_row(pdf, work_days, year, month, weekday_names, start_x, layout)

        # Draw data rows (2 rows per person)
        for monthly in attendance_list:
            self._draw_person_rows(pdf, monthly, work_days, start_x, num_work_days, layout)

    def _draw_header_row(
        self,
        pdf: AttendancePdf,
        work_days: List[int],
        year: int,
        month: int,
        weekday_names: List[str],
        start_x: float,
        layout: Dict[str, float]
    ) -> None:
        """Draw the header row with date labels."""
        name_w = layout['name_col_width']
        day_w = layout['day_col_width']
        remarks_w = layout['remarks_col_width']
        actual_w = layout['actual_col_width']
        rate_w = layout['rate_col_width']
        header_h = 20  # Increased header height

        pdf.set_font(pdf.font_family_name, '', 8)
        pdf.set_fill_color(*self.COLORS['header'])
        pdf.set_text_color(255, 255, 255)
        pdf.set_line_width(self.THIN_LINE)

        x = start_x
        y = pdf.get_y()

        # Name header
        pdf.set_xy(x, y)
        pdf.cell(name_w, header_h, "姓名", border=1, align='C', fill=True)
        x += name_w

        # Date headers
        for day in work_days:
            d = date(year, month, day)
            weekday_str = weekday_names[d.weekday()]
            short_label = f"{day}\n({weekday_str})"

            pdf.set_xy(x, y)
            pdf.multi_cell(day_w, header_h / 2, short_label, border=1, align='C', fill=True)
            x += day_w

        # Summary headers
        pdf.set_xy(x, y)
        pdf.cell(remarks_w, header_h, "備註", border=1, align='C', fill=True)
        x += remarks_w

        pdf.set_xy(x, y)
        pdf.multi_cell(actual_w, header_h / 2, "實際\n出勤", border=1, align='C', fill=True)
        x += actual_w

        pdf.set_xy(x, y)
        pdf.cell(rate_w, header_h, "出席率", border=1, align='C', fill=True)

        # Move to next row
        pdf.set_xy(start_x, y + header_h)
        pdf.set_text_color(0, 0, 0)

    def _draw_person_rows(
        self,
        pdf: AttendancePdf,
        monthly: MonthlyAttendance,
        work_days: List[int],
        start_x: float,
        num_work_days: int,
        layout: Dict[str, float]
    ) -> None:
        """
        Draw two rows for a single person.
        - Name, Remarks, Actual, Rate: merged (span 2 rows), vertically centered
        - Day columns: separate rows for check-in / check-out
        """
        name_w = layout['name_col_width']
        day_w = layout['day_col_width']
        remarks_w = layout['remarks_col_width']
        actual_w = layout['actual_col_width']
        rate_w = layout['rate_col_width']
        row_h = 10  # Increased row height for larger text

        records_by_day = {r.date.day: r for r in monthly.records}
        staff = monthly.staff

        y_start = pdf.get_y()
        merged_h = row_h * 2  # Height for merged cells

        # Check if we need a new page
        if y_start + merged_h > self.PAGE_HEIGHT - 15:
            pdf.add_page()
            y_start = pdf.get_y() + 5

        # Prepare cell data
        in_values: List[Tuple[str, Optional[Tuple[int, int, int]]]] = []
        out_values: List[Tuple[str, Optional[Tuple[int, int, int]]]] = []

        for day in work_days:
            record = records_by_day.get(day)
            in_text, in_color, out_text, out_color = self._get_cell_data(record)
            in_values.append((in_text, in_color))
            out_values.append((out_text, out_color))

        # Calculate remarks
        late_count = 0
        early_count = 0
        overtime_count = 0

        for day in work_days:
            record = records_by_day.get(day)
            if not record:
                continue
            if record.status == AttendanceStatus.LATE:
                late_count += 1
            elif record.status == AttendanceStatus.EARLY_LEAVE:
                early_count += 1
            if record.check_out and record.remark == "下班延遲打卡":
                overtime_count += 1

        remark_parts = []
        if late_count > 0:
            remark_parts.append(f"遲到{late_count}天")
        if early_count > 0:
            remark_parts.append(f"早退{early_count}天")
        if overtime_count > 0:
            remark_parts.append(f"超時{overtime_count}天")
        remarks_text = ", ".join(remark_parts)

        x = start_x

        # ─────────────────────────────────────────────────────────────────────
        # 1. Name column (merged, vertically centered)
        # ─────────────────────────────────────────────────────────────────────
        self._draw_merged_cell(
            pdf, x, y_start, name_w, merged_h, staff.name,
            is_left=True, is_right=False, font_size=10
        )
        x += name_w

        # ─────────────────────────────────────────────────────────────────────
        # 2. Day columns (NOT merged - separate rows)
        # ─────────────────────────────────────────────────────────────────────
        for i, day in enumerate(work_days):
            in_text, in_color = in_values[i]
            out_text, out_color = out_values[i]
            is_first_day = (i == 0)
            is_last_day = (i == num_work_days - 1)

            # Check-in row (top)
            self._draw_cell_with_border(
                pdf, x, y_start, day_w, row_h, in_text,
                self.THICK_LINE, self.THIN_LINE,
                self.THIN_LINE, self.THIN_LINE,
                in_color, align='C', font_size=9
            )
            # Check-out row (bottom)
            self._draw_cell_with_border(
                pdf, x, y_start + row_h, day_w, row_h, out_text,
                self.THIN_LINE, self.THICK_LINE,
                self.THIN_LINE, self.THIN_LINE,
                out_color, align='C', font_size=9
            )
            x += day_w

        # ─────────────────────────────────────────────────────────────────────
        # 3. Remarks column (merged, vertically centered)
        # ─────────────────────────────────────────────────────────────────────
        self._draw_merged_cell(
            pdf, x, y_start, remarks_w, merged_h, remarks_text,
            is_left=False, is_right=False, font_size=9, align='L'
        )
        x += remarks_w

        # ─────────────────────────────────────────────────────────────────────
        # 4. Actual days column (merged, vertically centered)
        # ─────────────────────────────────────────────────────────────────────
        self._draw_merged_cell(
            pdf, x, y_start, actual_w, merged_h, str(monthly.actual_days),
            is_left=False, is_right=False, font_size=10
        )
        x += actual_w

        # ─────────────────────────────────────────────────────────────────────
        # 5. Rate column (merged, vertically centered, with color)
        # ─────────────────────────────────────────────────────────────────────
        rate_color = self._get_rate_color(monthly.rate_color)
        self._draw_merged_cell(
            pdf, x, y_start, rate_w, merged_h,
            f"{monthly.attendance_rate:.1f}%",
            is_left=False, is_right=True, font_size=10,
            fill_color=rate_color
        )

        pdf.set_xy(start_x, y_start + merged_h)

    def _draw_merged_cell(
        self,
        pdf: AttendancePdf,
        x: float, y: float,
        width: float, height: float,
        text: str,
        is_left: bool = False,
        is_right: bool = False,
        font_size: int = 9,
        align: str = 'C',
        fill_color: Optional[Tuple[int, int, int]] = None
    ) -> None:
        """Draw a merged cell spanning two rows with vertically centered text."""
        # Fill background
        if fill_color:
            pdf.set_fill_color(*fill_color)
            pdf.rect(x, y, width, height, style='F')

        # Draw thick border around entire merged cell
        left_w = self.THICK_LINE if is_left else self.THIN_LINE
        right_w = self.THICK_LINE if is_right else self.THIN_LINE

        pdf.set_line_width(self.THICK_LINE)
        pdf.line(x, y, x + width, y)  # Top
        pdf.line(x, y + height, x + width, y + height)  # Bottom

        pdf.set_line_width(left_w)
        pdf.line(x, y, x, y + height)  # Left

        pdf.set_line_width(right_w)
        pdf.line(x + width, y, x + width, y + height)  # Right

        # Text color
        if fill_color and self._is_fill_dark(fill_color):
            pdf.set_text_color(255, 255, 255)
        else:
            pdf.set_text_color(0, 0, 0)

        # Draw text vertically centered
        pdf.set_font(pdf.font_family_name, '', font_size)
        text_h = font_size * 0.35  # Approximate text height in mm
        text_y = y + (height - text_h) / 2
        pdf.set_xy(x, text_y)
        pdf.cell(width, text_h, text, align=align)

    def _draw_cell_with_border(
        self,
        pdf: AttendancePdf,
        x: float, y: float,
        width: float, height: float,
        text: str,
        top_w: float, bottom_w: float,
        left_w: float, right_w: float,
        fill_color: Optional[Tuple[int, int, int]],
        align: str = 'C',
        font_size: int = 9
    ) -> None:
        """Draw a cell with custom border widths and optional fill color."""
        # Fill background
        if fill_color:
            pdf.set_fill_color(*fill_color)
            pdf.rect(x, y, width, height, style='F')

        # Draw borders
        pdf.set_line_width(top_w)
        pdf.line(x, y, x + width, y)
        pdf.set_line_width(bottom_w)
        pdf.line(x, y + height, x + width, y + height)
        pdf.set_line_width(left_w)
        pdf.line(x, y, x, y + height)
        pdf.set_line_width(right_w)
        pdf.line(x + width, y, x + width, y + height)

        # Text color
        if fill_color and self._is_fill_dark(fill_color):
            pdf.set_text_color(255, 255, 255)
        else:
            pdf.set_text_color(0, 0, 0)

        # Draw text vertically centered
        pdf.set_font(pdf.font_family_name, '', font_size)
        text_h = font_size * 0.35
        text_y = y + (height - text_h) / 2
        pdf.set_xy(x, text_y)
        pdf.cell(width, text_h, text, align=align)

    def _is_fill_dark(self, rgb: Tuple[int, int, int]) -> bool:
        """Check if RGB color is dark (needs white text)."""
        r, g, b = rgb
        # Simple luminance check
        luminance = (0.299 * r + 0.587 * g + 0.114 * b)
        return luminance < 128

    def _get_cell_data(
        self,
        record: Optional[AttendanceRecord]
    ) -> Tuple[str, Optional[Tuple[int, int, int]], str, Optional[Tuple[int, int, int]]]:
        """
        Get cell text and color for check-in and check-out.

        Returns: (in_text, in_color, out_text, out_color)
        """
        if not record:
            # No record = absent
            in_text = self._color_logic.absent_text
            out_text = self._color_logic.absent_text
            in_color = self._get_rgb(self._color_logic.absent_color)
            out_color = self._get_rgb(self._color_logic.absent_color)
            return in_text, in_color, out_text, out_color

        has_in = record.check_in is not None
        has_out = record.check_out is not None

        if has_in and has_out:
            in_text = record.check_in.strftime('%H:%M')
            out_text = record.check_out.strftime('%H:%M')
            in_color, out_color = self._get_status_colors(record)

        elif has_in and not has_out:
            in_text = record.check_in.strftime('%H:%M')
            out_text = self._color_logic.missing_punch_text
            in_color, _ = self._get_status_colors(record)
            out_color = self._get_rgb(self._color_logic.missing_punch_color)

        elif not has_in and has_out:
            in_text = self._color_logic.missing_punch_text
            out_text = record.check_out.strftime('%H:%M')
            in_color = self._get_rgb(self._color_logic.missing_punch_color)
            _, out_color = self._get_status_colors(record)

        else:
            # Both missing = absent
            in_text = self._color_logic.absent_text
            out_text = self._color_logic.absent_text
            in_color = self._get_rgb(self._color_logic.absent_color)
            out_color = self._get_rgb(self._color_logic.absent_color)

        return in_text, in_color, out_text, out_color

    def _get_status_colors(
        self,
        record: AttendanceRecord
    ) -> Tuple[Optional[Tuple[int, int, int]], Optional[Tuple[int, int, int]]]:
        """Get colors based on attendance status (mirrors ExcelWriter logic)."""
        in_color = None
        out_color = None

        if record.status == AttendanceStatus.NORMAL:
            if record.check_in:
                in_color = self._get_rgb(self._color_logic.normal_in_color)
            if record.check_out:
                out_color = self._get_rgb(self._color_logic.normal_out_color)

        elif record.status == AttendanceStatus.LATE:
            in_color = self._get_rgb(self._color_logic.abnormal_in_color)
            if record.check_out:
                out_color = self._get_rgb(self._color_logic.normal_out_color)

        elif record.status == AttendanceStatus.EARLY_LEAVE:
            if record.check_in:
                in_color = self._get_rgb(self._color_logic.normal_in_color)
            out_color = self._get_rgb(self._color_logic.abnormal_out_color)

        elif record.status in (AttendanceStatus.ABNORMAL, AttendanceStatus.ABSENT):
            if record.check_in:
                in_color = self._get_rgb(self._color_logic.abnormal_in_color)
            if record.check_out:
                out_color = self._get_rgb(self._color_logic.abnormal_out_color)

        return in_color, out_color

    def _get_rate_color(self, rate_tier: RateColorTier) -> Tuple[int, int, int]:
        """Map RateColorTier enum to RGB color tuple."""
        if rate_tier == RateColorTier.GREEN:
            return self.COLORS['green']
        elif rate_tier == RateColorTier.YELLOW:
            return self.COLORS['yellow']
        else:
            return self.COLORS['red']


# ==============================================================================
# Utility Functions
# ==============================================================================
def format_filename(pattern: str, year: int, month: int) -> str:
    """Format filename pattern with placeholders."""
    return pattern.format(
        year=year,
        month=f"{month:02d}"
    )
