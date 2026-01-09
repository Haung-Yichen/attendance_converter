"""
PDF Writer Module

Generates formatted PDF attendance reports using fpdf2.
Supports Chinese text via system fonts with cross-platform compatibility.
"""

import sys
from pathlib import Path
from typing import List, Optional, Tuple

from fpdf import FPDF

from domain.entities import MonthlyAttendance, RateColorTier
from infrastructure.logger import get_logger

logger = get_logger("PdfWriter")


# ==============================================================================
# Font Configuration
# ==============================================================================
# Fallback chain for Chinese-capable fonts on Windows.
WINDOWS_FONT_PATHS: List[Path] = [
    Path("C:/Windows/Fonts/msjh.ttc"),       # 微軟正黑體 (Microsoft JhengHei)
    Path("C:/Windows/Fonts/msyh.ttc"),       # 微軟雅黑 (Microsoft YaHei)
    Path("C:/Windows/Fonts/simsun.ttc"),     # 宋體 (SimSun)
    Path("C:/Windows/Fonts/mingliu.ttc"),    # 細明體 (MingLiU)
]

# Fallback chain for Chinese-capable fonts on macOS.
MACOS_FONT_PATHS: List[Path] = [
    Path("/System/Library/Fonts/STHeiti Light.ttc"),      # 黑體-繁
    Path("/System/Library/Fonts/STHeiti Medium.ttc"),     # 黑體-繁 中
    Path("/System/Library/Fonts/PingFang.ttc"),           # 蘋方
    Path("/Library/Fonts/Arial Unicode.ttf"),             # Arial Unicode MS
    Path("/System/Library/Fonts/Supplemental/Songti.ttc"), # 宋體
]

# Fallback chain for Linux (common CJK fonts)
LINUX_FONT_PATHS: List[Path] = [
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
    Path("/usr/share/fonts/truetype/droid/DroidSansFallback.ttf"),
]

FALLBACK_FONT = "Helvetica"  # Last resort if no Chinese font found


def find_chinese_font(custom_font_path: Optional[str] = None) -> Optional[Path]:
    """
    Search for an available Chinese font with cross-platform support.
    
    Priority:
    1. Custom font path from config (if provided and exists)
    2. Platform-specific font search
    3. Matplotlib font_manager fallback (if available)
    
    Args:
        custom_font_path: Optional custom font path from configuration
        
    Returns:
        Path to the first available font, or None if none found.
    """
    # Priority 1: Custom font path from config
    if custom_font_path:
        custom_path = Path(custom_font_path)
        if custom_path.exists():
            logger.info(f"使用自訂字型: {custom_path}")
            return custom_path
        else:
            logger.warning(f"自訂字型路徑不存在: {custom_path}")
    
    # Priority 2: Platform-specific font paths
    platform_fonts = _get_platform_fonts()
    for font_path in platform_fonts:
        if font_path.exists():
            logger.debug(f"找到系統字型: {font_path}")
            return font_path
    
    # Priority 3: Try matplotlib font_manager if available
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
        # Linux and other Unix-like systems
        return LINUX_FONT_PATHS


def _try_matplotlib_font() -> Optional[Path]:
    """
    Try to find a Chinese font using matplotlib's font_manager.
    
    This is a fallback for when platform-specific paths don't work.
    """
    try:
        from matplotlib import font_manager
        
        # Try common Chinese font names
        font_names = [
            'Microsoft JhengHei',  # Windows Traditional Chinese
            'Microsoft YaHei',     # Windows Simplified Chinese
            'SimHei',              # Windows 黑體
            'PingFang TC',         # macOS Traditional Chinese
            'PingFang SC',         # macOS Simplified Chinese
            'Noto Sans CJK TC',    # Google Noto fonts
            'Noto Sans CJK SC',
            'WenQuanYi Micro Hei', # Linux common
        ]
        
        for font_name in font_names:
            font_path = font_manager.findfont(
                font_manager.FontProperties(family=font_name),
                fallback_to_default=False
            )
            if font_path and Path(font_path).exists():
                # Verify it's not the default fallback
                if 'DejaVu' not in font_path and 'Helvetica' not in font_path:
                    logger.info(f"透過 matplotlib 找到字型: {font_path}")
                    return Path(font_path)
        
    except ImportError:
        logger.debug("matplotlib 未安裝，跳過 font_manager 字型搜尋")
    except Exception as e:
        logger.debug(f"matplotlib font_manager 搜尋失敗: {e}")
    
    return None


# ==============================================================================
# AttendancePdf Class
# ==============================================================================
class AttendancePdf(FPDF):
    """
    Custom FPDF class with Chinese font support, header, and footer.
    
    This class handles:
    - Loading and registering Chinese fonts
    - Drawing page headers with report title
    - Drawing page footers with page numbers
    """
    
    # Font family name after registration
    _font_family: str = FALLBACK_FONT
    _font_loaded: bool = False
    
    def __init__(self, title: str = "", custom_font_path: Optional[str] = None):
        """
        Initialize AttendancePdf.
        
        Args:
            title: Report title to display in header
            custom_font_path: Optional custom font path from configuration
        """
        # Portrait A4 for cleaner summary layout
        super().__init__(orientation='P', unit='mm', format='A4')
        self.title_text = title
        self._setup_chinese_font(custom_font_path)
    
    def _setup_chinese_font(self, custom_font_path: Optional[str] = None) -> None:
        """Load Chinese font if available, otherwise fall back to Helvetica."""
        font_path = find_chinese_font(custom_font_path)
        
        if font_path:
            try:
                # Register the font with a simple name
                self.add_font("ChineseFont", "", str(font_path), uni=True)
                self._font_family = "ChineseFont"
                self._font_loaded = True
                logger.info(f"成功載入中文字型: {font_path.name}")
            except Exception as e:
                # Font loading failed, use fallback
                logger.warning(f"無法載入中文字型 {font_path}: {e}")
                self._font_family = FALLBACK_FONT
                self._font_loaded = False
        else:
            logger.warning(
                "無法找到中文字型，PDF 中文將無法正確顯示。\n"
                "建議：\n"
                "  - Windows: 確認 C:\\Windows\\Fonts 中有中文字型\n"
                "  - macOS: 確認系統已安裝中文字型\n"
                "  - 或在設定中指定自訂字型路徑 (custom_font_path)"
            )
            self._font_family = FALLBACK_FONT
            self._font_loaded = False
    
    @property
    def font_family_name(self) -> str:
        """Return the active font family name."""
        return self._font_family
    
    def header(self) -> None:
        """Draw page header with centered title."""
        self.set_font(self._font_family, '', 16)
        self.cell(0, 12, self.title_text, align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(8)
    
    def footer(self) -> None:
        """Draw page footer with page number."""
        self.set_y(-15)
        self.set_font(self._font_family, '', 9)
        self.cell(0, 10, f'第 {self.page_no()}/{{nb}} 頁', align='C')


# ==============================================================================
# PdfWriter Class
# ==============================================================================
class PdfWriter:
    """
    Generates PDF attendance summary reports.
    
    Output format (simplified):
    - 姓名 (Name)
    - 實際出勤天數 (Actual Days)
    - 出席率 (Attendance Rate)
    
    Responsibilities:
    - Coordinating PDF creation
    - Delegating drawing to specialized methods
    - NOT performing data processing (SRP)
    """
    
    # RGB Color definitions
    COLORS = {
        'green': (144, 238, 144),    # Light green - good rate
        'yellow': (255, 215, 0),      # Gold - warning rate
        'red': (255, 107, 107),       # Light red - poor rate
        'header': (68, 114, 196),     # Blue header background
        'white': (255, 255, 255),
        'black': (0, 0, 0),
        'gray': (245, 245, 245),      # Alternating row background
    }
    
    # Table layout (mm) - optimized for 3-column summary
    # A4 Portrait width = 210mm, margins = 10mm each side → 190mm available
    MARGIN = 10
    TABLE_WIDTH = 190
    COL_WIDTHS = {
        'name': 80,          # 姓名
        'actual_days': 55,   # 實際出勤天數
        'rate': 55,          # 出席率
    }
    ROW_HEIGHT = 10
    
    def __init__(self, custom_font_path: Optional[str] = None):
        """
        Initialize PdfWriter.
        
        Args:
            custom_font_path: Optional custom font path for PDF generation
        """
        self._custom_font_path = custom_font_path
    
    def create_report(
        self,
        attendance_list: List[MonthlyAttendance],
        year: int,
        month: int,
        output_path: Path,
        report_type: str = "internal"
    ) -> None:
        """
        Create PDF summary report.
        
        Args:
            attendance_list: List of monthly attendance records
            year: Report year
            month: Report month
            output_path: Output file path (.pdf)
            report_type: "internal" or "external" (for title display)
        """
        if not attendance_list:
            return
        
        type_label = "內勤" if report_type == "internal" else "外勤"
        title = f"{year}年{month}月 {type_label}出勤報表"
        
        pdf = AttendancePdf(title=title, custom_font_path=self._custom_font_path)
        pdf.alias_nb_pages()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Draw table
        self._draw_table_header(pdf)
        self._draw_table_body(pdf, attendance_list)
        
        # Save
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))
    
    def create_combined_report(
        self,
        internal_list: List[MonthlyAttendance],
        external_list: List[MonthlyAttendance],
        year: int,
        month: int,
        output_path: Path
    ) -> None:
        """
        Create combined PDF report with both internal and external staff.
        
        Args:
            internal_list: List of internal staff attendance records
            external_list: List of external staff attendance records
            year: Report year
            month: Report month
            output_path: Output file path (.pdf)
        """
        if not internal_list and not external_list:
            return
        
        title = f"{year}年{month}月 出勤報表"
        pdf = AttendancePdf(title=title, custom_font_path=self._custom_font_path)
        pdf.alias_nb_pages()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Internal staff section
        if internal_list:
            pdf.add_page()
            pdf.set_font(pdf.font_family_name, '', 14)
            pdf.cell(0, 10, "【內勤】", new_x='LMARGIN', new_y='NEXT')
            pdf.ln(5)
            self._draw_table_header(pdf)
            self._draw_table_body(pdf, internal_list)
        
        # External staff section
        if external_list:
            pdf.add_page()
            pdf.set_font(pdf.font_family_name, '', 14)
            pdf.cell(0, 10, "【外勤】", new_x='LMARGIN', new_y='NEXT')
            pdf.ln(5)
            self._draw_table_header(pdf)
            self._draw_table_body(pdf, external_list)
        
        # Save
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pdf.output(str(output_path))
    
    def _draw_table_header(self, pdf: AttendancePdf) -> None:
        """Draw the table header row."""
        pdf.set_font(pdf.font_family_name, '', 11)
        pdf.set_fill_color(*self.COLORS['header'])
        pdf.set_text_color(*self.COLORS['white'])
        
        # Center the table
        x_start = (210 - self.TABLE_WIDTH) / 2
        pdf.set_x(x_start)
        
        headers = [
            ('姓名', self.COL_WIDTHS['name']),
            ('實際出勤天數', self.COL_WIDTHS['actual_days']),
            ('出席率', self.COL_WIDTHS['rate']),
        ]
        
        for text, width in headers:
            pdf.cell(width, self.ROW_HEIGHT, text, border=1, align='C', fill=True)
        
        pdf.ln()
        pdf.set_text_color(*self.COLORS['black'])
    
    def _draw_table_body(
        self, 
        pdf: AttendancePdf, 
        attendance_list: List[MonthlyAttendance]
    ) -> None:
        """Draw all data rows."""
        pdf.set_font(pdf.font_family_name, '', 10)
        x_start = (210 - self.TABLE_WIDTH) / 2
        
        for idx, attendance in enumerate(attendance_list):
            # Alternating row background
            use_gray_bg = (idx % 2 == 1)
            
            pdf.set_x(x_start)
            self._draw_data_row(pdf, attendance, use_gray_bg)
    
    def _draw_data_row(
        self, 
        pdf: AttendancePdf, 
        attendance: MonthlyAttendance,
        use_gray_bg: bool
    ) -> None:
        """
        Draw a single data row.
        
        Args:
            pdf: The PDF instance
            attendance: MonthlyAttendance data for this row
            use_gray_bg: Whether to use gray background for alternating rows
        """
        # Prepare cell data
        name = attendance.staff.name
        actual_days = str(attendance.actual_days)
        rate_str = f"{attendance.attendance_rate:.1f}%"
        rate_color = self._get_rate_color(attendance.rate_color)
        
        # Row background
        bg_color = self.COLORS['gray'] if use_gray_bg else self.COLORS['white']
        
        # Name cell
        pdf.set_fill_color(*bg_color)
        pdf.cell(
            self.COL_WIDTHS['name'], self.ROW_HEIGHT, 
            name, border=1, align='L', fill=True
        )
        
        # Actual days cell
        pdf.cell(
            self.COL_WIDTHS['actual_days'], self.ROW_HEIGHT,
            actual_days, border=1, align='C', fill=True
        )
        
        # Rate cell - colored by tier
        pdf.set_fill_color(*rate_color)
        pdf.cell(
            self.COL_WIDTHS['rate'], self.ROW_HEIGHT,
            rate_str, border=1, align='C', fill=True
        )
        
        pdf.ln()
    
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
