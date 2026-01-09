"""
Unit tests for PdfWriter including combined report generation.
"""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
import tempfile

import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from infrastructure.pdf_writer import PdfWriter, format_filename, AttendancePdf


class TestFormatFilename:
    """Tests for format_filename utility function."""
    
    def test_basic_formatting(self):
        """Test basic year/month formatting."""
        result = format_filename("Report_{year}_{month}.pdf", 2025, 12)
        assert result == "Report_2025_12.pdf"
    
    def test_month_padding(self):
        """Test that month is zero-padded."""
        result = format_filename("Report_{year}_{month}.pdf", 2025, 1)
        assert result == "Report_2025_01.pdf"
    
    def test_chinese_pattern(self):
        """Test Chinese filename pattern."""
        result = format_filename("內勤出勤報表_{year}_{month}.pdf", 2025, 6)
        assert result == "內勤出勤報表_2025_06.pdf"
    
    def test_combined_pattern(self):
        """Test combined report filename pattern."""
        result = format_filename("出勤報表_{year}_{month}.pdf", 2026, 1)
        assert result == "出勤報表_2026_01.pdf"


class TestPdfWriter:
    """Tests for PdfWriter class."""
    
    def test_create_report_empty_list(self):
        """Test that empty attendance list returns early."""
        writer = PdfWriter()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "test.pdf"
            
            # Should not raise, should return early
            writer.create_report([], 2025, 12, output_path, "internal")
            
            # File should not exist
            assert not output_path.exists()
    
    def test_create_combined_report_empty_lists(self):
        """Test that empty attendance lists return early for combined report."""
        writer = PdfWriter()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "combined.pdf"
            
            # Should not raise, should return early
            writer.create_combined_report([], [], 2025, 12, output_path)
            
            # File should not exist
            assert not output_path.exists()


class TestAttendancePdf:
    """Tests for AttendancePdf class."""
    
    def test_initialization(self):
        """Test AttendancePdf initialization."""
        pdf = AttendancePdf(title="Test Report")
        assert pdf.title_text == "Test Report"
    
    def test_font_family_name(self):
        """Test font family name property."""
        pdf = AttendancePdf(title="Test")
        
        # Should return either ChineseFont or Helvetica
        assert pdf.font_family_name in ["ChineseFont", "Helvetica"]


class TestPdfWriterIntegration:
    """Integration tests for PDF generation (may require actual font files)."""
    
    @pytest.fixture
    def mock_attendance(self):
        """Create mock MonthlyAttendance object."""
        from domain.entities import MonthlyAttendance, Staff, StaffType, RateColorTier
        
        staff = Staff(name="測試員工", staff_type=StaffType.INTERNAL)
        attendance = MonthlyAttendance(
            staff=staff,
            year=2025,
            month=12,
            required_days=22,
            actual_days=20,
            attendance_rate=90.91,
            rate_color=RateColorTier.GREEN,
            records=[]
        )
        return attendance
    
    def test_create_report_generates_file(self, mock_attendance):
        """Test that create_report generates a PDF file."""
        writer = PdfWriter()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "output" / "test.pdf"
            
            writer.create_report(
                [mock_attendance], 2025, 12, 
                output_path, "internal"
            )
            
            assert output_path.exists()
            assert output_path.stat().st_size > 0
    
    def test_create_combined_report_generates_file(self, mock_attendance):
        """Test that create_combined_report generates a PDF file."""
        from domain.entities import MonthlyAttendance, Staff, StaffType, RateColorTier
        
        external_staff = Staff(name="外勤員工", staff_type=StaffType.EXTERNAL)
        external_attendance = MonthlyAttendance(
            staff=external_staff,
            year=2025,
            month=12,
            required_days=22,
            actual_days=18,
            attendance_rate=81.82,
            rate_color=RateColorTier.YELLOW,
            records=[]
        )
        
        writer = PdfWriter()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "combined.pdf"
            
            writer.create_combined_report(
                [mock_attendance], [external_attendance],
                2025, 12, output_path
            )
            
            assert output_path.exists()
            assert output_path.stat().st_size > 0
    
    def test_create_combined_internal_only(self, mock_attendance):
        """Test combined report with only internal staff."""
        writer = PdfWriter()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "internal_only.pdf"
            
            writer.create_combined_report(
                [mock_attendance], [],
                2025, 12, output_path
            )
            
            assert output_path.exists()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
