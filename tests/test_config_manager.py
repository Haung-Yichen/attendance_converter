"""
Unit tests for ConfigManager and ColorLogic/OutputSettings dataclasses.
"""

import pytest
import json
import tempfile
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from config.config_manager import (
    ConfigManager, AppConfig, ColorLogic, OutputSettings,
    TimeRule, TimeRules, Paths, Holidays, UIPrefs
)


class TestColorLogic:
    """Tests for ColorLogic dataclass."""
    
    def test_default_values(self):
        """Test default color logic values."""
        cl = ColorLogic()
        
        # New color string fields
        assert cl.normal_in_color == "green"
        assert cl.normal_out_color == "green"
        assert cl.abnormal_in_color == "red"
        assert cl.abnormal_out_color == "red"
        assert cl.missing_punch_color == "black"
        assert cl.missing_punch_text == "*"
        assert cl.absent_color == "none"
        assert cl.absent_text == "-"
        
        # Legacy bool fields (backward compatibility)
        assert cl.green_normal_in is True
        assert cl.green_normal_out is True
        assert cl.red_abnormal_in is True
        assert cl.red_abnormal_out is True
    
    def test_custom_values(self):
        """Test custom color logic values."""
        cl = ColorLogic(
            normal_in_color="blue",
            normal_out_color="purple",
            abnormal_in_color="orange",
            abnormal_out_color="yellow",
            missing_punch_color="pink",
            missing_punch_text="?",
            absent_color="red",
            absent_text="X"
        )
        
        assert cl.normal_in_color == "blue"
        assert cl.normal_out_color == "purple"
        assert cl.abnormal_in_color == "orange"
        assert cl.abnormal_out_color == "yellow"
        assert cl.missing_punch_color == "pink"
        assert cl.missing_punch_text == "?"
        assert cl.absent_color == "red"
        assert cl.absent_text == "X"


class TestOutputSettings:
    """Tests for OutputSettings dataclass."""
    
    def test_default_values(self):
        """Test default output settings values."""
        os = OutputSettings()
        
        assert os.output_dir == ""
        assert os.filename_pattern == "Attendance_{year}_{month}.xlsx"
        assert os.generate_pdf is True
        assert os.separate_pdf is True
        assert os.pdf_output_dir == ""
        assert os.pdf_filename_pattern == "出勤報表_{year}_{month}.pdf"
        assert os.internal_pdf_pattern == "內勤出勤報表_{year}_{month}.pdf"
        assert os.external_pdf_pattern == "外勤出勤報表_{year}_{month}.pdf"
    
    def test_combined_pdf_settings(self):
        """Test settings for combined PDF generation."""
        os = OutputSettings(
            separate_pdf=False,
            pdf_output_dir="C:/output/pdf",
            pdf_filename_pattern="combined_{year}_{month}.pdf"
        )
        
        assert os.separate_pdf is False
        assert os.pdf_output_dir == "C:/output/pdf"
        assert os.pdf_filename_pattern == "combined_{year}_{month}.pdf"


class TestConfigManager:
    """Tests for ConfigManager class."""
    
    def test_load_default_config(self):
        """Test loading default config when file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.json"
            manager = ConfigManager(config_path)
            config = manager.load()
            
            assert isinstance(config, AppConfig)
            assert config.ui_prefs.color_logic.normal_in_color == "green"
            assert config.output_settings.separate_pdf is True
    
    def test_save_and_load_config(self):
        """Test saving and loading config with new fields."""
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.json"
            manager = ConfigManager(config_path)
            
            # Load default and modify
            config = manager.load()
            config.ui_prefs.color_logic.normal_in_color = "blue"
            config.ui_prefs.color_logic.missing_punch_text = "?"
            config.output_settings.separate_pdf = False
            config.output_settings.pdf_output_dir = "/custom/path"
            
            # Save
            manager.save()
            
            # Reload
            manager2 = ConfigManager(config_path)
            config2 = manager2.load()
            
            assert config2.ui_prefs.color_logic.normal_in_color == "blue"
            assert config2.ui_prefs.color_logic.missing_punch_text == "?"
            assert config2.output_settings.separate_pdf is False
            assert config2.output_settings.pdf_output_dir == "/custom/path"
    
    def test_backward_compatibility(self):
        """Test loading old config format with only bool fields."""
        old_config_data = {
            "paths": {"staff_csv": "", "leave_list": "", "last_source_file": ""},
            "holidays": {"use_auto_fetch": True, "custom_dates": []},
            "time_rules": {
                "internal": {"in_start": "09:00", "in_end": "09:30", "out_start": "18:00", "out_end": "18:30"},
                "external": {"in_start": "09:30", "in_end": "10:00", "out_start": "10:30", "out_end": "12:00"}
            },
            "ui_prefs": {
                "color_logic": {
                    "green_normal_in": True,
                    "green_normal_out": False,
                    "red_abnormal_in": True,
                    "red_abnormal_out": False
                },
                "rate_threshold": 80,
                "theme_name": "Dark Mode"
            },
            "output_settings": {
                "output_dir": "",
                "filename_pattern": "Attendance_{year}_{month}.xlsx",
                "generate_pdf": False,
                "internal_pdf_pattern": "內勤出勤報表_{year}_{month}.pdf",
                "external_pdf_pattern": "外勤出勤報表_{year}_{month}.pdf"
            }
        }
        
        with tempfile.TemporaryDirectory() as tmpdir:
            config_path = Path(tmpdir) / "config.json"
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(old_config_data, f)
            
            manager = ConfigManager(config_path)
            config = manager.load()
            
            # New fields should have defaults
            assert config.ui_prefs.color_logic.normal_in_color == "green"
            assert config.ui_prefs.color_logic.missing_punch_text == "*"
            assert config.output_settings.separate_pdf is True
            assert config.output_settings.pdf_output_dir == ""
            
            # Legacy bool fields should be loaded
            assert config.ui_prefs.color_logic.green_normal_in is True
            assert config.ui_prefs.color_logic.green_normal_out is False


class TestColorOptions:
    """Tests to verify valid color option values."""
    
    VALID_COLORS = ["red", "orange", "yellow", "green", "blue", "purple", "pink", "black", "none"]
    
    def test_default_colors_are_valid(self):
        """Test that default ColorLogic values are in valid set."""
        cl = ColorLogic()
        
        assert cl.normal_in_color in self.VALID_COLORS
        assert cl.normal_out_color in self.VALID_COLORS
        assert cl.abnormal_in_color in self.VALID_COLORS
        assert cl.abnormal_out_color in self.VALID_COLORS
        assert cl.missing_punch_color in self.VALID_COLORS
        assert cl.absent_color in self.VALID_COLORS


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
