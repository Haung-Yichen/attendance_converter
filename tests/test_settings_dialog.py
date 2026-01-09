"""
Unit tests for SettingsDialog UI component.
"""

import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from config.config_manager import AppConfig, ColorLogic, OutputSettings, UIPrefs


class TestSettingsDialogColorOptions:
    """Tests for SettingsDialog color options mapping."""
    
    # Import the COLOR_OPTIONS from settings_dialog
    from ui.settings_dialog import COLOR_OPTIONS, VALUE_TO_DISPLAY
    
    def test_color_options_has_required_colors(self):
        """Test that all required colors are available."""
        required_colors = ["紅色", "橙色", "黃色", "綠色", "藍色", "紫色", "粉色", "無"]
        
        for color in required_colors:
            assert color in self.COLOR_OPTIONS, f"Missing color: {color}"
    
    def test_color_options_values(self):
        """Test that color values are correct."""
        expected = {
            "紅色": "red",
            "橙色": "orange",
            "黃色": "yellow",
            "綠色": "green",
            "藍色": "blue",
            "紫色": "purple",
            "粉色": "pink",
            "無": "none"
        }
        
        for display_name, expected_value in expected.items():
            actual_value = self.COLOR_OPTIONS[display_name][0]
            assert actual_value == expected_value
    
    def test_value_to_display_reverse_lookup(self):
        """Test that reverse lookup works correctly."""
        assert self.VALUE_TO_DISPLAY["green"] == "綠色"
        assert self.VALUE_TO_DISPLAY["red"] == "紅色"
        assert self.VALUE_TO_DISPLAY["none"] == "無"
    
    def test_black_color_included(self):
        """Test that black color is included for missing punch."""
        assert "黑色" in self.COLOR_OPTIONS
        assert self.COLOR_OPTIONS["黑色"][0] == "black"


class TestSettingsDialogConfigMapping:
    """Tests for SettingsDialog configuration mapping."""
    
    def test_config_to_dialog_mapping(self):
        """Test that config values map correctly to dialog fields."""
        config = AppConfig()
        
        # Verify default values
        cl = config.ui_prefs.color_logic
        assert cl.normal_in_color == "green"
        assert cl.missing_punch_text == "*"
        assert cl.absent_text == "-"
        
        os = config.output_settings
        assert os.separate_pdf is True
        assert os.pdf_output_dir == ""
    
    def test_dialog_to_config_mapping(self):
        """Test that dialog values can update config correctly."""
        config = AppConfig()
        
        # Simulate dialog changing values
        config.ui_prefs.color_logic.normal_in_color = "blue"
        config.ui_prefs.color_logic.missing_punch_color = "pink"
        config.ui_prefs.color_logic.missing_punch_text = "?"
        config.output_settings.separate_pdf = False
        config.output_settings.pdf_output_dir = "C:/custom/path"
        
        # Verify changes
        assert config.ui_prefs.color_logic.normal_in_color == "blue"
        assert config.ui_prefs.color_logic.missing_punch_color == "pink"
        assert config.ui_prefs.color_logic.missing_punch_text == "?"
        assert config.output_settings.separate_pdf is False
        assert config.output_settings.pdf_output_dir == "C:/custom/path"


class TestColorLogicValidation:
    """Tests for ColorLogic value validation."""
    
    VALID_COLORS = ["red", "orange", "yellow", "green", "blue", "purple", "pink", "black", "none"]
    
    def test_all_color_fields_accept_valid_values(self):
        """Test that all color fields accept valid color values."""
        for color in self.VALID_COLORS:
            cl = ColorLogic(
                normal_in_color=color,
                normal_out_color=color,
                abnormal_in_color=color,
                abnormal_out_color=color,
                missing_punch_color=color,
                absent_color=color
            )
            
            assert cl.normal_in_color == color
            assert cl.normal_out_color == color



class TestSettingsDialogUI:
    """Tests for SettingsDialog UI components setup."""
    
    @pytest.fixture
    def mock_app(self):
        """Create a mock QApplication instance."""
        from PyQt6.QtWidgets import QApplication
        app = QApplication.instance()
        if not app:
            app = QApplication([])
        return app
    
    def test_combobox_delegates(self, mock_app):
        """Test that combo boxes use the ColorDelegate."""
        from ui.settings_dialog import SettingsDialog, ColorDelegate
        
        config = AppConfig()
        dialog = SettingsDialog(config)
        
        # Check if delegates are set correctly
        combos = [
            dialog.cmb_normal_in,
            dialog.cmb_normal_out,
            dialog.cmb_abnormal_in,
            dialog.cmb_abnormal_out,
            dialog.cmb_missing_punch,
            dialog.cmb_absent
        ]
        
        for combo in combos:
            delegate = combo.itemDelegate()
            assert isinstance(delegate, ColorDelegate), f"Combo {combo} should use ColorDelegate"
            
        dialog.close()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
