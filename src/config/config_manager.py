"""
Configuration Manager Module

Handles loading, saving, and managing application configuration.
Provides bi-directional mapping between UI state and JSON persistence.
"""

import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional
import os


@dataclass
class ColorLogic:
    """Color logic settings for attendance marking.
    
    Color values: 'red', 'orange', 'yellow', 'green', 'blue', 'purple', 'pink', 'black', 'none'
    """
    # 打卡顏色設定 (改為顏色字串)
    normal_in_color: str = "green"       # 正常上班打卡顏色
    normal_out_color: str = "green"      # 正常下班打卡顏色
    abnormal_in_color: str = "red"       # 異常上班打卡顏色
    abnormal_out_color: str = "red"      # 異常下班打卡顏色
    
    # 缺卡與曠職設定
    missing_punch_color: str = "black"   # 當日缺少打卡紀錄顏色
    missing_punch_text: str = "*"        # 缺少打卡紀錄標記文字
    absent_color: str = "none"           # 曠職標記顏色
    absent_text: str = "-"               # 曠職標記文字
    
    # 保留舊欄位以維持向後相容 (讀取舊 config 時自動轉換)
    green_normal_in: bool = True
    green_normal_out: bool = True
    red_abnormal_in: bool = True
    red_abnormal_out: bool = True


@dataclass
class TimeRule:
    """Time rule for check-in/check-out boundaries."""
    in_start: str = "09:00"
    in_end: str = "09:30"
    out_start: str = "18:00"
    out_end: str = "18:30"


@dataclass
class TimeRules:
    """Combined time rules for internal and external staff."""
    internal: TimeRule = field(default_factory=TimeRule)
    external: TimeRule = field(default_factory=lambda: TimeRule(
        in_start="09:30",
        in_end="10:00",
        out_start="10:30",
        out_end="12:00"
    ))


@dataclass
class Paths:
    """File paths configuration."""
    staff_csv: str = ""
    leave_list: str = ""
    last_source_file: str = ""
    custom_font_path: str = ""  # Custom font path for PDF generation


@dataclass
class Holidays:
    """Holiday settings."""
    use_auto_fetch: bool = True
    custom_dates: list = field(default_factory=list)


@dataclass
class UIPrefs:
    """UI preferences including color logic and rate threshold."""
    color_logic: ColorLogic = field(default_factory=ColorLogic)
    rate_threshold: int = 80
    theme_name: str = "Dark Mode"


@dataclass
class OutputSettings:
    """Output settings for generated report."""
    output_dir: str = ""  # Default empty = project root
    filename_pattern: str = "Attendance_{year}_{month}.xlsx"
    generate_pdf: bool = True  # 預設開啟
    
    # 輸出排序設定
    sort_by: str = "attendance_rate"  # 排序依據: "attendance_rate" 或 "name_strokes"
    
    # PDF 輸出設定
    separate_pdf: bool = True  # True=分開生成內外勤PDF, False=合併為一份
    pdf_output_dir: str = ""   # PDF 輸出路徑，空字串=跟 xlsx 同目錄
    pdf_filename_pattern: str = "出勤報表_{year}_{month}.pdf"  # 合併時使用的檔名
    
    # PDF filename patterns (支援 {year}, {month} 佔位符)
    internal_pdf_pattern: str = "內勤出勤報表_{year}_{month}.pdf"
    external_pdf_pattern: str = "外勤出勤報表_{year}_{month}.pdf"


@dataclass
class AppConfig:
    """Main application configuration container."""
    paths: Paths = field(default_factory=Paths)
    holidays: Holidays = field(default_factory=Holidays)
    time_rules: TimeRules = field(default_factory=TimeRules)
    ui_prefs: UIPrefs = field(default_factory=UIPrefs)
    output_settings: OutputSettings = field(default_factory=OutputSettings)


class ConfigManager:
    """
    Manages application configuration with JSON persistence.
    
    Responsibilities:
    - Load configuration from JSON file
    - Save configuration to JSON file
    - Provide default configuration
    - Convert between dataclass and dict representations
    """
    
    DEFAULT_CONFIG_PATH = Path(__file__).parent.parent / "config.json"
    
    def __init__(self, config_path: Optional[Path] = None):
        self.config_path = config_path or self.DEFAULT_CONFIG_PATH
        self._config: AppConfig = AppConfig()
    
    @property
    def config(self) -> AppConfig:
        """Get current configuration."""
        return self._config
    
    def load(self) -> AppConfig:
        """Load configuration from JSON file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self._config = self._dict_to_config(data)
            except (json.JSONDecodeError, KeyError) as e:
                print(f"Warning: Failed to load config, using defaults. Error: {e}")
                self._config = AppConfig()
        else:
            self._config = AppConfig()
        return self._config
    
    def save(self) -> None:
        """Save current configuration to JSON file."""
        data = self._config_to_dict(self._config)
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def update(self, **kwargs) -> None:
        """Update specific configuration values."""
        for key, value in kwargs.items():
            if hasattr(self._config, key):
                setattr(self._config, key, value)
        self.save()
    
    def _config_to_dict(self, config: AppConfig) -> dict:
        """Convert AppConfig dataclass to dictionary."""
        return {
            "paths": {
                "staff_csv": config.paths.staff_csv,
                "leave_list": config.paths.leave_list,
                "last_source_file": config.paths.last_source_file,
                "custom_font_path": config.paths.custom_font_path
            },
            "holidays": {
                "use_auto_fetch": config.holidays.use_auto_fetch,
                "custom_dates": config.holidays.custom_dates
            },
            "time_rules": {
                "internal": {
                    "in_start": config.time_rules.internal.in_start,
                    "in_end": config.time_rules.internal.in_end,
                    "out_start": config.time_rules.internal.out_start,
                    "out_end": config.time_rules.internal.out_end
                },
                "external": {
                    "in_start": config.time_rules.external.in_start,
                    "in_end": config.time_rules.external.in_end,
                    "out_start": config.time_rules.external.out_start,
                    "out_end": config.time_rules.external.out_end
                }
            },
            "ui_prefs": {
                "color_logic": {
                    "normal_in_color": config.ui_prefs.color_logic.normal_in_color,
                    "normal_out_color": config.ui_prefs.color_logic.normal_out_color,
                    "abnormal_in_color": config.ui_prefs.color_logic.abnormal_in_color,
                    "abnormal_out_color": config.ui_prefs.color_logic.abnormal_out_color,
                    "missing_punch_color": config.ui_prefs.color_logic.missing_punch_color,
                    "missing_punch_text": config.ui_prefs.color_logic.missing_punch_text,
                    "absent_color": config.ui_prefs.color_logic.absent_color,
                    "absent_text": config.ui_prefs.color_logic.absent_text,
                    # 保留舊欄位以維持向後相容
                    "green_normal_in": config.ui_prefs.color_logic.green_normal_in,
                    "green_normal_out": config.ui_prefs.color_logic.green_normal_out,
                    "red_abnormal_in": config.ui_prefs.color_logic.red_abnormal_in,
                    "red_abnormal_out": config.ui_prefs.color_logic.red_abnormal_out
                },
                "rate_threshold": config.ui_prefs.rate_threshold,
                "theme_name": config.ui_prefs.theme_name
            },
            "output_settings": {
                "output_dir": config.output_settings.output_dir,
                "filename_pattern": config.output_settings.filename_pattern,
                "generate_pdf": config.output_settings.generate_pdf,
                "sort_by": config.output_settings.sort_by,
                "separate_pdf": config.output_settings.separate_pdf,
                "pdf_output_dir": config.output_settings.pdf_output_dir,
                "pdf_filename_pattern": config.output_settings.pdf_filename_pattern,
                "internal_pdf_pattern": config.output_settings.internal_pdf_pattern,
                "external_pdf_pattern": config.output_settings.external_pdf_pattern
            }
        }
    
    def _dict_to_config(self, data: dict) -> AppConfig:
        """Convert dictionary to AppConfig dataclass."""
        paths_data = data.get("paths", {})
        holidays_data = data.get("holidays", {})
        time_rules_data = data.get("time_rules", {})
        ui_prefs_data = data.get("ui_prefs", {})
        
        # Build Paths
        paths = Paths(
            staff_csv=paths_data.get("staff_csv", ""),
            leave_list=paths_data.get("leave_list", ""),
            last_source_file=paths_data.get("last_source_file", ""),
            custom_font_path=paths_data.get("custom_font_path", "")
        )
        
        # Build Holidays
        holidays = Holidays(
            use_auto_fetch=holidays_data.get("use_auto_fetch", True),
            custom_dates=holidays_data.get("custom_dates", [])
        )
        
        # Build TimeRules
        internal_data = time_rules_data.get("internal", {})
        external_data = time_rules_data.get("external", {})
        
        time_rules = TimeRules(
            internal=TimeRule(
                in_start=internal_data.get("in_start", "09:00"),
                in_end=internal_data.get("in_end", "09:30"),
                out_start=internal_data.get("out_start", "18:00"),
                out_end=internal_data.get("out_end", "18:30")
            ),
            external=TimeRule(
                in_start=external_data.get("in_start", "09:30"),
                in_end=external_data.get("in_end", "10:00"),
                out_start=external_data.get("out_start", "10:30"),
                out_end=external_data.get("out_end", "12:00")
            )
        )
        
        # Build UIPrefs
        color_logic_data = ui_prefs_data.get("color_logic", {})
        ui_prefs = UIPrefs(
            color_logic=ColorLogic(
                normal_in_color=color_logic_data.get("normal_in_color", "green"),
                normal_out_color=color_logic_data.get("normal_out_color", "green"),
                abnormal_in_color=color_logic_data.get("abnormal_in_color", "red"),
                abnormal_out_color=color_logic_data.get("abnormal_out_color", "red"),
                missing_punch_color=color_logic_data.get("missing_punch_color", "black"),
                missing_punch_text=color_logic_data.get("missing_punch_text", "*"),
                absent_color=color_logic_data.get("absent_color", "none"),
                absent_text=color_logic_data.get("absent_text", "-"),
                # 保留舊欄位以維持向後相容
                green_normal_in=color_logic_data.get("green_normal_in", True),
                green_normal_out=color_logic_data.get("green_normal_out", True),
                red_abnormal_in=color_logic_data.get("red_abnormal_in", True),
                red_abnormal_out=color_logic_data.get("red_abnormal_out", True)
            ),
            rate_threshold=ui_prefs_data.get("rate_threshold", 80),
            theme_name=ui_prefs_data.get("theme_name", "Dark Mode")
        )
        
        # Build OutputSettings
        output_settings_data = data.get("output_settings", {})
        output_settings = OutputSettings(
            output_dir=output_settings_data.get("output_dir", ""),
            filename_pattern=output_settings_data.get("filename_pattern", "Attendance_{year}_{month}.xlsx"),
            generate_pdf=output_settings_data.get("generate_pdf", True),
            sort_by=output_settings_data.get("sort_by", "attendance_rate"),
            separate_pdf=output_settings_data.get("separate_pdf", True),
            pdf_output_dir=output_settings_data.get("pdf_output_dir", ""),
            pdf_filename_pattern=output_settings_data.get("pdf_filename_pattern", "出勤報表_{year}_{month}.pdf"),
            internal_pdf_pattern=output_settings_data.get("internal_pdf_pattern", "內勤出勤報表_{year}_{month}.pdf"),
            external_pdf_pattern=output_settings_data.get("external_pdf_pattern", "外勤出勤報表_{year}_{month}.pdf")
        )
        
        return AppConfig(
            paths=paths,
            holidays=holidays,
            time_rules=time_rules,
            ui_prefs=ui_prefs,
            output_settings=output_settings
        )
