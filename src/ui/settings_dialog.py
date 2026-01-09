"""
Settings Dialog Module

PyQt6 dialog for application settings including:
- Output style settings (punch card colors, missing/absent markers)
- PDF generation settings (separate/combined, output path)
"""

from pathlib import Path
from typing import Dict, Tuple

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QComboBox, QLineEdit, QCheckBox,
    QPushButton, QFileDialog, QDialogButtonBox, QStyledItemDelegate,
    QStyle, QWidget, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPainter, QColor, QPalette

from config.config_manager import AppConfig
from ui.styles import ThemeManager


# Color options for combo boxes
COLOR_OPTIONS: Dict[str, Tuple[str, str]] = {
    # display_name: (value, hex_color for preview)
    "紅色": ("red", "#FF6B6B"),
    "橙色": ("orange", "#FFA500"),
    "黃色": ("yellow", "#FFD700"),
    "綠色": ("green", "#90EE90"),
    "藍色": ("blue", "#6B8CFF"),
    "紫色": ("purple", "#DDA0DD"),
    "粉色": ("pink", "#FFB6C1"),
    "黑色": ("black", "#333333"),
    "無": ("none", "transparent"),
}

# Reverse lookup: value -> display_name
VALUE_TO_DISPLAY = {v[0]: k for k, v in COLOR_OPTIONS.items()}


class ColorDelegate(QStyledItemDelegate):
    """Custom delegate to paint color swatches in QComboBox."""
    
    def paint(self, painter: QPainter, option, index):
        painter.save()
        
        # 繪製背景 (Selected state)
        # 繪製背景 (Selected state)
        if option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())
        
        # 取得顏色代碼
        hex_color = index.data(Qt.ItemDataRole.UserRole)
        text = index.data(Qt.ItemDataRole.DisplayRole)
        
        # 定義顏色方塊區域
        rect = option.rect
        color_rect_size = 20  # 加大色塊
        color_rect_x = rect.left() + 10
        color_rect_y = rect.top() + (rect.height() - color_rect_size) // 2
        
        # 繪製顏色方塊
        if hex_color and hex_color != "transparent":
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(hex_color))
            painter.drawRoundedRect(color_rect_x, color_rect_y, color_rect_size, color_rect_size, 4, 4)
        
        # 繪製文字
        text_x = color_rect_x + color_rect_size + 12
        text_rect = rect.adjusted(text_x - rect.left(), 0, 0, 0)
        
        # 設定較大的字體 (Remove hardcoded size, rely on theme or slight adjustment)
        font = painter.font()
        # font.setPointSize(11) # Maintain standard font size from theme
        painter.setFont(font)
        
        # Use palette text color
        text_color = option.palette.highlightedText().color() if (option.state & QStyle.StateFlag.State_Selected) else option.palette.text().color()
        painter.setPen(text_color)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, text)
        
        painter.restore()


class SettingsDialog(QDialog):
    """Settings dialog with output style and PDF options."""
    
    def __init__(self, config: AppConfig, parent=None):
        super().__init__(parent)
        self.config = config
        self._init_ui()
        self._load_config_to_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """Initialize the dialog UI."""
        self.setWindowTitle("設定")
        self.resize(800, 600)  # 再加大視窗尺寸
        self.setMinimumWidth(700)
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Output Style Settings Group
        style_group = self._create_style_group()
        layout.addWidget(style_group)
        
        # Output Data Settings Group (排序設定)
        output_data_group = self._create_output_data_group()
        layout.addWidget(output_data_group)
        
        # PDF Settings Group
        pdf_group = self._create_pdf_group()
        layout.addWidget(pdf_group)
        
        # Dialog buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self._on_accept)
        button_box.rejected.connect(self.reject)
        
        # Button styling layout
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(button_box)
        layout.addLayout(btn_layout)
        
        # Apply styling
        self._apply_styles()


    def _create_style_group(self) -> QGroupBox:
        """Create the output style settings group."""
        group = QGroupBox("輸出樣式設定")
        layout = QGridLayout(group)
        layout.setContentsMargins(15, 25, 15, 15)  # 調整邊距，top加大保留給標題
        layout.setSpacing(15)  # 增加格線間距
        
        row = 0
        
        # Normal check-in color
        layout.addWidget(QLabel("正常上班打卡標記"), row, 0)
        self.cmb_normal_in = self._create_color_combo()
        layout.addWidget(self.cmb_normal_in, row, 1)
        row += 1
        
        # Normal check-out color
        layout.addWidget(QLabel("正常下班打卡標記"), row, 0)
        self.cmb_normal_out = self._create_color_combo()
        layout.addWidget(self.cmb_normal_out, row, 1)
        row += 1
        
        # Abnormal check-in color
        layout.addWidget(QLabel("異常上班打卡標記"), row, 0)
        self.cmb_abnormal_in = self._create_color_combo()
        layout.addWidget(self.cmb_abnormal_in, row, 1)
        row += 1
        
        # Abnormal check-out color
        layout.addWidget(QLabel("異常下班打卡標記"), row, 0)
        self.cmb_abnormal_out = self._create_color_combo()
        layout.addWidget(self.cmb_abnormal_out, row, 1)
        row += 1
        
        # Missing punch color and text
        layout.addWidget(QLabel("缺少打卡紀錄標記"), row, 0)
        missing_layout = QHBoxLayout()
        missing_layout.setSpacing(10)
        self.cmb_missing_punch = self._create_color_combo()
        missing_layout.addWidget(self.cmb_missing_punch)
        missing_layout.addWidget(QLabel("文字:"))
        self.txt_missing_punch = QLineEdit()
        self.txt_missing_punch.setMaximumWidth(80)
        self.txt_missing_punch.setMaxLength(5)
        missing_layout.addWidget(self.txt_missing_punch)
        missing_layout.addStretch()
        layout.addLayout(missing_layout, row, 1)
        row += 1
        
        # Absent color and text
        layout.addWidget(QLabel("曠職標記"), row, 0)
        absent_layout = QHBoxLayout()
        absent_layout.setSpacing(10)
        self.cmb_absent = self._create_color_combo()
        absent_layout.addWidget(self.cmb_absent)
        absent_layout.addWidget(QLabel("文字:"))
        self.txt_absent = QLineEdit()
        self.txt_absent.setMaximumWidth(80)
        absent_layout.addWidget(self.txt_absent)
        absent_layout.addStretch()
        layout.addLayout(absent_layout, row, 1)
        
        return group
    
    def _create_output_data_group(self) -> QGroupBox:
        """Create the output data settings group (sorting options)."""
        group = QGroupBox("輸出資料設定")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(15, 25, 15, 15)
        layout.setSpacing(10)
        
        # Sorting label
        layout.addWidget(QLabel("排序依據:"))
        
        # Sort options list
        self.list_sort_by = QListWidget()
        self.list_sort_by.setMaximumHeight(70)
        self.list_sort_by.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        
        # Add sort options
        item1 = QListWidgetItem("1. 出席率")
        item1.setData(Qt.ItemDataRole.UserRole, "attendance_rate")
        self.list_sort_by.addItem(item1)
        
        item2 = QListWidgetItem("2. 姓氏筆畫")
        item2.setData(Qt.ItemDataRole.UserRole, "name_strokes")
        self.list_sort_by.addItem(item2)
        
        layout.addWidget(self.list_sort_by)
        
        return group
    
    def _create_pdf_group(self) -> QGroupBox:
        """Create the PDF settings group."""
        group = QGroupBox("PDF 設定")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(15, 25, 15, 15)
        layout.setSpacing(15)
        
        # Separate PDF checkbox
        self.chk_separate_pdf = QCheckBox("分別生成內外勤報表")
        layout.addWidget(self.chk_separate_pdf)
        
        # PDF output path
        path_layout = QHBoxLayout()
        path_layout.setSpacing(10)
        path_layout.addWidget(QLabel("PDF 輸出路徑:"))
        self.txt_pdf_path = QLineEdit()
        self.txt_pdf_path.setPlaceholderText("留空則與 Excel 同目錄")
        path_layout.addWidget(self.txt_pdf_path, stretch=1)
        
        self.btn_browse_pdf = QPushButton("瀏覽")
        self.btn_browse_pdf.setMinimumWidth(80)
        path_layout.addWidget(self.btn_browse_pdf)
        
        layout.addLayout(path_layout)
        
        return group
    
    def _create_color_combo(self) -> QComboBox:
        """Create a color selection combo box."""
        combo = QComboBox()
        combo.setMinimumWidth(120)
        combo.setItemDelegate(ColorDelegate(combo))  # 使用自定義 Delegate
        
        for display_name, (value, hex_color) in COLOR_OPTIONS.items():
            combo.addItem(display_name, value)
            # 設定 Item Data UserRole 為顏色代碼，供 Delegate 繪製
            idx = combo.count() - 1
            combo.setItemData(idx, hex_color, Qt.ItemDataRole.UserRole)
        
        return combo
    
    # 移除 _update_combo_style，改用 Delegate
    def _update_combo_style(self, combo: QComboBox):
        pass

    def _set_combo_value(self, combo: QComboBox, value: str):
        """Set combo box to the given color value."""
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)
        else:
            combo.setCurrentIndex(0)
    
    def _get_combo_value(self, combo: QComboBox) -> str:
        """Get the color value from combo box."""
        return combo.currentData() or "none"
    
    def _load_config_to_ui(self):
        """Load configuration values into UI controls."""
        cl = self.config.ui_prefs.color_logic
        os = self.config.output_settings
        
        # Color settings
        self._set_combo_value(self.cmb_normal_in, cl.normal_in_color)
        self._set_combo_value(self.cmb_normal_out, cl.normal_out_color)
        self._set_combo_value(self.cmb_abnormal_in, cl.abnormal_in_color)
        self._set_combo_value(self.cmb_abnormal_out, cl.abnormal_out_color)
        self._set_combo_value(self.cmb_missing_punch, cl.missing_punch_color)
        self._set_combo_value(self.cmb_absent, cl.absent_color)
        
        # Text settings
        self.txt_missing_punch.setText(cl.missing_punch_text)
        self.txt_absent.setText(cl.absent_text)
        
        # PDF settings
        self.chk_separate_pdf.setChecked(os.separate_pdf)
        self.txt_pdf_path.setText(os.pdf_output_dir)
        
        # Sort settings - select the matching item
        for i in range(self.list_sort_by.count()):
            item = self.list_sort_by.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == os.sort_by:
                item.setSelected(True)
                break
    
    def _save_ui_to_config(self):
        """Save UI values to configuration."""
        cl = self.config.ui_prefs.color_logic
        os = self.config.output_settings
        
        # Color settings
        cl.normal_in_color = self._get_combo_value(self.cmb_normal_in)
        cl.normal_out_color = self._get_combo_value(self.cmb_normal_out)
        cl.abnormal_in_color = self._get_combo_value(self.cmb_abnormal_in)
        cl.abnormal_out_color = self._get_combo_value(self.cmb_abnormal_out)
        cl.missing_punch_color = self._get_combo_value(self.cmb_missing_punch)
        cl.absent_color = self._get_combo_value(self.cmb_absent)
        
        # Text settings
        cl.missing_punch_text = self.txt_missing_punch.text() or "*"
        cl.absent_text = self.txt_absent.text() or "-"
        
        # PDF settings
        os.separate_pdf = self.chk_separate_pdf.isChecked()
        os.pdf_output_dir = self.txt_pdf_path.text()
        
        # Sort settings
        selected_items = self.list_sort_by.selectedItems()
        if selected_items:
            os.sort_by = selected_items[0].data(Qt.ItemDataRole.UserRole)
        else:
            os.sort_by = "attendance_rate"  # Default
    
    def _connect_signals(self):
        """Connect UI signals."""
        self.btn_browse_pdf.clicked.connect(self._on_browse_pdf)
    
    def _on_browse_pdf(self):
        """Handle browse PDF output path."""
        current_path = self.txt_pdf_path.text()
        start_dir = current_path if current_path else str(Path.cwd())
        
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "選擇 PDF 輸出目錄",
            start_dir
        )
        if dir_path:
            self.txt_pdf_path.setText(dir_path)
    
    def _on_accept(self):
        """Handle OK button click."""
        self._save_ui_to_config()
        self.accept()
    
    def _apply_styles(self):
        """Apply dialog styling from global theme."""
        theme_name = self.config.ui_prefs.theme_name
        theme = ThemeManager.get_theme(theme_name)
        self.setStyleSheet(theme.stylesheet)


