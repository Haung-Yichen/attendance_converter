"""
Style Management Module

Handles application theming and style definitions.
Follows SOLID principles:
- SRP: Only responsible for serving style definitions.
- OCP: New themes can be added without modifying existing logic.
"""

from abc import ABC, abstractmethod
from typing import Dict, Type

class Theme(ABC):
    """Abstract base class for Themes."""
    
    @property
    @abstractmethod
    def name(self) -> str:
        pass
        
    @property
    @abstractmethod
    def stylesheet(self) -> str:
        """Returns the fully compiled stylesheet string."""
        pass


class DarkTheme(Theme):
    """The default dark theme for the application."""
    
    @property
    def name(self) -> str:
        return "Dark Mode"
        
    @property
    def stylesheet(self) -> str:
        return """
            QMainWindow, QDialog {
                background-color: #1e1e1e;
            }
            QWidget {
                font-family: 'Segoe UI', system-ui, sans-serif;
                font-size: 13px;
                color: #e0e0e0;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #3e3e42;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 16px;
                background-color: #252526;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 5px;
                color: #4ec9b0;
            }
            QListWidget {
                background-color: #1e1e1e;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 4px;
                outline: none;
            }
            QListWidget::item:selected {
                background-color: #37373d;
                border-radius: 2px;
            }
            QListWidget::item {
                 color: #e0e0e0;
            }
            QLabel {
                color: #e0e0e0;
            }
            QCheckBox {
                spacing: 8px;
                color: #e0e0e0;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 2px solid #6e6e6e;
                border-radius: 3px;
                background: #1e1e1e;
            }
            QCheckBox::indicator:checked {
                background-color: #4ec9b0;
                border-color: #4ec9b0;
            }
            QLineEdit, QTimeEdit, QSpinBox {
                background-color: #3c3c3c;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 4px 8px;
                padding-right: 20px; /* Make room for buttons */
                color: #f0f0f0;
                selection-background-color: #264f78;
            }
            QLineEdit:focus, QTimeEdit:focus, QSpinBox:focus {
                border: 1px solid #0078d4;
            }
            QSpinBox::up-button, QTimeEdit::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 20px;
                height: 14px;
                border-width: 0px;
                background: transparent;
                margin-top: 1px;
                margin-right: 1px;
            }
            QSpinBox::up-button:hover, QTimeEdit::up-button:hover {
                background-color: #505050;
                border-radius: 2px;
            }
            QSpinBox::down-button, QTimeEdit::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 20px;
                height: 14px;
                border-width: 0px;
                background: transparent;
                margin-bottom: 1px;
                margin-right: 1px;
            }
            QSpinBox::down-button:hover, QTimeEdit::down-button:hover {
                background-color: #505050;
                border-radius: 2px;
            }
            QSpinBox::up-arrow, QTimeEdit::up-arrow {
                width: 0; 
                height: 0; 
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-bottom: 5px solid #e0e0e0;
            }
            QSpinBox::down-arrow, QTimeEdit::down-arrow {
                width: 0; 
                height: 0; 
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #e0e0e0;
            }
            QComboBox {
                background-color: #3c3c3c;
                border: 1px solid #3e3e42;
                border-radius: 4px;
                padding: 4px 8px;
                color: #f0f0f0;
                selection-background-color: #264f78;
            }
            QComboBox:hover {
                border: 1px solid #0078d4;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #e0e0e0;
            }
            QComboBox QAbstractItemView {
                background-color: #252526;
                color: #e0e0e0;
                selection-background-color: #37373d;
                border: 1px solid #3e3e42;
                outline: none;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QMenuBar {
                background-color: #252526;
                color: #cccccc;
                border-bottom: 1px solid #3e3e42;
            }
            QMenuBar::item:selected {
                background-color: #37373d;
            }
            QMenu {
                background-color: #252526;
                border: 1px solid #3e3e42;
                color: #e0e0e0;
            }
            QMenu::item:selected {
                background-color: #37373d;
            }
        """


class ClassicWhiteTheme(Theme):
    """A minimal, Windows-native like light theme."""
    
    @property
    def name(self) -> str:
        return "Classic White"
        
    @property
    def stylesheet(self) -> str:
        return """
            QMainWindow, QDialog {
                background-color: #f0f0f0;
            }
            QWidget {
                font-family: 'Segoe UI', system-ui, sans-serif;
                font-size: 13px;
                color: #000000;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #b0b0b0;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 16px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 5px;
                color: #003399;
                background-color: #ffffff;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                padding: 4px;
                outline: none;
            }
            QListWidget::item:selected {
                background-color: #cce8ff;
                color: #000000;
                border-radius: 2px;
            }
            QListWidget::item {
                 color: #000000;
            }
            QLabel {
                color: #000000;
            }
            QCheckBox {
                spacing: 8px;
                color: #000000;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border: 1px solid #808080;
                border-radius: 3px;
                background: #ffffff;
            }
            QCheckBox::indicator:checked {
                background-color: #0078d4;
                border-color: #0078d4;
            }
            QLineEdit, QTimeEdit, QSpinBox {
                background-color: #ffffff;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                padding: 4px 8px;
                padding-right: 20px; /* Make room for buttons */
                color: #000000;
                selection-background-color: #cce8ff;
                selection-color: black;
            }
            QLineEdit:focus, QTimeEdit:focus, QSpinBox:focus {
                border: 1px solid #0078d4;
            }
            QSpinBox::up-button, QTimeEdit::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                width: 20px;
                height: 14px;
                border-width: 0px;
                background: transparent;
                margin-top: 1px;
                margin-right: 1px;
            }
            QSpinBox::up-button:hover, QTimeEdit::up-button:hover {
                background-color: #e5f1fb;
                border-radius: 2px;
            }
            QSpinBox::down-button, QTimeEdit::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                width: 20px;
                height: 14px;
                border-width: 0px;
                background: transparent;
                margin-bottom: 1px;
                margin-right: 1px;
            }
            QSpinBox::down-button:hover, QTimeEdit::down-button:hover {
                background-color: #e5f1fb;
                border-radius: 2px;
            }
            QSpinBox::up-arrow, QTimeEdit::up-arrow {
                width: 0; 
                height: 0; 
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-bottom: 5px solid #444444;
            }
            QSpinBox::down-arrow, QTimeEdit::down-arrow {
                width: 0; 
                height: 0; 
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #444444;
            }
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                padding: 4px 8px;
                color: #000000;
                selection-background-color: #cce8ff;
            }
            QComboBox:hover {
                border: 1px solid #0078d4;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #444444;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #000000;
                selection-background-color: #cce8ff;
                selection-color: #000000;
                border: 1px solid #c0c0c0;
                outline: none;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: 1px solid #005a9e;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QMenuBar {
                background-color: #f9f9f9;
                color: #000000;
                border-bottom: 1px solid #d0d0d0;
            }
            QMenuBar::item:selected {
                background-color: #e5f3ff;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #cccccc;
                color: #000000;
            }
            QMenu::item:selected {
                background-color: #e5f3ff;
                color: #000000;
            }
        """


class ThemeManager:
    """
    Factory and manager for application themes.
    Singleton pattern usage is implied by typical PyQt usage, but this class is stateless.
    """
    
    _themes: Dict[str, Type[Theme]] = {
        "Dark Mode": DarkTheme,
        "Classic White": ClassicWhiteTheme
    }
    
    @classmethod
    def get_theme(cls, theme_name: str) -> Theme:
        """Factory method to get a theme instance by name."""
        theme_cls = cls._themes.get(theme_name)
        if not theme_cls:
            # Fallback to default if theme name not found
            return DarkTheme()
        return theme_cls()
    
    @classmethod
    def get_available_themes(cls) -> list[str]:
        """Returns a list of available theme names."""
        return list(cls._themes.keys())
