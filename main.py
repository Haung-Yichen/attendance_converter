"""
701Client Attendance Report Converter

A PyQt6 application for converting raw Excel attendance exports
into formatted Monthly Attendance Reports.
"""

import sys
from pathlib import Path

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from ui.main_window import run_app


def main():
    """Application entry point."""
    run_app()


if __name__ == "__main__":
    main()
