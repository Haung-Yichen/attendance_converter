"""
Annual Report Worker Module

Provides a ``QThread``-based background worker for annual report
generation. Emits typed signals for progress, completion, and warnings
so the UI thread stays responsive.
"""

from __future__ import annotations

from pathlib import Path
from typing import List

from PyQt6.QtCore import QThread, pyqtSignal

from application.annual_report_service import (
    AnnualReportParams,
    AnnualReportResult,
    AnnualReportService,
)


class AnnualReportWorker(QThread):
    """
    Background worker that runs ``AnnualReportService.generate()``
    off the main / UI thread.

    Signals:
        progress(int, int, str):
            ``(current_step, total_steps, description)`` – emitted once
            per month processed.
        finished(AnnualReportResult):
            Emitted when the entire operation is done (success **or**
            failure).  The caller inspects ``result.success`` to decide
            how to present the outcome.
        warning(str):
            Emitted for non-fatal issues (e.g. a month file is missing
            or unparseable).  The UI may collect these and display them
            after the run completes.
    """

    # ---- Signals ---- #
    progress = pyqtSignal(int, int, str)        # current, total, message
    finished_result = pyqtSignal(object)         # AnnualReportResult
    warning = pyqtSignal(str)                    # warning text

    def __init__(
        self,
        params: AnnualReportParams,
        service: AnnualReportService | None = None,
        parent=None,
    ) -> None:
        """
        Args:
            params: Fully-populated ``AnnualReportParams``.
            service: Optional pre-configured service instance
                (useful for testing with injected fakes).
            parent: QObject parent.
        """
        super().__init__(parent)
        self._params = params
        self._service = service or AnnualReportService()

    # ------------------------------------------------------------------ #
    # QThread override
    # ------------------------------------------------------------------ #

    def run(self) -> None:  # noqa: D401 – Qt naming convention
        """Execute the annual report generation on the worker thread."""
        try:
            result = self._service.generate(
                self._params,
                on_progress=self._on_progress,
            )

            # Re-emit collected warnings individually so the UI can
            # display them in a list widget or log pane.
            for msg in result.warnings:
                self.warning.emit(msg)

            self.finished_result.emit(result)

        except Exception as exc:
            # Catch-all so the thread never crashes silently.
            error_result = AnnualReportResult(
                success=False,
                year=self._params.year,
                output_path=self._params.output_path,
                error_message=f"非預期錯誤: {exc}",
            )
            self.finished_result.emit(error_result)

    # ------------------------------------------------------------------ #
    # Private callback bridging service → Qt signal
    # ------------------------------------------------------------------ #

    def _on_progress(self, current: int, total: int, message: str) -> None:
        """Bridge the service's progress callback to a Qt signal."""
        self.progress.emit(current, total, message)
