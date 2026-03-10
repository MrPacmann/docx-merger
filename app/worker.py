"""Фоновый worker для выполнения объединения без блокировки GUI."""

from __future__ import annotations

from time import perf_counter
from threading import Event

from PySide6.QtCore import QObject, Signal, Slot

from .logger import get_logger
from .merger import DocxMerger, MergeCancelledError, MergeProgress, MergeRequest
from .utils import build_user_friendly_error


class MergeWorker(QObject):
    """Выполняет объединение документов в рабочем потоке и уведомляет UI сигналами."""

    progress_changed = Signal(int, int, int, str, str)
    status_changed = Signal(str)
    finished = Signal(str)
    failed = Signal(str, str)
    cancelled = Signal()

    def __init__(self, request: MergeRequest, merger: DocxMerger) -> None:
        super().__init__()
        self._request = request
        self._merger = merger
        self._cancel_event = Event()
        self._logger = get_logger(__name__)
        self._started_at: float | None = None

    @Slot()
    def run(self) -> None:
        """Запускает объединение документов и обрабатывает штатные сценарии завершения."""
        self._started_at = perf_counter()
        self.status_changed.emit("Подготовка фоновой задачи объединения...")
        try:
            output_path = self._merger.merge(
                request=self._request,
                cancel_event=self._cancel_event,
                progress_callback=self._on_progress,
            )
        except MergeCancelledError:
            self._logger.info("Операция объединения отменена.")
            self.status_changed.emit("Операция отменена.")
            self.cancelled.emit()
        except Exception as exc:
            self._logger.exception("Ошибка при объединении файлов.")
            title, message = build_user_friendly_error(exc)
            self.status_changed.emit(f"{title}: {message}")
            self.failed.emit(title, message)
        else:
            self.status_changed.emit("Объединение успешно завершено.")
            self.finished.emit(str(output_path))

    def cancel(self) -> None:
        """Устанавливает флаг отмены для фоновой операции."""
        self._cancel_event.set()
        self.status_changed.emit("Получен запрос на отмену операции.")

    def _on_progress(self, progress: MergeProgress) -> None:
        """Преобразует внутреннюю модель прогресса в Qt-сигнал."""
        percent = self._calculate_percent(progress.current, progress.total)
        eta_text = self._estimate_eta(progress.current, progress.total)
        self.progress_changed.emit(
            progress.current,
            progress.total,
            percent,
            eta_text,
            progress.message,
        )

    @staticmethod
    def _calculate_percent(current: int, total: int) -> int:
        """Преобразует шаговый прогресс в проценты."""
        if total <= 0:
            return 0
        return max(0, min(100, int(current / total * 100)))

    def _estimate_eta(self, current: int, total: int) -> str:
        """Оценивает приблизительное оставшееся время по уже выполненным шагам."""
        if self._started_at is None or total <= 1 or current <= 0 or current >= total:
            return "Осталось: меньше минуты"

        elapsed_seconds = max(0.0, perf_counter() - self._started_at)
        average_step_seconds = elapsed_seconds / current if current else 0.0
        remaining_steps = total - current
        remaining_seconds = max(0, int(round(average_step_seconds * remaining_steps)))
        return f"Осталось: {self._format_duration(remaining_seconds)}"

    @staticmethod
    def _format_duration(seconds: int) -> str:
        """Форматирует длительность в человекочитаемый вид."""
        if seconds < 60:
            return "меньше минуты"

        minutes, secs = divmod(seconds, 60)
        if minutes < 60:
            return f"{minutes} мин {secs} сек"

        hours, minutes = divmod(minutes, 60)
        return f"{hours} ч {minutes} мин"
