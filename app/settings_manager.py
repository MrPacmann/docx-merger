"""Чтение и сохранение пользовательских настроек приложения через QSettings."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QByteArray, QPoint, QSettings, QSize

from .app_info import APP_NAME, APP_ORGANIZATION
from .platform_utils import get_default_save_dir, normalize_path


class SettingsManager:
    """Инкапсулирует работу с QSettings и хранением пользовательских параметров."""

    KEY_WINDOW_GEOMETRY = "window/geometry"
    KEY_WINDOW_SIZE = "window/size"
    KEY_WINDOW_POSITION = "window/position"
    KEY_LAST_OPEN_DIR = "paths/last_open_dir"
    KEY_LAST_SAVE_DIR = "paths/last_save_dir"
    KEY_LAST_OUTPUT_NAME = "output/filename"
    KEY_LAST_MERGE_MODE = "output/merge_mode"
    KEY_RECENT_FILES = "files/recent"

    DEFAULT_OUTPUT_NAME = "merged.docx"
    DEFAULT_MERGE_MODE = "page_break"

    def __init__(self) -> None:
        self._settings = QSettings(APP_ORGANIZATION, APP_NAME)

    def load_window_geometry(self) -> QByteArray | None:
        """Возвращает сохраненную геометрию окна, если она есть."""
        geometry = self._settings.value(self.KEY_WINDOW_GEOMETRY, None)
        if isinstance(geometry, QByteArray):
            return geometry
        return None

    def save_window_geometry(self, geometry: QByteArray) -> None:
        """Сохраняет геометрию главного окна."""
        self._settings.setValue(self.KEY_WINDOW_GEOMETRY, geometry)

    def load_window_size(self) -> QSize | None:
        """Возвращает сохраненный размер окна, если он был записан отдельно."""
        size = self._settings.value(self.KEY_WINDOW_SIZE, None)
        if isinstance(size, QSize):
            return size
        return None

    def save_window_size(self, size: QSize) -> None:
        """Сохраняет размер окна."""
        self._settings.setValue(self.KEY_WINDOW_SIZE, size)

    def load_window_position(self) -> QPoint | None:
        """Возвращает сохраненную позицию окна, если она была записана отдельно."""
        position = self._settings.value(self.KEY_WINDOW_POSITION, None)
        if isinstance(position, QPoint):
            return position
        return None

    def save_window_position(self, position: QPoint) -> None:
        """Сохраняет позицию окна."""
        self._settings.setValue(self.KEY_WINDOW_POSITION, position)

    def load_last_open_dir(self) -> Path:
        """Возвращает последнюю директорию выбора входных файлов."""
        raw_value = self._settings.value(self.KEY_LAST_OPEN_DIR, "")
        if raw_value:
            return normalize_path(str(raw_value))
        return get_default_save_dir()

    def save_last_open_dir(self, path: str | Path) -> None:
        """Сохраняет последнюю директорию открытия документов."""
        self._settings.setValue(self.KEY_LAST_OPEN_DIR, str(normalize_path(path)))

    def load_last_save_dir(self) -> Path:
        """Возвращает последнюю директорию сохранения результата."""
        raw_value = self._settings.value(self.KEY_LAST_SAVE_DIR, "")
        if raw_value:
            return normalize_path(str(raw_value))
        return get_default_save_dir()

    def save_last_save_dir(self, path: str | Path) -> None:
        """Сохраняет последнюю директорию сохранения."""
        self._settings.setValue(self.KEY_LAST_SAVE_DIR, str(normalize_path(path)))

    def load_last_output_name(self) -> str:
        """Возвращает последнее введенное имя итогового файла."""
        return str(self._settings.value(self.KEY_LAST_OUTPUT_NAME, self.DEFAULT_OUTPUT_NAME))

    def save_last_output_name(self, filename: str) -> None:
        """Сохраняет последнее имя итогового файла."""
        self._settings.setValue(self.KEY_LAST_OUTPUT_NAME, filename)

    def load_last_merge_mode(self) -> str:
        """Возвращает последний выбранный режим объединения."""
        return str(self._settings.value(self.KEY_LAST_MERGE_MODE, self.DEFAULT_MERGE_MODE))

    def save_last_merge_mode(self, merge_mode: str) -> None:
        """Сохраняет последний выбранный режим объединения."""
        self._settings.setValue(self.KEY_LAST_MERGE_MODE, merge_mode)

    def load_recent_files(self) -> list[Path]:
        """Возвращает список последних файлов в сохраненном порядке."""
        raw_value = self._settings.value(self.KEY_RECENT_FILES, [])

        if isinstance(raw_value, str):
            values = [raw_value]
        elif isinstance(raw_value, list):
            values = [str(item) for item in raw_value]
        else:
            values = []

        recent_files: list[Path] = []
        seen: set[Path] = set()
        for value in values:
            path = normalize_path(value)
            if path in seen:
                continue
            seen.add(path)
            recent_files.append(path)

        return recent_files

    def save_recent_files(self, files: list[str | Path]) -> None:
        """Сохраняет список последних файлов в порядке, указанном пользователем."""
        normalized_files: list[str] = []
        seen: set[Path] = set()

        for file_path in files:
            path = normalize_path(file_path)
            if path in seen:
                continue
            seen.add(path)
            normalized_files.append(str(path))

        self._settings.setValue(self.KEY_RECENT_FILES, normalized_files)

    def clear_recent_files(self) -> None:
        """Очищает сохраненный список последних файлов."""
        self._settings.remove(self.KEY_RECENT_FILES)

    def sync(self) -> None:
        """Принудительно записывает настройки в хранилище."""
        self._settings.sync()
