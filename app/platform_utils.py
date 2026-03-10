"""Кроссплатформенные утилиты и тонкий слой абстракции между macOS и Windows."""

from __future__ import annotations

import os
import platform
import subprocess
from pathlib import Path

from .app_info import APP_STORAGE_DIR_NAME


class PlatformActionError(Exception):
    """Ошибка платформенного действия: открыть файл, папку или показать объект в системе."""


def get_platform_name() -> str:
    """Возвращает упрощенное имя текущей платформы."""
    system_name = platform.system().lower()
    if system_name.startswith("darwin"):
        return "macos"
    if system_name.startswith("windows"):
        return "windows"
    return system_name


def is_macos() -> bool:
    """Проверяет, что приложение запущено на macOS."""
    return get_platform_name() == "macos"


def is_windows() -> bool:
    """Проверяет, что приложение запущено на Windows."""
    return get_platform_name() == "windows"


def get_app_data_dir() -> Path:
    """Возвращает директорию приложения для логов и служебных файлов."""
    home = Path.home()

    if is_windows():
        local_app_data = os.environ.get("LOCALAPPDATA")
        base_dir = Path(local_app_data) if local_app_data else home / "AppData" / "Local"
        return base_dir / APP_STORAGE_DIR_NAME

    if is_macos():
        return home / "Library" / "Application Support" / APP_STORAGE_DIR_NAME

    return home / ".local" / "share" / APP_STORAGE_DIR_NAME


def normalize_path(path: str | Path) -> Path:
    """Приводит путь к абсолютному виду без обращения к файловой системе."""
    return Path(path).expanduser().resolve(strict=False)


def get_default_save_dir() -> Path:
    """Возвращает безопасную директорию по умолчанию для сохранения результата."""
    documents_dir = Path.home() / "Documents"
    if documents_dir.exists():
        return documents_dir
    return Path.home()


def get_invalid_filename_characters() -> set[str]:
    """Возвращает набор запрещенных символов имени файла для текущей ОС."""
    if is_windows():
        return {'<', '>', ':', '"', '/', '\\', '|', '?', '*'}
    return {'/', ':'}


def get_reserved_windows_names() -> set[str]:
    """Возвращает набор зарезервированных имен файлов Windows."""
    return {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        "COM1",
        "COM2",
        "COM3",
        "COM4",
        "COM5",
        "COM6",
        "COM7",
        "COM8",
        "COM9",
        "LPT1",
        "LPT2",
        "LPT3",
        "LPT4",
        "LPT5",
        "LPT6",
        "LPT7",
        "LPT8",
        "LPT9",
    }


def open_directory(path: str | Path) -> None:
    """Открывает папку в Finder, Explorer или другом файловом менеджере."""
    directory = normalize_path(path)
    if not directory.exists() or not directory.is_dir():
        raise PlatformActionError(f"Папка не найдена: {directory}")

    _open_path(directory)


def open_file(path: str | Path) -> None:
    """Открывает файл приложением по умолчанию."""
    file_path = normalize_path(path)
    if not file_path.exists() or not file_path.is_file():
        raise PlatformActionError(f"Файл не найден: {file_path}")

    _open_path(file_path)


def reveal_in_file_manager(path: str | Path) -> None:
    """Показывает файл или папку в Finder или Explorer."""
    target = normalize_path(path)
    if not target.exists():
        raise PlatformActionError(f"Файл или папка не найдены: {target}")

    try:
        if is_windows():
            subprocess.run(
                ["explorer", f"/select,{str(target)}"],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return

        if is_macos():
            subprocess.run(
                ["open", "-R", str(target)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return

        parent = target if target.is_dir() else target.parent
        _open_path(parent)
    except (OSError, subprocess.SubprocessError) as exc:
        raise PlatformActionError(
            f"Не удалось показать объект в файловом менеджере: {target}"
        ) from exc


def _open_path(target: Path) -> None:
    """Открывает путь платформенным способом по умолчанию."""
    try:
        if is_windows():
            os.startfile(str(target))  # type: ignore[attr-defined]
            return

        if is_macos():
            subprocess.run(
                ["open", str(target)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return

        subprocess.run(
            ["xdg-open", str(target)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (AttributeError, OSError, subprocess.SubprocessError) as exc:
        raise PlatformActionError(f"Не удалось открыть: {target}") from exc
