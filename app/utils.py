"""Общие утилиты валидации и подготовки данных."""

from __future__ import annotations

from pathlib import Path

from .merger import (
    CorruptedDocumentError,
    DocumentMergeError,
    EmptyFileListError,
    FileAccessError,
    MergeCancelledError,
    OutputWriteError,
    UnsupportedFormatError,
)
from .platform_utils import (
    PlatformActionError,
    get_invalid_filename_characters,
    get_reserved_windows_names,
    is_windows,
    normalize_path,
)


def ensure_docx_extension(filename: str) -> str:
    """Добавляет расширение .docx, если пользователь его не указал."""
    cleaned = filename.strip()
    if cleaned.lower().endswith(".docx"):
        return cleaned
    return f"{cleaned}.docx"


def sanitize_output_filename(filename: str) -> str:
    """Удаляет пробелы по краям и запрещенные завершающие символы."""
    cleaned = ensure_docx_extension(filename.strip())
    if is_windows():
        cleaned = cleaned.rstrip(" .")
    return cleaned


def is_valid_output_filename(filename: str) -> tuple[bool, str]:
    """Проверяет корректность имени выходного файла для текущей платформы."""
    cleaned = sanitize_output_filename(filename)

    if cleaned in {".docx", ""}:
        return False, "Укажите имя итогового файла."

    invalid_characters = get_invalid_filename_characters()
    if any(char in invalid_characters for char in cleaned):
        invalid_display = "".join(sorted(invalid_characters))
        return False, f"Имя файла содержит недопустимые символы: {invalid_display}"

    stem = Path(cleaned).stem
    if is_windows() and stem.upper() in get_reserved_windows_names():
        return False, "Имя файла использует зарезервированное системное имя Windows."

    return True, ""


def normalize_files(paths: list[str]) -> list[Path]:
    """Преобразует список строковых путей в нормализованные Path."""
    return [normalize_path(path) for path in paths]


def validate_source_files(paths: list[Path]) -> tuple[bool, str]:
    """Проверяет, что список входных документов пригоден для обработки."""
    if len(paths) < 2:
        return False, "Добавьте как минимум два DOCX-файла."

    for path in paths:
        if not path.exists():
            return False, f"Файл не найден: {path}"
        if path.suffix.lower() != ".docx":
            return False, f"Поддерживаются только файлы DOCX: {path.name}"
        if not path.is_file():
            return False, f"Ожидается файл, а не папка: {path}"

    return True, ""


def build_output_path(directory: str | Path, filename: str) -> Path:
    """Формирует абсолютный путь итогового файла."""
    return normalize_path(directory) / sanitize_output_filename(filename)


def build_user_friendly_error(error: Exception | str) -> tuple[str, str]:
    """Преобразует внутреннее исключение в понятный для пользователя заголовок и текст."""
    if isinstance(error, str):
        return "Ошибка", error

    if isinstance(error, MergeCancelledError):
        return "Операция отменена", "Объединение документов было остановлено пользователем."

    if isinstance(error, EmptyFileListError):
        return "Нет файлов", "Добавьте хотя бы один DOCX-файл для объединения."

    if isinstance(error, UnsupportedFormatError):
        return "Неподдерживаемый формат", str(error)

    if isinstance(error, CorruptedDocumentError):
        return "Поврежденный документ", str(error)

    if isinstance(error, FileAccessError):
        return "Ошибка доступа к файлу", str(error)

    if isinstance(error, OutputWriteError):
        return "Ошибка сохранения", str(error)

    if isinstance(error, PlatformActionError):
        return "Не удалось открыть файл или папку", str(error)

    if isinstance(error, DocumentMergeError):
        return "Ошибка объединения", str(error)

    if isinstance(error, ValueError):
        return "Некорректные параметры", str(error)

    return "Непредвиденная ошибка", str(error) or "Произошла неизвестная ошибка."
