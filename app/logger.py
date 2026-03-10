"""Настройка, маршрутизация и получение логгера приложения."""

from __future__ import annotations

import logging
import sys
from collections.abc import Callable
from pathlib import Path

from .platform_utils import get_app_data_dir


APP_LOGGER_NAME = "docx_merger"

_log_listeners: list[Callable[[str], None]] = []


class InterfaceLogHandler(logging.Handler):
    """Переадресует форматированные записи логгера в UI-подписчиков."""

    def emit(self, record: logging.LogRecord) -> None:
        """Передает сообщение всем зарегистрированным слушателям интерфейса."""
        try:
            message = self.format(record)
        except Exception:
            return

        for listener in list(_log_listeners):
            try:
                listener(message)
            except Exception:
                continue


def configure_logging() -> Path:
    """Настраивает файловый, консольный и интерфейсный логгеры приложения."""
    log_dir = get_app_data_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"

    logger = logging.getLogger(APP_LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if logger.handlers:
        return log_file

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    interface_handler = InterfaceLogHandler()
    interface_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.addHandler(interface_handler)

    return log_file


def get_logger(name: str | None = None) -> logging.Logger:
    """Возвращает дочерний логгер приложения."""
    logger_name = APP_LOGGER_NAME if not name else f"{APP_LOGGER_NAME}.{name}"
    return logging.getLogger(logger_name)


def register_log_listener(listener: Callable[[str], None]) -> None:
    """Регистрирует callback для текстового лога в интерфейсе."""
    if listener not in _log_listeners:
        _log_listeners.append(listener)


def unregister_log_listener(listener: Callable[[str], None]) -> None:
    """Удаляет callback интерфейсного лога."""
    if listener in _log_listeners:
        _log_listeners.remove(listener)


def install_excepthook() -> None:
    """Устанавливает глобальный обработчик непойманных исключений."""

    def _handle_exception(exc_type: type[BaseException], exc_value: BaseException, exc_traceback: object) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return

        logger = get_logger("unhandled")
        logger.exception(
            "Неперехваченное исключение приложения.",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    sys.excepthook = _handle_exception
