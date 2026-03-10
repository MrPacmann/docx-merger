"""Точка входа в desktop-приложение объединения DOCX-файлов."""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from app.app_info import APP_NAME, APP_ORGANIZATION, APP_ORGANIZATION_DOMAIN
from app.logger import configure_logging, get_logger, install_excepthook
from app.settings_manager import SettingsManager
from app.ui import MainWindow


def main() -> int:
    """Создает QApplication, инициализирует зависимости и запускает GUI."""
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName(APP_ORGANIZATION)
    app.setOrganizationDomain(APP_ORGANIZATION_DOMAIN)

    configure_logging()
    install_excepthook()
    logger = get_logger(__name__)
    settings = SettingsManager()

    logger.info("Приложение запущено.")

    window = MainWindow(settings_manager=settings)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
