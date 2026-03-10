"""Единое место для типовых и служебных диалогов приложения."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QMessageBox,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from .app_info import APP_AUTHOR, APP_GITHUB_URL, APP_NAME, APP_VERSION


class InfoDialog(QDialog):
    """Базовый служебный диалог для статической информации."""

    def __init__(self, title: str, body: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(560, 420)
        self.setMinimumSize(440, 320)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        content = QPlainTextEdit(self)
        content.setReadOnly(True)
        content.setPlainText(body)
        content.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        content.setStyleSheet(
            """
            QPlainTextEdit {
                background: #ffffff;
                border: 1px solid #d6dde5;
                border-radius: 10px;
                padding: 10px;
                color: #243b53;
                selection-background-color: #dbeafe;
            }
            """
        )

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok, parent=self)
        buttons.accepted.connect(self.accept)

        layout.addWidget(content)
        layout.addWidget(buttons)


def choose_docx_files(parent: QWidget, start_dir: Path) -> list[str]:
    """Открывает диалог выбора исходных DOCX-файлов."""
    files, _ = QFileDialog.getOpenFileNames(
        parent,
        "Выберите DOCX-файлы",
        str(start_dir),
        "Документы Word (*.docx)",
    )
    return files


def choose_output_dir(parent: QWidget, start_dir: Path) -> str:
    """Открывает диалог выбора папки сохранения результата."""
    return QFileDialog.getExistingDirectory(
        parent,
        "Выберите папку сохранения",
        str(start_dir),
    )


def show_error(parent: QWidget, title: str, message: str) -> None:
    """Показывает диалог ошибки."""
    QMessageBox.critical(parent, title, message)


def show_info(parent: QWidget, title: str, message: str) -> None:
    """Показывает информационный диалог."""
    QMessageBox.information(parent, title, message)


def ask_overwrite(parent: QWidget, path: Path) -> bool:
    """Запрашивает подтверждение на перезапись существующего файла."""
    result = QMessageBox.question(
        parent,
        "Файл уже существует",
        f"Файл\n{path}\nуже существует.\nПерезаписать его?",
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No,
    )
    return result == QMessageBox.StandardButton.Yes


def show_about_dialog(parent: QWidget | None = None) -> None:
    """Показывает окно 'О программе'."""
    body = (
        f"{APP_NAME}\n"
        f"Версия: {APP_VERSION}\n\n"
        "Приложение для объединения DOCX-документов с максимальным сохранением форматирования.\n\n"
        f"Автор: {APP_AUTHOR}\n"
        f"GitHub: {APP_GITHUB_URL}"
    )
    dialog = InfoDialog("О программе", body, parent=parent)
    dialog.exec()


def show_instruction_dialog(parent: QWidget | None = None) -> None:
    """Показывает окно с инструкцией по использованию приложения."""
    body = (
        "Как пользоваться приложением\n\n"
        "1. Как добавить файлы\n"
        "- Нажмите кнопку «Добавить файлы» или перетащите DOCX-файлы в окно приложения.\n"
        "- Поддерживаются только локальные файлы с расширением .docx.\n\n"
        "2. Как поменять порядок\n"
        "- Перетаскивайте файлы мышью внутри списка.\n"
        "- Для точного управления используйте кнопки «Вверх» и «Вниз».\n\n"
        "3. Как выбрать режим объединения\n"
        "- В блоке параметров сохранения выберите режим:\n"
        "  • разрыв страницы;\n"
        "  • разрыв раздела;\n"
        "  • без разрыва.\n"
        "- Выбранный режим влияет на то, как следующий документ присоединяется к предыдущему.\n\n"
        "4. Как сохранить итоговый файл\n"
        "- Укажите папку сохранения.\n"
        "- Введите имя итогового файла.\n"
        "- Нажмите кнопку «Объединить».\n\n"
        "5. Как отменить операцию\n"
        "- Во время выполнения нажмите «Отменить».\n"
        "- Остановка выполняется безопасно на ближайшей допустимой точке.\n\n"
        "6. Ограничения приложения\n"
        "- Поддерживаются только файлы .docx.\n"
        "- Поврежденные или недоступные документы не будут обработаны.\n"
        "- Точность прогресса и оставшегося времени приблизительная.\n"
        "- Итог зависит от особенностей структуры исходных документов и возможностей бесплатного движка объединения.\n"
        "- Колонтитулы и часть секционных настроек ориентируются в первую очередь на первый документ."
    )
    dialog = InfoDialog("Инструкция", body, parent=parent)
    dialog.exec()
