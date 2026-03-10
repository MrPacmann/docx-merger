"""Главное окно приложения и координация пользовательского сценария."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtGui import QCloseEvent, QDragEnterEvent, QDragMoveEvent, QDropEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QPlainTextEdit,
    QPushButton,
    QProgressBar,
    QMessageBox,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from .app_info import APP_NAME, APP_VERSION
from .dialogs import (
    ask_overwrite,
    choose_docx_files,
    choose_output_dir,
    show_about_dialog,
    show_error,
    show_info,
    show_instruction_dialog,
)
from .logger import get_logger, register_log_listener, unregister_log_listener
from .merger import DocxMerger, MergeRequest
from .platform_utils import open_directory, open_file, reveal_in_file_manager
from .settings_manager import SettingsManager
from .utils import (
    build_output_path,
    build_user_friendly_error,
    is_valid_output_filename,
    normalize_files,
    validate_source_files,
)
from .worker import MergeWorker


class FileListWidget(QListWidget):
    """Список файлов с поддержкой внешнего drag-and-drop и внутренней перестановки."""

    files_dropped = Signal(list)
    order_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDragDropOverwriteMode(False)
        self.viewport().setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802 - Qt API
        """Принимает перенос DOCX-файлов извне или элементов списка внутри виджета."""
        if event.source() is self:
            event.acceptProposedAction()
            return

        if self._extract_docx_paths(event) or self._contains_file_urls(event):
            event.acceptProposedAction()
            return

        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:  # noqa: N802 - Qt API
        """Поддерживает индикацию позиции при переносе файлов."""
        if event.source() is self:
            event.acceptProposedAction()
            return

        if self._extract_docx_paths(event):
            event.acceptProposedAction()
            return

        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802 - Qt API
        """Обрабатывает внутреннюю перестановку и внешний импорт DOCX-файлов."""
        if event.source() is self:
            super().dropEvent(event)
            self.order_changed.emit()
            event.acceptProposedAction()
            return

        dropped_files = self._extract_docx_paths(event)
        if dropped_files:
            self.files_dropped.emit(dropped_files)
            event.acceptProposedAction()
            return

        event.ignore()

    @staticmethod
    def _contains_file_urls(event: QDragEnterEvent | QDragMoveEvent | QDropEvent) -> bool:
        """Проверяет, содержит ли событие URL-адреса файлов."""
        return event.mimeData().hasUrls()

    @staticmethod
    def _extract_docx_paths(event: QDragEnterEvent | QDragMoveEvent | QDropEvent) -> list[str]:
        """Извлекает только локальные DOCX-файлы из события drag-and-drop."""
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            return []

        paths: list[str] = []
        for url in mime_data.urls():
            if not url.isLocalFile():
                continue

            path = Path(url.toLocalFile())
            if path.is_file() and path.suffix.lower() == ".docx":
                paths.append(str(path))

        return paths


class LogRelay(QObject):
    """Безопасно доставляет сообщения логгера в GUI-поток."""

    message_received = Signal(str)


class MainWindow(QMainWindow):
    """Главное окно приложения объединения DOCX-файлов."""

    MERGE_MODES: tuple[tuple[str, str], ...] = (
        ("Разрыв страницы", "page_break"),
        ("Разрыв раздела", "section_break"),
        ("Без разрыва", "no_break"),
    )

    def __init__(self, settings_manager: SettingsManager) -> None:
        super().__init__()
        self._logger = get_logger(__name__)
        self._settings = settings_manager
        self._merger = DocxMerger()

        self._worker_thread: QThread | None = None
        self._worker: MergeWorker | None = None
        self._last_output_path: Path | None = None
        self._close_after_worker = False
        self._log_relay = LogRelay(self)

        self._build_ui()
        self._log_relay.message_received.connect(self._append_log_from_logger)
        register_log_listener(self._log_relay.message_received.emit)
        self._restore_settings()

    def _build_ui(self) -> None:
        """Создает основной интерфейс главного окна."""
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(920, 620)
        self.setMinimumSize(820, 560)
        self.setAcceptDrops(True)
        self._build_menu()

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(16)

        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)
        root_layout.addWidget(splitter, stretch=1)

        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(16)

        files_group = QGroupBox("Файлы для объединения")
        files_layout = QGridLayout(files_group)
        files_layout.setContentsMargins(14, 18, 14, 14)
        files_layout.setHorizontalSpacing(10)
        files_layout.setVerticalSpacing(10)

        self.files_list = FileListWidget()
        self.files_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.files_list.setAlternatingRowColors(True)
        self.files_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.files_list.setToolTip(
            "Список документов в порядке объединения. Можно перетаскивать DOCX-файлы в окно и менять порядок внутри списка."
        )

        files_layout.addWidget(self.files_list, 0, 0, 7, 1)

        self.add_files_button = QPushButton("Добавить файлы")
        self.remove_file_button = QPushButton("Удалить выбранный")
        self.move_up_button = QPushButton("Вверх")
        self.move_down_button = QPushButton("Вниз")
        self.clear_button = QPushButton("Очистить")

        self.add_files_button.setToolTip("Добавить один или несколько DOCX-файлов в список.")
        self.remove_file_button.setToolTip("Удалить один или несколько выделенных документов из списка.")
        self.move_up_button.setToolTip("Переместить выделенные файлы выше в порядке объединения.")
        self.move_down_button.setToolTip("Переместить выделенные файлы ниже в порядке объединения.")
        self.clear_button.setToolTip("Полностью очистить список файлов.")

        file_buttons = [
            self.add_files_button,
            self.remove_file_button,
            self.move_up_button,
            self.move_down_button,
            self.clear_button,
        ]
        for row, button in enumerate(file_buttons):
            button.setMinimumHeight(38)
            button.setMinimumWidth(154)
            files_layout.addWidget(button, row, 1)
        files_layout.setRowStretch(6, 1)

        self.file_count_label = QLabel("Файлов выбрано: 0")
        self.file_count_label.setToolTip("Количество документов, которые войдут в итоговый файл.")
        self.file_count_label.setProperty("secondary", True)
        files_layout.addWidget(self.file_count_label, 7, 0, 1, 2)

        output_group = QGroupBox("Параметры сохранения")
        output_layout = QFormLayout(output_group)
        output_layout.setContentsMargins(14, 18, 14, 14)
        output_layout.setHorizontalSpacing(12)
        output_layout.setVerticalSpacing(12)
        output_layout.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("Выберите папку для сохранения")
        self.output_dir_edit.setToolTip("Папка, в которую будет сохранен итоговый DOCX-файл.")

        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("Например: merged.docx")
        self.output_name_edit.setToolTip("Имя итогового файла. Расширение .docx добавится автоматически.")

        self.merge_mode_combo = QComboBox()
        self.merge_mode_combo.setToolTip(
            "Выберите способ соединения документов: через разрыв страницы, разрыв раздела или без разрыва."
        )
        for title, value in self.MERGE_MODES:
            self.merge_mode_combo.addItem(title, value)

        choose_dir_layout = QHBoxLayout()
        choose_dir_layout.setContentsMargins(0, 0, 0, 0)
        choose_dir_layout.setSpacing(8)
        choose_dir_layout.addWidget(self.output_dir_edit)

        self.choose_output_dir_button = QPushButton("Обзор...")
        self.choose_output_dir_button.setMinimumHeight(38)
        self.choose_output_dir_button.setToolTip("Открыть диалог выбора папки сохранения.")
        choose_dir_layout.addWidget(self.choose_output_dir_button)

        output_layout.addRow("Папка сохранения:", self._wrap_layout(choose_dir_layout))
        output_layout.addRow("Имя файла:", self.output_name_edit)
        output_layout.addRow("Режим объединения:", self.merge_mode_combo)

        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(8)

        self.progress_label = QLabel("Готово к работе")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.progress_label.setStyleSheet("font-weight: 600;")

        self.progress_details_label = QLabel("Прогресс не запущен")
        self.progress_details_label.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        self.progress_details_label.setToolTip(
            "Показывает процент выполнения и приблизительное оставшееся время."
        )
        self.progress_details_label.setProperty("secondary", True)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setToolTip("Индикатор выполнения операции объединения.")

        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_details_label)
        progress_layout.addWidget(self.progress_bar)

        top_layout.addWidget(files_group, stretch=3)
        top_layout.addWidget(output_group, stretch=0)
        top_layout.addLayout(progress_layout)

        logs_group = QGroupBox("Журнал")
        logs_layout = QVBoxLayout(logs_group)
        logs_layout.setContentsMargins(14, 18, 14, 14)
        logs_layout.setSpacing(10)

        logs_hint = QLabel("Служебные сообщения, статусы выполнения и диагностическая информация.")
        logs_hint.setProperty("secondary", True)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setPlaceholderText("Здесь будут отображаться сообщения о ходе работы приложения.")
        self.log_output.setToolTip("Текстовый журнал событий интерфейса и операций объединения.")
        self.log_output.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.log_output.setMaximumBlockCount(2000)
        logs_layout.addWidget(logs_hint)
        logs_layout.addWidget(self.log_output)

        splitter.addWidget(top_widget)
        splitter.addWidget(logs_group)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 2)

        action_layout = QHBoxLayout()
        action_layout.setSpacing(10)

        self.start_button = QPushButton("Объединить")
        self.cancel_button = QPushButton("Отменить")
        self.open_folder_button = QPushButton("Открыть папку")
        self.open_file_button = QPushButton("Открыть файл")
        self.reveal_file_button = QPushButton("Показать в системе")
        self.cancel_button.setEnabled(False)
        self.open_folder_button.setEnabled(False)
        self.open_file_button.setEnabled(False)
        self.reveal_file_button.setEnabled(False)
        self.start_button.setToolTip("Запустить объединение документов с текущими параметрами.")
        self.cancel_button.setToolTip("Отменить выполняемую операцию после безопасной точки остановки.")
        self.open_folder_button.setToolTip("Открыть папку, в которой находится итоговый файл.")
        self.open_file_button.setToolTip("Открыть итоговый файл приложением по умолчанию.")
        self.reveal_file_button.setToolTip("Показать итоговый файл в Finder или Проводнике.")

        self.start_button.setMinimumHeight(40)
        self.cancel_button.setMinimumHeight(40)
        self.open_folder_button.setMinimumHeight(40)
        self.open_file_button.setMinimumHeight(40)
        self.reveal_file_button.setMinimumHeight(40)
        self.start_button.setMinimumWidth(132)
        self.cancel_button.setMinimumWidth(118)

        action_layout.addStretch(1)
        action_layout.addWidget(self.open_folder_button)
        action_layout.addWidget(self.open_file_button)
        action_layout.addWidget(self.reveal_file_button)
        action_layout.addWidget(self.cancel_button)
        action_layout.addWidget(self.start_button)

        root_layout.addLayout(action_layout)

        self.setStatusBar(QStatusBar(self))
        self.statusBar().showMessage("Добавьте DOCX-файлы для объединения.")
        self._append_log("Интерфейс готов к работе.")

        self._apply_styles()

        self.add_files_button.clicked.connect(self._add_files)
        self.remove_file_button.clicked.connect(self._remove_selected_file)
        self.move_up_button.clicked.connect(self._move_selected_up)
        self.move_down_button.clicked.connect(self._move_selected_down)
        self.clear_button.clicked.connect(self._clear_files)
        self.choose_output_dir_button.clicked.connect(self._choose_output_directory)
        self.start_button.clicked.connect(self._start_merge)
        self.cancel_button.clicked.connect(self._cancel_merge)
        self.open_folder_button.clicked.connect(self._open_result_folder)
        self.open_file_button.clicked.connect(self._open_result_file)
        self.reveal_file_button.clicked.connect(self._reveal_result_file)
        self.files_list.itemSelectionChanged.connect(self._update_buttons_state)
        self.files_list.files_dropped.connect(self._handle_dropped_files)
        self.files_list.order_changed.connect(self._on_list_order_changed)
        self.output_name_edit.editingFinished.connect(self._save_form_settings)
        self.output_dir_edit.editingFinished.connect(self._save_form_settings)
        self.merge_mode_combo.currentIndexChanged.connect(self._save_form_settings)

        self._update_buttons_state()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802 - Qt API
        """Разрешает перетаскивание DOCX-файлов в окно приложения."""
        dropped_files = FileListWidget._extract_docx_paths(event)
        if dropped_files or FileListWidget._contains_file_urls(event):
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:  # noqa: N802 - Qt API
        """Поддерживает визуальное сопровождение внешнего перетаскивания."""
        dropped_files = FileListWidget._extract_docx_paths(event)
        if dropped_files:
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802 - Qt API
        """Добавляет DOCX-файлы, брошенные на любую область главного окна."""
        dropped_files = FileListWidget._extract_docx_paths(event)
        if not dropped_files:
            self._append_log("Перетаскивание проигнорировано: поддерживаются только локальные файлы .docx.")
            event.ignore()
            return

        self._add_file_paths(dropped_files, source="drop")
        event.acceptProposedAction()

    def _wrap_layout(self, layout: QHBoxLayout) -> QWidget:
        """Помещает layout в QWidget, чтобы использовать его в форме."""
        container = QWidget()
        container.setLayout(layout)
        return container

    def _build_menu(self) -> None:
        """Создает верхнее меню приложения."""
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("Файл")
        settings_menu = menu_bar.addMenu("Настройки")
        help_menu = menu_bar.addMenu("Справка")

        add_files_action = file_menu.addAction("Добавить файлы...")
        add_files_action.triggered.connect(self._add_files)

        choose_dir_action = file_menu.addAction("Выбрать папку сохранения...")
        choose_dir_action.triggered.connect(self._choose_output_directory)

        file_menu.addSeparator()

        self.open_folder_action = file_menu.addAction("Открыть папку результата")
        self.open_folder_action.triggered.connect(self._open_result_folder)

        self.open_file_action = file_menu.addAction("Открыть итоговый файл")
        self.open_file_action.triggered.connect(self._open_result_file)

        self.reveal_file_action = file_menu.addAction("Показать итоговый файл в системе")
        self.reveal_file_action.triggered.connect(self._reveal_result_file)

        file_menu.addSeparator()

        exit_action = file_menu.addAction("Выход")
        exit_action.triggered.connect(self.close)

        reset_view_action = settings_menu.addAction("Сбросить размер и положение окна")
        reset_view_action.triggered.connect(self._reset_window_layout)

        clear_recent_action = settings_menu.addAction("Очистить список последних файлов")
        clear_recent_action.triggered.connect(self._clear_recent_files_and_list)

        instruction_action = help_menu.addAction("Инструкция")
        instruction_action.triggered.connect(lambda: show_instruction_dialog(self))

        about_action = help_menu.addAction("О программе")
        about_action.triggered.connect(lambda: show_about_dialog(self))

    def _apply_styles(self) -> None:
        """Применяет аккуратный нейтральный стиль, одинаково подходящий для macOS и Windows."""
        self.setStyleSheet(
            """
            QMainWindow {
                background: #eef3f7;
            }
            QMenuBar {
                background: #ffffff;
                border-bottom: 1px solid #d7dde4;
                padding: 4px 8px;
                color: #243b53;
            }
            QMenuBar::item {
                padding: 6px 10px;
                border-radius: 6px;
                background: transparent;
            }
            QMenuBar::item:selected {
                background: #e8edf4;
            }
            QMenu {
                background: #ffffff;
                border: 1px solid #d7dde4;
                padding: 6px;
            }
            QMenu::item {
                padding: 7px 24px 7px 12px;
                border-radius: 6px;
            }
            QMenu::item:selected {
                background: #e8edf4;
            }
            QGroupBox {
                background: #ffffff;
                border: 1px solid #d7dde4;
                border-radius: 12px;
                margin-top: 14px;
                font-weight: 600;
                color: #26323f;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 4px;
            }
            QListWidget, QLineEdit, QComboBox, QPlainTextEdit {
                background: #ffffff;
                border: 1px solid #c7d0da;
                border-radius: 8px;
                padding: 7px 9px;
                color: #1f2933;
            }
            QListWidget::item {
                padding: 7px 6px;
            }
            QListWidget::item:selected {
                background: #dbeafe;
                color: #102a43;
                border-radius: 4px;
            }
            QPushButton {
                background: #edf2f7;
                border: 1px solid #c7d0da;
                border-radius: 8px;
                padding: 8px 12px;
                color: #243b53;
            }
            QPushButton:hover {
                background: #e5ebf1;
            }
            QPushButton:pressed {
                background: #d9e2ec;
            }
            QPushButton:disabled {
                color: #7b8794;
                background: #f1f5f8;
            }
            QPushButton[text="Объединить"] {
                background: #2563eb;
                border-color: #1d4ed8;
                color: #ffffff;
                font-weight: 600;
            }
            QPushButton[text="Объединить"]:hover {
                background: #1d4ed8;
            }
            QPushButton[text="Объединить"]:pressed {
                background: #1e40af;
            }
            QProgressBar {
                background: #e8edf2;
                border: 1px solid #c7d0da;
                border-radius: 8px;
                text-align: center;
                min-height: 20px;
                color: #102a43;
            }
            QProgressBar::chunk {
                background: #3b82f6;
                border-radius: 7px;
            }
            QStatusBar {
                background: #ffffff;
                border-top: 1px solid #d7dde4;
                color: #243b53;
            }
            QLabel {
                color: #243b53;
            }
            QLabel[secondary="true"] {
                color: #627d98;
            }
            """
        )

    def _restore_settings(self) -> None:
        """Восстанавливает настройки окна и последних пользовательских значений."""
        geometry = self._settings.load_window_geometry()
        if geometry is not None:
            self.restoreGeometry(geometry)
        else:
            window_size = self._settings.load_window_size()
            if window_size is not None:
                self.resize(window_size)

            window_position = self._settings.load_window_position()
            if window_position is not None:
                self.move(window_position)

        self.output_dir_edit.setText(str(self._settings.load_last_save_dir()))
        self.output_name_edit.setText(self._settings.load_last_output_name())
        self._restore_merge_mode()
        self._restore_recent_files()

    def _save_form_settings(self) -> None:
        """Сохраняет текущие значения формы, которые должны переживать перезапуск."""
        output_dir = self.output_dir_edit.text().strip()
        if output_dir:
            self._settings.save_last_save_dir(output_dir)

        output_name = self.output_name_edit.text().strip()
        if output_name:
            self._settings.save_last_output_name(output_name)

        self._settings.save_last_merge_mode(self._current_merge_mode())
        self._settings.sync()

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802 - Qt API
        """Сохраняет геометрию окна и корректно завершает закрытие."""
        if self._worker is not None:
            should_cancel = QMessageBox.question(
                self,
                "Операция выполняется",
                "Объединение документов еще выполняется.\nОтменить операцию и закрыть приложение?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if should_cancel != QMessageBox.StandardButton.Yes:
                event.ignore()
                return

            self._close_after_worker = True
            self._cancel_merge()
            event.ignore()
            return

        unregister_log_listener(self._log_relay.message_received.emit)
        self._settings.save_window_geometry(self.saveGeometry())
        self._settings.save_window_size(self.size())
        self._settings.save_window_position(self.pos())
        self._persist_recent_files()
        self._save_form_settings()
        super().closeEvent(event)

    def _add_files(self) -> None:
        """Открывает диалог выбора файлов и добавляет новые элементы в список."""
        start_dir = self._settings.load_last_open_dir()
        selected_files = choose_docx_files(self, start_dir)
        if not selected_files:
            self._append_log("Выбор файлов отменен пользователем.")
            return

        self._add_file_paths(selected_files, source="dialog")

    def _remove_selected_file(self) -> None:
        """Удаляет один или несколько выделенных файлов из списка."""
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return

        rows = sorted((self.files_list.row(item) for item in selected_items), reverse=True)
        removed_paths: list[str] = []

        for row in rows:
            item = self.files_list.takeItem(row)
            removed_paths.append(str(item.data(Qt.ItemDataRole.UserRole)))

        for path in reversed(removed_paths):
            self._append_log(f"Удален файл из списка: {path}")

        self._persist_recent_files()
        self._update_buttons_state()

    def _move_selected_up(self) -> None:
        """Перемещает выделенные элементы списка на одну позицию вверх."""
        selected_rows = self._selected_rows()
        if not selected_rows or selected_rows[0] == 0:
            return

        for row in selected_rows:
            item = self.files_list.takeItem(row)
            self.files_list.insertItem(row - 1, item)

        self._restore_selection([row - 1 for row in selected_rows])
        self._append_log("Выделенные файлы перемещены вверх.")
        self._persist_recent_files()
        self._update_buttons_state()

    def _move_selected_down(self) -> None:
        """Перемещает выделенные элементы списка на одну позицию вниз."""
        selected_rows = self._selected_rows()
        if not selected_rows or selected_rows[-1] >= self.files_list.count() - 1:
            return

        for row in reversed(selected_rows):
            item = self.files_list.takeItem(row)
            self.files_list.insertItem(row + 1, item)

        self._restore_selection([row + 1 for row in selected_rows])
        self._append_log("Выделенные файлы перемещены вниз.")
        self._persist_recent_files()
        self._update_buttons_state()

    def _clear_files(self) -> None:
        """Очищает список входных документов."""
        if self.files_list.count() > 0:
            self._append_log("Список файлов очищен.")
        self.files_list.clear()
        self._persist_recent_files()
        self._update_buttons_state()

    def _choose_output_directory(self) -> None:
        """Позволяет пользователю выбрать папку сохранения результата."""
        current_text = self.output_dir_edit.text().strip()
        start_dir = Path(current_text) if current_text else self._settings.load_last_save_dir()
        selected_dir = choose_output_dir(self, start_dir)
        if not selected_dir:
            self._append_log("Выбор папки сохранения отменен пользователем.")
            return

        self.output_dir_edit.setText(selected_dir)
        self._settings.save_last_save_dir(selected_dir)
        self._settings.sync()
        self._append_log(f"Выбрана папка сохранения: {selected_dir}")

    def _collect_source_files(self) -> list[Path]:
        """Извлекает текущий порядок входных файлов из списка."""
        files: list[Path] = []
        for index in range(self.files_list.count()):
            item = self.files_list.item(index)
            files.append(Path(item.data(Qt.ItemDataRole.UserRole)))
        return files

    def _selected_rows(self) -> list[int]:
        """Возвращает отсортированные индексы выделенных строк списка."""
        rows = {self.files_list.row(item) for item in self.files_list.selectedItems()}
        return sorted(rows)

    def _restore_selection(self, rows: list[int]) -> None:
        """Восстанавливает выделение элементов после программной перестановки."""
        self.files_list.clearSelection()
        for row in rows:
            item = self.files_list.item(row)
            if item is not None:
                item.setSelected(True)

        if rows:
            self.files_list.setCurrentRow(rows[0])

    def _add_file_paths(self, paths: list[str], source: str) -> None:
        """Добавляет в список только уникальные DOCX-файлы и логирует результат."""
        normalized = normalize_files(paths)
        unique_existing = {
            str(self.files_list.item(index).data(Qt.ItemDataRole.UserRole))
            for index in range(self.files_list.count())
        }

        added_count = 0
        duplicate_count = 0
        invalid_count = 0
        first_added_parent: Path | None = None

        for path in normalized:
            path_text = str(path)
            if path.suffix.lower() != ".docx" or not path.is_file():
                invalid_count += 1
                self._append_log(f"Пропущен неподдерживаемый файл: {path_text}")
                continue

            if path_text in unique_existing:
                duplicate_count += 1
                self._append_log(f"Пропущен дубликат файла: {path_text}")
                continue

            item = QListWidgetItem(path_text)
            item.setToolTip(path_text)
            item.setData(Qt.ItemDataRole.UserRole, path_text)
            self.files_list.addItem(item)
            unique_existing.add(path_text)
            added_count += 1
            if first_added_parent is None:
                first_added_parent = path.parent
            self._append_log(f"Добавлен файл: {path_text}")

        if first_added_parent is not None:
            self._settings.save_last_open_dir(first_added_parent)
            self._persist_recent_files()
            self._settings.sync()

        if source == "drop" and added_count == 0 and invalid_count == 0 and duplicate_count == 0:
            self._append_log("Перетаскивание не добавило новых файлов.")

        self._update_buttons_state()

    def _handle_dropped_files(self, paths: list[str]) -> None:
        """Добавляет файлы, брошенные непосредственно в область списка."""
        self._add_file_paths(paths, source="drop")

    def _on_list_order_changed(self) -> None:
        """Обрабатывает завершение внутренней перестановки файлов в списке."""
        self._append_log("Порядок файлов изменен перетаскиванием.")
        self._persist_recent_files()
        self._update_buttons_state()

    def _current_merge_mode(self) -> str:
        """Возвращает код выбранного режима объединения."""
        return str(self.merge_mode_combo.currentData())

    def _start_merge(self) -> None:
        """Проверяет ввод, запускает worker и обновляет состояние интерфейса."""
        source_files = self._collect_source_files()
        valid_sources, source_error = validate_source_files(source_files)
        if not valid_sources:
            self._append_log(f"Ошибка проверки входных файлов: {source_error}")
            show_error(self, "Некорректные входные данные", source_error)
            return

        output_dir = self.output_dir_edit.text().strip()
        output_name = self.output_name_edit.text().strip()

        is_name_valid, name_error = is_valid_output_filename(output_name)
        if not is_name_valid:
            self._append_log(f"Ошибка проверки имени файла: {name_error}")
            show_error(self, "Некорректное имя файла", name_error)
            return

        if not output_dir:
            self._append_log("Папка сохранения не выбрана.")
            show_error(self, "Не выбрана папка", "Укажите папку, в которую нужно сохранить итоговый файл.")
            return

        output_path = build_output_path(output_dir, output_name)
        if output_path in source_files:
            message = "Итоговый файл не должен совпадать ни с одним из исходных документов."
            self._append_log(message)
            show_error(self, "Некорректный путь сохранения", message)
            return

        if output_path.exists() and not ask_overwrite(self, output_path):
            self._append_log(f"Перезапись файла отменена: {output_path}")
            return

        self._save_form_settings()
        self._run_worker(
            MergeRequest(
                source_files=source_files,
                output_file=output_path,
                merge_mode=self._current_merge_mode(),
            )
        )

    def _run_worker(self, request: MergeRequest) -> None:
        """Создает рабочий поток и подключает сигналы фоновой операции."""
        self._set_busy_state(True)
        self.progress_label.setText("Подготовка к объединению...")
        self.progress_details_label.setText("0% • выполняется подготовка")
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.statusBar().showMessage("Выполняется объединение документов...")
        self._append_log(
            "Запущено объединение: "
            f"{len(request.source_files)} файлов, режим={request.merge_mode}, выход={request.output_file}"
        )

        self._worker_thread = QThread(self)
        self._worker = MergeWorker(request=request, merger=self._merger)
        self._worker.moveToThread(self._worker_thread)

        self._worker_thread.started.connect(self._worker.run)
        self._worker.progress_changed.connect(self._on_progress_changed)
        self._worker.status_changed.connect(self._on_worker_status_changed)
        self._worker.finished.connect(self._on_merge_finished)
        self._worker.failed.connect(self._on_merge_failed)
        self._worker.cancelled.connect(self._on_merge_cancelled)

        self._worker.finished.connect(self._cleanup_worker)
        self._worker.failed.connect(self._cleanup_worker)
        self._worker.cancelled.connect(self._cleanup_worker)

        self._worker_thread.start()

    def _cancel_merge(self) -> None:
        """Передает worker запрос на кооперативную отмену операции."""
        if self._worker is None:
            return
        self._worker.cancel()
        self.cancel_button.setEnabled(False)
        self.progress_label.setText("Отмена операции...")
        self.progress_details_label.setText("Ожидание безопасной остановки...")
        self.statusBar().showMessage("Ожидание безопасной остановки...")
        self._append_log("Запрошена отмена операции.")

    def _cleanup_worker(self) -> None:
        """Корректно завершает рабочий поток и освобождает ссылки."""
        if self._worker_thread is not None:
            self._worker_thread.quit()
            self._worker_thread.wait()
            self._worker_thread.deleteLater()
            self._worker_thread = None

        if self._worker is not None:
            self._worker.deleteLater()
            self._worker = None

        self._set_busy_state(False)
        if self._close_after_worker:
            self._close_after_worker = False
            self.close()

    def _on_progress_changed(
        self,
        current: int,
        total: int,
        percent: int,
        eta_text: str,
        message: str,
    ) -> None:
        """Обновляет индикатор прогресса по сигналам worker."""
        self.progress_bar.setValue(percent)
        self.progress_bar.setFormat(f"{percent}%")
        self.progress_label.setText(message)
        self.progress_details_label.setText(f"{percent}% • {current}/{total} • {eta_text}")
        self.statusBar().showMessage(message)
        self._append_log(message)

    def _on_worker_status_changed(self, message: str) -> None:
        """Получает вспомогательные статусы от worker и пишет их в журнал и строку состояния."""
        self.statusBar().showMessage(message)
        self._append_log(message)

    def _on_merge_finished(self, output_path: str) -> None:
        """Показывает пользователю результат успешного объединения."""
        output = Path(output_path)
        self._last_output_path = output
        self._settings.save_last_save_dir(output.parent)
        self._settings.save_last_output_name(output.name)
        self._settings.sync()

        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("100%")
        self.progress_label.setText("Объединение успешно завершено")
        self.progress_details_label.setText("100% • готово")
        self.statusBar().showMessage(f"Готово: {output}")
        self._append_log(f"Объединение завершено успешно: {output}")
        self._update_result_action_buttons()
        show_info(self, "Готово", f"Итоговый файл сохранен:\n{output}")

    def _on_merge_failed(self, title: str, message: str) -> None:
        """Показывает ошибку, полученную из фоновой задачи."""
        self._logger.error("Объединение завершилось с ошибкой: %s | %s", title, message)
        self.progress_label.setText("Ошибка объединения")
        self.progress_details_label.setText("Выполнение прервано из-за ошибки")
        self.statusBar().showMessage("Операция завершилась с ошибкой.")
        self._append_log(f"{title}: {message}")
        show_error(self, title, message)

    def _on_merge_cancelled(self) -> None:
        """Обрабатывает штатную отмену операции пользователем."""
        self.progress_label.setText("Операция отменена")
        self.progress_details_label.setText("Операция остановлена безопасно")
        self.statusBar().showMessage("Операция была отменена пользователем.")
        self._append_log("Операция объединения отменена пользователем.")

    def _set_busy_state(self, is_busy: bool) -> None:
        """Блокирует или разблокирует элементы управления во время задачи."""
        selected_rows = self._selected_rows()
        self.add_files_button.setEnabled(not is_busy)
        self.remove_file_button.setEnabled(not is_busy and bool(selected_rows))
        self.move_up_button.setEnabled(not is_busy and bool(selected_rows) and selected_rows[0] > 0)
        self.move_down_button.setEnabled(
            not is_busy
            and bool(selected_rows)
            and selected_rows[-1] < self.files_list.count() - 1
        )
        self.clear_button.setEnabled(not is_busy and self.files_list.count() > 0)
        self.choose_output_dir_button.setEnabled(not is_busy)
        self.output_dir_edit.setEnabled(not is_busy)
        self.output_name_edit.setEnabled(not is_busy)
        self.merge_mode_combo.setEnabled(not is_busy)
        self.start_button.setEnabled(not is_busy)
        self.cancel_button.setEnabled(is_busy)
        self.files_list.setEnabled(not is_busy)
        self.open_folder_button.setEnabled(not is_busy and self._has_result_path())
        self.open_file_button.setEnabled(not is_busy and self._has_result_path())
        self.reveal_file_button.setEnabled(not is_busy and self._has_result_path())

    def _update_buttons_state(self) -> None:
        """Синхронизирует состояние кнопок с текущим выделением и режимом окна."""
        if self._worker is not None:
            self._set_busy_state(True)
            return

        selected_rows = self._selected_rows()
        has_selection = bool(selected_rows)
        has_files = self.files_list.count() > 0

        self.remove_file_button.setEnabled(has_selection)
        self.move_up_button.setEnabled(has_selection and selected_rows[0] > 0)
        self.move_down_button.setEnabled(
            has_selection and selected_rows[-1] < self.files_list.count() - 1
        )
        self.clear_button.setEnabled(has_files)
        self.cancel_button.setEnabled(False)
        self.file_count_label.setText(f"Файлов выбрано: {self.files_list.count()}")
        self._update_result_action_buttons()

    def _append_log(self, message: str) -> None:
        """Добавляет строку в текстовый журнал окна."""
        self.log_output.appendPlainText(message)

    def _append_log_from_logger(self, message: str) -> None:
        """Добавляет сообщение внутреннего логгера в интерфейс, избегая визуального конфликта."""
        self.log_output.appendPlainText(message)

    def _restore_merge_mode(self) -> None:
        """Восстанавливает последний выбранный режим объединения."""
        saved_merge_mode = self._settings.load_last_merge_mode()
        for index in range(self.merge_mode_combo.count()):
            if self.merge_mode_combo.itemData(index) == saved_merge_mode:
                self.merge_mode_combo.setCurrentIndex(index)
                return

    def _restore_recent_files(self) -> None:
        """Восстанавливает список последних файлов, которые еще доступны."""
        recent_files = self._settings.load_recent_files()
        if not recent_files:
            return

        restored_count = 0
        missing_count = 0
        for path in recent_files:
            if not path.exists() or not path.is_file() or path.suffix.lower() != ".docx":
                missing_count += 1
                continue

            item = QListWidgetItem(str(path))
            item.setToolTip(str(path))
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            self.files_list.addItem(item)
            restored_count += 1

        if restored_count:
            self._append_log(f"Восстановлено файлов из прошлой сессии: {restored_count}")
        if missing_count:
            self._append_log(f"Пропущено недоступных файлов из прошлой сессии: {missing_count}")

        self._persist_recent_files()
        self._update_buttons_state()

    def _persist_recent_files(self) -> None:
        """Сохраняет текущий список файлов в настройках."""
        self._settings.save_recent_files(
            [
                str(self.files_list.item(index).data(Qt.ItemDataRole.UserRole))
                for index in range(self.files_list.count())
            ]
        )
        self._settings.sync()

    def _has_result_path(self) -> bool:
        """Проверяет, что последний итоговый файл существует и доступен."""
        return self._last_output_path is not None and self._last_output_path.exists()

    def _update_result_action_buttons(self) -> None:
        """Синхронизирует кнопки открытия результата с текущим состоянием файла."""
        enabled = self._worker is None and self._has_result_path()
        self.open_folder_button.setEnabled(enabled)
        self.open_file_button.setEnabled(enabled)
        self.reveal_file_button.setEnabled(enabled)
        self.open_folder_action.setEnabled(enabled)
        self.open_file_action.setEnabled(enabled)
        self.reveal_file_action.setEnabled(enabled)

    def _open_result_folder(self) -> None:
        """Открывает папку с итоговым файлом."""
        if self._last_output_path is None:
            return
        try:
            open_directory(self._last_output_path.parent)
        except Exception as exc:
            self._show_runtime_error(exc)

    def _open_result_file(self) -> None:
        """Открывает итоговый файл приложением по умолчанию."""
        if self._last_output_path is None:
            return
        try:
            open_file(self._last_output_path)
        except Exception as exc:
            self._show_runtime_error(exc)

    def _reveal_result_file(self) -> None:
        """Показывает итоговый файл в Finder или Проводнике."""
        if self._last_output_path is None:
            return
        try:
            reveal_in_file_manager(self._last_output_path)
        except Exception as exc:
            self._show_runtime_error(exc)

    def _show_runtime_error(self, error: Exception) -> None:
        """Показывает пользователю понятную ошибку и пишет детали в лог."""
        title, message = build_user_friendly_error(error)
        self._logger.exception("Пользовательская операция завершилась ошибкой.")
        self._append_log(f"{title}: {message}")
        show_error(self, title, message)

    def _reset_window_layout(self) -> None:
        """Сбрасывает размер и положение окна к базовым значениям."""
        self.resize(920, 620)
        screen = self.screen() or self.windowHandle().screen() if self.windowHandle() else None
        if screen is not None:
            available = screen.availableGeometry()
            frame = self.frameGeometry()
            frame.moveCenter(available.center())
            self.move(frame.topLeft())
        self._append_log("Размер и положение окна сброшены.")

    def _clear_recent_files_and_list(self) -> None:
        """Очищает список последних файлов в интерфейсе и настройках."""
        self.files_list.clear()
        self._settings.clear_recent_files()
        self._settings.sync()
        self._append_log("Список последних файлов очищен.")
        self._update_buttons_state()
