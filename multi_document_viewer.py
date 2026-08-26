from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path

import fitz
from PySide6.QtCore import QEvent, QRunnable, QThreadPool, QTimer, Qt, Signal, QObject
from PySide6.QtGui import QAction, QCloseEvent, QDragEnterEvent, QDropEvent, QIcon, QImage, QKeySequence, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedLayout,
    QVBoxLayout,
    QWidget,
)

from native_office import NativeOfficeController, document_family
from telemetry import TelemetryClient
from version import APP_VERSION


APP_TITLE = "MultiDoc Sync｜多文档联动"
PROJECT_URL = "https://github.com/shenquanzhen/multidoc-sync"
SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx", ".ppt", ".pptx"}
FILE_FILTER = "支持的文档 (*.pdf *.doc *.docx *.ppt *.pptx);;所有文件 (*.*)"


def is_supported(path: str | Path) -> bool:
    return Path(path).suffix.lower() in SUPPORTED_EXTENSIONS


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


class WorkerSignals(QObject):
    finished = Signal(int, int, str, object, object)


class PdfOpenWorker(QRunnable):
    def __init__(
        self,
        pane_index: int,
        token: int,
        source: Path,
    ) -> None:
        super().__init__()
        self.pane_index = pane_index
        self.token = token
        self.source = source
        self.signals = WorkerSignals()

    def run(self) -> None:
        try:
            if self.source.suffix.lower() != ".pdf":
                raise RuntimeError("Word/PPT应通过原程序模式打开，不进行格式转换")
            pdf_path = self.source
            self.signals.finished.emit(
                self.pane_index, self.token, str(self.source), str(pdf_path), None
            )
        except Exception as exc:
            message = str(exc).strip() or exc.__class__.__name__
            self.signals.finished.emit(
                self.pane_index, self.token, str(self.source), None, message
            )


class PageScrollArea(QScrollArea):
    boundary_page = Signal(int)
    interacted = Signal()

    def wheelEvent(self, event) -> None:  # noqa: N802 - Qt naming
        self.interacted.emit()
        bar = self.verticalScrollBar()
        direction = -1 if event.angleDelta().y() > 0 else 1
        at_boundary = (
            (direction < 0 and bar.value() <= bar.minimum())
            or (direction > 0 and bar.value() >= bar.maximum())
        )
        if at_boundary:
            self.boundary_page.emit(direction)
            event.accept()
            return
        super().wheelEvent(event)


class DocumentPane(QFrame):
    view_changed = Signal(object, int, float)
    activated = Signal(object)
    link_changed = Signal()
    replace_requested = Signal(object)
    close_requested = Signal(object)
    files_dropped = Signal(object, object)

    def __init__(self, index: int, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.index = index
        self.load_token = 0
        self.source_path: Path | None = None
        self.pdf_path: Path | None = None
        self.document: fitz.Document | None = None
        self.page_index = 0
        self._suppress_view_signal = False
        self._last_ratio = 0.0
        self._requested_beyond_end = False

        self.setObjectName("documentPane")
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setAcceptDrops(True)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(6, 6, 6, 6)
        outer.setSpacing(5)

        header = QHBoxLayout()
        self.name_label = QLabel(f"第 {index + 1} 栏")
        self.name_label.setObjectName("fileName")
        self.name_label.setMinimumWidth(50)
        self.name_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.page_label = QLabel("— / —")
        self.page_label.setObjectName("pageInfo")

        self.link_button = QPushButton("联动")
        self.link_button.setCheckable(True)
        self.link_button.setChecked(True)
        self.link_button.setToolTip("点击后让本栏独立滚动")
        self.link_button.clicked.connect(self._on_link_toggled)

        self.open_button = QPushButton("原程序")
        self.open_button.setToolTip("用Word、PowerPoint或默认PDF程序打开原文件")
        self.open_button.clicked.connect(self.open_original)
        self.open_button.setEnabled(False)

        self.replace_button = QPushButton("替换")
        self.replace_button.clicked.connect(lambda: self.replace_requested.emit(self))

        self.close_button = QPushButton("×")
        self.close_button.setObjectName("closeButton")
        self.close_button.setFixedWidth(30)
        self.close_button.setToolTip("清空本栏")
        self.close_button.clicked.connect(lambda: self.close_requested.emit(self))

        header.addWidget(self.name_label, 1)
        header.addWidget(self.page_label)
        header.addWidget(self.link_button)
        header.addWidget(self.open_button)
        header.addWidget(self.replace_button)
        header.addWidget(self.close_button)
        outer.addLayout(header)

        self.stack_host = QWidget()
        self.stack = QStackedLayout(self.stack_host)
        self.stack.setContentsMargins(0, 0, 0, 0)

        self.message_label = QLabel("将文件拖到这里\n或点击“替换”选择文件")
        self.message_label.setObjectName("emptyMessage")
        self.message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.message_label.setWordWrap(True)

        self.page_image = QLabel()
        self.page_image.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.page_image.setStyleSheet("background: #dedede;")
        self.scroll_area = PageScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.scroll_area.setWidget(self.page_image)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.verticalScrollBar().valueChanged.connect(self._on_scroll_value_changed)
        self.scroll_area.boundary_page.connect(self._on_boundary_page)
        self.scroll_area.interacted.connect(lambda: self.activated.emit(self))

        self.stack.addWidget(self.message_label)
        self.stack.addWidget(self.scroll_area)
        outer.addWidget(self.stack_host, 1)

        self.resize_timer = QTimer(self)
        self.resize_timer.setSingleShot(True)
        self.resize_timer.setInterval(140)
        self.resize_timer.timeout.connect(self._rerender_after_resize)
        self.show_empty()

    @property
    def linked(self) -> bool:
        return self.link_button.isChecked()

    @property
    def page_count(self) -> int:
        return self.document.page_count if self.document is not None else 0

    def set_linked(self, linked: bool) -> None:
        self.link_button.setChecked(linked)
        self._refresh_link_button()

    def _on_link_toggled(self) -> None:
        self._refresh_link_button()
        self.link_changed.emit()

    def _refresh_link_button(self) -> None:
        if self.linked:
            self.link_button.setText("联动")
            self.link_button.setToolTip("本栏跟随其他已联动栏；点击后独立滚动")
        else:
            self.link_button.setText("独立")
            self.link_button.setToolTip("本栏独立滚动；点击后重新加入联动")

    def show_loading(self, source: Path) -> None:
        self.clear_document(increment_token=False)
        self.source_path = source
        self.name_label.setText(source.name)
        self.name_label.setToolTip(str(source))
        self.page_label.setText("读取中…")
        self.message_label.setText(f"正在准备\n{source.name}")
        self.message_label.setProperty("state", "loading")
        self.message_label.style().unpolish(self.message_label)
        self.message_label.style().polish(self.message_label)
        self.stack.setCurrentWidget(self.message_label)
        self.open_button.setEnabled(True)

    def show_error(self, source: Path, message: str) -> None:
        self.source_path = source
        self.name_label.setText(source.name)
        self.name_label.setToolTip(str(source))
        self.page_label.setText("打开失败")
        self.message_label.setText(f"无法打开此文件\n\n{message}")
        self.message_label.setProperty("state", "error")
        self.message_label.style().unpolish(self.message_label)
        self.message_label.style().polish(self.message_label)
        self.stack.setCurrentWidget(self.message_label)
        self.open_button.setEnabled(True)

    def show_empty(self) -> None:
        self.name_label.setText(f"第 {self.index + 1} 栏")
        self.name_label.setToolTip("")
        self.page_label.setText("— / —")
        self.message_label.setText("将文件拖到这里\n或点击“替换”选择文件")
        self.message_label.setProperty("state", "empty")
        self.message_label.style().unpolish(self.message_label)
        self.message_label.style().polish(self.message_label)
        self.stack.setCurrentWidget(self.message_label)
        self.open_button.setEnabled(False)

    def load_pdf(self, source: Path, pdf_path: Path) -> None:
        self.clear_document(increment_token=False)
        try:
            document = fitz.open(str(pdf_path))
            if document.needs_pass:
                document.close()
                raise RuntimeError("文件已加密，需要密码，当前只读工具无法打开")
            if document.page_count < 1:
                document.close()
                raise RuntimeError("文件中没有可显示的页面")
        except Exception as exc:
            self.show_error(source, str(exc))
            return

        self.source_path = source
        self.pdf_path = pdf_path
        self.document = document
        self.page_index = 0
        self._last_ratio = 0.0
        self.name_label.setText(source.name)
        self.name_label.setToolTip(str(source))
        self.open_button.setEnabled(True)
        self.stack.setCurrentWidget(self.scroll_area)
        self.render_current_page(0.0, suppress_signal=True)

    def clear_document(self, increment_token: bool = True) -> None:
        if increment_token:
            self.load_token += 1
        if self.document is not None:
            try:
                self.document.close()
            except Exception:
                pass
        self.document = None
        self.source_path = None
        self.pdf_path = None
        self.page_image.clear()
        self.page_index = 0
        self._last_ratio = 0.0

    def close_and_reset(self) -> None:
        self.clear_document()
        self.show_empty()

    def open_original(self) -> None:
        if self.source_path is None:
            return
        try:
            os.startfile(str(self.source_path))
        except Exception as exc:
            QMessageBox.warning(self, APP_TITLE, f"无法调用原程序：\n{exc}")

    def set_page(self, requested_page: int, ratio: float = 0.0, suppress_signal: bool = False) -> None:
        if self.document is None:
            return
        requested_page = max(0, requested_page)
        target = min(requested_page, self.page_count - 1)
        self._requested_beyond_end = requested_page >= self.page_count
        self.page_index = target
        self.render_current_page(ratio, suppress_signal=suppress_signal)

    def apply_synced_view(self, page_index: int, ratio: float) -> None:
        self.set_page(page_index, ratio, suppress_signal=True)

    def render_current_page(self, ratio: float | None = None, suppress_signal: bool = False) -> None:
        if self.document is None:
            return
        if ratio is None:
            ratio = self.current_ratio()
        ratio = max(0.0, min(1.0, float(ratio)))
        self._last_ratio = ratio
        self._suppress_view_signal = self._suppress_view_signal or suppress_signal

        try:
            page = self.document.load_page(self.page_index)
            viewport_width = max(180, self.scroll_area.viewport().width() - 4)
            dpr = max(1.0, float(self.devicePixelRatioF()))
            logical_scale = viewport_width / max(1.0, page.rect.width)
            render_scale = max(0.25, min(4.0, logical_scale * dpr))
            pix = page.get_pixmap(matrix=fitz.Matrix(render_scale, render_scale), alpha=False)
            image = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format.Format_RGB888,
            ).copy()
            qt_pixmap = QPixmap.fromImage(image)
            qt_pixmap.setDevicePixelRatio(dpr)
            self.page_image.setPixmap(qt_pixmap)
            self.page_image.setFixedSize(qt_pixmap.deviceIndependentSize().toSize())
            self._update_page_label()
        except Exception as exc:
            source = self.source_path or Path("未知文件")
            self.show_error(source, f"第 {self.page_index + 1} 页渲染失败：{exc}")
            return

        QTimer.singleShot(0, lambda: self._restore_scroll(ratio, suppress_signal))

    def _restore_scroll(self, ratio: float, suppress_signal: bool) -> None:
        bar = self.scroll_area.verticalScrollBar()
        self._suppress_view_signal = self._suppress_view_signal or suppress_signal
        bar.setValue(round(bar.maximum() * ratio))
        self._last_ratio = ratio
        self._suppress_view_signal = False

    def _update_page_label(self) -> None:
        if self.document is None:
            self.page_label.setText("— / —")
            return
        suffix = "  已到末页" if self._requested_beyond_end else ""
        self.page_label.setText(f"{self.page_index + 1} / {self.page_count}{suffix}")
        self.page_label.setProperty("clamped", self._requested_beyond_end)
        self.page_label.style().unpolish(self.page_label)
        self.page_label.style().polish(self.page_label)

    def current_ratio(self) -> float:
        bar = self.scroll_area.verticalScrollBar()
        if bar.maximum() <= 0:
            return self._last_ratio
        return bar.value() / bar.maximum()

    def _on_scroll_value_changed(self) -> None:
        if self.document is None or self._suppress_view_signal:
            return
        self._last_ratio = self.current_ratio()
        self.activated.emit(self)
        self.view_changed.emit(self, self.page_index, self._last_ratio)

    def _on_boundary_page(self, direction: int) -> None:
        if self.document is None:
            return
        target = self.page_index + direction
        if target < 0 or target >= self.page_count:
            return
        ratio = 1.0 if direction < 0 else 0.0
        self.activated.emit(self)
        self.set_page(target, ratio, suppress_signal=False)
        self.view_changed.emit(self, self.page_index, ratio)

    def _rerender_after_resize(self) -> None:
        if self.document is None:
            return
        self.render_current_page(self.current_ratio(), suppress_signal=True)

    def resizeEvent(self, event) -> None:  # noqa: N802 - Qt naming
        super().resizeEvent(event)
        if self.document is not None:
            self.resize_timer.start()

    def event(self, event) -> bool:
        if event.type() == QEvent.Type.DevicePixelRatioChange and self.document is not None:
            self.resize_timer.start()
        return super().event(event)

    def enterEvent(self, event) -> None:  # noqa: N802 - Qt naming
        self.activated.emit(self)
        super().enterEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        paths = [url.toLocalFile() for url in event.mimeData().urls()]
        if paths and all(is_supported(path) for path in paths):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        paths = [path for path in paths if is_supported(path)]
        if paths:
            self.files_dropped.emit(self, paths)
            event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self, initial_paths: list[str]) -> None:
        super().__init__()
        self.thread_pool = QThreadPool.globalInstance()
        # A small limit avoids launching four heavy Office instances at once.
        self.thread_pool.setMaxThreadCount(2)
        self.panes: list[DocumentPane] = []
        self.active_pane: DocumentPane | None = None
        self._closing = False
        self._screen_signal_connected = False
        self.native_controller: NativeOfficeController | None = None
        self.telemetry = TelemetryClient()

        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(1100, 650)
        self.setAcceptDrops(True)

        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(7)

        toolbar = QHBoxLayout()
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        toolbar.addWidget(title)
        toolbar.addStretch(1)

        self.choose_button = QPushButton("选择3～4个文件")
        self.choose_button.clicked.connect(self.choose_files)
        toolbar.addWidget(self.choose_button)

        self.global_link_button = QPushButton("全部独立")
        self.global_link_button.setToolTip("当前全部联动；点击后让各栏独立滚动")
        self.global_link_button.clicked.connect(self.toggle_all_links)
        toolbar.addWidget(self.global_link_button)

        self.about_button = QPushButton("关于")
        self.about_button.clicked.connect(self.show_about)
        toolbar.addWidget(self.about_button)

        self.privacy_button = QPushButton("隐私")
        self.privacy_button.clicked.connect(lambda: self.telemetry.show_privacy_settings(self))
        toolbar.addWidget(self.privacy_button)

        self.previous_button = QPushButton("上一页")
        self.previous_button.clicked.connect(lambda: self.navigate_relative(-1))
        toolbar.addWidget(self.previous_button)
        self.next_button = QPushButton("下一页")
        self.next_button.clicked.connect(lambda: self.navigate_relative(1))
        toolbar.addWidget(self.next_button)

        toolbar.addWidget(QLabel("页码"))
        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, 1)
        self.page_spin.setFixedWidth(72)
        self.page_spin.setKeyboardTracking(False)
        self.page_spin.valueChanged.connect(self.jump_to_page)
        toolbar.addWidget(self.page_spin)
        self.active_total_label = QLabel("/ —")
        toolbar.addWidget(self.active_total_label)
        root.addLayout(toolbar)

        self.columns_widget = QWidget()
        self.columns_layout = QHBoxLayout(self.columns_widget)
        self.columns_layout.setContentsMargins(0, 0, 0, 0)
        self.columns_layout.setSpacing(7)
        root.addWidget(self.columns_widget, 1)
        self.setCentralWidget(central)

        self.statusBar().showMessage("可选择或拖入3～4个文件；每栏可单独切换联动/独立")
        self._install_shortcuts()
        self._apply_style()

        valid_initial = [path for path in initial_paths if is_supported(path)]
        initial_count = len(valid_initial) if len(valid_initial) in (3, 4) else 3
        self.rebuild_panes(initial_count)
        if initial_paths:
            QTimer.singleShot(0, lambda: self.load_batch(initial_paths))
        else:
            QTimer.singleShot(250, self.choose_files)
        QTimer.singleShot(600, lambda: self.telemetry.ask_first_run(self))

    def _install_shortcuts(self) -> None:
        open_action = QAction(self)
        open_action.setShortcut(QKeySequence.StandardKey.Open)
        open_action.triggered.connect(self.choose_files)
        self.addAction(open_action)

        previous_action = QAction(self)
        previous_action.setShortcut(QKeySequence(Qt.Key.Key_PageUp))
        previous_action.triggered.connect(lambda: self.navigate_relative(-1))
        self.addAction(previous_action)

        next_action = QAction(self)
        next_action.setShortcut(QKeySequence(Qt.Key.Key_PageDown))
        next_action.triggered.connect(lambda: self.navigate_relative(1))
        self.addAction(next_action)

    def show_about(self) -> None:
        QMessageBox.about(
            self,
            "关于 MultiDoc Sync",
            f"<h3>MultiDoc Sync {APP_VERSION}</h3>"
            "<p>同时并排查看3～4个同类文档，并控制联动或独立浏览。</p>"
            "<p>PDF直接显示；Word和PowerPoint使用本机原程序，不进行格式转换。</p>"
            f'<p>开源地址：<a href="{PROJECT_URL}">{PROJECT_URL}</a></p>',
        )

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background: #f4f6f8; color: #202428; font-family: "Microsoft YaHei UI"; font-size: 12px; }
            QLabel#appTitle { font-size: 18px; font-weight: 700; color: #17324d; }
            QFrame#documentPane { background: #ffffff; border: 1px solid #c8d0d8; border-radius: 7px; }
            QLabel#fileName { font-weight: 600; color: #17324d; }
            QLabel#pageInfo { color: #5b6570; }
            QLabel#pageInfo[clamped="true"] { color: #c05a18; font-weight: 600; }
            QLabel#emptyMessage { background: #eef2f5; color: #687582; border: 1px dashed #aeb9c3; border-radius: 5px; padding: 18px; }
            QLabel#emptyMessage[state="loading"] { color: #2369a6; }
            QLabel#emptyMessage[state="error"] { color: #a33a32; background: #fff1f0; border-color: #e2aaa5; }
            QPushButton { background: #ffffff; border: 1px solid #aeb8c2; border-radius: 4px; padding: 5px 9px; }
            QPushButton:hover { background: #e8f1fa; border-color: #6b99c3; }
            QPushButton:checked { background: #dceeff; border-color: #4386bf; color: #14558c; }
            QPushButton#closeButton { color: #a33a32; font-weight: 700; }
            QScrollArea { border: 0; background: #dedede; }
            QStatusBar { background: #edf1f4; color: #506070; }
            """
        )

    def rebuild_panes(self, count: int) -> None:
        count = 4 if count == 4 else 3
        for pane in self.panes:
            pane.clear_document()
            pane.setParent(None)
            pane.deleteLater()
        self.panes.clear()
        self.active_pane = None

        while self.columns_layout.count():
            item = self.columns_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for index in range(count):
            pane = DocumentPane(index)
            pane.view_changed.connect(self.on_view_changed)
            pane.activated.connect(self.set_active_pane)
            pane.link_changed.connect(self.update_global_link_button)
            pane.replace_requested.connect(self.replace_pane)
            pane.close_requested.connect(self.close_pane)
            pane.files_dropped.connect(self.on_pane_files_dropped)
            self.columns_layout.addWidget(pane, 1)
            self.panes.append(pane)
        self.update_global_link_button()
        self.update_navigation_controls()

    def choose_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择3～4个文档", "", FILE_FILTER)
        if not paths:
            return
        self.load_batch(paths)

    def load_batch(self, paths: list[str]) -> None:
        paths = [str(Path(path)) for path in paths]
        if len(paths) not in (3, 4):
            QMessageBox.information(self, APP_TITLE, "请一次选择3个或4个文件。")
            return
        invalid = [path for path in paths if not is_supported(path)]
        if invalid:
            QMessageBox.warning(
                self,
                APP_TITLE,
                "以下文件格式不受支持：\n" + "\n".join(invalid),
            )
            return
        family = document_family(paths)
        if family is None:
            QMessageBox.information(
                self,
                APP_TITLE,
                "请一次选择同一种格式：\n"
                "• 全部为PDF\n"
                "• 全部为Word（DOC/DOCX可混用）\n"
                "• 全部为PowerPoint（PPT/PPTX可混用）",
            )
            return
        if family in {"word", "powerpoint"}:
            self.start_native_office_session(family, [Path(path).resolve() for path in paths])
            return
        if len(self.panes) != len(paths):
            self.rebuild_panes(len(paths))
        for pane, path in zip(self.panes, paths):
            self.load_path_into_pane(pane, Path(path))
        self.statusBar().showMessage(f"正在准备 {len(paths)} 个文件…")

    def start_native_office_session(self, family: str, paths: list[Path]) -> None:
        if self.native_controller is not None:
            self.native_controller.close()
        self.native_controller = NativeOfficeController(family, paths)
        self.native_controller.finished.connect(self._native_session_finished)
        self.native_controller.show()
        self.hide()

    def _native_session_finished(self) -> None:
        self.native_controller = None
        self.showMaximized()
        self.raise_()
        self.activateWindow()
        self.statusBar().showMessage("原生Office并排会话已结束，可重新选择文件")

    def load_path_into_pane(self, pane: DocumentPane, source: Path) -> None:
        source = source.expanduser().resolve()
        pane.load_token += 1
        token = pane.load_token
        pane.show_loading(source)
        worker = PdfOpenWorker(pane.index, token, source)
        worker.signals.finished.connect(self.on_conversion_finished)
        self.thread_pool.start(worker)

    def on_conversion_finished(
        self,
        pane_index: int,
        token: int,
        source_text: str,
        pdf_text: object,
        error: object,
    ) -> None:
        if self._closing or pane_index >= len(self.panes):
            return
        pane = self.panes[pane_index]
        if token != pane.load_token:
            return
        source = Path(source_text)
        if error is not None or pdf_text is None:
            pane.show_error(source, str(error))
        else:
            pane.load_pdf(source, Path(str(pdf_text)))
            if self.active_pane is None:
                self.set_active_pane(pane)
        loaded = sum(1 for item in self.panes if item.document is not None)
        errors = sum(
            1 for item in self.panes if item.source_path is not None and item.document is None
        )
        self.statusBar().showMessage(
            f"已打开 {loaded} 个文件" + (f"，{errors} 个失败" if errors else "")
        )
        self.update_navigation_controls()

    def on_view_changed(self, source: DocumentPane, page_index: int, ratio: float) -> None:
        self.set_active_pane(source)
        if source.linked:
            for pane in self.panes:
                if pane is not source and pane.linked and pane.document is not None:
                    pane.apply_synced_view(page_index, ratio)
        self.update_navigation_controls()

    def set_active_pane(self, pane: DocumentPane) -> None:
        if pane.document is not None:
            self.active_pane = pane
            self.update_navigation_controls()

    def update_navigation_controls(self) -> None:
        pane = self.active_pane
        if pane is None or pane.document is None:
            pane = next((item for item in self.panes if item.document is not None), None)
        enabled = pane is not None
        self.previous_button.setEnabled(enabled)
        self.next_button.setEnabled(enabled)
        self.page_spin.setEnabled(enabled)
        if not enabled or pane is None:
            self.page_spin.blockSignals(True)
            self.page_spin.setRange(1, 1)
            self.page_spin.setValue(1)
            self.page_spin.blockSignals(False)
            self.active_total_label.setText("/ —")
            return
        self.active_pane = pane
        self.page_spin.blockSignals(True)
        self.page_spin.setRange(1, pane.page_count)
        self.page_spin.setValue(pane.page_index + 1)
        self.page_spin.blockSignals(False)
        self.active_total_label.setText(f"/ {pane.page_count}")

    def navigate_relative(self, delta: int) -> None:
        pane = self._navigation_anchor()
        if pane is None:
            return
        target = max(0, min(pane.page_count - 1, pane.page_index + delta))
        pane.set_page(target, 0.0, suppress_signal=False)
        pane.view_changed.emit(pane, target, 0.0)

    def jump_to_page(self, one_based_page: int) -> None:
        pane = self._navigation_anchor()
        if pane is None:
            return
        target = max(0, one_based_page - 1)
        pane.set_page(target, 0.0, suppress_signal=False)
        pane.view_changed.emit(pane, target, 0.0)

    def _navigation_anchor(self) -> DocumentPane | None:
        if self.active_pane is not None and self.active_pane.document is not None:
            return self.active_pane
        return next((pane for pane in self.panes if pane.document is not None), None)

    def toggle_all_links(self) -> None:
        populated = [pane for pane in self.panes if pane.source_path is not None]
        targets = populated or self.panes
        all_linked = all(pane.linked for pane in targets)
        new_state = not all_linked
        for pane in self.panes:
            pane.set_linked(new_state)
        self.update_global_link_button()

    def update_global_link_button(self) -> None:
        if self.panes and all(pane.linked for pane in self.panes):
            self.global_link_button.setText("全部独立")
            self.global_link_button.setToolTip("当前全部联动；点击后让各栏独立滚动")
        else:
            self.global_link_button.setText("全部联动")
            self.global_link_button.setToolTip("点击后让所有栏按相同页码联动")

    def replace_pane(self, pane: DocumentPane) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择替换文件", "", FILE_FILTER)
        if path:
            self.load_path_into_pane(pane, Path(path))

    def close_pane(self, pane: DocumentPane) -> None:
        pane.close_and_reset()
        if self.active_pane is pane:
            self.active_pane = None
        self.update_navigation_controls()
        self.statusBar().showMessage("本栏已清空，可拖入或选择另一个文件")

    def on_pane_files_dropped(self, pane: DocumentPane, paths: list[str]) -> None:
        if len(paths) == 1:
            self.load_path_into_pane(pane, Path(paths[0]))
        elif len(paths) in (3, 4):
            self.load_batch(paths)
        else:
            QMessageBox.information(self, APP_TITLE, "请拖入1个替换文件，或一次拖入3～4个文件。")

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        paths = [url.toLocalFile() for url in event.mimeData().urls()]
        if paths and all(is_supported(path) for path in paths):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.load_batch(paths)
            event.acceptProposedAction()

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        self._closing = True
        self.telemetry.stop()
        for pane in self.panes:
            pane.clear_document()
        self.thread_pool.waitForDone(15000)
        event.accept()

    def showEvent(self, event) -> None:  # noqa: N802 - Qt naming
        super().showEvent(event)
        if not self._screen_signal_connected and self.windowHandle() is not None:
            self.windowHandle().screenChanged.connect(self._on_screen_changed)
            self._screen_signal_connected = True

    def _on_screen_changed(self, _screen=None) -> None:
        for pane in self.panes:
            if pane.document is not None:
                pane.resize_timer.start()


def parse_initial_paths(arguments: list[str]) -> list[str]:
    return [str(Path(arg).expanduser()) for arg in arguments if arg.strip()]


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("LocalTools")
    icon_path = resource_path("assets/MultiDocSync.png")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    initial_paths = parse_initial_paths(sys.argv[1:])
    window = MainWindow(initial_paths)
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        detail = traceback.format_exc()
        try:
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, APP_TITLE, f"程序启动失败：\n\n{detail}")
        finally:
            raise
