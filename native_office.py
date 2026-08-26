from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QObject, QTimer, Qt, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)


WORD_EXTENSIONS = {".doc", ".docx"}
POWERPOINT_EXTENSIONS = {".ppt", ".pptx"}


class OfficeDependencyError(RuntimeError):
    pass


def document_family(paths: list[str | Path]) -> str | None:
    suffixes = {Path(path).suffix.lower() for path in paths}
    if suffixes and suffixes <= WORD_EXTENSIONS:
        return "word"
    if suffixes and suffixes <= POWERPOINT_EXTENSIONS:
        return "powerpoint"
    if suffixes == {".pdf"}:
        return "pdf"
    return None


class BaseOfficeBackend:
    state_suffix = ""

    def __init__(self, family: str, paths: list[Path]) -> None:
        self.family = family
        self.paths = paths
        self.linked = [True] * len(paths)
        self.last_states: list[int | None] = [None] * len(paths)
        self.last_active_index = 0

    def open(self, progress: Callable[[str], None]) -> None:
        raise NotImplementedError

    def arrange(self, left: int, top: int, width: int, height: int, dpi: float) -> None:
        raise NotImplementedError

    def poll_and_sync(self) -> list[int | None]:
        return self.last_states

    def nudge(self, direction: int) -> None:
        raise NotImplementedError

    def close(self, close_documents: bool) -> None:
        raise NotImplementedError

    def state_text(self, state: int | None, index: int) -> str:
        if state is None:
            return "窗口不可用"
        return f"{state}{self.state_suffix}"


class WindowsOfficeBackend(BaseOfficeBackend):
    def __init__(self, family: str, paths: list[Path]) -> None:
        super().__init__(family, paths)
        self.app = None
        self.documents: list[object] = []
        self.windows: list[object] = []
        self.pythoncom = None
        self.win32gui = None
        self._com_initialized = False
        self._syncing = False
        self._pending_sync: tuple[int, int, float] | None = None
        self.state_suffix = "%" if family == "word" else "页"

    def open(self, progress: Callable[[str], None]) -> None:
        try:
            import pythoncom
            import win32com.client
            import win32gui
        except ImportError as exc:
            raise OfficeDependencyError(
                "Windows Office控制组件缺失。请重新下载完整版，或安装 pywin32。"
            ) from exc

        self.pythoncom = pythoncom
        self.win32gui = win32gui
        pythoncom.CoInitialize()
        self._com_initialized = True
        prog_id = "Word.Application" if self.family == "word" else "PowerPoint.Application"
        product_name = "Microsoft Word" if self.family == "word" else "Microsoft PowerPoint"
        try:
            self.app = win32com.client.DispatchEx(prog_id)
        except Exception as exc:
            self._release_com()
            raise OfficeDependencyError(
                f"未检测到 {product_name}。请先安装 Microsoft Office 后再打开这些文件。"
            ) from exc

        try:
            self.app.Visible = True
            try:
                self.app.DisplayAlerts = 0 if self.family == "word" else 1
            except Exception:
                pass
            try:
                self.app.AutomationSecurity = 3
            except Exception:
                pass
            if self.family == "word":
                try:
                    self.app.Options.UpdateLinksAtOpen = False
                except Exception:
                    pass

            for index, path in enumerate(self.paths):
                progress(f"正在用原程序打开 {index + 1}/{len(self.paths)}：{path.name}")
                QApplication.processEvents()
                if self.family == "word":
                    document = self.app.Documents.Open(
                        str(path),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Revert=False,
                        Visible=True,
                        OpenAndRepair=True,
                        NoEncodingDialog=True,
                        PasswordDocument="",
                        WritePasswordDocument="",
                    )
                    window = document.ActiveWindow
                else:
                    document = self.app.Presentations.Open(
                        str(path), ReadOnly=True, Untitled=False, WithWindow=True
                    )
                    window = document.Windows(1)
                self.documents.append(document)
                self.windows.append(window)
            self.last_states = [self._read_state(i) for i in range(len(self.windows))]
        except Exception:
            self.close(close_documents=True)
            raise

    def arrange(self, left: int, top: int, width: int, height: int, dpi: float) -> None:
        if not self.windows:
            return
        scale = 72.0 / max(72.0, dpi)
        toolbar_height = 92
        usable_top = top + toolbar_height
        usable_height = max(300, height - toolbar_height)
        column_width = max(280, width // len(self.windows))
        for index, window in enumerate(self.windows):
            try:
                window.WindowState = 0 if self.family == "word" else 1
                window.Left = round((left + index * column_width) * scale)
                window.Top = round(usable_top * scale)
                target_width = column_width if index < len(self.windows) - 1 else width - index * column_width
                window.Width = round(target_width * scale)
                window.Height = round(usable_height * scale)
                if self.family == "word":
                    try:
                        window.View.Zoom.PageFit = 2  # wdPageFitBestFit
                    except Exception:
                        pass
            except Exception:
                continue

    def _window_handle(self, index: int) -> int | None:
        try:
            return int(self.windows[index].Hwnd)
        except Exception:
            try:
                return int(self.windows[index].HWND)
            except Exception:
                return None

    def _read_state(self, index: int) -> int | None:
        try:
            window = self.windows[index]
            if self.family == "word":
                return int(window.VerticalPercentScrolled)
            return int(window.View.Slide.SlideIndex)
        except Exception:
            return None

    def _write_state(self, index: int, state: int) -> None:
        window = self.windows[index]
        if self.family == "word":
            window.Activate()
            window.VerticalPercentScrolled = max(0, min(100, int(state)))
        else:
            maximum = int(self.documents[index].Slides.Count)
            window.View.GotoSlide(max(1, min(maximum, int(state))))

    def _active_index(self) -> int | None:
        if self.win32gui is None:
            return None
        foreground = int(self.win32gui.GetForegroundWindow())
        for index in range(len(self.windows)):
            if self._window_handle(index) == foreground:
                self.last_active_index = index
                return index
        return None

    def poll_and_sync(self) -> list[int | None]:
        if self._syncing:
            return self.last_states
        states = [self._read_state(i) for i in range(len(self.windows))]
        active = self._active_index()
        changed = [
            i
            for i, (old, new) in enumerate(zip(self.last_states, states))
            if old is not None and new is not None and old != new
        ]
        source = active if active in changed else (changed[0] if len(changed) == 1 else None)
        now = time.monotonic()
        if source is not None and self.linked[source] and states[source] is not None:
            if self.family == "word":
                # Word only applies scroll assignments reliably to its active
                # window. Debounce until the user pauses, then update peers and
                # restore focus to the source window to avoid focus thrashing.
                self._pending_sync = (source, int(states[source]), now)
            else:
                self._sync_linked(source, int(states[source]))
                states = [self._read_state(i) for i in range(len(self.windows))]
        elif self.family == "word" and self._pending_sync is not None:
            pending_source, pending_state, changed_at = self._pending_sync
            if now - changed_at >= 0.35:
                self._sync_linked(pending_source, pending_state)
                self._pending_sync = None
                states = [self._read_state(i) for i in range(len(self.windows))]
        self.last_states = states
        return states

    def _sync_linked(self, source: int, state: int) -> None:
        if not self.linked[source]:
            return
        self._syncing = True
        try:
            for index in range(len(self.windows)):
                if index != source and self.linked[index] and self._read_state(index) is not None:
                    self._write_state(index, state)
            if self.family == "word":
                try:
                    self.windows[source].Activate()
                except Exception:
                    pass
        finally:
            self._syncing = False

    def nudge(self, direction: int) -> None:
        if not self.windows:
            return
        source = min(self.last_active_index, len(self.windows) - 1)
        try:
            if self.family == "word":
                if direction < 0:
                    self.windows[source].LargeScroll(Up=1)
                else:
                    self.windows[source].LargeScroll(Down=1)
            else:
                current = self._read_state(source) or 1
                self._write_state(source, current + direction)
            self.last_states[source] = self._read_state(source)
            state = self.last_states[source]
            if self.linked[source] and state is not None:
                for index in range(len(self.windows)):
                    if index != source and self.linked[index]:
                        self._write_state(index, state)
                self.last_states = [self._read_state(i) for i in range(len(self.windows))]
        except Exception:
            pass

    def close(self, close_documents: bool) -> None:
        if close_documents:
            for document in reversed(self.documents):
                try:
                    if self.family == "word":
                        document.Close(SaveChanges=0)
                    else:
                        document.Close()
                except Exception:
                    pass
            if self.app is not None:
                try:
                    if self.family == "word":
                        self.app.Quit(SaveChanges=0)
                    else:
                        self.app.Quit()
                except Exception:
                    pass
        self.documents.clear()
        self.windows.clear()
        self.app = None
        self._release_com()

    def _release_com(self) -> None:
        if self._com_initialized and self.pythoncom is not None:
            try:
                self.pythoncom.CoUninitialize()
            except Exception:
                pass
        self._com_initialized = False


class MacOfficeBackend(BaseOfficeBackend):
    """Native Mac launcher and window tiler.

    Automatic Office scroll observation is not exposed consistently by Office for
    Mac. The controller still opens and tiles originals, while synchronized
    previous/next commands use macOS accessibility keystrokes.
    """

    def __init__(self, family: str, paths: list[Path]) -> None:
        super().__init__(family, paths)
        self.app_name = "Microsoft Word" if family == "word" else "Microsoft PowerPoint"
        self.process_name = self.app_name
        self.state_suffix = ""

    def open(self, progress: Callable[[str], None]) -> None:
        app_path = Path("/Applications") / f"{self.app_name}.app"
        if not app_path.exists():
            raise OfficeDependencyError(
                f"未检测到 {self.app_name}。请先安装 Microsoft Office for Mac。"
            )
        progress(f"正在用 {self.app_name} 打开 {len(self.paths)} 个原文件…")
        completed = subprocess.run(
            ["open", "-a", self.app_name, *[str(path) for path in self.paths]],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or "无法启动Office应用")
        time.sleep(2)
        self.last_states = [None] * len(self.paths)

    @staticmethod
    def _apple_string(value: str) -> str:
        return value.replace("\\", "\\\\").replace('"', '\\"')

    def arrange(self, left: int, top: int, width: int, height: int, dpi: float) -> None:
        toolbar_height = 92
        usable_height = max(300, height - toolbar_height)
        column_width = max(280, width // len(self.paths))
        commands = []
        for index, path in enumerate(self.paths):
            x = left + index * column_width
            w = column_width if index < len(self.paths) - 1 else width - index * column_width
            title = self._apple_string(path.stem)
            commands.append(
                f'''repeat with candidateWindow in windows
                    if name of candidateWindow contains "{title}" then
                        set position of candidateWindow to {{{x}, {top + toolbar_height}}}
                        set size of candidateWindow to {{{w}, {usable_height}}}
                        exit repeat
                    end if
                end repeat'''
            )
        script = f'''tell application "System Events"
            tell process "{self.process_name}"
                {os.linesep.join(commands)}
            end tell
        end tell'''
        completed = subprocess.run(
            ["osascript", "-e", script], capture_output=True, text=True, timeout=20
        )
        if completed.returncode != 0:
            raise OfficeDependencyError(
                "macOS需要“辅助功能”权限才能自动排列窗口。请在“系统设置 → 隐私与安全性 → 辅助功能”中允许本程序。"
            )

    def nudge(self, direction: int) -> None:
        key_code = 116 if direction < 0 else 121  # Page Up / Page Down
        # Apply the navigation key to matching windows one by one.
        titles = [self._apple_string(path.stem) for i, path in enumerate(self.paths) if self.linked[i]]
        clauses = []
        for title in titles:
            clauses.append(
                f'''repeat with candidateWindow in windows
                    if name of candidateWindow contains "{title}" then
                        perform action "AXRaise" of candidateWindow
                        key code {key_code}
                        exit repeat
                    end if
                end repeat'''
            )
        script = f'''tell application "System Events"
            tell process "{self.process_name}"
                set frontmost to true
                {os.linesep.join(clauses)}
            end tell
        end tell'''
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=20)

    def close(self, close_documents: bool) -> None:
        if not close_documents:
            return
        titles = [self._apple_string(path.stem) for path in self.paths]
        clauses = []
        for title in titles:
            clauses.append(
                f'''repeat with candidateWindow in windows
                    if name of candidateWindow contains "{title}" then
                        click button 1 of candidateWindow
                        exit repeat
                    end if
                end repeat'''
            )
        script = f'''tell application "System Events"
            tell process "{self.process_name}"
                {os.linesep.join(clauses)}
            end tell
        end tell'''
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=20)


def create_backend(family: str, paths: list[Path]) -> BaseOfficeBackend:
    if sys.platform == "win32":
        return WindowsOfficeBackend(family, paths)
    if sys.platform == "darwin":
        return MacOfficeBackend(family, paths)
    raise OfficeDependencyError("原生Word/PPT模式目前支持 Windows 和 macOS。")


class NativeOfficeController(QDialog):
    finished = Signal()

    def __init__(self, family: str, paths: list[Path], parent=None) -> None:
        super().__init__(parent)
        self.family = family
        self.paths = paths
        self.backend = create_backend(family, paths)
        self._closing_accepted = False
        self.setWindowTitle("MultiDoc Sync 控制栏")
        self.setWindowFlags(
            Qt.WindowType.Tool
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowCloseButtonHint
        )
        self.setMinimumWidth(760)

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        self.status_label = QLabel("准备打开原文件…")
        self.status_label.setStyleSheet("font-weight: 600; color: #17324d;")
        top.addWidget(self.status_label, 1)
        self.previous_button = QPushButton("上一页")
        self.next_button = QPushButton("下一页")
        self.arrange_button = QPushButton("重新排列")
        self.all_link_button = QPushButton("全部独立")
        self.end_button = QPushButton("结束")
        top.addWidget(self.previous_button)
        top.addWidget(self.next_button)
        top.addWidget(self.arrange_button)
        top.addWidget(self.all_link_button)
        top.addWidget(self.end_button)
        root.addLayout(top)

        rows = QHBoxLayout()
        self.link_boxes: list[QCheckBox] = []
        self.state_labels: list[QLabel] = []
        for index, path in enumerate(paths):
            item = QVBoxLayout()
            name = QLabel(path.name)
            name.setToolTip(str(path))
            name.setMaximumWidth(250)
            box = QCheckBox(f"第 {index + 1} 栏联动")
            box.setChecked(True)
            state = QLabel("准备中")
            box.toggled.connect(lambda checked, i=index: self._set_linked(i, checked))
            item.addWidget(name)
            item.addWidget(box)
            item.addWidget(state)
            rows.addLayout(item, 1)
            self.link_boxes.append(box)
            self.state_labels.append(state)
        root.addLayout(rows)

        self.previous_button.clicked.connect(lambda: self.backend.nudge(-1))
        self.next_button.clicked.connect(lambda: self.backend.nudge(1))
        self.arrange_button.clicked.connect(self.arrange_windows)
        self.all_link_button.clicked.connect(self.toggle_all)
        self.end_button.clicked.connect(self.close)

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(250)
        self.poll_timer.timeout.connect(self.poll)
        QTimer.singleShot(0, self.start_session)

    def start_session(self) -> None:
        try:
            self.backend.open(self.status_label.setText)
            self.arrange_windows()
            self.poll_timer.start()
            family_name = "Word" if self.family == "word" else "PowerPoint"
            if sys.platform == "darwin":
                self.status_label.setText(
                    f"{family_name} 原文件已并排；Mac版请用控制栏同步翻页"
                )
            else:
                self.status_label.setText(f"{family_name} 原文件已并排并启用联动")
        except OfficeDependencyError as exc:
            QMessageBox.warning(self, "缺少依赖", str(exc))
            self._closing_accepted = True
            self.close()
        except Exception as exc:
            QMessageBox.critical(self, "打开失败", f"无法打开原文件：\n{exc}")
            self._closing_accepted = True
            self.close()

    def arrange_windows(self) -> None:
        screen = self.screen() or QApplication.primaryScreen()
        geometry = screen.availableGeometry()
        try:
            self.backend.arrange(
                geometry.left(),
                geometry.top(),
                geometry.width(),
                geometry.height(),
                screen.logicalDotsPerInch(),
            )
            self.move(geometry.left() + 12, geometry.top() + 8)
            self.resize(min(geometry.width() - 24, 1100), self.sizeHint().height())
        except OfficeDependencyError as exc:
            QMessageBox.warning(self, "需要系统权限", str(exc))

    def poll(self) -> None:
        states = self.backend.poll_and_sync()
        for index, state in enumerate(states):
            self.state_labels[index].setText(self.backend.state_text(state, index))

    def _set_linked(self, index: int, checked: bool) -> None:
        self.backend.linked[index] = checked
        self._refresh_all_link_button()

    def toggle_all(self) -> None:
        new_state = not all(box.isChecked() for box in self.link_boxes)
        for box in self.link_boxes:
            box.setChecked(new_state)
        self._refresh_all_link_button()

    def _refresh_all_link_button(self) -> None:
        self.all_link_button.setText(
            "全部独立" if all(box.isChecked() for box in self.link_boxes) else "全部联动"
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._closing_accepted:
            self.poll_timer.stop()
            self.backend.close(close_documents=True)
            event.accept()
            self.finished.emit()
            return
        answer = QMessageBox.question(
            self,
            "结束并排阅读",
            "是否同时关闭本次打开的Office文件？\n\n选择“否”会保留原文件窗口。",
            QMessageBox.StandardButton.Yes
            | QMessageBox.StandardButton.No
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Yes,
        )
        if answer == QMessageBox.StandardButton.Cancel:
            event.ignore()
            return
        self.poll_timer.stop()
        self.backend.close(close_documents=answer == QMessageBox.StandardButton.Yes)
        event.accept()
        self.finished.emit()
