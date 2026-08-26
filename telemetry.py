from __future__ import annotations

import json
import os
import platform
import threading
import time
import urllib.request
import uuid
from datetime import datetime, timezone

from PySide6.QtCore import QSettings, QTimer
from PySide6.QtWidgets import QMessageBox, QWidget

from telemetry_config import TELEMETRY_ENDPOINT
from version import APP_VERSION


class TelemetryClient:
    """Small, opt-in, anonymous usage reporter.

    No document metadata, paths, account data, or hardware identifiers are
    collected. Network failures never interrupt the reader.
    """

    def __init__(self) -> None:
        self.settings = QSettings("MultiDocSync", "MultiDocSync")
        self.endpoint = os.environ.get("MULTIDOCSYNC_TELEMETRY_ENDPOINT", TELEMETRY_ENDPOINT).strip()
        self.session_id = str(uuid.uuid4())
        self.started_at = time.monotonic()
        self.timer = QTimer()
        self.timer.setInterval(5 * 60 * 1000)
        self.timer.timeout.connect(lambda: self.send("heartbeat"))

    @property
    def has_choice(self) -> bool:
        return self.settings.contains("telemetry/consent")

    @property
    def enabled(self) -> bool:
        return bool(self.endpoint) and self.settings.value(
            "telemetry/consent", False, type=bool
        )

    def ask_first_run(self, parent: QWidget) -> None:
        if not self.endpoint or self.has_choice:
            self.start_if_enabled()
            return
        answer = QMessageBox.question(
            parent,
            "匿名使用统计",
            "是否允许发送匿名使用统计？\n\n"
            "仅发送：随机安装编号、应用版本、操作系统、启动次数和使用时长。\n"
            "不会发送文件名、文件路径、文档内容、用户名或硬件编号。\n"
            "您以后可以在“隐私”中随时关闭。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        self.settings.setValue("telemetry/consent", answer == QMessageBox.StandardButton.Yes)
        self.start_if_enabled()

    def show_privacy_settings(self, parent: QWidget) -> None:
        if not self.endpoint:
            QMessageBox.information(
                parent,
                "匿名使用统计",
                "当前版本尚未配置统计服务器，不会发送任何使用数据。",
            )
            return
        current = "已开启" if self.enabled else "已关闭"
        answer = QMessageBox.question(
            parent,
            "匿名使用统计",
            f"当前状态：{current}\n\n"
            "统计仅包含随机安装编号、版本、系统和会话时长；不包含任何文档信息。\n\n"
            "选择“是”开启，选择“否”关闭。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes if self.enabled else QMessageBox.StandardButton.No,
        )
        new_value = answer == QMessageBox.StandardButton.Yes
        was_enabled = self.enabled
        self.settings.setValue("telemetry/consent", new_value)
        if new_value and not was_enabled:
            self.start_if_enabled()
        elif not new_value:
            self.timer.stop()

    def start_if_enabled(self) -> None:
        if not self.enabled:
            return
        if not self.settings.value("telemetry/install_id", ""):
            self.settings.setValue("telemetry/install_id", str(uuid.uuid4()))
        self.send("session_start")
        self.timer.start()

    def stop(self) -> None:
        self.timer.stop()
        if self.enabled:
            self.send("session_end", timeout=1.5, background=False)

    def send(self, event: str, timeout: float = 3.0, background: bool = True) -> None:
        if not self.enabled:
            return
        install_id = self.settings.value("telemetry/install_id", "")
        if not install_id:
            return
        payload = {
            "event": event,
            "install_id": install_id,
            "session_id": self.session_id,
            "version": APP_VERSION,
            "platform": platform.system().lower(),
            "architecture": platform.machine().lower(),
            "elapsed_seconds": max(0, round(time.monotonic() - self.started_at)),
            "sent_at": datetime.now(timezone.utc).isoformat(),
        }

        def transmit() -> None:
            try:
                request = urllib.request.Request(
                    self.endpoint,
                    data=json.dumps(payload, separators=(",", ":")).encode("utf-8"),
                    headers={
                        "Content-Type": "application/json",
                        "User-Agent": f"MultiDocSync/{APP_VERSION}",
                    },
                    method="POST",
                )
                with urllib.request.urlopen(request, timeout=timeout) as response:
                    response.read(64)
            except Exception:
                pass

        if background:
            threading.Thread(target=transmit, daemon=True).start()
        else:
            transmit()
