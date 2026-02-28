from __future__ import annotations

import logging
import traceback
from pathlib import Path

from PyQt6.QtCore import QObject, QThread, QTimer, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QMenu,
    QPushButton,
    QSlider,
    QSpinBox,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from media_player import MediaPlayerController
from pipewire_controller import PipeWireController, PipeWireError, PipeWireSnapshot
from state_manager import RouteAction, RoutingStateManager
from app_config import AppConfig, load_config, save_config


class PipeWireWorker(QObject):
    snapshot_ready = pyqtSignal(object)
    actions_failed = pyqtSignal(str)

    def __init__(self, controller: PipeWireController) -> None:
        super().__init__()
        self._controller = controller

    @pyqtSlot()
    def poll_snapshot(self) -> None:
        try:
            snapshot = self._controller.snapshot()
            self.snapshot_ready.emit(snapshot)
        except Exception as exc:
            self.actions_failed.emit(f"PipeWire refresh failed: {exc}")

    @pyqtSlot(object)
    def apply_actions(self, actions: object) -> None:
        if not isinstance(actions, list) or not actions:
            return

        try:
            snapshot = self._controller.snapshot()
        except Exception as exc:
            self.actions_failed.emit(f"Failed to prepare routing actions: {exc}")
            return

        errors: list[str] = []
        for action in actions:
            if not isinstance(action, RouteAction):
                continue
            try:
                if action.op == "link":
                    self._controller.create_link_by_key(action.source_key, action.target_key, snapshot)
                elif action.op == "unlink":
                    self._controller.remove_link_by_key(action.source_key, action.target_key, snapshot)
            except PipeWireError as exc:
                errors.append(str(exc))

        if errors:
            self.actions_failed.emit("; ".join(errors[:3]))

    @pyqtSlot(object)
    def set_virtual_mic_volume(self, desired_percent: object) -> None:
        if not isinstance(desired_percent, (int, float)):
            return
        try:
            snapshot = self._controller.snapshot()
            self._controller.apply_target_volume_percent_by_keys(
                [self._controller.virtual_mic_sink_key()],
                snapshot,
                float(desired_percent),
            )
        except PipeWireError as exc:
            self.actions_failed.emit(str(exc))


class CheckListWidget(QListWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)

    def mousePressEvent(self, event):  # type: ignore[override]
        item = self.itemAt(event.pos())
        if (
            item is not None
            and self.isEnabled()
            and event.button() == Qt.MouseButton.LeftButton
        ):
            item.setCheckState(
                Qt.CheckState.Unchecked
                if item.checkState() == Qt.CheckState.Checked
                else Qt.CheckState.Checked
            )
            event.accept()
            return
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    USER_ROLE = int(Qt.ItemDataRole.UserRole)
    # Internal exclude lists: add exact node keys (node.name values) here.
    EXCLUDED_SOURCE_KEYS: list[str] = [
        "input.audiolink_virtual_mic",
        "gsr-default_input",
        "gsr-default_output",
    ]
    EXCLUDED_TARGET_KEYS: list[str] = [
        "input.audiolink_virtual_mic",
        "gsr-default_output",
    ]

    request_poll = pyqtSignal()
    request_apply_actions = pyqtSignal(object)
    request_set_virtual_mic_volume = pyqtSignal(object)

    not_available_color = QColor(155, 80, 80)

    def __init__(self, controller: PipeWireController, media: MediaPlayerController) -> None:
        super().__init__()
        self._logger = logging.getLogger("audiolink.ui")
        self.controller = controller
        self.media = media
        self.state = RoutingStateManager()
        self.snapshot = PipeWireSnapshot(nodes={}, links=[])
        self._updating_lists = False
        self._poll_in_flight = False
        self._virtual_mic_sink_key = self.controller.virtual_mic_sink_key()
        self._virtual_mic_source_key = self.controller.virtual_mic_source_key()
        self._pending_virtual_mic_percent: float = 100.0
        self._cfg: AppConfig = load_config()
        self._auto_select_source_keys: set[str] = set(self._cfg.auto_select_sources)
        self._auto_select_target_keys: set[str] = set(self._cfg.auto_select_targets)
        self.state.selected_sources.update(self._auto_select_source_keys)
        self.state.selected_targets.update(self._auto_select_target_keys)

        self.setWindowTitle("AudioLink")
        self.resize(800, 600)
        self._logger.info("MainWindow initialized with size 800x600")

        self._build_ui()
        self._build_worker()
        self.controller.ensure_virtual_microphone()

        self._timer = QTimer(self)
        self._timer.setInterval(500)
        self._timer.timeout.connect(self._request_poll)
        self._timer.start()
        self._request_poll()

        self._volume_timer = QTimer(self)
        self._volume_timer.setSingleShot(True)
        self._volume_timer.setInterval(100)
        self._volume_timer.timeout.connect(self._flush_virtual_mic_percent)

    def closeEvent(self, event):  # type: ignore[override]
        self._timer.stop()
        self._volume_timer.stop()
        self._save_config()

        # Best-effort cleanup: unlink managed routes on shutdown.
        try:
            snap = self.controller.snapshot()
            self.state.set_streaming_active(False)
            actions = self.state.compute_actions(
                snap,
                virtual_sink_key=self._virtual_mic_sink_key,
                virtual_source_key=self._virtual_mic_source_key,
            )
            if actions:
                self._apply_actions_sync(actions)
        except Exception as exc:
            self._logger.warning("Shutdown unlink cleanup failed: %s", exc)
        finally:
            try:
                self.controller.teardown_virtual_microphone()
            except Exception as exc:
                self._logger.warning("Virtual microphone teardown failed: %s", exc)

        self._worker_thread.quit()
        self._worker_thread.wait(1500)
        super().closeEvent(event)

    def _build_ui(self) -> None:
        root = QWidget(self)
        self.setCentralWidget(root)
        main_layout = QVBoxLayout(root)

        top_row = QHBoxLayout()

        source_group = QGroupBox("SOURCE APPS")
        target_group = QGroupBox("TARGET APPS")

        source_layout = QVBoxLayout(source_group)
        target_layout = QVBoxLayout(target_group)

        self.sources_list = CheckListWidget()
        self.targets_list = CheckListWidget()
        self.sources_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.targets_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

        source_layout.addWidget(self.sources_list)
        target_layout.addWidget(self.targets_list)

        top_row.addWidget(source_group, 1)
        top_row.addWidget(target_group, 1)

        main_layout.addLayout(top_row)

        row1 = QHBoxLayout()
        self.toggle_streaming_btn = QPushButton("Toggle Streaming")
        self.toggle_streaming_btn.setCheckable(True)
        self.clear_capturing_btn = QPushButton("Clear Capturing")
        self.clear_streaming_btn = QPushButton("Clear Streaming")
        row1.addWidget(self.toggle_streaming_btn)
        row1.addWidget(self.clear_capturing_btn)
        row1.addWidget(self.clear_streaming_btn)

        row2 = QHBoxLayout()
        self.auto_capture_btn = QPushButton("Auto Capture")
        self.auto_capture_btn.setCheckable(True)
        self.auto_streaming_btn = QPushButton("Auto Streaming")
        self.auto_streaming_btn.setCheckable(True)
        row2.addWidget(self.auto_capture_btn)
        row2.addWidget(self.auto_streaming_btn)

        volume_row = QHBoxLayout()
        self.volume_down_btn = QPushButton("-5%")
        self.reset_volume_btn = QPushButton("Reset Volume")
        self.volume_up_btn = QPushButton("+5%")
        self.volume_spin = QSpinBox()
        self.volume_spin.setRange(0, 200)
        self.volume_spin.setSuffix(" %")
        self.volume_spin.setValue(100)
        self.volume_slider = QSlider(Qt.Orientation.Horizontal)
        self.volume_slider.setRange(0, 200)
        self.volume_slider.setValue(100)
        volume_row.addWidget(QLabel("Virtual Mic Volume"))
        volume_row.addWidget(self.volume_down_btn)
        volume_row.addWidget(self.reset_volume_btn)
        volume_row.addWidget(self.volume_up_btn)
        volume_row.addWidget(self.volume_spin)
        volume_row.addWidget(self.volume_slider, 1)

        main_layout.addLayout(volume_row)
        main_layout.addLayout(row1)
        main_layout.addLayout(row2)

        self.sources_list.itemChanged.connect(self._on_source_item_changed)
        self.targets_list.itemChanged.connect(self._on_target_item_changed)
        self.sources_list.customContextMenuRequested.connect(
            lambda pos: self._open_item_menu(self.sources_list, pos, source_list=True)
        )
        self.targets_list.customContextMenuRequested.connect(
            lambda pos: self._open_item_menu(self.targets_list, pos, source_list=False)
        )
        self.toggle_streaming_btn.toggled.connect(self._toggle_streaming)
        self.clear_capturing_btn.clicked.connect(self._clear_capturing)
        self.clear_streaming_btn.clicked.connect(self._clear_streaming)
        self.auto_capture_btn.toggled.connect(self._toggle_auto_capture)
        self.auto_streaming_btn.toggled.connect(self._toggle_auto_streaming)
        self.volume_down_btn.clicked.connect(lambda: self._change_virtual_mic_volume(-5))
        self.reset_volume_btn.clicked.connect(self._reset_virtual_mic_volume)
        self.volume_up_btn.clicked.connect(lambda: self._change_virtual_mic_volume(5))
        self.volume_spin.valueChanged.connect(self._on_volume_spin_changed)
        self.volume_slider.valueChanged.connect(self._on_volume_slider_changed)

        self.toggle_streaming_btn.setChecked(True)
        self.setStatusBar(QStatusBar())

    def _build_worker(self) -> None:
        self._worker_thread = QThread(self)
        self._worker = PipeWireWorker(self.controller)
        self._worker.moveToThread(self._worker_thread)

        self.request_poll.connect(self._worker.poll_snapshot)
        self.request_apply_actions.connect(self._worker.apply_actions)
        self.request_set_virtual_mic_volume.connect(self._worker.set_virtual_mic_volume)
        self._worker.snapshot_ready.connect(self._on_snapshot)
        self._worker.actions_failed.connect(self._show_status)

        self._worker_thread.start()

    @pyqtSlot(object)
    def _on_snapshot(self, snapshot: object) -> None:
        self._poll_in_flight = False
        if not isinstance(snapshot, PipeWireSnapshot):
            self._logger.warning("Worker emitted non-snapshot payload: %r", type(snapshot))
            return

        self.snapshot = snapshot
        sources = [
            node
            for node in self.controller.application_sources(snapshot)
            if node.name not in self.EXCLUDED_SOURCE_KEYS
        ]
        targets = [
            node
            for node in self.controller.application_targets(snapshot)
            if node.name not in self.EXCLUDED_TARGET_KEYS
        ]
        self.state.update_available(
            sources=sources,
            targets=targets,
        )

        self._refresh_lists()
        self._apply_routing_actions()

        self._show_status(
            f"Sources: {len(self.state.available_sources)} | Targets: {len(self.state.available_targets)}"
        )

    def _request_poll(self) -> None:
        if self._poll_in_flight:
            return
        self._poll_in_flight = True
        self.request_poll.emit()

    def _refresh_lists(self) -> None:
        self._updating_lists = True
        try:
            self.sources_list.clear()
            for entry in self.state.source_entries():
                node = self.state.available_sources.get(entry.key)
                label = f"{entry.label} [auto]" if entry.key in self._auto_select_source_keys else entry.label
                item = QListWidgetItem(label)
                item.setData(self.USER_ROLE, entry.key)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if entry.selected else Qt.CheckState.Unchecked)
                if not entry.available:
                    item.setBackground(self.not_available_color)
                self.sources_list.addItem(item)

            self.targets_list.clear()
            for entry in self.state.target_entries():
                label = f"{entry.label} [auto]" if entry.key in self._auto_select_target_keys else entry.label
                item = QListWidgetItem(label)
                item.setData(self.USER_ROLE, entry.key)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if entry.selected else Qt.CheckState.Unchecked)
                if not entry.available:
                    item.setBackground(self.not_available_color)
                self.targets_list.addItem(item)
        finally:
            self._updating_lists = False

        self.sources_list.setEnabled(not self.state.auto_capture)
        self.targets_list.setEnabled(not self.state.auto_streaming)

    def _apply_routing_actions(self) -> None:
        actions = self.state.compute_actions(
            self.snapshot,
            virtual_sink_key=self._virtual_mic_sink_key,
            virtual_source_key=self._virtual_mic_source_key,
        )
        if actions:
            self.request_apply_actions.emit(actions)

    def _apply_actions_sync(self, actions: list[RouteAction]) -> None:
        if not actions:
            return
        try:
            snapshot = self.controller.snapshot()
        except PipeWireError as exc:
            self._logger.warning("Unable to refresh snapshot for sync actions: %s", exc)
            return

        for action in actions:
            try:
                if action.op == "link":
                    self.controller.create_link_by_key(action.source_key, action.target_key, snapshot)
                elif action.op == "unlink":
                    self.controller.remove_link_by_key(action.source_key, action.target_key, snapshot)
            except PipeWireError as exc:
                self._logger.warning("Sync action failed (%s %s -> %s): %s", action.op, action.source_key, action.target_key, exc)

    def _checked_keys(self, widget: QListWidget) -> set[str]:
        keys: set[str] = set()
        for i in range(widget.count()):
            item = widget.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
            key = item.data(self.USER_ROLE)
            if isinstance(key, str):
                keys.add(key)
        return keys

    def _on_source_item_changed(self, _item: QListWidgetItem) -> None:
        if self._updating_lists:
            return
        self.state.set_source_selection(self._checked_keys(self.sources_list))
        self._apply_routing_actions()

    def _on_target_item_changed(self, _item: QListWidgetItem) -> None:
        if self._updating_lists:
            return
        self.state.set_target_selection(self._checked_keys(self.targets_list))
        self._apply_routing_actions()

    def _toggle_streaming(self, enabled: bool) -> None:
        self.state.set_streaming_active(enabled)
        self._set_streaming_indicator(enabled)
        self._apply_routing_actions()

    def _clear_capturing(self) -> None:
        self.state.clear_sources()
        self._refresh_lists()
        self._apply_routing_actions()

    def _clear_streaming(self) -> None:
        self.state.clear_targets()
        self._refresh_lists()
        self._apply_routing_actions()

    def _toggle_auto_capture(self, enabled: bool) -> None:
        self.state.set_auto_capture(enabled)
        self._refresh_lists()
        self._apply_routing_actions()

    def _toggle_auto_streaming(self, enabled: bool) -> None:
        self.state.set_auto_streaming(enabled)
        self._refresh_lists()
        self._apply_routing_actions()

    def _set_streaming_indicator(self, enabled: bool) -> None:
        if enabled:
            self.toggle_streaming_btn.setStyleSheet("background-color: #4caf50; color: white; font-weight: 600;")
        else:
            self.toggle_streaming_btn.setStyleSheet("")

    def _change_virtual_mic_volume(self, delta: int) -> None:
        target = max(self.volume_spin.minimum(), min(self.volume_spin.maximum(), self.volume_spin.value() + delta))
        self.volume_spin.setValue(target)

    def _reset_virtual_mic_volume(self) -> None:
        self.volume_spin.setValue(100)

    def _on_volume_spin_changed(self, value: int) -> None:
        slider_value = max(self.volume_slider.minimum(), min(self.volume_slider.maximum(), value))
        if self.volume_slider.value() != slider_value:
            self.volume_slider.blockSignals(True)
            self.volume_slider.setValue(slider_value)
            self.volume_slider.blockSignals(False)
        self._schedule_virtual_mic_percent(float(value))

    def _on_volume_slider_changed(self, value: int) -> None:
        if self.volume_spin.value() != value:
            self.volume_spin.blockSignals(True)
            self.volume_spin.setValue(value)
            self.volume_spin.blockSignals(False)
        self._schedule_virtual_mic_percent(float(value))

    def _schedule_virtual_mic_percent(self, value: float) -> None:
        self._pending_virtual_mic_percent = value
        self._volume_timer.start()

    def _flush_virtual_mic_percent(self) -> None:
        self.request_set_virtual_mic_volume.emit(self._pending_virtual_mic_percent)

    def _open_item_menu(self, widget: QListWidget, pos, source_list: bool) -> None:
        item = widget.itemAt(pos)
        if item is None:
            return
        key = item.data(self.USER_ROLE)
        if not isinstance(key, str):
            return

        auto_set = self._auto_select_source_keys if source_list else self._auto_select_target_keys
        menu = QMenu(self)
        copy_name_action = menu.addAction("Copy Name")
        menu.addSeparator()
        toggle_label = "Unmark Auto Select" if key in auto_set else "Mark Auto Select"
        toggle_auto_action = menu.addAction(toggle_label)

        chosen = menu.exec(widget.viewport().mapToGlobal(pos))
        if chosen is copy_name_action:
            QApplication.clipboard().setText(key)
        elif chosen is toggle_auto_action:
            if key in auto_set:
                auto_set.remove(key)
            else:
                auto_set.add(key)
                if source_list:
                    self.state.selected_sources.add(key)
                else:
                    self.state.selected_targets.add(key)
            self._save_config()
            self._refresh_lists()

    def _save_config(self) -> None:
        try:
            save_config(
                AppConfig(
                    auto_select_sources=set(self._auto_select_source_keys),
                    auto_select_targets=set(self._auto_select_target_keys),
                )
            )
        except OSError as exc:
            self._logger.warning("Failed to save config.json: %s", exc)

    def load_media_file(self, file_path: str) -> None:
        path = Path(file_path).expanduser().resolve()
        self.media.load_file(str(path))

    def route_loaded_media_to_targets(self) -> None:
        player_sources = self.controller.find_sources_by_pid(self.media.process_id, self.snapshot)
        if not player_sources:
            self._show_status("No media player source detected yet. Start playback first.")
            return

        media_source_key = player_sources[0].name
        actions = self.state.route_media_to_targets_actions(
            media_source_key,
            virtual_sink_key=self._virtual_mic_sink_key,
            virtual_source_key=self._virtual_mic_source_key,
        )
        if actions:
            self.request_apply_actions.emit(actions)

    def _show_error(self, message: str) -> None:
        self._logger.error(message)
        self.statusBar().showMessage(message, 7000)
        QMessageBox.critical(self, "AudioLink", message)

    def _show_status(self, message: str) -> None:
        self._logger.info(message)
        if message.startswith("PipeWire refresh failed:") or message.startswith("Failed to prepare routing actions:"):
            self._poll_in_flight = False
        self.statusBar().showMessage(message, 5000)


def build_window(controller: PipeWireController, media: MediaPlayerController) -> MainWindow:
    try:
        return MainWindow(controller=controller, media=media)
    except Exception as exc:
        traceback.print_exc()
        raise RuntimeError(f"Failed to initialize UI: {exc}") from exc
