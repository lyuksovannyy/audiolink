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
    QTreeWidget,
    QTreeWidgetItem,
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


class SourceTreeWidget(QTreeWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setHeaderHidden(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)

    def mousePressEvent(self, event):  # type: ignore[override]
        item = self.itemAt(event.pos())
        if (
            item is not None
            and self.isEnabled()
            and event.button() == Qt.MouseButton.LeftButton
        ):
            item.setCheckState(
                0,
                Qt.CheckState.Unchecked
                if item.checkState(0) == Qt.CheckState.Checked
                else Qt.CheckState.Checked,
            )
            event.accept()
            return
        super().mousePressEvent(event)


class MainWindow(QMainWindow):
    USER_ROLE = int(Qt.ItemDataRole.UserRole)
    TITLE_ROLE = USER_ROLE + 1
    NAME_ROLE = USER_ROLE + 2
    APP_ROLE = USER_ROLE + 3
    APP_MARKER_ROLE = USER_ROLE + 4
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
        self._auto_select_source_apps: set[str] = set(self._cfg.auto_select_sources)
        self._auto_select_source_items: set[str] = set(self._cfg.auto_select_source_items)
        self._auto_select_target_names: set[str] = set(self._cfg.auto_select_targets)

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

        self.sources_list = SourceTreeWidget()
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
            lambda pos: self._open_source_item_menu(pos)
        )
        self.targets_list.customContextMenuRequested.connect(
            lambda pos: self._open_target_item_menu(pos)
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
        for key, node in self.state.available_sources.items():
            app_name = self._source_group_name(node)
            app_marker = self._source_group_marker(node)
            marker = self._source_item_marker(
                app_marker=app_marker,
                title=(node.media_name or node.description or node.name),
                raw_name=node.name,
            )
            if (
                app_marker in self._auto_select_source_apps
                or app_name in self._auto_select_source_apps  # backward compatibility with older config values
                or marker in self._auto_select_source_items
            ):
                self.state.selected_sources.add(key)
        for key, node in self.state.available_targets.items():
            if node.name in self._auto_select_target_names:
                self.state.selected_targets.add(key)

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
            self._refresh_source_tree()

            self.targets_list.clear()
            for entry in self.state.target_entries():
                node = self.state.available_targets.get(entry.key)
                label = self._target_display_label(node, entry.label)
                if node is not None and node.name in self._auto_select_target_names:
                    label = f"{label} [auto]"
                item = QListWidgetItem(label)
                item.setData(self.USER_ROLE, entry.key)
                item.setData(self.TITLE_ROLE, label)
                item.setData(self.NAME_ROLE, node.name if node is not None else entry.key)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked if entry.selected else Qt.CheckState.Unchecked)
                if not entry.available:
                    item.setBackground(self.not_available_color)
                self.targets_list.addItem(item)
        finally:
            self._updating_lists = False

        self.sources_list.setEnabled(not self.state.auto_capture)
        self.targets_list.setEnabled(not self.state.auto_streaming)

    def _refresh_source_tree(self) -> None:
        self.sources_list.clear()
        grouped: dict[str, list[tuple[str, str, bool, bool]]] = {}
        group_labels: dict[str, str] = {}
        for entry in self.state.source_entries():
            node = self.state.available_sources.get(entry.key)
            app_name = self._source_group_name(node) if node is not None else "Unavailable"
            app_marker = self._source_group_marker(node) if node is not None else f"missing:{entry.key}"
            grouped.setdefault(app_marker, []).append((entry.key, entry.label, entry.available, entry.selected))
            group_labels[app_marker] = app_name

        for app_marker in sorted(grouped.keys(), key=str.lower):
            app_name = group_labels.get(app_marker, app_marker)
            group_items = sorted(grouped[app_marker], key=lambda x: x[1].lower())
            if len(group_items) == 1:
                key, child_label, available, selected = group_items[0]
                node = self.state.available_sources.get(key)
                label = f"{app_name} - {child_label}" if child_label and child_label != app_name else app_name
                item_marker = self._source_item_marker(
                    app_marker=app_marker,
                    title=(child_label or app_name),
                    raw_name=(node.name if node is not None else key),
                )
                if (
                    app_marker in self._auto_select_source_apps
                    or app_name in self._auto_select_source_apps  # backward compatibility with older config values
                    or item_marker in self._auto_select_source_items
                ):
                    label = f"{label} [auto]"
                leaf = QTreeWidgetItem([label])
                leaf.setData(0, self.USER_ROLE, key)
                leaf.setData(0, self.TITLE_ROLE, child_label or app_name)
                leaf.setData(0, self.NAME_ROLE, node.name if node is not None else key)
                leaf.setData(0, self.APP_ROLE, app_name)
                leaf.setData(0, self.APP_MARKER_ROLE, app_marker)
                leaf.setFlags(leaf.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                leaf.setCheckState(0, Qt.CheckState.Checked if selected else Qt.CheckState.Unchecked)
                if not available:
                    leaf.setBackground(0, self.not_available_color)
                self.sources_list.addTopLevelItem(leaf)
                continue

            parent_label = (
                f"{app_name} [auto]"
                if (app_marker in self._auto_select_source_apps or app_name in self._auto_select_source_apps)
                else app_name
            )
            parent = QTreeWidgetItem([parent_label])
            parent.setData(0, self.USER_ROLE, app_marker)
            parent.setData(0, self.TITLE_ROLE, app_name)
            parent.setData(0, self.NAME_ROLE, app_name)
            parent.setData(0, self.APP_ROLE, app_name)
            parent.setData(0, self.APP_MARKER_ROLE, app_marker)
            parent.setFlags(parent.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            self.sources_list.addTopLevelItem(parent)

            child_states: list[Qt.CheckState] = []
            for key, label, available, selected in group_items:
                node = self.state.available_sources.get(key)
                item_marker = self._source_item_marker(
                    app_marker=app_marker,
                    title=label,
                    raw_name=(node.name if node is not None else key),
                )
                child_label = f"{label} [auto]" if item_marker in self._auto_select_source_items else label
                child = QTreeWidgetItem([label])
                child.setData(0, self.USER_ROLE, key)
                child.setData(0, self.TITLE_ROLE, label)
                child.setText(0, child_label)
                child.setData(0, self.NAME_ROLE, node.name if node is not None else key)
                child.setData(0, self.APP_ROLE, app_name)
                child.setData(0, self.APP_MARKER_ROLE, app_marker)
                child.setFlags(child.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                state = Qt.CheckState.Checked if selected else Qt.CheckState.Unchecked
                child.setCheckState(0, state)
                child_states.append(state)
                if not available:
                    child.setBackground(0, self.not_available_color)
                parent.addChild(child)

            if child_states and all(s == Qt.CheckState.Checked for s in child_states):
                parent.setCheckState(0, Qt.CheckState.Checked)
            elif any(s == Qt.CheckState.Checked for s in child_states):
                parent.setCheckState(0, Qt.CheckState.PartiallyChecked)
            else:
                parent.setCheckState(0, Qt.CheckState.Unchecked)
            parent.setExpanded(True)

    def _apply_routing_actions(self) -> None:
        actions = self.state.compute_actions(
            self.snapshot,
            virtual_sink_key=self._virtual_mic_sink_key,
            virtual_source_key=self._virtual_mic_source_key,
        )
        if actions:
            self._logger.debug(
                "Dispatching %d routing action(s): %s",
                len(actions),
                ", ".join(f"{a.op}:{a.source_key}->{a.target_key}" for a in actions[:8]),
            )
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

    def _on_source_item_changed(self, item: QTreeWidgetItem, _column: int) -> None:
        if self._updating_lists:
            return
        self._updating_lists = True
        try:
            if item.parent() is None and item.childCount() > 0:
                desired = item.checkState(0)
                if desired == Qt.CheckState.PartiallyChecked:
                    desired = Qt.CheckState.Checked
                for i in range(item.childCount()):
                    child = item.child(i)
                    child.setCheckState(0, desired)
            elif item.parent() is not None:
                parent = item.parent()
                checked = 0
                for i in range(parent.childCount()):
                    if parent.child(i).checkState(0) == Qt.CheckState.Checked:
                        checked += 1
                if checked == 0:
                    parent.setCheckState(0, Qt.CheckState.Unchecked)
                elif checked == parent.childCount():
                    parent.setCheckState(0, Qt.CheckState.Checked)
                else:
                    parent.setCheckState(0, Qt.CheckState.PartiallyChecked)
        finally:
            self._updating_lists = False

        selected: set[str] = set()
        for i in range(self.sources_list.topLevelItemCount()):
            parent = self.sources_list.topLevelItem(i)
            if parent.childCount() == 0:
                if parent.checkState(0) == Qt.CheckState.Checked:
                    key = parent.data(0, self.USER_ROLE)
                    if isinstance(key, str):
                        selected.add(key)
                continue
            for j in range(parent.childCount()):
                child = parent.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    key = child.data(0, self.USER_ROLE)
                    if isinstance(key, str):
                        selected.add(key)
        self.state.set_source_selection(selected)
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

    def _open_source_item_menu(self, pos) -> None:
        item = self.sources_list.itemAt(pos)
        if item is None:
            return
        key = item.data(0, self.USER_ROLE)
        title = item.data(0, self.TITLE_ROLE)
        raw_name = item.data(0, self.NAME_ROLE)
        if not isinstance(key, str):
            return
        if not isinstance(title, str):
            title = key
        if not isinstance(raw_name, str):
            raw_name = key
        is_group = item.parent() is None and item.childCount() > 0
        group_child_keys: list[str] = []
        if is_group:
            for i in range(item.childCount()):
                child = item.child(i)
                child_key = child.data(0, self.USER_ROLE)
                if isinstance(child_key, str):
                    group_child_keys.append(child_key)
        app_name = item.data(0, self.APP_ROLE)
        if not isinstance(app_name, str):
            app_name = self._source_group_name(self.state.available_sources.get(key))
        app_marker = item.data(0, self.APP_MARKER_ROLE)
        if not isinstance(app_marker, str):
            app_marker = app_name
        auto_set = self._auto_select_source_apps
        auto_item_set = self._auto_select_source_items
        menu = QMenu(self)
        copy_name_action = menu.addAction("Copy Name")
        menu.addSeparator()
        marker_key = app_marker
        item_marker = self._source_item_marker(app_marker=app_marker, title=title, raw_name=raw_name)
        if is_group:
            marked = marker_key in auto_set or app_name in auto_set
        else:
            marked = item_marker in auto_item_set
        toggle_label = "Unmark Auto Select" if marked else "Mark Auto Select"
        toggle_auto_action = menu.addAction(toggle_label)

        chosen = menu.exec(self.sources_list.viewport().mapToGlobal(pos))
        if chosen is copy_name_action:
            QApplication.clipboard().setText(app_name if is_group else raw_name)
        elif chosen is toggle_auto_action:
            if is_group:
                if marker_key in auto_set:
                    auto_set.remove(marker_key)
                else:
                    auto_set.add(marker_key)
                # Cleanup legacy value if present to avoid dual markers in config.
                auto_set.discard(app_name)
            else:
                if item_marker in auto_item_set:
                    auto_item_set.remove(item_marker)
                else:
                    auto_item_set.add(item_marker)
                    self.state.selected_sources.add(key)
            if is_group:
                for child_key in group_child_keys:
                    self.state.selected_sources.add(child_key)
            else:
                self.state.selected_sources.add(key)
            self._save_config()
            self._refresh_lists()

    def _open_target_item_menu(self, pos) -> None:
        item = self.targets_list.itemAt(pos)
        if item is None:
            return
        key = item.data(self.USER_ROLE)
        raw_name = item.data(self.NAME_ROLE)
        if not isinstance(key, str):
            return
        if not isinstance(raw_name, str):
            raw_name = key

        marker_key = raw_name
        auto_set = self._auto_select_target_names
        menu = QMenu(self)
        copy_name_action = menu.addAction("Copy Name")
        menu.addSeparator()
        toggle_label = "Unmark Auto Select" if marker_key in auto_set else "Mark Auto Select"
        toggle_auto_action = menu.addAction(toggle_label)

        chosen = menu.exec(self.targets_list.viewport().mapToGlobal(pos))
        if chosen is copy_name_action:
            QApplication.clipboard().setText(raw_name)
        elif chosen is toggle_auto_action:
            if marker_key in auto_set:
                auto_set.remove(marker_key)
            else:
                auto_set.add(marker_key)
                for cur_key, node in self.state.available_targets.items():
                    if node.name == marker_key:
                        self.state.selected_targets.add(cur_key)
            self._save_config()
            self._refresh_lists()

    def _save_config(self) -> None:
        try:
            save_config(
                AppConfig(
                    auto_select_sources=set(self._auto_select_source_apps),
                    auto_select_source_items=set(self._auto_select_source_items),
                    auto_select_targets=set(self._auto_select_target_names),
                )
            )
        except OSError as exc:
            self._logger.warning("Failed to save config.json: %s", exc)

    @staticmethod
    def _source_group_name(node) -> str:
        if node is None:
            return "Unavailable"
        app = (node.application_name or "").strip()
        if app:
            return app
        if node.process_id is not None:
            return f"PID {node.process_id}"
        return node.description or node.name

    @staticmethod
    def _source_group_marker(node) -> str:
        if node is None:
            return "unavailable"
        app = (node.application_name or "").strip()
        if app and node.process_id is not None:
            return f"app:{app}|pid:{node.process_id}"
        if app:
            return f"app:{app}"
        if node.process_id is not None:
            return f"pid:{node.process_id}"
        return f"name:{node.name}"

    @staticmethod
    def _target_display_label(node, fallback_label: str) -> str:
        if node is None:
            return fallback_label
        app = (node.application_name or "").strip()
        stream = (node.media_name or "").strip()
        if app and stream and stream.lower() != app.lower():
            return f"{app} - {stream}"
        if app:
            return app
        if stream:
            return stream
        return fallback_label

    @staticmethod
    def _source_item_marker(app_marker: str, title: str, raw_name: str) -> str:
        return "|".join([app_marker.strip(), title.strip(), raw_name.strip()])

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
