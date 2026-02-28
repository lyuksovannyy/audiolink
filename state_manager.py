from __future__ import annotations

from dataclasses import dataclass

from pipewire_controller import Node, PipeWireSnapshot


@dataclass(slots=True)
class ListEntry:
    key: str
    label: str
    available: bool
    selected: bool


@dataclass(slots=True)
class RouteAction:
    op: str  # "link" or "unlink"
    source_key: str
    target_key: str


class RoutingStateManager:
    def __init__(self) -> None:
        self.streaming_active = False
        self.auto_capture = False
        self.auto_streaming = False

        self.selected_sources: set[str] = set()
        self.selected_targets: set[str] = set()

        self.available_sources: dict[str, Node] = {}
        self.available_targets: dict[str, Node] = {}
        self._source_labels: dict[str, str] = {}
        self._target_labels: dict[str, str] = {}

        # Desired pairs while streaming is enabled. Includes temporarily missing nodes.
        self._desired_pairs: set[tuple[str, str]] = set()

    def update_available(self, sources: list[Node], targets: list[Node]) -> None:
        self.available_sources = {self._key(n): n for n in sources}
        self.available_targets = {self._key(n): n for n in targets}

        for key, node in self.available_sources.items():
            self._source_labels[key] = self._label(node)
        for key, node in self.available_targets.items():
            self._target_labels[key] = self._label(node)

        # Keep manual checkbox state unchanged; auto modes are routing-only.

    def set_source_selection(self, keys: set[str]) -> None:
        if self.auto_capture:
            return
        self.selected_sources = set(keys)

    def set_target_selection(self, keys: set[str]) -> None:
        if self.auto_streaming:
            return
        self.selected_targets = set(keys)

    def clear_sources(self) -> None:
        if not self.auto_capture:
            self.selected_sources.clear()

    def clear_targets(self) -> None:
        if not self.auto_streaming:
            self.selected_targets.clear()

    def set_auto_capture(self, enabled: bool) -> None:
        self.auto_capture = enabled

    def set_auto_streaming(self, enabled: bool) -> None:
        self.auto_streaming = enabled

    def set_streaming_active(self, enabled: bool) -> None:
        self.streaming_active = enabled

    def source_entries(self) -> list[ListEntry]:
        return self._build_entries(
            selected=self.selected_sources,
            available=self.available_sources,
            labels=self._source_labels,
        )

    def target_entries(self) -> list[ListEntry]:
        return self._build_entries(
            selected=self.selected_targets,
            available=self.available_targets,
            labels=self._target_labels,
        )

    def compute_actions(
        self,
        snapshot: PipeWireSnapshot,
        virtual_sink_key: str | None = None,
        virtual_source_key: str | None = None,
    ) -> list[RouteAction]:
        linked_pairs = self._linked_pairs(snapshot)
        available_snapshot_sources = {self._key(n) for n in snapshot.sources}
        available_snapshot_targets = {self._key(n) for n in snapshot.sinks}

        source_pool = set(self.available_sources.keys()) if self.auto_capture else set(self.selected_sources)
        target_pool = set(self.available_targets.keys()) if self.auto_streaming else set(self.selected_targets)
        if virtual_sink_key is not None and virtual_source_key is not None:
            desired_pairs = {(source_key, virtual_sink_key) for source_key in source_pool}
            desired_pairs |= {(virtual_source_key, target_key) for target_key in target_pool}
        else:
            desired_pairs = {(s, t) for s in source_pool for t in target_pool}

        actions: list[RouteAction] = []

        if self.streaming_active:
            for source_key, target_key in sorted(self._desired_pairs - desired_pairs):
                actions.append(RouteAction("unlink", source_key, target_key))

            for source_key, target_key in sorted(desired_pairs):
                source_available = source_key in available_snapshot_sources
                target_available = target_key in available_snapshot_targets
                if not (source_available and target_available):
                    continue
                if (source_key, target_key) not in linked_pairs:
                    actions.append(RouteAction("link", source_key, target_key))

            self._desired_pairs = desired_pairs
        else:
            for source_key, target_key in sorted(self._desired_pairs):
                actions.append(RouteAction("unlink", source_key, target_key))
            self._desired_pairs.clear()

        return actions

    def route_media_to_targets_actions(
        self,
        media_source_key: str,
        virtual_sink_key: str | None = None,
        virtual_source_key: str | None = None,
    ) -> list[RouteAction]:
        target_pool = set(self.available_targets.keys()) if self.auto_streaming else set(self.selected_targets)
        actions: list[RouteAction] = []
        if virtual_sink_key is not None and virtual_source_key is not None:
            actions.append(RouteAction("link", media_source_key, virtual_sink_key))
            for target_key in sorted(target_pool):
                if target_key in self.available_targets:
                    actions.append(RouteAction("link", virtual_source_key, target_key))
            return actions

        for target_key in sorted(target_pool):
            if target_key in self.available_targets:
                actions.append(RouteAction("link", media_source_key, target_key))
        return actions

    def selected_target_keys(self) -> list[str]:
        target_pool = set(self.available_targets.keys()) if self.auto_streaming else set(self.selected_targets)
        return sorted(k for k in target_pool if k in self.available_targets)

    def _build_entries(
        self,
        selected: set[str],
        available: dict[str, Node],
        labels: dict[str, str],
    ) -> list[ListEntry]:
        keys = set(available.keys()) | {k for k in selected if k not in available}
        entries: list[ListEntry] = []
        for key in sorted(keys, key=str.lower):
            is_available = key in available
            base_label = labels.get(key, key)
            label = base_label if is_available else f"{base_label} (unavailable)"
            entries.append(
                ListEntry(
                    key=key,
                    label=label,
                    available=is_available,
                    selected=key in selected,
                )
            )
        return entries

    def _linked_pairs(self, snapshot: PipeWireSnapshot) -> set[tuple[str, str]]:
        nodes = snapshot.nodes
        by_id = {node.id: self._key(node) for node in nodes.values()}
        linked: set[tuple[str, str]] = set()

        for link in snapshot.links:
            source_key = by_id.get(link.output_node_id)
            target_key = by_id.get(link.input_node_id)
            if source_key is None or target_key is None:
                continue
            linked.add((source_key, target_key))

        return linked

    @staticmethod
    def _key(node: Node) -> str:
        return node.name

    @staticmethod
    def _label(node: Node) -> str:
        return node.description
