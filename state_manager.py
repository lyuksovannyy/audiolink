from __future__ import annotations

from dataclasses import dataclass

from pipewire_controller import Node, PipeWireSnapshot


@dataclass(slots=True)
class ListEntry:
    key: str
    label: str
    available: bool
    active: bool
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
        self._source_match_signatures: dict[str, str] = {}
        self._source_match_bases: dict[str, str] = {}

        # Desired pairs while streaming is enabled. Includes temporarily missing nodes.
        self._desired_pairs: set[tuple[str, str]] = set()

    def update_available(self, sources: list[Node], targets: list[Node]) -> None:
        self.available_sources = {self._source_key(n): n for n in sources}
        self.available_targets = {self._target_key(n): n for n in targets}
        available_by_signature: dict[str, str] = {}
        available_by_base: dict[str, str] = {}
        source_meta: dict[str, tuple[str, str, int | None]] = {}
        base_counts: dict[str, int] = {}

        for key, node in self.available_sources.items():
            self._source_labels[key] = self._label(node)
            app_name, title, pid = self._source_match_parts(node)
            source_meta[key] = (app_name, title, pid)
            base = self._source_base_signature(app_name, title)
            base_counts[base] = base_counts.get(base, 0) + 1
        for key, (app_name, title, pid) in source_meta.items():
            base = self._source_base_signature(app_name, title)
            signature = self._source_match_signature(app_name, title, pid, base_counts)
            self._source_match_bases[key] = base
            self._source_match_signatures[key] = signature
            available_by_signature.setdefault(signature, key)
            available_by_base.setdefault(base, key)
        for key, node in self.available_targets.items():
            self._target_labels[key] = self._label(node)

        # If a previously selected/unavailable source reappears with a new key
        # (for example new stream node id), remap selection by APPNAME - TITLE.
        remapped_sources: set[str] = set()
        kept_unavailable_bases: set[str] = set()
        for key in self.selected_sources:
            if key in self.available_sources:
                remapped_sources.add(key)
                continue
            old_signature = self._source_match_signatures.get(key)
            old_base = self._source_match_bases.get(key)
            if old_signature is None:
                remapped_sources.add(key)
                continue
            replacement = available_by_signature.get(old_signature)
            if replacement is None and old_base is not None:
                replacement = available_by_base.get(old_base)
            if replacement is not None:
                remapped_sources.add(replacement)
                continue
            # Keep only one stale unavailable key per APPNAME - TITLE base signature.
            base_for_dedupe = old_base or old_signature
            if base_for_dedupe in kept_unavailable_bases:
                continue
            kept_unavailable_bases.add(base_for_dedupe)
            remapped_sources.add(key)
        self.selected_sources = remapped_sources

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
        available_snapshot_sources = {self._source_key(n) for n in snapshot.sources}
        available_snapshot_targets = {self._target_key(n) for n in snapshot.sinks}

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
        def sort_key(key: str) -> tuple[int, str]:
            is_available = key in available
            if not is_available:
                return (2, key.lower())
            node = available[key]
            is_active = (node.state or "").lower() == "running"
            return (0 if is_active else 1, key.lower())

        entries: list[ListEntry] = []
        for key in sorted(keys, key=sort_key):
            node = available.get(key)
            is_available = node is not None
            is_active = bool(node is not None and (node.state or "").lower() == "running")
            base_label = labels.get(key, key)
            label = base_label if is_available else f"{base_label} (unavailable)"
            entries.append(
                ListEntry(
                    key=key,
                    label=label,
                    available=is_available,
                    active=is_active,
                    selected=key in selected,
                )
            )
        return entries

    def _linked_pairs(self, snapshot: PipeWireSnapshot) -> set[tuple[str, str]]:
        nodes = snapshot.nodes
        by_id_source = {node.id: self._source_key(node) for node in nodes.values()}
        by_id_target = {node.id: self._target_key(node) for node in nodes.values()}
        linked: set[tuple[str, str]] = set()

        for link in snapshot.links:
            source_key = by_id_source.get(link.output_node_id)
            target_key = by_id_target.get(link.input_node_id)
            if source_key is None or target_key is None:
                continue
            linked.add((source_key, target_key))

        return linked

    @staticmethod
    def _source_key(node: Node) -> str:
        media_class = (node.media_class or "").lower()
        # Keep app stream nodes unique, but keep device/virtual sources name-keyed
        # so virtual mic bridge links can be tracked correctly.
        if media_class.startswith("stream/output/audio"):
            return f"source:{node.id}"
        return node.name

    @staticmethod
    def _target_key(node: Node) -> str:
        media_class = (node.media_class or "").lower()
        if media_class.startswith("stream/input/audio"):
            return f"target:{node.id}"
        return node.name

    @staticmethod
    def _label(node: Node) -> str:
        if node.media_name:
            return node.media_name
        return node.description

    @staticmethod
    def _source_match_parts(node: Node) -> tuple[str, str, int | None]:
        app = (node.application_name or "").strip()
        if not app:
            if node.process_id is not None:
                app = f"PID {node.process_id}"
            else:
                app = (node.description or node.name).strip()
        title = (node.media_name or node.description or node.name).strip()
        return app, title, node.process_id

    @staticmethod
    def _source_base_signature(app_name: str, title: str) -> str:
        return f"{app_name} - {title}"

    @classmethod
    def _source_match_signature(
        cls,
        app_name: str,
        title: str,
        pid: int | None,
        base_counts: dict[str, int],
    ) -> str:
        base = cls._source_base_signature(app_name, title)
        if base_counts.get(base, 0) > 1 and pid is not None:
            return f"{base} - {pid}"
        return base
