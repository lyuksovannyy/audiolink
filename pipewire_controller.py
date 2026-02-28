from __future__ import annotations

import json
import logging
import math
import re
import shutil
import subprocess
from dataclasses import dataclass, field
from typing import Dict, Iterable, List


class PipeWireError(RuntimeError):
    """Raised when a PipeWire command fails."""


@dataclass(slots=True)
class Port:
    id: int
    name: str
    direction: str
    node_id: int
    node_name: str


@dataclass(slots=True)
class Node:
    id: int
    name: str
    description: str
    media_class: str
    process_id: int | None
    ports: List[Port] = field(default_factory=list)

    @property
    def has_output(self) -> bool:
        return any(p.direction == "out" for p in self.ports)

    @property
    def has_input(self) -> bool:
        return any(p.direction == "in" for p in self.ports)

    @property
    def output_ports(self) -> List[Port]:
        return [p for p in self.ports if p.direction == "out"]

    @property
    def input_ports(self) -> List[Port]:
        return [p for p in self.ports if p.direction == "in"]


@dataclass(slots=True)
class Link:
    id: int
    output_node_id: int
    output_port_id: int
    input_node_id: int
    input_port_id: int


@dataclass(slots=True)
class PipeWireSnapshot:
    nodes: Dict[int, Node]
    links: List[Link]

    @property
    def sources(self) -> List[Node]:
        return sorted((n for n in self.nodes.values() if n.has_output), key=lambda n: n.description.lower())

    @property
    def sinks(self) -> List[Node]:
        return sorted((n for n in self.nodes.values() if n.has_input), key=lambda n: n.description.lower())


class PipeWireController:
    VIRTUAL_MIC_SINK_NAME = "audiolink_virtual_mic_sink"
    VIRTUAL_MIC_SOURCE_NAME = "audiolink_virtual_mic"

    def __init__(self) -> None:
        self._logger = logging.getLogger("audiolink.pipewire")
        self._require_command("pw-dump")
        self._require_command("pw-link")
        self._require_command("pactl")
        self._virtual_sink_module_id: int | None = None
        self._virtual_source_module_id: int | None = None

    def _require_command(self, cmd: str) -> None:
        if shutil.which(cmd) is None:
            raise PipeWireError(f"Missing required command: {cmd}")

    def _has_command(self, cmd: str) -> bool:
        return shutil.which(cmd) is not None

    def _run(self, args: Iterable[str]) -> str:
        cmd = list(args)
        self._logger.debug("Running command: %s", " ".join(cmd))
        try:
            result = subprocess.run(
                cmd,
                check=True,
                text=True,
                capture_output=True,
            )
        except FileNotFoundError as exc:
            self._logger.exception("Command not found: %s", " ".join(cmd))
            raise PipeWireError(f"Command not found: {cmd}") from exc
        except subprocess.CalledProcessError as exc:
            stderr = exc.stderr.strip() if exc.stderr else ""
            lower_stderr = stderr.lower()
            if "host is down" in lower_stderr or "failed to connect" in lower_stderr:
                self._logger.error("PipeWire connectivity failure for command: %s | %s", " ".join(cmd), stderr)
                raise PipeWireError(
                    "PipeWire is not reachable for this process. Run AudioLink as the same desktop user "
                    "that owns the active PipeWire session (not root)."
                ) from exc
            self._logger.error("Command failed: %s | %s", " ".join(cmd), stderr)
            raise PipeWireError(f"Command failed: {' '.join(cmd)}; {stderr}") from exc
        self._logger.debug("Command succeeded: %s", " ".join(cmd))
        return result.stdout

    def snapshot(self) -> PipeWireSnapshot:
        raw = self._run(["pw-dump"])
        data = json.loads(raw)

        nodes: Dict[int, Node] = {}
        ports_by_node: Dict[int, List[Port]] = {}
        links: List[Link] = []

        for obj in data:
            info = obj.get("info", {})
            props = info.get("props", {})
            obj_id = obj.get("id")
            obj_type = obj.get("type", "")

            if obj_type.endswith("Node"):
                if not isinstance(obj_id, int):
                    continue
                node_name = props.get("node.name") or props.get("object.path") or f"node-{obj_id}"
                desc = (
                    props.get("node.description")
                    or props.get("node.nick")
                    or props.get("application.name")
                    or node_name
                )
                media_class = props.get("media.class", "")
                pid_raw = props.get("application.process.id")
                try:
                    pid = int(pid_raw) if pid_raw is not None else None
                except (TypeError, ValueError):
                    pid = None

                nodes[obj_id] = Node(
                    id=obj_id,
                    name=str(node_name),
                    description=str(desc),
                    media_class=str(media_class),
                    process_id=pid,
                )

            elif obj_type.endswith("Port"):
                if not isinstance(obj_id, int):
                    continue
                direction = (
                    info.get("direction")
                    or props.get("direction")
                    or props.get("port.direction")
                )
                node_id = self._as_int(
                    info.get("node.id")
                    or props.get("node.id")
                    or obj.get("node.id")
                )
                if direction not in {"in", "out"} or not isinstance(node_id, int):
                    continue

                port_name = (
                    props.get("port.alias")
                    or props.get("port.name")
                    or props.get("object.path")
                    or f"port-{obj_id}"
                )
                port = Port(
                    id=obj_id,
                    name=str(port_name),
                    direction=str(direction),
                    node_id=node_id,
                    node_name="",
                )
                ports_by_node.setdefault(node_id, []).append(port)

            elif obj_type.endswith("Link"):
                if not isinstance(obj_id, int):
                    continue
                output_node_id = self._as_int(
                    info.get("output-node-id") or props.get("link.output.node")
                )
                output_port_id = self._as_int(
                    info.get("output-port-id") or props.get("link.output.port")
                )
                input_node_id = self._as_int(
                    info.get("input-node-id") or props.get("link.input.node")
                )
                input_port_id = self._as_int(
                    info.get("input-port-id") or props.get("link.input.port")
                )

                if not all(isinstance(v, int) for v in [output_node_id, output_port_id, input_node_id, input_port_id]):
                    continue
                links.append(
                    Link(
                        id=obj_id,
                        output_node_id=output_node_id,
                        output_port_id=output_port_id,
                        input_node_id=input_node_id,
                        input_port_id=input_port_id,
                    )
                )

        for node_id, node_ports in ports_by_node.items():
            node = nodes.get(node_id)
            if node is None:
                continue
            for port in node_ports:
                port.node_name = node.name
            node.ports.extend(node_ports)

        if sum(len(n.ports) for n in nodes.values()) == 0:
            self._logger.debug("No ports attached from pw-dump, attempting pw-link port fallback")
            self._attach_ports_from_pw_link(nodes)

        attached_ports = sum(len(n.ports) for n in nodes.values())
        self._logger.debug(
            "Snapshot parsed: nodes=%d ports(attached)=%d links=%d",
            len(nodes),
            attached_ports,
            len(links),
        )

        return PipeWireSnapshot(nodes=nodes, links=links)

    def create_link(self, source: Node, sink: Node) -> None:
        source_port = self._pick_audio_port(source.output_ports)
        sink_port = self._pick_audio_port(sink.input_ports)
        self._run(["pw-link", f"{source.name}:{source_port.name}", f"{sink.name}:{sink_port.name}"])

    def remove_link(self, source: Node, sink: Node) -> None:
        source_port = self._pick_audio_port(source.output_ports)
        sink_port = self._pick_audio_port(sink.input_ports)
        self._run(["pw-link", "-d", f"{source.name}:{source_port.name}", f"{sink.name}:{sink_port.name}"])

    def remove_link_by_ports(self, output_port: Port, input_port: Port) -> None:
        self._run(["pw-link", "-d", f"{output_port.node_name}:{output_port.name}", f"{input_port.node_name}:{input_port.name}"])

    def is_linked(self, source: Node, sink: Node, snapshot: PipeWireSnapshot) -> bool:
        out_ids = {p.id for p in source.output_ports}
        in_ids = {p.id for p in sink.input_ports}
        for link in snapshot.links:
            if link.output_port_id in out_ids and link.input_port_id in in_ids:
                return True
        return False

    def find_sources_by_pid(self, pid: int, snapshot: PipeWireSnapshot) -> List[Node]:
        return [n for n in snapshot.sources if n.process_id == pid]

    def application_sources(self, snapshot: PipeWireSnapshot) -> List[Node]:
        filtered_sources = [n for n in snapshot.sources if not self._is_managed_virtual_node(n)]
        preferred = [n for n in filtered_sources if self._is_application_audio_node(n)]
        if preferred:
            return sorted(preferred, key=lambda n: n.description.lower())
        # Fallback: some setups don't expose full media.class/process metadata.
        fallback = [n for n in filtered_sources if self._looks_routable_node(n)]
        if fallback:
            return sorted(fallback, key=lambda n: n.description.lower())
        return sorted(filtered_sources, key=lambda n: n.description.lower())

    def application_targets(self, snapshot: PipeWireSnapshot) -> List[Node]:
        filtered_targets = [n for n in snapshot.sinks if not self._is_managed_virtual_node(n)]
        preferred = [n for n in filtered_targets if self._is_application_audio_node(n)]
        if preferred:
            return sorted(preferred, key=lambda n: n.description.lower())
        fallback = [n for n in filtered_targets if self._looks_routable_node(n)]
        if fallback:
            return sorted(fallback, key=lambda n: n.description.lower())
        return sorted(filtered_targets, key=lambda n: n.description.lower())

    def virtual_mic_sink_key(self) -> str:
        return self.VIRTUAL_MIC_SINK_NAME

    def virtual_mic_source_key(self) -> str:
        return self.VIRTUAL_MIC_SOURCE_NAME

    def ensure_virtual_microphone(self) -> None:
        self._teardown_named_virtual_microphone()

        sink_id_out = self._run(
            [
                "pactl",
                "load-module",
                "module-null-sink",
                f"sink_name={self.VIRTUAL_MIC_SINK_NAME}",
                "sink_properties=device.description=AudioLink Virtual Mic Sink",
            ]
        ).strip()
        source_id_out = self._run(
            [
                "pactl",
                "load-module",
                "module-remap-source",
                f"master={self.VIRTUAL_MIC_SINK_NAME}.monitor",
                f"source_name={self.VIRTUAL_MIC_SOURCE_NAME}",
                "source_properties=device.description=AudioLink Virtual Microphone",
            ]
        ).strip()

        self._virtual_sink_module_id = self._parse_module_id(sink_id_out, "module-null-sink")
        self._virtual_source_module_id = self._parse_module_id(source_id_out, "module-remap-source")

    def teardown_virtual_microphone(self) -> None:
        # Remove explicitly tracked modules first.
        for module_id in [self._virtual_source_module_id, self._virtual_sink_module_id]:
            if module_id is None:
                continue
            try:
                self._run(["pactl", "unload-module", str(module_id)])
            except PipeWireError:
                self._logger.warning("Failed to unload virtual microphone module id=%s", module_id)

        self._virtual_source_module_id = None
        self._virtual_sink_module_id = None

        # Also remove any stale modules from previous runs.
        self._teardown_named_virtual_microphone()

    def create_link_by_key(self, source_key: str, target_key: str, snapshot: PipeWireSnapshot) -> None:
        source = self._find_node_by_key(source_key, snapshot, source=True)
        target = self._find_node_by_key(target_key, snapshot, source=False)
        self.create_link(source, target)

    def remove_link_by_key(self, source_key: str, target_key: str, snapshot: PipeWireSnapshot) -> None:
        source = self._find_node_by_key(source_key, snapshot, source=True)
        target = self._find_node_by_key(target_key, snapshot, source=False)
        self.remove_link(source, target)

    def set_volume_by_keys(self, source_keys: List[str], snapshot: PipeWireSnapshot, percent: int) -> None:
        if not self._has_command("wpctl"):
            raise PipeWireError("Missing required command for volume control: wpctl")

        volume = max(0, min(100, percent)) / 100.0
        by_name = {n.name: n for n in snapshot.sources}
        errors: list[str] = []
        for key in source_keys:
            node = by_name.get(key)
            if node is None:
                continue
            try:
                self._run(["wpctl", "set-volume", str(node.id), f"{volume:.2f}"])
            except PipeWireError as exc:
                errors.append(f"{node.description}: {exc}")
        if errors:
            raise PipeWireError("; ".join(errors[:3]))

    def apply_target_db_offset_by_keys(
        self,
        target_keys: List[str],
        snapshot: PipeWireSnapshot,
        desired_db: float,
        baselines: Dict[str, float],
    ) -> None:
        if not self._has_command("wpctl"):
            raise PipeWireError("Missing required command for target dB offset control: wpctl")

        by_name = {n.name: n for n in snapshot.sinks}
        errors: list[str] = []
        for key in target_keys:
            node = by_name.get(key)
            if node is None:
                continue
            try:
                baseline = baselines.get(key)
                if baseline is None:
                    baseline = self._get_node_volume_linear(node.id)
                    baselines[key] = baseline

                gain = math.pow(10.0, desired_db / 20.0)
                target_linear = max(0.0, baseline * gain)
                self._run(["wpctl", "set-volume", str(node.id), "--", f"{target_linear:.4f}"])
            except PipeWireError as exc:
                errors.append(f"{node.description}: {exc}")
        if errors:
            raise PipeWireError("; ".join(errors[:3]))

    def _get_node_volume_linear(self, node_id: int) -> float:
        out = self._run(["wpctl", "get-volume", str(node_id)])
        # Example: "Volume: 0.73"
        m = re.search(r"Volume:\s*([0-9]*\.?[0-9]+)", out)
        if not m:
            raise PipeWireError(f"Could not parse current volume for node {node_id}: {out.strip()}")
        try:
            return float(m.group(1))
        except ValueError as exc:
            raise PipeWireError(f"Invalid volume value from wpctl for node {node_id}") from exc

    def _find_node_by_key(self, key: str, snapshot: PipeWireSnapshot, source: bool) -> Node:
        candidates = snapshot.sources if source else snapshot.sinks
        for node in candidates:
            if node.name == key:
                return node
        raise PipeWireError(f"Node not available: {key}")

    @staticmethod
    def _is_application_audio_node(node: Node) -> bool:
        media_class = node.media_class.lower()
        if "audio" not in media_class:
            return False
        # Prefer real app streams, but also accept stream-like app nodes with weak metadata.
        if "stream" in media_class:
            return True
        desc = node.description.lower()
        name = node.name.lower()
        if any(token in desc or token in name for token in ["firefox", "chrome", "discord", "spotify", "vlc"]):
            return True
        return node.process_id is not None

    @classmethod
    def _is_managed_virtual_node(cls, node: Node) -> bool:
        return node.name in {cls.VIRTUAL_MIC_SINK_NAME, cls.VIRTUAL_MIC_SOURCE_NAME}

    @staticmethod
    def _looks_routable_node(node: Node) -> bool:
        # Last-resort display filter: keep nodes that expose at least one audio-ish port label.
        for port in node.ports:
            low = port.name.lower()
            if any(token in low for token in ["audio", "playback", "capture", "input", "output", "monitor"]):
                return True
        return bool(node.ports)

    @staticmethod
    def _as_int(value: object) -> int | None:
        try:
            if value is None:
                return None
            return int(value)
        except (TypeError, ValueError):
            return None

    def _attach_ports_from_pw_link(self, nodes: Dict[int, Node]) -> None:
        synthetic_id = -1

        for direction, cmd in (("out", ["pw-link", "-o"]), ("in", ["pw-link", "-i"])):
            try:
                out = self._run(cmd)
            except PipeWireError:
                self._logger.exception("Failed to query ports with: %s", " ".join(cmd))
                continue

            for raw_line in out.splitlines():
                line = raw_line.strip()
                if not line or ":" not in line:
                    continue

                token = self._extract_port_token(line)
                if token is None or ":" not in token:
                    continue

                node_name, port_name = token.split(":", 1)
                node_id = self._find_node_id_by_name(node_name, nodes)
                if node_id is None:
                    continue

                node = nodes[node_id]
                if any(p.name == port_name and p.direction == direction for p in node.ports):
                    continue

                node.ports.append(
                    Port(
                        id=synthetic_id,
                        name=port_name,
                        direction=direction,
                        node_id=node_id,
                        node_name=node.name,
                    )
                )
                synthetic_id -= 1

        self._logger.debug(
            "pw-link fallback attached ports: %d",
            sum(len(n.ports) for n in nodes.values()),
        )

    @staticmethod
    def _extract_port_token(line: str) -> str | None:
        # Typical pw-link formats include one token with "node:port"; keep parsing tolerant.
        for token in line.split():
            if ":" in token and not token.startswith(("->", "<-")):
                return token
        # Fallback: use the whole line if it still looks like node:port.
        return line if ":" in line else None

    @staticmethod
    def _find_node_id_by_name(name: str, nodes: Dict[int, Node]) -> int | None:
        for node in nodes.values():
            if node.name == name:
                return node.id
        lname = name.lower()
        for node in nodes.values():
            nlow = node.name.lower()
            if nlow == lname or nlow.endswith(lname) or lname.endswith(nlow):
                return node.id
        return None

    @staticmethod
    def _pick_audio_port(ports: List[Port]) -> Port:
        if not ports:
            raise PipeWireError("No compatible ports available")

        # Prefer ports that look like audio channels.
        for p in ports:
            low = p.name.lower()
            if "audio" in low or "monitor" in low or "playback" in low or "capture" in low:
                return p
        return ports[0]

    @staticmethod
    def _parse_module_id(raw: str, module_name: str) -> int:
        try:
            return int(raw.strip())
        except ValueError as exc:
            raise PipeWireError(f"Unexpected module id output for {module_name}: {raw!r}") from exc

    def _teardown_named_virtual_microphone(self) -> None:
        modules = self._run(["pactl", "list", "short", "modules"])
        sink_module_ids: list[int] = []
        source_module_ids: list[int] = []

        for line in modules.splitlines():
            # Format: "<id>\t<module_name>\t<args>"
            parts = line.split("\t", 2)
            if len(parts) < 3:
                continue
            module_id_raw, module_name, args = parts
            try:
                module_id = int(module_id_raw)
            except ValueError:
                continue
            if module_name == "module-remap-source" and f"source_name={self.VIRTUAL_MIC_SOURCE_NAME}" in args:
                source_module_ids.append(module_id)
            elif module_name == "module-null-sink" and f"sink_name={self.VIRTUAL_MIC_SINK_NAME}" in args:
                sink_module_ids.append(module_id)

        # Unload source modules before sink modules due to dependency on sink monitor.
        for module_id in source_module_ids + sink_module_ids:
            self._run(["pactl", "unload-module", str(module_id)])
