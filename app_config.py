from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class AppConfig:
    auto_select_sources: set[str]
    auto_select_source_items: set[str]
    auto_select_targets: set[str]


def _config_path() -> Path:
    return Path(__file__).resolve().parent / "config.json"


def load_config() -> AppConfig:
    path = _config_path()
    if not path.exists():
        return AppConfig(
            auto_select_sources=set(),
            auto_select_source_items=set(),
            auto_select_targets=set(),
        )

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return AppConfig(
            auto_select_sources=set(),
            auto_select_source_items=set(),
            auto_select_targets=set(),
        )

    src = data.get("auto_select_sources", [])
    src_items = data.get("auto_select_source_items", [])
    dst = data.get("auto_select_targets", [])
    return AppConfig(
        auto_select_sources={s for s in src if isinstance(s, str)},
        auto_select_source_items={s for s in src_items if isinstance(s, str)},
        auto_select_targets={s for s in dst if isinstance(s, str)},
    )


def save_config(cfg: AppConfig) -> None:
    path = _config_path()
    payload = {
        "auto_select_sources": sorted(cfg.auto_select_sources),
        "auto_select_source_items": sorted(cfg.auto_select_source_items),
        "auto_select_targets": sorted(cfg.auto_select_targets),
    }
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
