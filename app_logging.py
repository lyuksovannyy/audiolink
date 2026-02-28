from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import QtMsgType, qInstallMessageHandler


def configure_logging() -> Path:
    log_dir = Path.home() / ".cache" / "audiolink"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "audiolink.log"

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    # Avoid duplicate handlers on repeated in-process runs.
    for h in list(root.handlers):
        root.removeHandler(h)

    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)

    root.addHandler(fh)
    root.addHandler(ch)

    logging.getLogger(__name__).info("Logging initialized: %s", log_file)
    return log_file


def install_global_hooks() -> None:
    import sys
    import traceback

    logger = logging.getLogger("audiolink.hooks")

    def _excepthook(exc_type, exc_value, exc_tb):
        logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))
        traceback.print_exception(exc_type, exc_value, exc_tb)

    sys.excepthook = _excepthook

    def _qt_handler(mode: QtMsgType, context, message: Optional[str]) -> None:
        text = message or ""
        if mode == QtMsgType.QtDebugMsg:
            logger.debug("Qt: %s", text)
        elif mode == QtMsgType.QtInfoMsg:
            logger.info("Qt: %s", text)
        elif mode == QtMsgType.QtWarningMsg:
            logger.warning("Qt: %s", text)
        elif mode == QtMsgType.QtCriticalMsg:
            logger.error("Qt: %s", text)
        else:
            logger.critical("Qt: %s", text)

    qInstallMessageHandler(_qt_handler)


def runtime_context() -> dict[str, str]:
    return {
        "uid": str(os.geteuid()) if hasattr(os, "geteuid") else "n/a",
        "user": os.environ.get("USER", ""),
        "display": os.environ.get("DISPLAY", ""),
        "wayland": os.environ.get("WAYLAND_DISPLAY", ""),
        "xdg_runtime_dir": os.environ.get("XDG_RUNTIME_DIR", ""),
    }
