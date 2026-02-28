from __future__ import annotations

import os
import sys
import logging

from PyQt6.QtCore import QTimer
from PyQt6.QtWidgets import QApplication, QMessageBox

from app_logging import configure_logging, install_global_hooks, runtime_context
from media_player import MediaPlayerController
from pipewire_controller import PipeWireController, PipeWireError
from ui import build_window


def main() -> int:
    log_file = configure_logging()
    install_global_hooks()
    logger = logging.getLogger("audiolink.main")
    logger.info("AudioLink starting")
    logger.info("Runtime context: %s", runtime_context())
    logger.info("Python: %s", sys.version.replace("\n", " "))
    logger.info("Log file: %s", log_file)

    if hasattr(os, "geteuid") and os.geteuid() == 0:
        logger.warning(
            "Running as root. PipeWire is per-user, so routing usually fails unless root has its own session."
        )

    app = QApplication(sys.argv)
    logger.info("QApplication created")

    logger.info("Creating PipeWireController")
    try:
        controller = PipeWireController()
    except PipeWireError as exc:
        logger.exception("PipeWire initialization failed")
        QMessageBox.critical(None, "AudioLink", f"PipeWire initialization failed: {exc}")
        return 1

    logger.info("PipeWireController ready")

    logger.info("Creating MediaPlayerController")
    media = MediaPlayerController()
    logger.info("MediaPlayerController ready")
    logger.info("Building MainWindow")
    try:
        window = build_window(controller=controller, media=media)
    except Exception as exc:
        logger.exception("UI initialization failed")
        QMessageBox.critical(None, "AudioLink", f"UI initialization failed: {exc}")
        return 1

    logger.info("Main window built, showing")
    window.show()
    app.processEvents()
    logger.info("Window visible=%s active=%s", window.isVisible(), window.isActiveWindow())
    QTimer.singleShot(
        1000,
        lambda: logger.info(
            "Window probe: visible=%s minimized=%s geometry=%s",
            window.isVisible(),
            window.isMinimized(),
            window.geometry().getRect(),
        ),
    )

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
