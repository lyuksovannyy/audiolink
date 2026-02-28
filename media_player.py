from __future__ import annotations

import logging
import os
from pathlib import Path

from PyQt6.QtCore import QUrl
from PyQt6.QtMultimedia import QAudioOutput, QMediaPlayer


class MediaPlayerController:
    def __init__(self) -> None:
        self._logger = logging.getLogger("audiolink.media")
        self._player: QMediaPlayer | None = None
        self._audio_output: QAudioOutput | None = None
        self._current_file: Path | None = None

    @property
    def player(self) -> QMediaPlayer:
        return self._ensure_backend()

    @property
    def current_file(self) -> Path | None:
        return self._current_file

    @property
    def process_id(self) -> int:
        return os.getpid()

    def _ensure_backend(self) -> QMediaPlayer:
        if self._player is not None:
            return self._player

        self._logger.info("Initializing Qt multimedia backend")
        player = QMediaPlayer()
        audio_output = QAudioOutput()
        player.setAudioOutput(audio_output)

        self._player = player
        self._audio_output = audio_output
        self._logger.info("Qt multimedia backend initialized")
        return player

    def load_file(self, file_path: str) -> None:
        path = Path(file_path).expanduser().resolve()
        self._ensure_backend().setSource(QUrl.fromLocalFile(str(path)))
        self._current_file = path

    def play(self) -> None:
        self._ensure_backend().play()

    def pause(self) -> None:
        self._ensure_backend().pause()

    def stop(self) -> None:
        self._ensure_backend().stop()

    def set_volume(self, value: int) -> None:
        self._ensure_backend()
        assert self._audio_output is not None
        # QAudioOutput volume is 0.0-1.0.
        self._audio_output.setVolume(max(0.0, min(1.0, value / 100.0)))
