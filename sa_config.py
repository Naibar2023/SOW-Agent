"""Loads settings from the .env file next to the program."""

from __future__ import annotations

import os
from pathlib import Path

_HERE = Path(__file__).resolve().parent
ROOT = _HERE.parent if _HERE.name == "sow_agent" else _HERE


def load_env(path=None):
    """Read KEY=value lines from .env into the environment (existing variables win)."""
    path = Path(path) if path else ROOT / ".env"
    if not path.is_file():
        return False
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key, value = key.strip(), value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        if key and value and key not in os.environ:
            os.environ[key] = value
    return True


def workspace_dir():
    """Folder for projects and Word files. Falls back to a temp folder if the program folder is read-only."""
    chosen = Path(os.environ.get("SOW_WORKSPACE") or ROOT / "workspace")
    try:
        chosen.mkdir(parents=True, exist_ok=True)
        probe = chosen / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return chosen
    except OSError:
        import tempfile
        fallback = Path(tempfile.gettempdir()) / "sow_agent_workspace"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback
