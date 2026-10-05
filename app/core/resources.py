"""Localização de recursos no código-fonte e no aplicativo empacotado."""
import sys
from pathlib import Path


def resource_base() -> Path:
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS)
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def resource_path(relative_path: str) -> str:
    relative = Path(relative_path.replace("\\", "/"))
    base = resource_base()
    candidate = base / relative
    if candidate.exists():
        return str(candidate)
    if getattr(sys, "frozen", False):
        fallback = Path(sys.executable).resolve().parent / relative
        if fallback.exists():
            return str(fallback)
    return str(candidate)
