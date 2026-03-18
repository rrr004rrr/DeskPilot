"""
Script Manager Module
Handles saving, loading, and listing of JSON scripts and window configs.
"""
import json
import os
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent.parent / "data" / "scripts"


class ScriptManager:
    def __init__(self, scripts_dir: Path | None = None):
        self._dir = Path(scripts_dir) if scripts_dir else SCRIPTS_DIR
        self._dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Scripts
    # ------------------------------------------------------------------
    def save_script(self, name: str, events: list[dict]) -> Path:
        path = self._dir / f"{name}.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(events, f, ensure_ascii=False, indent=2)
        return path

    def load_script(self, name: str) -> list[dict]:
        path = self._dir / f"{name}.json"
        if not path.exists():
            raise FileNotFoundError(f"Script not found: {name}")
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def load_script_from_path(self, path: str | Path) -> list[dict]:
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def list_scripts(self) -> list[str]:
        return sorted(
            p.stem for p in self._dir.glob("*.json")
            if not p.name.endswith(".window.json")
        )

    def delete_script(self, name: str):
        path = self._dir / f"{name}.json"
        if path.exists():
            path.unlink()

    def script_path(self, name: str) -> Path:
        return self._dir / f"{name}.json"

    # ------------------------------------------------------------------
    # Window config (stored alongside scripts)
    # ------------------------------------------------------------------
    def save_window_config(self, name: str, config: dict) -> Path:
        path = self._dir / f"{name}.window.json"
        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return path

    def load_window_config(self, name: str) -> dict | None:
        path = self._dir / f"{name}.window.json"
        if not path.exists():
            return None
        with open(path, encoding="utf-8") as f:
            return json.load(f)

    def list_window_configs(self) -> list[str]:
        return sorted(
            p.name.removesuffix(".window.json")
            for p in self._dir.glob("*.window.json")
        )

    def delete_window_config(self, name: str):
        path = self._dir / f"{name}.window.json"
        if path.exists():
            path.unlink()
