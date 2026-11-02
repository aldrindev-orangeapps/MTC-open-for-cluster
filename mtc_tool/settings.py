from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

from .paths import default_output_folder, settings_path


@dataclass
class Settings:
    reference_path: str = ""
    output_folder: str = ""

    @classmethod
    def load(cls, path: Path | None = None) -> Settings:
        try:
            data = json.loads((path or settings_path()).read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("Invalid settings")
            return cls(**{k: v for k, v in data.items()
                          if k in cls.__dataclass_fields__ and isinstance(v, str)})
        except (OSError, ValueError, TypeError):
            return cls()

    @property
    def output_directory(self) -> Path:
        return Path(self.output_folder) if self.output_folder else default_output_folder()

    def save(self, path: Path | None = None) -> None:
        target = path or settings_path()
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_suffix(".tmp")
        temporary.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")
        temporary.replace(target)
