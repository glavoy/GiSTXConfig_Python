from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from models import SurveyManifest


def clean_none(value: Any) -> Any:
    if isinstance(value, list):
        return [clean_none(v) for v in value if v is not None]
    if hasattr(value, "__dataclass_fields__"):
        result = {}
        for key, v in asdict(value).items():
            if v is None:
                continue
            result[key] = clean_none(v)
        return result
    if isinstance(value, dict):
        return {k: clean_none(v) for k, v in value.items() if v is not None}
    return value


class JsonGenerator:
    @staticmethod
    def write_manifest(path: Path, manifest: SurveyManifest) -> None:
        payload = clean_none(manifest)
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
