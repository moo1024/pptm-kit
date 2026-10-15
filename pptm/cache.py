from __future__ import annotations

import hashlib
import json
from pathlib import Path


class ArtifactCache:
    """산출물 옆에 .meta.json 사이드카를 두고 입력 해시로 캐시 유효성을 판단."""

    def __init__(self, directory: Path) -> None:
        self.dir = Path(directory)

    def _meta_path(self, artifact: str) -> Path:
        return self.dir / (Path(artifact).name + ".meta.json")

    @staticmethod
    def _hash(inputs: dict) -> str:
        payload = json.dumps(inputs, sort_keys=True, ensure_ascii=False)
        return hashlib.sha256(payload.encode()).hexdigest()

    def is_fresh(self, artifact: str, *, inputs: dict, force: bool = False) -> bool:
        """산출물 파일 + 메타 모두 존재하고 입력 해시가 일치하면 True."""
        if force:
            return False
        meta_path = self._meta_path(artifact)
        artifact_path = self.dir / artifact
        if not meta_path.exists() or not artifact_path.exists():
            return False
        try:
            meta = json.loads(meta_path.read_text())
        except (json.JSONDecodeError, OSError):
            return False
        return meta.get("input_hash") == self._hash(inputs)

    def save_meta(self, artifact: str, *, inputs: dict, usage: dict) -> None:
        """산출물 생성 후 사이드카 기록."""
        self.dir.mkdir(parents=True, exist_ok=True)
        meta = {"input_hash": self._hash(inputs), "usage": usage}
        self._meta_path(artifact).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2)
        )
