"""페이지 목록 — 한 프로젝트에 어떤 면이 있는지 판정하는 유일한 자리.

한 곳에서만 센다. 세는 코드가 흩어지면 **자마다 다른 덱을 보게 된다.**

    2026-08-27 실측. `p[0-9][0-9]` 라는 두 자리 가정이 네 군데에 각자 박혀 있었고
    (렌더러·게이트들), 156장짜리 덱에서
    p100~p156 **57장이 어느 게이트에도 안 걸렸다.** 그런데 화면에는
    `✓ 정렬 검수 통과 (0건)` 이 찍혔다 — 봐서 0건이 아니라 못 봐서 0건이다.
    게이트가 조용히 통과하는 게 게이트가 없는 것보다 나쁘다.

정렬은 **숫자 순**이다. 문자열 정렬은 세 자리가 들어오는 순간 깨진다
(p01, p100, p101, p11, …). 두 자리 시절엔 우연히 맞았을 뿐이라 버그가 안 보였다.
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# pNN·pNNN 전부 받는다. 자릿수를 고정하지 않는 것이 이 모듈의 존재 이유다.
PAGE_RE = re.compile(r"^p(\d+)$")
SPEC_NAME = "layout_spec.json"


def _proj(target: str | Path) -> Path:
    """슬러그든 프로젝트 경로든 프로젝트 디렉터리로 만든다."""
    if isinstance(target, Path):
        return target
    if "/" in target or "\\" in target:
        return Path(target)
    return ROOT / "projects" / target


def page_dirs(target: str | Path) -> list[Path]:
    """면 디렉터리 목록, 면 번호 오름차순. 없으면 빈 리스트."""
    base = _proj(target) / "pages"
    if not base.is_dir():
        return []
    found = []
    for d in base.iterdir():
        m = PAGE_RE.match(d.name)
        if m and d.is_dir():
            found.append((int(m.group(1)), d))
    return [d for _, d in sorted(found)]


def page_numbers(target: str | Path) -> list[int]:
    """면 번호 목록 (p04 → 4), 오름차순."""
    return [int(PAGE_RE.match(d.name).group(1)) for d in page_dirs(target)]


def spec_files(target: str | Path) -> list[Path]:
    """layout_spec.json 이 **실제로 있는** 면의 스펙 경로, 면 번호 오름차순.

    디렉터리만 있고 스펙이 없는 면은 빠진다 — 스펙을 읽는 자가 쓴다.
    """
    return [sp for d in page_dirs(target) if (sp := d / SPEC_NAME).is_file()]


def count(target: str | Path) -> int:
    """면 디렉터리 개수. 썸네일 번호 상한처럼 '몇 장인가'만 필요할 때."""
    return len(page_dirs(target))
