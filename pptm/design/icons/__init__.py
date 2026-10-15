"""라인 아이콘(Lucide, ISC) → 테마색 리컬러 + 고해상도 PNG 래스터.

번들 `lucide.json`(name → SVG 문자열, 1998종)을 읽어 `currentColor`를 지정 색으로
치환하고 cairosvg로 투명배경 PNG를 굽는다. 결과는 프로젝트 assets에 캐시한다.
아이콘 폰트(TTF)와 달리 래스터가 pptx에 박히므로 폰트 없는 PC에서도 깨지지 않는다.

브랜드 로고(Claude·OpenAI·Figma 등)는 별도 번들 `brand_icons/`에 있다 —
`brand_available`/`brand_png`/`brand_mono_png` 참조. 출처·라이선스는
`docs/design-refs-brand-icons.md`.
"""
from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path

_BUNDLE = Path(__file__).with_name("lucide.json")
_FALLBACK_BUNDLE = Path(__file__).with_name("tabler.json")
_BRAND_DIR = Path(__file__).resolve().parents[1] / "brand_icons"
_FLAG_DIR = Path(__file__).resolve().parents[1] / "flags"


@lru_cache(maxsize=1)
def _icons() -> dict[str, str]:
    return json.loads(_BUNDLE.read_text(encoding="utf-8"))


@lru_cache(maxsize=1)
def _fallback_icons() -> dict[str, str]:
    """Tabler Icons(MIT) — Lucide에 없는 이름일 때만 쓰는 폴백. 문법이 같아
    `currentColor` 치환 코드를 그대로 공유한다."""
    return json.loads(_FALLBACK_BUNDLE.read_text(encoding="utf-8")) \
        if _FALLBACK_BUNDLE.exists() else {}


@lru_cache(maxsize=1)
def _brand_manifest() -> dict:
    man = _BRAND_DIR / "manifest.json"
    return json.loads(man.read_text(encoding="utf-8")) if man.exists() else {"color": {}, "mono": {}}


def available(name: str) -> bool:
    return name in _icons() or name in _fallback_icons()


def names() -> list[str]:
    return sorted(_icons())


def fallback_names() -> list[str]:
    """Lucide에 없어서 Tabler 폴백으로만 잡히는 이름들."""
    return sorted(set(_fallback_icons()) - set(_icons()))


def brand_available(slug: str, mono: bool = False) -> bool:
    kind = "mono" if mono else "color"
    return slug in _brand_manifest().get(kind, {})


def brand_names(mono: bool = False) -> list[str]:
    kind = "mono" if mono else "color"
    return sorted(_brand_manifest().get(kind, {}))


def brand_png(assets_dir: Path, slug: str, variant: str = "light",
             px: int = 256) -> Path | None:
    """브랜드 로고(풀컬러) PNG — 이미 완성색이라 리컬러하지 않고 리사이즈만 한다."""
    variant = "dark" if variant == "dark" else "light"
    src = _BRAND_DIR / variant / f"{slug}.png"
    if not src.exists():
        return None
    from PIL import Image
    key = f"brand_{slug}_{variant}_{px}"
    out = assets_dir / f"{key}.png"
    if out.exists():
        return out
    assets_dir.mkdir(parents=True, exist_ok=True)
    im = Image.open(src).convert("RGBA")
    im = im.resize((px, px), Image.LANCZOS)
    im.save(out)
    return out


def brand_mono_png(assets_dir: Path, slug: str, color: str = "1E2A3A",
                   px: int = 256) -> Path | None:
    """브랜드 로고(단색, Simple Icons CC0) → 테마색 리컬러 PNG."""
    src = _BRAND_DIR / "mono" / f"{slug}.svg"
    if not src.exists():
        return None
    import cairosvg

    hexc = _norm_color(color)
    key = f"brandmono_{slug}_{hexc.lstrip('#')}_{px}"
    out = assets_dir / f"{key}.png"
    if out.exists():
        return out

    svg = src.read_text(encoding="utf-8")
    # Simple Icons 는 currentColor 를 안 쓰고 path에 fill 속성이 없다 —
    # 루트 <svg> 에 fill 을 얹으면 자식 path 전체에 상속된다.
    svg = svg.replace("<svg ", f'<svg fill="{hexc}" ', 1)

    assets_dir.mkdir(parents=True, exist_ok=True)
    cairosvg.svg2png(bytestring=svg.encode("utf-8"), write_to=str(out),
                     output_width=px, output_height=px)
    return out


def flag_available(cc: str) -> bool:
    return (_FLAG_DIR / "4x3" / f"{cc.lower()}.svg").exists()


def flag_png(assets_dir: Path, cc: str, shape: str = "4x3",
             px: int = 256) -> Path | None:
    """국기(flag-icons, MIT) — 원색 그대로 래스터. `shape`: 4x3(가로) | 1x1(정사각)."""
    shape = "1x1" if shape == "1x1" else "4x3"
    cc = cc.lower()
    src = _FLAG_DIR / shape / f"{cc}.svg"
    if not src.exists():
        return None
    import cairosvg

    key = f"flag_{shape}_{cc}_{px}"
    out = assets_dir / f"{key}.png"
    if out.exists():
        return out
    assets_dir.mkdir(parents=True, exist_ok=True)
    h = px if shape == "1x1" else round(px * 3 / 4)
    cairosvg.svg2png(url=str(src), write_to=str(out),
                     output_width=px, output_height=h)
    return out


def _norm_color(color: str) -> str:
    c = color.lstrip("#").strip()
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    return "#" + c.upper()


def icon_png(assets_dir: Path, name: str, color: str = "1E2A3A",
             px: int = 256, weight: float | None = None) -> Path | None:
    """아이콘 이름→리컬러 PNG 경로. Lucide에 없으면 Tabler(MIT)를 본다.

    둘 다 24×24·`stroke="currentColor"`·`stroke-width 2` 문법이라 같은 치환
    코드로 리컬러된다 — 한 덱에 섞여도 선 굵기·둥근 정도가 튀지 않는다.
    """
    svg = _icons().get(name)
    src = "lucide"
    if svg is None:
        svg = _fallback_icons().get(name)
        src = "tabler"
    if svg is None:
        return None
    import cairosvg  # 지연 import: 아이콘 안 쓰면 의존성 불필요

    hexc = _norm_color(color)
    key = f"icon_{src}_{name}_{hexc.lstrip('#')}_{px}"
    if weight is not None:
        key += f"_w{weight:g}"
    out = assets_dir / f"{key}.png"
    if out.exists():
        return out

    svg = svg.replace("currentColor", hexc)
    if weight is not None:
        svg = re.sub(r'stroke-width="[^"]*"', f'stroke-width="{weight:g}"', svg)

    assets_dir.mkdir(parents=True, exist_ok=True)
    cairosvg.svg2png(bytestring=svg.encode("utf-8"), write_to=str(out),
                     output_width=px, output_height=px)
    return out
