"""도형 병합(`merge` 아이템) — PowerPoint '도형 병합'(합치기·빼기·교차·조각·결합)을 피마2 문법으로.

유튜브 강좌 기법 `shape-merge`(17개 영상) · 스포트라이트 · 마스크 드러내기 · 글자 모양 창(글자 안 영상)의 재료.
결과는 편집 가능한 자유형 도형 하나(조각이면 여럿) — PowerPoint 가 병합 결과로 남기는 것과 같은 custGeom 이다.
구멍은 한 도형 안의 바깥 윤곽 + 안쪽 윤곽(여러 contour)으로 그린다.

    {"type": "merge", "op": "subtract", "fill": "111111", "alpha": 0.75, "morph_id": "spot",
     "shapes": [{"kind": "rect", "x": 0, "y": 0, "w": 1, "h": 1},
                {"kind": "ellipse", "x": .4, "y": .3, "w": .2, "h": .36}]}

kind — rect(+radius 짧은 변 비율) · ellipse · poly(points) · text(text·size pt·font 이름 또는 파일)
op   — union(합치기) · subtract(빼기: 첫 도형에서 나머지) · intersect(교차) · xor(결합) · fragment(조각)
"""
from __future__ import annotations

import subprocess
from functools import lru_cache
from pathlib import Path

from shapely import affinity
from shapely.geometry import MultiPolygon, Point, Polygon, box
from shapely.ops import unary_union

OPS = ("union", "subtract", "intersect", "xor", "fragment")


@lru_cache(maxsize=None)
def _font_file(name: str) -> str:
    p = Path(name)
    if p.suffix.lower() in (".ttf", ".otf") and p.exists():
        return str(p)
    out = subprocess.run(["fc-match", "-f", "%{file}", name], capture_output=True, text=True).stdout.strip()
    if not out:
        raise ValueError(f"merge text 글꼴 {name!r} 를 못 찾았다(fc-match)")
    return out


class _Flatten:
    """fontTools 펜 — 곡선을 잘게 꺾어 윤곽(점 목록)으로."""

    def __init__(self, glyphset, n=8):
        from fontTools.pens.basePen import BasePen

        outer = self

        class Pen(BasePen):
            def _moveTo(self, pt):
                outer.cur = [pt]
                outer.rings.append(outer.cur)

            def _lineTo(self, pt):
                outer.cur.append(pt)

            def _curveToOne(self, p1, p2, p3):
                p0 = outer.cur[-1]
                for i in range(1, n + 1):
                    t = i / n
                    mt = 1 - t
                    outer.cur.append((mt ** 3 * p0[0] + 3 * mt * mt * t * p1[0] + 3 * mt * t * t * p2[0] + t ** 3 * p3[0],
                                      mt ** 3 * p0[1] + 3 * mt * mt * t * p1[1] + 3 * mt * t * t * p2[1] + t ** 3 * p3[1]))

            def _qCurveToOne(self, p1, p2):
                p0 = outer.cur[-1]
                for i in range(1, n + 1):
                    t = i / n
                    mt = 1 - t
                    outer.cur.append((mt * mt * p0[0] + 2 * mt * t * p1[0] + t * t * p2[0],
                                      mt * mt * p0[1] + 2 * mt * t * p1[1] + t * t * p2[1]))

            def _closePath(self):
                pass

        self.rings, self.cur = [], None
        self.pen = Pen(glyphset)


def _text_geom(sh, W, H):
    from fontTools.ttLib import TTFont
    font = TTFont(_font_file(sh.get("font", "Pretendard:weight=800")), fontNumber=0)
    upm = font["head"].unitsPerEm
    gs = font.getGlyphSet()
    cmap = font.getBestCmap()
    size_emu = float(sh.get("size", 120)) * 12700
    k = size_emu / upm
    asc = font["hhea"].ascent
    x0, y0 = sh.get("x", 0) * W, sh.get("y", 0) * H + asc * k
    tracking = float(sh.get("tracking", 0)) * 12700
    polys = []
    pen_x = 0.0
    for ch in str(sh["text"]):
        gname = cmap.get(ord(ch))
        if gname is None:
            continue
        fl = _Flatten(gs)
        gs[gname].draw(fl.pen)
        g = None
        for ring in fl.rings:
            if len(ring) < 3:
                continue
            pts = [(x0 + (pen_x + px) * k, y0 - py * k) for px, py in ring]
            poly = Polygon(pts).buffer(0)
            g = poly if g is None else g.symmetric_difference(poly)      # 안쪽 윤곽 = 구멍(ㅇ·ㅁ·O)
        if g is not None:
            polys.append(g)
        pen_x += gs[gname].width + tracking / k
    return unary_union(polys) if polys else Polygon()


def geom(sh, W, H):
    kind = sh.get("kind", "rect")
    if kind == "text":
        return _text_geom(sh, W, H)
    if kind == "poly":
        return Polygon([(px * W, py * H) for px, py in sh["points"]]).buffer(0)
    x, y, w, h = sh["x"] * W, sh["y"] * H, sh["w"] * W, sh["h"] * H
    if kind == "ellipse":
        c = Point(x + w / 2, y + h / 2).buffer(1.0, quad_segs=32)
        return affinity.scale(c, w / 2, h / 2)
    if kind == "rect":
        r = float(sh.get("radius", 0)) * min(w, h) / 2
        if r > 0:
            return box(x + r, y + r, x + w - r, y + h - r).buffer(r, quad_segs=12)
        return box(x, y, x + w, y + h)
    raise ValueError(f"merge shape kind {kind!r} — rect · ellipse · poly · text")


def combine(item, W, H) -> list:
    op = item.get("op", "union")
    if op not in OPS:
        raise ValueError(f"merge op {op!r} — 허용: {OPS}")
    gs = [geom(s, W, H) for s in item["shapes"]]
    if not gs:
        raise ValueError("merge 는 shapes 가 하나 이상 있어야 한다")
    if op == "union":
        return [unary_union(gs)]
    if op == "subtract":
        return [gs[0].difference(unary_union(gs[1:]))] if len(gs) > 1 else [gs[0]]
    if op == "intersect":
        out = gs[0]
        for g in gs[1:]:
            out = out.intersection(g)
        return [out]
    if op == "xor":
        out = gs[0]
        for g in gs[1:]:
            out = out.symmetric_difference(g)
        return [out]
    # fragment — 겹친 모든 조각을 따로(PowerPoint '조각')
    from shapely.ops import polygonize
    lines = unary_union([g.boundary for g in gs])
    whole = unary_union(gs)
    return [p for p in polygonize(lines) if whole.contains(p.representative_point())]


def _rings(g):
    polys = list(g.geoms) if isinstance(g, MultiPolygon) else [g] if isinstance(g, Polygon) else \
        [p for p in getattr(g, "geoms", []) if isinstance(p, Polygon)]
    for p in polys:
        if p.is_empty:
            continue
        yield list(p.exterior.coords)[:-1]
        for hole in p.interiors:
            yield list(hole.coords)[:-1]


def add_merge(slide, item, W, H):
    """병합 결과 → 자유형 도형(구멍은 같은 경로 안의 안쪽 윤곽). 만든 도형 목록을 돌려준다."""
    out = []
    for g in combine(item, W, H):
        rings = [r for r in _rings(g) if len(r) >= 3]
        if not rings:
            continue
        r0 = rings[0]
        fb = slide.shapes.build_freeform(int(r0[0][0]), int(r0[0][1]), scale=1.0)
        fb.add_line_segments([(int(a), int(b)) for a, b in r0[1:]], close=True)
        for r in rings[1:]:
            fb.move_to(int(r[0][0]), int(r[0][1]))
            fb.add_line_segments([(int(a), int(b)) for a, b in r[1:]], close=True)
        out.append(fb.convert_to_shape())
    return out
