"""도식 골격 생성기 — `figure` 아이템을 기존 프리미티브로 전개한다.

AntV Infographic(antvis/Infographic)의 **구조(structure) × 아이템(item)** 조합 문법을
우리 파이프라인에 이식한 것이다. 원전은 SVG를 그리지만 그건 pptx에서 이미지가 되어
텍스트 편집이 죽는다 — 그래서 **문법만** 가져오고, 결과물은 전부 네이티브 개체다.

    {"type": "figure", "figure": "funnel", "item": "compact_card",
     "x": 0.07, "y": 0.28, "w": 0.86, "h": 0.50,
     "nodes": [{"title": "인지", "value": "12,000", "caption": "노출에서 시작"}, ...]}

전개 결과는 card/text/chip/poly/arrow/hairline/pill 같은 **이미 있는 아이템**이라
렌더러도 게이트도 손댈 필요가 없다. 좌표는 전부 계산이다 — 이 파일의 존재 이유가
"손좌표를 없애는 것"이므로, 여기에 눈대중 상수를 넣지 마라.

설계 규약 (어기면 게이트가 잡는다)
  · 면 안의 글자는 **면과 같은 박스 + anchor:"middle"** 한 덩어리로 낸다 (절대규칙 6).
  · 노드 하나당 텍스트 아이템은 **하나**다 — 제목·수치·캡션은 runs 의 줄로 쌓는다.
  · 도식은 주어진 bbox 를 **채운다**. 위쪽에만 붙고 아래가 비면 그게 반려 사유다.
  · 가로 분배는 판면 폭 균등분배 — 좌우 여백이 자동으로 같아진다 (절대규칙 8).
"""
from __future__ import annotations

import math

__all__ = ["FIGURES", "SKINS", "expand_figures", "dark_view"]

GAP = 0.024          # 카드 사이 기본 간극
CHIP_D = 0.046       # 넘버 서클 지름 (절대규칙 8 표준값)
ARROW_H = 0.012
PAD = 0.018          # 면 안쪽 여백


# ── 색 ──────────────────────────────────────────────────────────────────────
def _mix(a: str, b: str, t: float) -> str:
    """a→b 를 t(0~1)만큼 섞는다. 단계별 틴트 램프에 쓴다."""
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def _lum(h: str) -> float:
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _solid(fill: str, T: dict, need: float = 4.5) -> str:
    """글자를 얹을 채움 면 — 흰 글자가 읽힐 때까지 면을 어둡게 내린다.

    밝은 채도색(주황 등)을 그대로 면으로 쓰면 흰 글자는 3:1 도 안 나오고,
    잉크를 얹으면 색이 탁해진다. 그래서 **면을 내린다** — 색 정체성은 유지되고
    흰 글자가 또렷해진다 (color-system §2-2 의 ACC_SOLID 역할).
    중립·아주 밝은 면(카드 배경)은 손대지 않는다.
    """
    r, g, b = (int(fill[i:i + 2], 16) for i in (0, 2, 4))
    if max(r, g, b) - min(r, g, b) < 26:        # 무채색 면은 그대로
        return fill
    # 밝은 틴트(연노랑 FEF9C3 류)는 흰 글자용 채도면이 아니라 **잉크 글자용 카드 면**이다.
    # 여기서 어둡게 내리면 스펙에 없는 올리브 다크(77745B)가 만들어진다 —
    # 흰색과 거의 같은 명도면 그대로 둔다.
    if _contrast("FFFFFF", fill) < 1.25:
        return fill
    if _contrast("FFFFFF", fill) >= need:
        return fill
    for i in range(1, 15):
        c = _mix(fill, "000000", i / 15)
        if _contrast("FFFFFF", c) >= need:
            return c
    return "000000"


def _on(bg: str, T: dict) -> str:
    """그 면 위에서 더 잘 읽히는 글자색 — 밝은 면엔 잉크, 어두운 면엔 반전색.

    다크 테마는 ink·on_accent 가 둘 다 흰색이라 밝은 강조 면 위에서 글자가
    사라진다(fb021 실측: vidifx 그라데이션 카드 위 흰 제목). 그래서 테마 후보가
    전부 죽는 면에서는 중립 잉크를 폴백으로 둔다 — 대비가 가장 큰 쪽이 이긴다.
    """
    cands = (T["on_accent"], T["ink"], "16161C")
    return max(cands, key=lambda c: _contrast(c, bg))


def _fit(color: str, bg: str, T: dict, need: float = 3.2) -> str:
    """악센트를 그 면 위에서 읽히는 데까지 어둡게(또는 밝게) 민다.

    수치를 악센트색으로 쓰면 밝은 카드 위에서 대비가 3:1을 못 넘기는 일이 잦다
    (실측 E8622C on EFECE7 = 2.9:1 → 납품물 게이트 [L]). 색 정체성은 유지하고
    명도만 조정한다.
    """
    if _contrast(color, bg) >= need:
        return color
    # **양쪽으로 다 밀어 보고 더 잘 읽히는 쪽을 고른다.** 밝기만 보고 한쪽을 정하면
    # 중간 밝기 면(주황 등)에서 흰색으로 밀었다가 3:1 도 못 넘긴다(실측 2.8:1).
    best, best_c = None, -1.0
    for target in ("000000", "FFFFFF"):
        for i in range(1, 13):
            c = _mix(color, target, i / 12)
            k = _contrast(c, bg)
            if k >= need:
                return c
            if k > best_c:
                best, best_c = c, k
    return best or color


def _ramp(T: dict, n: int, base: str | None = None,
          focus: int | None = None) -> list[str]:
    """단계 면 n단 — **중립으로 흐르고 강조 한 단만 악센트.**

    예전엔 악센트→카드색 유채 램프였다. 그러면 지면의 유채 면적이 40%를 넘어
    악센트가 악센트가 아니게 된다 — `docs/design-refs-color-system.md` §2-1 이
    실측으로 반려 판정한 바로 그 패턴(mint 안 44.5% 반려 / carbon 안 15.9% 채택).
    """
    ramp = [_mix(T["card"], T["ink"], 0 if n == 1 else i / (n - 1) * 0.20)
            for i in range(n)]
    if focus is not None and 0 <= focus < n:
        ramp[focus] = base or T["accent"]
    return ramp


# ── 기하 ────────────────────────────────────────────────────────────────────
def _even(x: float, w: float, n: int, gap: float = GAP) -> list[tuple[float, float]]:
    """판면 폭을 n칸으로 균등분배 → [(x_i, w_i)]. 좌우 여백이 자동으로 같다."""
    cw = (w - gap * (n - 1)) / n
    return [(x + i * (cw + gap), cw) for i in range(n)]


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


# ── 노드 텍스트 ─────────────────────────────────────────────────────────────
def _runs(node: dict, T: dict, color: str, *, title: float = 13.5,
          value: float = 20, caption: float = 8.5,
          bg: str | None = None) -> list[list[dict]]:
    """제목·수치·캡션을 한 텍스트 아이템의 줄들로 쌓는다.

    노드당 텍스트를 하나로 유지하는 게 핵심 — 손좌표 행이 생기지 않고,
    세로중앙 정렬이 박스 하나로 끝난다.
    """
    # 글자색은 **얹히는 면(bg)에서 역산한다** — 하드코딩하면 밝은 면에서 대비가 깨진다.
    bg = bg or T["bg"]
    ink = _on(bg, T)
    gray = _mix(ink, bg, 0.34)
    if _contrast(gray, bg) < 4.5:
        gray = _mix(ink, bg, 0.12)
    lines: list[list[dict]] = []
    if node.get("title"):
        # 제목 글자만 악센트로 띄우는 문법 — **면을 칠하는 대신** 쓴다(R-ACCENT:
        # "색을 면에 칠하면 악센트가 죽는다"). 지정이 없으면 기존대로 잉크색이다.
        tc = node.get("title_color")
        lines.append([{"t": str(node["title"]), "size": title,
                       "font": "display_sb",
                       "color": _fit(tc, bg, T) if tc else ink}])
    if node.get("value") is not None:
        lines.append([{"t": str(node["value"]), "size": value, "font": "display_b",
                       "color": _fit(color, bg, T)}])
    if node.get("caption"):
        lines.append([{"t": str(node["caption"]), "size": caption, "color": gray}])
    return lines or [[{"t": "", "size": title, "color": ink}]]


def _text_in(rect, node, T, color, *, align="left", pad=PAD, bg=None,
             **kw) -> dict:
    """면 안에 앉는 글자 — 면과 같은 y·h + anchor:middle (절대규칙 6)."""
    x, y, w, h = rect
    return {"type": "text", "style": "body", "spacing": 1.3,
            "x": round(x + pad, 4), "y": round(y, 4),
            "w": round(max(w - pad * 2, 0.02), 4), "h": round(h, 4),
            "anchor": "middle", "align": align,
            "runs": _runs(node, T, color, bg=bg, **kw)}


# ── 아이템 스킨 (AntV 의 item) ───────────────────────────────────────────────
# skin(rect, node, ctx) -> [프리미티브...]   rect = (x, y, w, h)
def _skin_plain(rect, node, ctx):
    return [_text_in(rect, node, ctx["T"], ctx["color"], align=ctx["align"])]


def _est_lines(text: str, size: float, w: float) -> int:
    """그 폭에서 몇 줄이 되는지 — 세로 배치를 계산으로 잡기 위한 추정."""
    if not text:
        return 0
    est = sum(0.0010 if ord(ch) >= 0x1100 else 0.00055 for ch in str(text)) * size
    return max(1, -(-int(est * 100) // int(max(w, 0.01) * 100)))


def _icon_in(rect, node, ctx, bg):
    """카드 머리에 얹는 단독 아이콘.

    **아이콘+글자를 한 덩어리로 보고 그 덩어리를 카드 세로중앙에 앉힌다.**
    예전엔 아이콘을 위에 붙이고 글자를 top 정렬해서 카드 아래 3분의 2가 통째로
    비었다(2026-08-12 반려: "여백도 너무 많고"). 게이트 [P]는 글자 총량만 보느라
    이걸 못 잡았다 — 밀도는 양이 아니라 **분포**의 문제이기도 하다.
    """
    if not node.get("icon"):
        return [], rect
    x, y, w, h = rect
    size = ctx.get("icon_size", 0.030)
    pad = PAD
    tw = w - pad * 2
    lines = (_est_lines(node.get("title"), 13.5, tw)
             + _est_lines(node.get("caption"), 8.5, tw)
             + (1 if node.get("value") is not None else 0))
    text_h = lines * 0.0286                      # 13.5pt·spacing 1.3 기준 한 줄
    gap = 0.014
    block = size * (16 / 9) + gap + text_h
    if block > h - pad:
        return [], rect          # 칸이 좁으면 아이콘을 포기한다 — 넘치는 것보다 낫다
    top = y + (h - block) / 2    # 덩어리를 세로중앙에
    T = ctx["T"]
    on = _on(bg, T)
    # 면이 채도색이면 아이콘도 글자와 같은 색으로 간다 — 혼자 어두우면 안 보인다
    icol = on if _contrast(ctx["color"], bg) < 3.0 else _fit(ctx["color"], bg, T)
    ic = {"type": "icon", "name": node["icon"], "size": size,
          "x": round(x + pad, 4), "y": round(top, 4), "color": icol}
    ty = top + size * (16 / 9) + gap
    return [ic], (x, round(ty, 4), w, round(max(y + h - ty, 0.03), 4))


def _skin_compact_card(rect, node, ctx):
    x, y, w, h = rect
    T = ctx["T"]
    # 노드 fill 이 2색 리스트면 강조 그라데이션 면 — 잉크 대비 기준은 밝은 쪽(fb021)
    fill_spec = ctx.get("fill", T["card"])
    grad = list(fill_spec) if isinstance(fill_spec, (list, tuple)) else None
    bg = grad[1] if grad else _solid(fill_spec, T)
    icons, body = _icon_in(rect, node, ctx, bg)
    txt = _text_in(body, node, T, ctx["color"], align=ctx["align"], bg=bg)
    if icons:                      # 아이콘이 있으면 글자는 위에서 아래로 흐른다
        txt["anchor"] = "top"
    card = {"type": "card", "x": round(x, 4), "y": round(y, 4),
            "w": round(w, 4), "h": round(h, 4),
            "fill": bg, "radius": ctx.get("radius", 0.14)}
    if grad:
        card["grad"] = grad
        card["grad_ang"] = 135
    return [card, *icons, txt]


def _skin_outline_card(rect, node, ctx):
    x, y, w, h = rect
    T = ctx["T"]
    return [
        {"type": "card", "x": round(x, 4), "y": round(y, 4),
         "w": round(w, 4), "h": round(h, 4), "fill": "none",
         "outline": ctx["color"], "outline_w": 1.1, "radius": ctx.get("radius", 0.14)},
        _text_in(rect, node, T, ctx["color"], align=ctx["align"]),
    ]


def _skin_badge_card(rect, node, ctx):
    """카드 위 테두리에 넘버 배지가 걸친 형태 — 글자 자리를 먹지 않는다."""
    x, y, w, h = rect
    d = min(CHIP_D, h * 0.42)
    ar = ctx["ar"]
    body = (x, y + d * ar / 2, w, h - d * ar / 2)
    out = _skin_compact_card(body, node, ctx)
    out.append({"type": "chip", "d": round(d, 4),
                "x": round(x + w / 2 - d / 2, 4), "y": round(y, 4),
                "fill": ctx["color"], "text": ctx.get("badge", str(ctx["idx"] + 1))})
    return out


def _skin_indexed_card(rect, node, ctx):
    """좌측에 큰 인덱스, 우측에 내용 — 목록형 카드."""
    x, y, w, h = rect
    T = ctx["T"]
    iw = min(0.062, w * 0.22)
    fill_spec = ctx.get("fill", T["card"])
    grad = list(fill_spec) if isinstance(fill_spec, (list, tuple)) else None
    bg = grad[1] if grad else fill_spec
    card = {"type": "card", "x": round(x, 4), "y": round(y, 4),
            "w": round(w, 4), "h": round(h, 4),
            "fill": bg, "radius": ctx.get("radius", 0.14)}
    if grad:
        card["grad"] = grad
        card["grad_ang"] = 135
    out = [card,
           {"type": "text", "style": "body", "spacing": 1.0,
            "x": round(x + PAD, 4), "y": round(y, 4), "w": round(iw, 4),
            "h": round(h, 4), "anchor": "middle", "align": "left",
            "runs": [[{"t": ctx.get("badge", f"{ctx['idx'] + 1:02d}"), "size": 21,
                       "font": "display_b",
                       "color": _fit(ctx["color"], bg, T)}]]},
           _text_in((x + iw, y, w - iw, h), node, T, ctx["color"], align="left",
                    bg=bg)]
    return out


def _skin_underline_text(rect, node, ctx):
    """면 없이 얇은 룰 하나로 구획 — 여백형 톤에 맞는 스킨."""
    x, y, w, h = rect
    T = ctx["T"]
    return [
        {"type": "hairline", "x": round(x, 4), "y": round(y, 4),
         "w": round(w, 4), "color": ctx["color"], "thick": 1.4},
        _text_in((x, y + 0.012, w, h - 0.012), node, T, ctx["color"],
                 align=ctx["align"], pad=0.0),
    ]


def _skin_pill_badge(rect, node, ctx):
    """상단 필 + 아래 설명 — 라벨이 짧고 캡션이 긴 목록에."""
    x, y, w, h = rect
    T = ctx["T"]
    ph = 0.042
    body = dict(node)
    title = body.pop("title", "")
    out = [{"type": "pill", "x": round(x, 4), "y": round(y, 4),
            "h": ph, "w": round(w, 4), "text": str(title),
            "fill": ctx["color"]}]
    if body.get("value") is not None or body.get("caption"):
        out.append(_text_in((x, y + ph + 0.008, w, h - ph - 0.008), body, T,
                            ctx["color"], align=ctx["align"], pad=0.0))
    return out


def _skin_stat(rect, node, ctx):
    """라벨 → 빅넘버 → 얇은 룰 → 캡션. 수치가 주인공일 때.

    검증된 조판이다.
    면을 위아래로 꽉 채우므로 게이트 [P](면이 내용보다 크다)도 자연히 통과한다.
    """
    x, y, w, h = rect
    T = ctx["T"]
    fill_spec = ctx.get("fill", T["card"])
    grad = list(fill_spec) if isinstance(fill_spec, (list, tuple)) else None
    bg = grad[1] if grad else _solid(fill_spec, T)
    ink = _on(bg, T)
    num = str(node.get("value", ""))
    pad = min(PAD, h * 0.12)
    iw = w - pad * 2
    # 낮은 카드에서는 숫자를 줄인다 — 고정 크기로 쌓으면 캡션이 카드 밖으로 흘러넘친다
    # (2026-08-12 실측: h0.167 카드에서 캡션이 다음 카드를 덮었다).
    lab_h = 0.030 if node.get("title") else 0.0
    cap_lines = _est_lines(node.get("caption"), 10, iw) if node.get("caption") else 0
    cap_h = cap_lines * 0.0250
    avail = h - pad * 2 - lab_h - cap_h - 0.022
    # num_size 를 주면 행 전체가 같은 크기로 간다 — 자동 크기는 숫자 폭에 따라
    # 카드마다 달라져서 룰(밑줄) 높이가 들쭉날쭉해진다(수치 스트립에서 눈에 띈다).
    size = ctx.get("num_size") or (30 if len(num) <= 4 else 24)
    size = max(15, min(size, avail * 540 / 1.2))
    # **폭도 본다.** 높이만 맞추면 "200~300개"가 "200~30 / 0개"로 쪼개진다
    # (2026-08-12 포트폴리오 실측 — 카드 폭을 넘긴 숫자가 전부 깨졌다).
    # 디스플레이 볼드는 본문보다 넓다 — 계수를 크게 잡고 여유를 둔다
    est = sum(0.00122 if ord(ch) >= 0x1100 else 0.00064 for ch in num)
    if est > 0:
        size = max(11, min(size, iw / est * 0.94))
    out = [{"type": "card", "x": round(x, 4), "y": round(y, 4),
            "w": round(w, 4), "h": round(h, 4), "fill": bg,
            "radius": ctx.get("radius", 0.14)}]
    if grad:
        out[0]["grad"] = grad
        out[0]["grad_ang"] = 135
    ty = y + pad
    if node.get("title"):
        # 라벨 회색은 면 대비 4.5:1 을 지키는 선까지만 눕힌다 — 밝은 grad 면에서
        # 0.34 고정 혼합이 3.37:1 로 떨어졌다(fb021 r5 실측). _runs 의 가드와 동일.
        # grad 면은 무조건 진한 혼합 — 가드가 bg=밝은 스탑 기준으로 재서 어두운
        # 스탑 쪽에 앉은 라벨을 놓친다(fb024 r6 재발: 8675E8 구간 3.62:1).
        lab_color = _mix(ink, bg, 0.12 if grad else 0.34)
        if _contrast(lab_color, bg) < 4.5:
            lab_color = _mix(ink, bg, 0.12)
        out.append({"type": "text", "style": "body", "x": round(x + pad, 4),
                    "y": round(ty, 4), "w": round(iw, 4), "spacing": 1.0,
                    "runs": [[{"t": str(node["title"]), "size": 9.5,
                               "color": lab_color}]]})
        ty += lab_h
    # 면이 곧 악센트면 숫자를 악센트로 칠할 수 없다 — 그 면 위 잉크로 간다.
    # 그라데이션 강조 면도 같다 — _fit 으로 악센트를 중간까지만 밀면 탁한 회색이
    # 남는다(fb021 실측: A6A1D9 위 흰 악센트가 올리브빛 회색으로).
    num_color = (_on(bg, T) if (grad or _contrast(ctx["color"], bg) < 1.6)
                 else _fit(ctx["color"], bg, ctx["T"]))
    out.append({"type": "text", "style": "num_xl", "x": round(x + pad, 4),
                "y": round(ty, 4), "w": round(iw, 4), "spacing": 1.0,
                "runs": [[{"t": num, "size": round(size, 1), "font": "display_b",
                           "color": num_color}]]})
    ty += size * 1.2 / 540 + min(0.022, pad)
    out.append({"type": "hairline", "x": round(x + pad, 4), "y": round(ty, 4),
                "w": round(iw, 4), "color": _mix(ink, bg, 0.72)})
    if node.get("caption"):
        cy = ty + min(0.020, pad)
        bottom = y + h - pad * 0.4           # 카드 안쪽 바닥 — 절대 못 넘는다
        room = bottom - cy
        if room < 0.016:
            return out                       # 자리가 없으면 캡션을 아예 안 넣는다
        need = cap_lines * 0.0250
        cap_size = 10 if room >= need else max(7.5, 10 * room / max(need, 1e-6))
        out.append({"type": "text", "style": "body", "x": round(x + pad, 4),
                    "y": round(cy, 4), "w": round(iw, 4),
                    "h": round(room, 4), "spacing": 1.3,
                    "runs": [[{"t": str(node["caption"]), "size": round(cap_size, 1),
                               "color": _mix(ink, bg, 0.22)}]]})
    return out


def _node_dot(px, py, d, fill, ar):
    """축 위 노드 표식 — **글자 없는 작은 점.**

    예전엔 여기에 순번을 적은 넘버 서클이 앉았다. 반려됐다(2026-08-17). 순서는
    축과 배치가 이미 말하고 있어서 원 안 숫자는 아무 일도 안 한다 — 점만 남긴다.
    """
    dd = min(d, 0.016)
    return {"type": "shape", "shape": "oval", "fill": fill,
            "x": round(px - dd / 2, 4), "y": round(py - dd * ar / 2, 4),
            "w": round(dd, 4), "h": round(dd * ar, 4)}


SKINS = {
    "plain": _skin_plain,
    "stat": _skin_stat,
    "compact_card": _skin_compact_card,
    "outline_card": _skin_outline_card,
    "badge_card": _skin_badge_card,
    "indexed_card": _skin_indexed_card,
    "underline_text": _skin_underline_text,
    "pill_badge": _skin_pill_badge,
}

# 이름은 남겨둔다 — 쓰면 "unknown skin" 이 아니라 **금지 사유**를 보여주기 위해서다.
_BANNED_SKINS = {"badge_card"}


# ── 구조 (AntV 의 structure) ────────────────────────────────────────────────
# figure(spec, ctx) -> [프리미티브...]
def _f_grid(spec, ctx):
    """균등 그리드 — 가장 기본. n≤4는 한 행, 그 이상은 2행으로 접는다."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    cols = spec.get("cols") or (n if n <= 4 else math.ceil(n / 2))
    rows = math.ceil(n / cols)
    gap = spec.get("gap", GAP)
    rh = (h - gap * (rows - 1)) / rows
    out = []
    for i, node in enumerate(nodes):
        r, c = divmod(i, cols)
        cx, cw = _even(x, w, cols, gap)[c]
        out += ctx["skin"]((cx, y + r * (rh + gap), cw, rh), node, ctx["node"](i))
    return out


def _f_timeline(spec, ctx):
    """가로 흐름 — 카드 + 커넥터 화살표.

    **넘버 배지(동그라미 안 숫자)는 금지다.** 옵트인으로 남겨뒀다가 다시 반려됐다
    (2026-08-12 "동그라미 글자도 쓰레기" → 2026-08-17 "이런 동그라미로 순서도
    만드는 인포그래픽 만들지마"). `badge`/`badges` 키를 주면 렌더가 멈춘다 —
    순서는 카드 + 커넥터 화살표, 또는 단계명·날짜 자체로 읽히게 만든다.
    """
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    d = spec.get("d", CHIP_D)
    gap = spec.get("gap", GAP)
    cols = _even(x, w, n, gap)
    top = y
    card_h = h
    axis = top + card_h / 2                        # 커넥터가 앉는 축
    out = []
    for i, (cx, cw) in enumerate(cols):
        out += ctx["skin"]((cx, top, cw, card_h), nodes[i], ctx["node"](i))
    for i, (cx, cw) in enumerate(cols):
        if i < n - 1:
            nx = cols[i + 1][0]
            g0, g1 = cx + cw, nx
            aw = min(0.022, max(0.008, (g1 - g0) * 0.42))
            out.append({"type": "arrow", "x": round((g0 + g1) / 2 - aw / 2, 4),
                        "y": round(axis - ARROW_H / 2, 4),
                        "w": round(aw, 4), "h": ARROW_H})
    return out


def _f_steps_up(spec, ctx):
    """상승 계단 — 왼쪽에서 오른쪽으로 높아진다. 바닥선은 하나로 맞는다."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    floor = spec.get("floor", 0.42)          # 첫 칸이 앉는 높이 비율
    cols = _even(x, w, n, spec.get("gap", GAP))
    # 디딤판(회색 블록)은 폐기했다 — 카드와 분리돼 빈 상자로 읽혔다
    # (2026-08-12 반려). 계단은 **카드 자리만** 층층이 올려서 만든다.
    card_h = min(spec.get("card_h", 0.145), h * 0.62)
    span = h - card_h
    out = []
    for i, (cx, cw) in enumerate(cols):
        out += ctx["skin"]((cx, y + span - span * (i / max(n - 1, 1)), cw, card_h),
                           nodes[i], ctx["node"](i))
    return out


def _f_waterfall(spec, ctx):
    """계단식 하강 — 단계가 내려가며 겹치듯 흐른다."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    drop = h * spec.get("drop", 0.45) / max(n - 1, 1)
    ch = h - drop * (n - 1)
    out = []
    for i, (cx, cw) in enumerate(_even(x, w, n, spec.get("gap", GAP))):
        out += ctx["skin"]((cx, y + i * drop, cw, ch), nodes[i], ctx["node"](i))
    return out


def _f_funnel(spec, ctx):
    """퍼널 — 폭이 줄어드는 사다리꼴 층. 수치가 있는 전환 흐름에.

    기본 강조는 **마지막 단**(=전환된 것). `focus`로 바꾼다.
    """
    return _taper(spec, ctx, narrow_at_bottom=True, default_focus=-1)


def _f_pyramid(spec, ctx):
    """피라미드 — 위가 좁은 층. 위계·구성비에. 기본 강조는 꼭대기."""
    return _taper(spec, ctx, narrow_at_bottom=False, default_focus=0)


def _taper(spec, ctx, *, narrow_at_bottom: bool, default_focus: int = 0):
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    T = ctx["T"]
    end = spec.get("shrink", 0.42)           # 좁은 쪽 폭 비율
    gap = spec.get("gap", 0.006)
    bh = (h - gap * (n - 1)) / n
    cx = x + w / 2
    colors = ctx["ramp"]
    if not spec.get("colors") and spec.get("focus") is None:
        colors = _ramp(T, n, ctx["color"](0), default_focus % n)

    def half(k):                              # k=0 위, k=n 아래
        t = k / n
        frac = 1 - (1 - end) * t if narrow_at_bottom else end + (1 - end) * t
        return w * frac / 2

    out = []
    for i, node in enumerate(nodes):
        top = y + i * (bh + gap)
        bot = top + bh
        ht, hb = half(i), half(i + 1)
        out.append({"type": "poly", "fill": colors[i],
                    "points": [(round(cx - ht, 4), round(top, 4)),
                               (round(cx + ht, 4), round(top, 4)),
                               (round(cx + hb, 4), round(bot, 4)),
                               (round(cx - hb, 4), round(bot, 4))]})
        tw = min(ht, hb) * 2 * 0.86
        out.append(_text_in((cx - tw / 2, top, tw, bh), node, T,
                            _on(colors[i], T), align="center", pad=0.0,
                            bg=colors[i], title=12.5, value=17, caption=8.5))
    return out


def _f_zigzag(spec, ctx):
    """가로 축 + 위아래 교차 카드 — 시간 흐름 위의 사건."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    ar = ctx["ar"]
    d = spec.get("d", 0.030)
    lane = (h - d * ar) / 2 - 0.014
    axis_y = y + lane + 0.014
    out = [{"type": "hairline", "x": round(x, 4),
            "y": round(axis_y + d * ar / 2, 4), "w": round(w, 4)}]
    for i, (cx, cw) in enumerate(_even(x, w, n, spec.get("gap", GAP))):
        up = i % 2 == 0
        out.append(_node_dot(cx + cw / 2, axis_y + d * ar / 2, d,
                             ctx["color"](i), ar))
        top = y if up else axis_y + d * ar + 0.014
        out += ctx["skin"]((cx, top, cw, lane), nodes[i], ctx["node"](i))
    return out


def _f_roadmap(spec, ctx):
    """세로 축 + 좌우 교차 카드 — 단계가 많고 설명이 긴 로드맵."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    ar = ctx["ar"]
    d = spec.get("d", 0.030)
    gap = spec.get("gap", GAP)
    rh = (h - gap * (n - 1)) / n
    cx = x + w / 2
    lane = w / 2 - d / 2 - 0.020
    out = [{"type": "card", "x": round(cx - 0.0015, 4), "y": round(y, 4),
            "w": 0.003, "h": round(h, 4), "fill": ctx["T"]["hair"], "radius": 0.5}]
    for i, node in enumerate(nodes):
        top = y + i * (rh + gap)
        out.append(_node_dot(cx, top + rh / 2, d, ctx["color"](i), ar))
        left = i % 2 == 0
        rx = x if left else cx + d / 2 + 0.020
        c = ctx["node"](i)
        c["align"] = "right" if left else "left"
        out += ctx["skin"]((rx, top, lane, rh), node, c)
    return out


def _f_snake(spec, ctx):
    """뱀형 스텝 — 한 행이 끝나면 아래로 꺾어 되돌아온다. 단계가 5개 이상일 때."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    cols = spec.get("cols") or math.ceil(n / 2)
    rows = math.ceil(n / cols)
    gap = spec.get("gap", GAP)
    rh = (h - gap * (rows - 1)) / rows
    grid = _even(x, w, cols, gap)
    out = []
    for i, node in enumerate(nodes):
        r = i // cols
        c = i % cols
        if r % 2:                              # 홀수 행은 오른쪽에서 왼쪽으로
            c = cols - 1 - c
        cx, cw = grid[c]
        top = y + r * (rh + gap)
        out += ctx["skin"]((cx, top, cw, rh), node, ctx["node"](i))
        if i == n - 1:
            continue
        nr = (i + 1) // cols
        if nr == r:                            # 같은 행 — 가로 커넥터
            nc = (i + 1) % cols
            if r % 2:
                nc = cols - 1 - nc
            nx = grid[nc][0]
            g0 = min(cx + cw, nx + grid[nc][1])
            g1 = max(cx, nx)
            aw = min(0.020, max(0.008, (g1 - g0) * 0.5))
            arrow = {"type": "arrow", "x": round((g0 + g1) / 2 - aw / 2, 4),
                     "y": round(top + rh / 2 - ARROW_H / 2, 4),
                     "w": round(aw, 4), "h": ARROW_H}
            if r % 2:                      # 되돌이 행은 흐름이 왼쪽이다
                arrow["rot"] = 180
            out.append(arrow)
        else:
            # 행 바꿈 — **가로 커넥터와 같은 언어로 그린다.**
            # 예전엔 hair 색 얇은 막대였는데, 방향이 꺾이는 유일한 지점이
            # 카드면과 명도차가 거의 없어 제일 흐렸다(2026-08-25 검수 major).
            # 아래를 향하는 화살표로 통일한다 — 색·굵기도 가로와 같다.
            #
            # **`rot:90` 이면 `w` 가 화면상 세로 길이가 된다.** `gap` 은 세로 비율인데
            # `w` 는 가로 비율이라, 환산 없이 쓰면 판형이 가로로 길수록 화살표가
            # 길어져 위아래 카드를 침범한다(16:9 덱에서 실측 0.16mm 침범).
            # `ar` 로 세로→가로 비율을 맞춘다.
            ar = ctx["ar"]
            # 세로 비율 → 가로 비율은 **나눈다**(`ar` = SLIDE_W/SLIDE_H).
            # 같은 파일 `max_pitch` 계산이 이미 그 관용구다. 곱하면 16:9 에서
            # 화살표가 1.78배 길어져 위아래 카드를 2mm 씩 침범한다.
            aw = min(0.016, max(0.008, gap * 0.6)) / ar
            out.append({"type": "arrow", "rot": 90,
                        "x": round(cx + cw / 2 - aw / 2, 4),
                        "y": round(top + rh + (gap - ARROW_H) / 2, 4),
                        "w": round(aw, 4), "h": ARROW_H})
    return out


def _f_quadrant(spec, ctx):
    """사분면 — SWOT·2×2 매트릭스. 노드 4개 기준.

    `axes` 를 주면 축 라벨이 붙는다. **축이 없으면 그냥 2×2 그리드로 읽힌다** —
    사분면인 이유(무엇과 무엇을 가르는가)는 축이 말한다.
        "axes": {"cols": ["내부 요인", "외부 요인"], "rows": ["도움", "방해"]}
    """
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    T = ctx["T"]
    gap = spec.get("gap", GAP)
    axes = spec.get("axes") or {}
    col_lab, row_lab = axes.get("cols") or [], axes.get("rows") or []
    top = 0.036 if col_lab else 0.0        # 열 라벨 띠
    left = 0.062 if row_lab else 0.0       # 행 라벨 열
    gx, gy = x + left, y + top
    gw, gh = w - left, h - top
    cw = (gw - gap) / 2
    rh = (gh - gap) / 2
    out = []
    for c, lab in enumerate(col_lab[:2]):
        out.append({"type": "text", "style": "caption", "x": round(gx + c * (cw + gap), 4),
                    "y": round(y, 4), "w": round(cw, 4), "h": round(top, 4),
                    "anchor": "middle", "align": "center", "text": str(lab),
                    "color": T["gray"], "size": 9.5})
    for r, lab in enumerate(row_lab[:2]):
        out.append({"type": "text", "style": "caption", "x": round(x, 4),
                    "y": round(gy + r * (rh + gap), 4), "w": round(left - 0.012, 4),
                    "h": round(rh, 4), "anchor": "middle", "align": "right",
                    "text": str(lab), "color": T["gray"], "size": 9.5})
    for i, node in enumerate(nodes[:4]):
        r, c = divmod(i, 2)
        out += ctx["skin"]((gx + c * (cw + gap), gy + r * (rh + gap), cw, rh),
                           node, ctx["node"](i))
    return out


def _f_circular(spec, ctx):
    """원형 시퀀스 — 링 위의 노드 점 + 방사 라벨. 순환하는 단계에.

    라벨을 늘 아래에 두면 좌우 노드가 허브 글자와 엉킨다. 각도를 보고
    좌·우 노드는 옆으로, 위·아래 노드는 위아래로 내보낸다 — 판면 폭도 그때 쓰인다.
    """
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    ar = ctx["ar"]
    T = ctx["T"]
    d = spec.get("d", 0.038)
    # 라벨 높이는 **내용에서 잡는다.** 고정값이면 긴 캡션이 좁은 상자를 넘쳐 흐른다.
    lw0 = min(0.24, w * 0.26)
    need = 1
    for nd in nodes:
        for txt, sz in ((nd.get("title"), 12), (nd.get("caption"), 8.5)):
            if not txt:
                continue
            est = sum(0.0010 if ord(ch) >= 0x1100 else 0.00055 for ch in str(txt)) * sz
            need = max(need, -(-int(est / lw0 * 100) // 100) if est > lw0 else 1)
    rows = sum(1 for k in ("title", "value", "caption")
               if any(nd.get(k) is not None for nd in nodes))
    lab_h = spec.get("label_h", max(0.058, (rows + need - 1) * 0.030))
    ry = h / 2 - lab_h - d * ar / 2
    rx = ry / ar
    cx, cy = x + w / 2, y + h / 2
    side_w = min(lw0, cx - rx - d / 2 - 0.016 - x)
    out = [{"type": "shape", "shape": "oval", "fill": "none",
            "outline": T["hair"], "outline_w": 1.2,
            "x": round(cx - rx, 4), "y": round(cy - ry, 4),
            "w": round(rx * 2, 4), "h": round(ry * 2, 4)}]
    if spec.get("hub"):
        out.append(_text_in((cx - rx * 0.72, cy - lab_h, rx * 1.44, lab_h * 2),
                            {"title": spec["hub"]}, T, T["accent"],
                            align="center", pad=0.0, title=14))
    for i, node in enumerate(nodes):
        th = -math.pi / 2 + 2 * math.pi * i / n
        px, py = cx + rx * math.cos(th), cy + ry * math.sin(th)
        out.append(_node_dot(px, py, d, ctx["color"](i), ar))
        if abs(math.cos(th)) > 0.5 and side_w > 0.08:      # 좌·우 노드는 옆으로
            right = math.cos(th) > 0
            lx = px + d / 2 + 0.014 if right else px - d / 2 - 0.014 - side_w
            box = (lx, py - lab_h / 2, side_w, lab_h)
            align = "left" if right else "right"
        else:                                              # 위·아래 노드는 위아래로
            lw = min(lw0, max(rx * 2, 0.16))
            up = math.sin(th) < 0
            ly = py - d * ar / 2 - lab_h if up else py + d * ar / 2
            box = (px - lw / 2, ly, lw, lab_h)
            align = "center"
        bx, by, bw, bh = box
        out.append({"type": "text", "style": "body", "spacing": 1.25,
                    "x": round(_clamp(bx, x, x + w - bw), 4),
                    "y": round(_clamp(by, y, y + h - bh), 4),
                    "w": round(bw, 4), "h": round(bh, 4),
                    "anchor": "middle", "align": align,
                    "runs": _runs(node, T, ctx["color"](i), title=12,
                                  value=15, caption=8.5)})
    return out


def _f_hub(spec, ctx):
    """허브 & 스포크 — 중심 하나에 위성이 붙는 관계도."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    ar = ctx["ar"]
    T = ctx["T"]
    gap = spec.get("gap", GAP)
    cols = 2
    rows = math.ceil(n / cols)
    rh = (h - gap * (rows - 1)) / rows
    side = (w - spec.get("hub_w", 0.24)) / 2 - gap
    hub_w = w - side * 2 - gap * 2
    out = []
    # 허브는 글자 한 줄짜리다 — 행 높이를 그대로 쓰면 텅 빈 큰 알약이 된다(게이트 [P]).
    hub_h = min(rh, spec.get("hub_h", 0.098))
    hub_box = (x + side + gap, y + h / 2 - hub_h / 2, hub_w, hub_h)
    hub_cy = y + h / 2
    hair = T["hair"]
    t = 0.003                                  # 스포크 두께
    bus = {0: x + side + gap / 2, 1: x + side + hub_w + gap * 1.5}
    lanes: dict[int, list[float]] = {0: [], 1: []}
    for i, node in enumerate(nodes):
        r, c = divmod(i, cols)
        rx = x if c == 0 else x + side + hub_w + gap * 2
        cy = y + r * (rh + gap) + rh / 2
        lanes[c].append(cy)
        stub0 = min(bus[c], rx + side) if c == 0 else max(bus[c], rx)
        stub1 = max(bus[c], rx + side) if c == 0 else min(bus[c], rx)
        out.append({"type": "card", "x": round(min(stub0, stub1), 4),
                    "y": round(cy - t / 2 * ar, 4),
                    "w": round(abs(stub1 - stub0), 4), "h": round(t * ar, 4),
                    "fill": hair, "radius": 0.5})
        out += ctx["skin"]((rx, y + r * (rh + gap), side, rh), node, ctx["node"](i))
    for c, ys in lanes.items():                # 세로 버스 + 허브까지 잇는 가로 줄기
        if not ys:
            continue
        top, bot = min(min(ys), hub_cy), max(max(ys), hub_cy)
        out.append({"type": "card", "x": round(bus[c] - t / 2, 4), "y": round(top, 4),
                    "w": t, "h": round(bot - top, 4), "fill": hair, "radius": 0.5})
        hx = x + side + gap if c == 0 else x + side + hub_w + gap
        out.append({"type": "card", "x": round(min(bus[c], hx), 4),
                    "y": round(hub_cy - t / 2 * ar, 4),
                    "w": round(abs(hx - bus[c]), 4), "h": round(t * ar, 4),
                    "fill": hair, "radius": 0.5})
    hx0, hy0, hw, hh = hub_box                 # 허브는 스포크 위에 얹힌다
    out.append({"type": "card", "x": round(hx0, 4), "y": round(hy0, 4),
                "w": round(hw, 4), "h": round(hh, 4),
                "fill": ctx["color"](0), "radius": 0.5})
    out.append(_text_in(hub_box, {"title": spec.get("hub", "")}, T,
                        _on(ctx["color"](0), T), align="center",
                        bg=ctx["color"](0), title=13.5))
    return out


def _f_tree(spec, ctx):
    """계층 트리 — 루트 하나 + 하위 행. 조직도·구성도."""
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    T = ctx["T"]
    root, children = nodes[0], nodes[1:] or nodes
    gap = spec.get("gap", GAP)
    root_h = h * spec.get("root_h", 0.30)
    bus_y = y + root_h + gap / 2
    child_y = y + root_h + gap
    ch = h - root_h - gap
    rw = min(w * 0.44, 0.34)
    rx = x + w / 2 - rw / 2
    c0 = ctx["node"](0)
    c0["align"] = "center"
    out = ctx["skin"]((rx, y, rw, root_h), root, c0)
    cols = _even(x, w, len(children), gap)
    if len(children) > 1:
        first = cols[0][0] + cols[0][1] / 2
        last = cols[-1][0] + cols[-1][1] / 2
        out.append({"type": "hairline", "x": round(first, 4), "y": round(bus_y, 4),
                    "w": round(last - first, 4), "color": T["hair"]})
    out.append({"type": "card", "x": round(x + w / 2 - 0.0015, 4),
                "y": round(y + root_h, 4), "w": 0.003,
                "h": round(gap / 2, 4), "fill": T["hair"], "radius": 0.5})
    for i, (cx, cw) in enumerate(cols):
        out.append({"type": "card", "x": round(cx + cw / 2 - 0.0015, 4),
                    "y": round(bus_y, 4), "w": 0.003,
                    "h": round(gap / 2, 4), "fill": T["hair"], "radius": 0.5})
        c = ctx["node"](i + 1)
        c["align"] = "center"
        out += ctx["skin"]((cx, child_y, cw, ch), children[i], c)
    return out


def _f_gallery(spec, ctx):
    """작업물 썸네일 격자 — 포트폴리오의 기본. 노드에 `image` 를 준다.

    포트폴리오 공통 문법: 대표 1장을 크게, 나머지를 격자로.
    `feature: true` 면 첫 노드를 좌측에 크게 두고 나머지를 우측 격자로 흘린다.
    """
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    T = ctx["T"]
    gap = spec.get("gap", 0.012)
    cap_h = spec.get("cap_h", 0.030) if any(n.get("title") for n in nodes) else 0.0
    out = []

    def cell(rect, node):
        cx, cy, cw, ch = rect
        ih = ch - cap_h
        if node.get("image"):
            out.append({"type": "image", "path": node["image"],
                        "x": round(cx, 4), "y": round(cy, 4),
                        "w": round(cw, 4), "h": round(ih, 4)})
        elif node.get("fill"):
            # 색 타일은 shape 로 낸다 — card 로 내면 글자 없는 큰 면이라 [P]에 걸린다
            out.append({"type": "shape", "shape": "rounded", "adj": [0.05],
                        "x": round(cx, 4), "y": round(cy, 4),
                        "w": round(cw, 4), "h": round(ih, 4), "fill": node["fill"]})
        else:
            raise ValueError(
                "gallery 노드에는 image(또는 fill)가 있어야 한다 — 빈 자리표시 금지"
                " (절대규칙 9). 아직 그림이 없으면 fill 로 색 타일을 두고 나중에 갈아낀다")
        if node.get("title"):
            out.append({"type": "text", "style": "caption",
                        "x": round(cx, 4), "y": round(cy + ih + 0.004, 4),
                        "w": round(cw, 4), "h": round(cap_h, 4), "anchor": "middle",
                        "runs": [[{"t": str(node["title"]), "size": 9,
                                   "color": _mix(_on(T["bg"], T), T["bg"], 0.30)}]]})

    if spec.get("feature") and len(nodes) >= 3:
        fw = spec.get("feature_w", 0.42) * w
        cell((x, y, fw, h), nodes[0])
        rest = nodes[1:]
        cols = spec.get("cols") or 2
        rows = math.ceil(len(rest) / cols)
        rw = w - fw - gap
        cwid = (rw - gap * (cols - 1)) / cols
        rh = (h - gap * (rows - 1)) / rows
        for i, node in enumerate(rest):
            r, c = divmod(i, cols)
            cell((x + fw + gap + c * (cwid + gap), y + r * (rh + gap), cwid, rh), node)
        return out

    n = len(nodes)
    cols = spec.get("cols") or (n if n <= 4 else math.ceil(n / 2))
    rows = math.ceil(n / cols)
    cwid = (w - gap * (cols - 1)) / cols
    rh = (h - gap * (rows - 1)) / rows
    for i, node in enumerate(nodes):
        r, c = divmod(i, cols)
        cell((x + c * (cwid + gap), y + r * (rh + gap), cwid, rh), node)
    return out


def _mockup_meta(device: str):
    """목업 프레임 크기 — 자산이 없으면 None (그때는 폭 기준으로만 배치)."""
    try:
        from pptm.design.gyoan import mockup_devices
        return mockup_devices().get(device)
    except Exception:
        return None


def _f_showcase(spec, ctx):
    """기기 목업을 계단식으로 겹쳐 늘어놓는다 — 작업물 나열의 정석.

    노드마다 `image`(화면) · `device`(기본 phone). 가운데가 가장 크고 바깥으로
    갈수록 작아지며 살짝 내려간다(부채꼴). 포트폴리오의 릴스 나열 문법.
    """
    x, y, w, h = ctx["box"]
    nodes = ctx["nodes"]
    n = len(nodes)
    dev = spec.get("device", "phone")
    overlap = spec.get("overlap", 0.22)          # 이웃과 겹치는 비율
    taper = spec.get("taper", 0.14)              # 바깥으로 갈수록 작아지는 정도
    pitch = w / (n - (n - 1) * overlap) if n > 1 else w
    # **주어진 높이를 넘지 않게 폭을 역산한다.** 예전엔 w 만 보고 배치해서 폰이
    # bbox 아래로 흘러 하단 카드와 겹쳤다(2026-08-12 IR 시연 실측).
    man = _mockup_meta(dev)
    if man:
        fw, fh = man["size"]
        drop_max = h * 0.10
        max_pitch = (h - drop_max) * (fw / fh) / ctx["ar"]
        if pitch > max_pitch:            # 높이를 넘을 때만 줄인다
            pitch = max_pitch
    unit = pitch * (1 - overlap) if n > 1 else pitch
    mid = (n - 1) / 2
    span = pitch + unit * (n - 1)          # 실제 차지하는 폭
    x0 = x + max(0.0, (w - span) / 2)      # 남으면 가운데로 모은다
    out = []
    order = sorted(range(n), key=lambda i: abs(i - mid), reverse=True)
    for i in order:                              # 바깥부터 그려야 가운데가 위로 온다
        k = 1.0 - taper * (abs(i - mid) / max(mid, 1))
        cw = pitch * k
        cx = x0 + i * unit + (pitch - cw) / 2
        drop = (h * 0.10) * (abs(i - mid) / max(mid, 1))
        node = nodes[i]
        # 화면은 image → grad → fill 순. grad 는 실물 캡처를 아직 못 받은 자리를
        # **의도된 면**으로 보이게 한다 — 단색 면은 빈 프레임처럼 읽힌다.
        if node.get("image"):
            screen = {"image": node["image"]}
        elif node.get("grad"):
            screen = {"grad": node["grad"]}
        else:
            screen = {"fill": node.get("fill", ctx["color"](i))}
        out.append({"type": "mockup", "device": node.get("device", dev),
                    "x": round(cx, 4), "y": round(y + drop, 4),
                    "w": round(cw, 4), "screen": screen,
                    "px": spec.get("px", 900)})
    return out


FIGURES = {
    "gallery": _f_gallery,
    "showcase": _f_showcase,
    "grid": _f_grid,
    "timeline": _f_timeline,
    "steps_up": _f_steps_up,
    "waterfall": _f_waterfall,
    "funnel": _f_funnel,
    "pyramid": _f_pyramid,
    "zigzag": _f_zigzag,
    "roadmap": _f_roadmap,
    "snake": _f_snake,
    "quadrant": _f_quadrant,
    "circular": _f_circular,
    "hub": _f_hub,
    "tree": _f_tree,
}

# 스킨을 바꿔도 골격이 무너지지 않는 조합만 기본값으로 둔다
_DEFAULT_SKIN = {
    "gallery": "plain", "showcase": "plain",
    "funnel": "plain", "pyramid": "plain", "circular": "plain",
    "timeline": "compact_card", "zigzag": "compact_card",
    "roadmap": "compact_card", "tree": "compact_card",
}


# ── 전개 ────────────────────────────────────────────────────────────────────
def dark_view(T: dict) -> dict:
    """다크 배리언트용 토큰 치환.

    figure 는 색을 **면에서 역산**하므로 지면이 다크면 bg/ink 도 다크 값이어야 한다.
    안 넘기면 어두운 지면에 어두운 글자를 앉힌다 — 실측 1.0:1 (게이트 [L]).

    **면 색은 지어내지 않는다.** 예전엔 dark_bg 를 흰색 쪽으로 섞어 카드색을 자동
    생성했는데, 근black 위 평면 중간회색 + 채도 높은 악센트가 나와 통째로 반려됐다
    (2026-08-12 "랩3의 색 조합은 쓰레기야, 앞으로 절대 쓰지마"). 테마가 `dark_card`
    를 정의했으면 그걸 쓰고, 없으면 스펙이 `fill` 을 **명시**해야 한다.
    """
    out = {**T, "bg": T["dark_bg"], "ink": T["dark_fg"], "gray": T["dark_sub"]}
    for key, tok in (("card", "dark_card"), ("card2", "dark_card2"),
                     ("hair", "dark_hair")):
        if T.get(tok):
            out[key] = T[tok]
    return out


def expand_figures(items: list[dict], theme: dict, ar: float,
                   variant: str = "paper") -> list[dict]:
    """items 안의 figure 를 프리미티브로 전개한 새 리스트를 돌려준다.

    theme: gyoan.THEMES 항목(색 토큰). ar: 슬라이드 가로/세로 비 — 원은 화면상
    정사각이어야 하므로 chip 높이 계산에 필요하다. variant: paper|dark.
    """
    if not any(it.get("type") == "figure" for it in items):
        return items
    if variant == "dark":
        theme = dark_view(theme)
        if not theme.get("dark_card"):
            bad = [it for it in items if it.get("type") == "figure"
                   and not it.get("fill") and not it.get("colors")]
            if bad:
                raise ValueError(
                    "다크 지면의 figure 는 면 색을 명시해야 한다 — 스펙에 fill(또는 "
                    "colors)을 주거나 테마에 dark_card 를 정의할 것. 자동 생성한 "
                    "회색 면은 2026-08-12 반려됐다(docs/design-refs-antv-infographic.md §5)")
    out: list[dict] = []
    for it in items:
        if it.get("type") != "figure":
            out.append(it)
            continue
        out += _expand_one(it, theme, ar)
    return out


def _expand_one(spec: dict, T: dict, ar: float) -> list[dict]:
    name = spec.get("figure")
    if name not in FIGURES:
        raise ValueError(f"unknown figure: {name} (지원: {sorted(FIGURES)})")
    skin_name = spec.get("item") or _DEFAULT_SKIN.get(name, "compact_card")
    if skin_name not in SKINS:
        raise ValueError(f"unknown item skin: {skin_name} (지원: {sorted(SKINS)})")
    # ── 넘버 서클 영구 금지 (2026-08-17 반려: "이런 동그라미로 순서도 만드는
    #    인포그래픽 만들지마"). 2026-08-12에 옵트인으로 낮췄다가 다시 나왔다.
    if spec.get("badge") or spec.get("badges") or skin_name in _BANNED_SKINS:
        raise ValueError(
            "넘버 서클(동그라미 안 숫자)은 금지다 — badge/badges/badge_card 를 쓰지 말 것. "
            "순서는 카드 + 커넥터 화살표로, 번호가 꼭 필요하면 카드 안 좌측 인덱스"
            "(item: 'indexed_card')로 간다. CLAUDE.md 절대규칙 8-2")

    nodes = spec.get("nodes") or []
    if not nodes:
        raise ValueError(f"figure {name}: nodes 가 비었다")
    n = len(nodes)
    palette = spec.get("colors")
    accent = spec.get("accent", T["accent"])

    def color(i):
        return palette[i % len(palette)] if palette else accent

    ctx = {
        "T": T, "ar": ar, "nodes": nodes,
        "box": (spec["x"], spec["y"], spec["w"], spec["h"]),
        "skin": SKINS[skin_name],
        "color": color,
        "ramp": spec.get("colors") or _ramp(T, n, accent, spec.get("focus")),
        "node": lambda i: {"T": T, "ar": ar, "idx": i, "n": n,
                           "color": color(i), "align": spec.get("align", "left"),
                           "radius": spec.get("radius", 0.14),
                           # fills 로 **노드별 면 색**을 준다 — 활성 탭 하나만 악센트로
                           # 띄우는 문법에 필요하다.
                           "fill": (spec.get("fills")
                                    or [spec.get("fill", T["card"])] * n)[i],
                           "badge": (spec.get("badges") or [None] * n)[i]
                                    or str(i + 1),
                           "num_size": spec.get("num_size")},
    }
    items = FIGURES[name](spec, ctx)
    rpx = spec.get("radius_px")         # 절대 px 라운드 — 노드 카드 전체에 전파 (fb021)
    for it in items:                    # 게이트가 "그린 것"으로 알아보게 표식을 남긴다
        it["_fig"] = name
        if rpx is not None and it.get("type") == "card":
            it["radius_px"] = rpx
    return items
