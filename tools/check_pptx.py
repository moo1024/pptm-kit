#!/usr/bin/env python3
"""납품 pptx 게이트 — **파일 자체를 열어** 반복 반려 사유를 기계로 잡는다.

`check_align.py`는 `layout_spec.json`만 읽는다. 그래서 `tools/gen_*.py`로 pptx를
직접 찍는 경로(포스터·명찰·OT 자료…)는 게이트를 통째로 우회했다.
이 검사기는 **어떻게 만들었든 상관없이 최종 산출물**을 검사한다.

차단(경고):
  F. 폰트 폴백      — 미설치 글꼴명 / 한글 run 에 ea 미지정 (PITFALLS §2)
  V. 세로중앙 누락  — 카드·밴드 안 글자에 anchor=middle 이 없음 (규칙 6)
  O. 넘침          — 줄바꿈 끈 채 상자 초과 / 채움 면 밖으로 글자가 삐져나감
  M. 제작 메모      — "네이티브 렌더" 같은 만든 방법이 덱 문구에 (규칙 9)
  L. 대비 3.0 미만  — 글자색 vs 밑면(도형 fill·배경 사진 픽셀 실측)
  C. 판면 이탈      — **글자가** 지면 밖으로 흘러나감 (상자만 큰 건 봐준다)
  A. 자간+알파 동시 — LibreOffice가 마지막 글자를 먹는다 (2026-08-10 실측)

참고(차단 안 함, 눈으로 확인):
  L 3.0~4.5 / E 빈 텍스트박스 / M 애매한 낱말(목업 등) / T 타입스케일 이탈

임계값은 **사용자가 승인한 덱 24종에서 차단 0건**이 되도록 역산했다.
소음 나는 게이트는 아무도 안 본다 — 오탐이 보이면 임계값을 고쳐라.

사용:
  python3 tools/check_pptx.py <deck.pptx> [deck2.pptx ...]
반환: 차단 경고 있으면 1 (완성 선언 금지), 0이면 통과.
"""
from __future__ import annotations

import json
import pathlib
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn

NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
FONTS_JSON = Path(__file__).resolve().parent / "data" / "win_fonts.json"

# 만든 방법을 덱 문구에 쓴 것 (규칙 9) — 문맥상 오해 없는 것만 차단한다
MEMO_HARD = ("네이티브 렌더", "네이티브렌더", "lorem ipsum", "제목을 입력",
             "여기에 내용", "플레이스홀더", "placeholder", "샘플 텍스트")
MEMO_SOFT = ("목업", "mockup", "TBD", "TODO")      # 본문 낱말일 수 있다 → 참고
# 덱 안에 쓰면 안 되는 **제작 메타** — 장수·발표시간·파일명은 덱의 내용이 아니다.
# (2026-08-24 반려: 표지에 "10분 발표 · 14장" 을 박았다. "발표하고 사람들 다 보는 건데")
MEMO_META_RE = (
    r"\d+\s*분\s*발표", r"발표\s*\d+\s*분", r"총\s*\d+\s*장",
    r"\d+\s*장\s*구성", r"슬라이드\s*\d+\s*장", r"\d+\s*페이지\s*구성",
)
MEMO_EXACT = ("IMAGE", "PHOTO", "이미지", "사진")   # 통짜로 이것만 = 자리표시

# 글자 폭 추정용 em 계수 (한글 전각 1.0 기준)
_W_HANGUL, _W_LATIN, _W_DIGIT, _W_SPACE, _W_NARROW = 1.0, 0.52, 0.55, 0.28, 0.30
_NARROW = set(".,:;!|'\"()[]·-–—")

DEFAULT_SIZE = 18.0      # pptx 기본값 (run에 sz가 없을 때)
LINE_FACTOR = 1.35       # 줄높이 = size × 이것
OVERFLOW_TOL = 1.10      # 추정치라 빠듯하게 잡으면 오탐
OVERFLOW_TOL_H = 1.25    # 높이는 줄 수 추정 오차가 곱해진다 → 더 느슨하게
EMU_PER_PT = 12700

BG_AREA = 0.55           # 지면 이 비율 넘게 덮으면 "배경 판" — [V] 숙주에서 제외
V_SLACK = 1.25           # 띠 높이가 글자 높이의 이 배 넘어야 세로중앙이 의미 있다
V_BAR_H = 0.15           # 지면 높이의 이 이하 = 칩·band·아이콘 행 (카드는 제외)
BLOCK_CONTRAST = 3.0     # 이 미만이면 차단, 3.0~4.5 는 참고
TYPE_SCALE = {8.5, 10, 11.5, 13.5, 23, 28, 30, 32, 35}


# ── 유틸 ───────────────────────────────────────────────────────────────────
def installed_fonts() -> set[str]:
    """설치 글꼴명 스냅샷(레지스트리 키 + 폰트파일 패밀리명 en/ko).

    갱신: tools/dump_win_fonts.py
    """
    if FONTS_JSON.exists():
        return set(json.loads(FONTS_JSON.read_text(encoding="utf-8")))
    return set()


def _est_w_pt(s: str, size: float) -> float:
    w = 0.0
    for ch in s:
        if ch in _NARROW:
            w += _W_NARROW
        elif ch == " ":
            w += _W_SPACE
        elif ch.isdigit():
            w += _W_DIGIT
        elif ord(ch) > 0x2000:      # 한글·전각
            w += _W_HANGUL
        else:
            w += _W_LATIN
    return w * size


def _lum(rgb):
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a, b):
    la, lb = _lum(a), _lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def _hex2rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ── 도형 훑기 ──────────────────────────────────────────────────────────────
def _group_map(sh, parent):
    """그룹 자식 좌표(child 공간) → 슬라이드 좌표. OOXML: off + (c - chOff) × ext/chExt."""
    xf = sh._element.grpSpPr.find(qn("a:xfrm"))
    if xf is None:
        return parent
    def pair(tag, a, b):
        el = xf.find(qn(tag))
        return (int(el.get(a)), int(el.get(b))) if el is not None else (0, 0)
    (ox, oy), (ex, ey) = pair("a:off", "x", "y"), pair("a:ext", "cx", "cy")
    (cox, coy), (cex, cey) = pair("a:chOff", "x", "y"), pair("a:chExt", "cx", "cy")
    sx, sy = (ex / cex if cex else 1.0), (ey / cey if cey else 1.0)
    px, py, psx, psy = parent
    # 부모 변환 위에 이 그룹 변환을 얹는다: X = px + psx·(ox + sx·(c − cox))
    return (px + psx * (ox - sx * cox), py + psy * (oy - sy * coy), psx * sx, psy * sy)


def _iter_shapes(shapes, tf=(0, 0, 1.0, 1.0)):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(sh.shapes, _group_map(sh, tf))
            continue
        try:
            tx, ty, kx, ky = tf
            box = (round(tx + kx * sh.left), round(ty + ky * sh.top),
                   round(kx * sh.width), round(ky * sh.height))
        except TypeError:
            continue
        yield sh, box


def _solid_fill_hex(sh):
    """도형의 단색 fill hex. 그라데·없음·테마색이면 None.

    spPr 은 `p:` 네임스페이스라 `a:spPr` 로 찾으면 늘 None 이다(초기 버그).
    """
    sp = getattr(sh._element, "spPr", None)
    if sp is None:
        return None
    fill = sp.find(f"{NS}solidFill")
    if fill is not None:
        srgb = fill.find(f"{NS}srgbClr")
        return srgb.get("val").upper() if srgb is not None else None
    # 그라데이션 면도 밑면이다 — 색 정지점의 평균으로 본다.
    # 예전엔 None 을 돌려줘 페이지 배경(흰색)으로 오판했고, 파란 그라데 위 흰 글자가
    # 1.0:1 로 잡혔다(2026-08-12 포트폴리오 커버 실측).
    grad = sp.find(f"{NS}gradFill")
    if grad is None:
        return None
    cols = [g.get("val") for g in grad.iter(f"{NS}srgbClr") if g.get("val")]
    if not cols:
        return None
    acc = [0, 0, 0]
    for c in cols:
        for i in range(3):
            acc[i] += int(c[i * 2:i * 2 + 2], 16)
    return "".join(f"{round(v / len(cols)):02X}" for v in acc)


def _run_color_hex(r):
    srgb = r._r.find(f"{NS}rPr/{NS}solidFill/{NS}srgbClr")
    return srgb.get("val").upper() if srgb is not None else None


def _run_alpha(r):
    a = r._r.find(f"{NS}rPr/{NS}solidFill/{NS}srgbClr/{NS}alpha")
    return None if a is None else int(a.get("val")) / 1000


def _run_typefaces(r):
    rPr = r._r.find(f"{NS}rPr")
    if rPr is None:
        return {}
    out = {}
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(f"{NS}{tag}")
        if el is not None and el.get("typeface"):
            out[tag] = el.get("typeface")
    return out


def _run_size(r):
    rPr = r._r.find(f"{NS}rPr")
    if rPr is not None and rPr.get("sz"):
        return int(rPr.get("sz")) / 100
    return DEFAULT_SIZE


def _para_size(p):
    return max([_run_size(r) for r in p.runs] or [DEFAULT_SIZE])


def _has_hangul(s):
    return any("가" <= c <= "힣" for c in s)


def _anchor(sh):
    try:
        return sh.text_frame.vertical_anchor
    except Exception:
        return None


def _est_text_h_pt(tf, avail_w_pt):
    lines = 0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs) or p.text
        if not txt:
            lines += 1
            continue
        if tf.word_wrap is False:
            lines += 1
        else:
            lines += max(1, int(-(-_est_w_pt(txt, _para_size(p)) // avail_w_pt)))
    size = max([_para_size(p) for p in tf.paragraphs] or [DEFAULT_SIZE])
    return lines * size * LINE_FACTOR, lines


def _picture_rgb(sh, box, region, under=None):
    """그림 위 글자의 배경색 — 해당 영역 픽셀 평균을 실제로 뽑는다.

    알파 PNG(글로우·스크림 등)는 원색이 아니라 **아래 색(under)과 합성한 결과**가
    실제 기판이다. RGB 로만 읽으면 투명 글로우의 원색(채도 높은 코어색)을 기판으로
    오판한다.
    """
    try:
        import io

        from PIL import Image
        img = Image.open(io.BytesIO(sh.image.blob)).convert("RGBA")
    except Exception:
        return None
    px, py, pw, ph = box
    if not pw or not ph:
        return None
    x0, y0, x1, y1 = region
    iw, ih = img.size
    cx0 = max(0, min(iw - 1, int((x0 - px) / pw * iw)))
    cx1 = max(cx0 + 1, min(iw, int((x1 - px) / pw * iw)))
    cy0 = max(0, min(ih - 1, int((y0 - py) / ph * ih)))
    cy1 = max(cy0 + 1, min(ih, int((y1 - py) / ph * ih)))
    crop = img.crop((cx0, cy0, cx1, cy1))
    n = crop.width * crop.height
    if not n:
        return None
    ur, ug, ub = under if under is not None else (255.0, 255.0, 255.0)
    tot = [0.0, 0.0, 0.0]
    for r, g, b, a in crop.getdata():
        f = a / 255.0
        tot[0] += ur * (1 - f) + r * f
        tot[1] += ug * (1 - f) + g * f
        tot[2] += ub * (1 - f) + b * f
    return tuple(v / n for v in tot)


# ── 검사 ───────────────────────────────────────────────────────────────────
def _is_oval(sh) -> bool:
    """원·타원 오토셰이프인가 — 넘버 서클 판정용."""
    if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        return False
    try:
        return sh.auto_shape_type in (MSO_SHAPE.OVAL, MSO_SHAPE.DONUT)
    except (ValueError, AttributeError):
        return False


#: 가운데가 뚫린 프리셋 — 상자는 글자를 덮어도 **밑에 깔린 색이 아니다**.
#: (도넛·블록아크로 그린 링 안에 놓인 숫자를 링 색 위 1:1 로 오판했다)
_HOLLOW_PRESETS = {MSO_SHAPE.DONUT, MSO_SHAPE.BLOCK_ARC, MSO_SHAPE.ARC,
                   MSO_SHAPE.NO_SYMBOL, MSO_SHAPE.FRAME,
                   MSO_SHAPE.HALF_FRAME}


def _is_hollow(sh) -> bool:
    if sh.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        return False
    try:
        return sh.auto_shape_type in _HOLLOW_PRESETS
    except (ValueError, AttributeError):
        return False


def _bare_num(s: str) -> bool:
    """'1' '02' '3.' '4)' 처럼 순번만 들었나 — 실측치(45%·3건)는 배지가 아니다.

    **첫 줄만 본다.** 배지 도형 안에 순번과 라벨이 함께 들어가는 문법이 있어서
    ("01\n내용을\n입력해주세요") 문자열 전체로 판정하면 통째로 빠져나간다 —
    링 8개(01~08)가 그렇게 무선언·무검출로 지나갔다
    (2026-08-25 검수). 순번이 **맨 앞 줄 단독**이면 배지로 본다.
    """
    raw = (s or "").strip()
    if not raw:                      # 글자 없는 점은 축 표식이다 — 배지가 아니다
        return False
    head = raw.splitlines()[0].strip().rstrip(".)")
    if not head:
        return False
    head = head.lstrip("0") or "0"
    return head.isdigit() and len(head) <= 2


def _slide_bg_rgb(slide):
    """슬라이드 배경 속성의 단색. 렌더러가 배경 도형 대신 이걸 쓴다(2026-08-19)."""
    try:
        el = slide.background.fill._xPr.find(f"{NS}solidFill")
    except Exception:
        return None
    if el is None:
        return None
    srgb = el.find(f"{NS}srgbClr")
    return _hex2rgb(srgb.get("val").upper()) if srgb is not None else None


def check_slide(slide, sw, sh_, fonts):
    warns, notes = [], []
    slide_bg = _slide_bg_rgb(slide)
    shapes = list(_iter_shapes(slide.shapes))
    page_area = sw * sh_

    # ── [N] 넘버 서클 금지 (2026-08-17 반려: "이런 동그라미로 순서도 만드는
    #    인포그래픽 만들지마"). 원(또는 원에 가까운 도형) 안에 순번만 든 배지.
    #    layout_spec 을 안 거치는 gen_*.py 경로도 여기서 걸린다.
    ovals = [(s, b) for s, b in shapes if _is_oval(s)]
    for osh, obox in ovals:
        own = osh.text_frame.text.strip() if osh.has_text_frame else ""
        hit = own if _bare_num(own) else None
        if hit is None:                       # 원 위에 따로 얹은 숫자 텍스트
            ocx, ocy = obox[0] + obox[2] / 2, obox[1] + obox[3] / 2
            for s2, b2 in shapes:
                if s2 is osh or not s2.has_text_frame:
                    continue
                t = s2.text_frame.text.strip()
                if not _bare_num(t):
                    continue
                if (abs(b2[0] + b2[2] / 2 - ocx) < obox[2] / 2
                        and abs(b2[1] + b2[3] / 2 - ocy) < obox[3] / 2):
                    hit = t
                    break
        if hit is not None:
            warns.append(
                f"[N] 넘버 서클 금지: 원(x{obox[0] / sw:.3f} y{obox[1] / sh_:.3f}) 안 "
                f"'{hit}' — 동그라미 안 순번 순서도는 반려된다. 카드+화살표로 가거나 "
                f"번호를 카드 안 좌측 인덱스로 (CLAUDE.md 절대규칙 8-2)")

    def _covers(obox, box):
        ox, oy, ow, oh = obox
        x, y, w, h = box
        return ox <= x and oy <= y and ox + ow >= x + w and oy + oh >= y + h

    def bg_rgb(box, self_sh):
        """글자 밑에 실제로 깔린 색 — 뒤에 그린 게 위(마지막 승)."""
        found = None
        x, y, w, h = box
        for other, obox in shapes:
            if other is self_sh or not _covers(obox, box):
                continue
            if _is_hollow(other):      # 링·도넛 안은 뚫려 있다
                continue
            if other.shape_type == MSO_SHAPE_TYPE.PICTURE:
                s = _picture_rgb(other, obox, (x, y, x + w, y + h),
                                 under=found if found is not None else slide_bg)
                if s:
                    found = s
                continue
            hexv = _solid_fill_hex(other)
            if hexv:
                found = _hex2rgb(hexv)
        # 아무 면도 안 깔렸으면 슬라이드 배경이 밑면이다.
        # (배경을 도형이 아니라 배경 속성으로 칠하므로 여기서 받아야 [L]이 산다.)
        return found if found is not None else slide_bg

    for sh, box in shapes:
        x, y, w, h = box
        if not sh.has_text_frame:
            continue
        tf = sh.text_frame
        body = tf.text
        own_fill = _solid_fill_hex(sh)

        if not body.strip():
            if own_fill is None and sh.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                notes.append(f"[E] 빈 텍스트박스 잔재: x{x / sw:.3f} y{y / sh_:.3f}")
            continue

        ml, mr = tf.margin_left or 0, tf.margin_right or 0
        mt, mb = tf.margin_top or 0, tf.margin_bottom or 0
        avail_w_pt = max(1, w - ml - mr) / EMU_PER_PT
        avail_h_pt = max(1, h - mt - mb) / EMU_PER_PT
        need_h, nlines = _est_text_h_pt(tf, avail_w_pt)

        # ── [M] 제작 메모
        low = body.lower()
        hard = next((m for m in MEMO_HARD if m.lower() in low), None)
        if hard is None and body.strip().upper() in MEMO_EXACT:
            hard = body.strip()
        if hard is None:
            hit = next((p for p in MEMO_META_RE if re.search(p, body)), None)
            if hit:
                hard = re.search(hit, body).group(0)
        if hard:
            warns.append(f"[M] 제작 메모가 덱 문구에: '{hard}' in '{body[:24]}…'")
        else:
            soft = next((m for m in MEMO_SOFT if m.lower() in low), None)
            if soft:
                notes.append(f"[M] 제작 메모일 수 있음: '{soft}' in '{body[:24]}…'")

        # ── [V] 칩·바·아이콘 행 안 글자는 anchor=middle (규칙 6)
        #
        # 겨냥 대상은 **자기 면을 가진 얇은 띠**(칩·band·아이콘 행)다. 줄간격이
        # 늘 아래로 밀어 손좌표로는 절대 못 맞추는 자리. 카드 제목·코드블록처럼
        # 위로 붙는 게 맞는 글자까지 잡으면 소음이 된다(승인 덱 실측 60건 오탐).
        # 컨테이너 카드 위 손좌표 텍스트는 layout_spec 단계에서 check_align [V]가 본다.
        if own_fill and h <= sh_ * V_BAR_H and nlines <= 2 \
                and _anchor(sh) not in (MSO_ANCHOR.MIDDLE, MSO_ANCHOR.BOTTOM) \
                and h > need_h * EMU_PER_PT * V_SLACK:
            warns.append(
                f"[V] 띠 안 글자에 anchor=middle 없음: '{body[:16]}…' "
                f"(y{y / sh_:.3f} h{h / sh_:.3f}) — "
                f"줄간격이 늘 아래로 민다, 손좌표로 맞추지 말 것")

        # ── run 단위
        for p in tf.paragraphs:
            for r in p.runs:
                if not r.text.strip():
                    continue
                size = _run_size(r)
                tfaces = _run_typefaces(r)

                for tag, name in tfaces.items():
                    if fonts and name not in fonts:
                        warns.append(
                            f"[F] 미설치 글꼴 '{name}' ({tag}) — '{r.text[:14]}…' "
                            f"→ 파워포인트가 다른 글꼴로 흘린다")
                if _has_hangul(r.text) and "ea" not in tfaces:
                    warns.append(
                        f"[F] 한글 run 에 ea 글꼴 미지정: '{r.text[:14]}…' "
                        f"(latin={tfaces.get('latin')}) — ea/cs 에도 같은 이름을")

                rPr = r._r.find(f"{NS}rPr")
                spc = rPr.get("spc") if rPr is not None else None
                if spc and _run_alpha(r) is not None:
                    warns.append(
                        f"[A] 자간(spc={spc})+알파 동시: '{r.text[:16]}…' — "
                        f"LibreOffice가 마지막 글자를 먹는다. 알파 대신 색으로")

                col = _run_color_hex(r)
                if col:
                    bg = _hex2rgb(own_fill) if own_fill else bg_rgb(box, sh)
                    if bg:
                        ratio = _contrast(_hex2rgb(col), bg)
                        need = 3.0 if size >= 18 else 4.5
                        msg = (f"[L] 대비 {ratio:.1f}:1 (기준 {need}) — "
                               f"'{r.text[:16]}…' #{col} on "
                               f"rgb{tuple(int(v) for v in bg)}")
                        if ratio < BLOCK_CONTRAST:
                            warns.append(msg)
                        elif ratio < need:
                            notes.append(msg)

                if size not in TYPE_SCALE:
                    notes.append(f"[T] 타입스케일 밖 {size}pt: '{r.text[:14]}…'")

        # ── [O] 넘침
        for p in tf.paragraphs:
            txt = "".join(r.text for r in p.runs) or p.text
            if not txt:
                continue
            est = _est_w_pt(txt, _para_size(p))
            if tf.word_wrap is False and est > avail_w_pt * OVERFLOW_TOL:
                warns.append(
                    f"[O] 넘침(줄바꿈 꺼짐): '{txt[:18]}…' {_para_size(p)}pt "
                    f"추정폭 {est:.0f}pt > 상자 {avail_w_pt:.0f}pt")
        # 높이는 **채움 면**일 때만 — 빈 텍스트박스는 넘쳐도 안 잘린다
        if own_fill and need_h > avail_h_pt * OVERFLOW_TOL_H:
            warns.append(
                f"[O] 글자가 면 밖으로: '{body[:18]}…' {nlines}줄 ⇒ "
                f"{need_h:.0f}pt > 면 {avail_h_pt:.0f}pt")

        # ── [C] 판면 이탈 — 상자가 아니라 **글자**가 나가야 이탈
        align = tf.paragraphs[0].alignment
        for p in tf.paragraphs:
            txt = "".join(r.text for r in p.runs) or p.text
            if not txt:
                continue
            est = min(_est_w_pt(txt, _para_size(p)) * EMU_PER_PT, w - ml - mr) \
                if tf.word_wrap is not False else _est_w_pt(
                    txt, _para_size(p)) * EMU_PER_PT
            if align == PP_ALIGN.CENTER:
                sx = x + (w - est) / 2
            elif align == PP_ALIGN.RIGHT:
                sx = x + w - mr - est
            else:
                sx = x + ml
            if sx < -sw * 0.01 or sx + est > sw * 1.01:
                warns.append(
                    f"[C] 글자가 판면 밖으로: '{txt[:18]}…' "
                    f"x{sx / sw:.3f}~{(sx + est) / sw:.3f}")
                break
        if y + h < -sh_ * 0.01 or y > sh_ * 1.01:
            warns.append(f"[C] 지면 밖 도형: '{body[:16]}…' y{y / sh_:.3f}")
    return warns, notes


def _sig(msg):
    """같은 원인끼리 묶기 위한 서명 — 인용문·좌표·수치를 지운 뼈대."""
    s = re.sub(r"'[^']*'", "'…'", msg)
    s = re.sub(r"[-+]?\d*\.?\d+", "#", s)
    return s


# ── [X] OOXML 값 범위 ────────────────────────────────────────────────
# PowerPoint 는 범위를 벗어난 수치를 만나면 파일 전체에 "복구" 를 건다.
# 복구는 그 도형 하나만 지우는 게 아니라 **그 뒤 도형을 통째로 날린다** —
# LibreOffice·PDF·썸네일은 멀쩡하니 눈검수로는 절대 안 잡힌다.
# (2026-08-24 실측: fill_alpha 를 0~1 이 아니라 퍼센트 44 로 줘서
#  <a:alpha val="4400000"/> 이 박혔고, 사용자가 열자 3개 장의 내용이 사라졌다.)
_OOXML_RANGES = {
    "alpha":   (0, 100000),      # 백분율 × 1000
    "lumMod":  (0, 1000000),
    "lumOff":  (0, 1000000),
    "shade":   (0, 100000),
    "tint":    (0, 100000),
    "satMod":  (0, 1000000),
    "spcPct":  (0, 13200000),
}


def check_package(path):
    """압축을 직접 열어 슬라이드 XML 의 수치 범위를 검사한다.

    python-pptx 객체 모델로는 안 보이는 층이라 여기서만 잡힌다.
    """
    import re
    import zipfile
    warns = []
    with zipfile.ZipFile(str(path)) as z:
        names = sorted((n for n in z.namelist()
                        if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=lambda n: int(re.findall(r"\d+", n)[0]))
        for n in names:
            pg = int(re.findall(r"\d+", n)[0])
            x = z.read(n).decode("utf-8")
            for tag, (lo, hi) in _OOXML_RANGES.items():
                for m in re.finditer(rf'<a:{tag} val="(-?\d+)"\s*/?>', x):
                    v = int(m.group(1))
                    if not (lo <= v <= hi):
                        warns.append(
                            f"p{pg:02d} [X] OOXML 값 범위 초과: <a:{tag} val=\"{v}\"/> "
                            f"— 허용 {lo}~{hi}. PowerPoint 가 이 파일에 「복구」를 걸고 "
                            f"뒤따르는 도형을 지운다 (fill_alpha 는 퍼센트가 아니라 0~1)")
            # 슬라이드 루트 그룹(<p:spTree> 바로 밑 <p:grpSpPr>)의 xfrm 은 도형이 아니다.
            # PowerPoint 는 저장할 때마다 여기에 ext 0 · chExt 0 을 표준으로 쓴다 —
            # 사용자가 PowerPoint 로 손본 파일을 받으면 전 장이 오탐으로 잡혔다
            # 루트 블록만 빼고 잰다.
            x_sz = re.sub(r"(<p:spTree>\s*<p:nvGrpSpPr>.*?</p:nvGrpSpPr>\s*)"
                          r"<p:grpSpPr>.*?</p:grpSpPr>", r"\1", x, count=1, flags=re.S)
            for m in re.finditer(r'<a:ext cx="(-?\d+)" cy="(-?\d+)"\s*/?>', x_sz):
                cx, cy = int(m.group(1)), int(m.group(2))
                if cx <= 0 or cy <= 0:
                    warns.append(f"p{pg:02d} [X] 크기가 0 이하인 도형: cx={cx} cy={cy} "
                                 f"— PowerPoint 복구 대상")
            for m in re.finditer(r'sz="(-?\d+)"', x):
                v = int(m.group(1))
                if not (100 <= v <= 400000):
                    warns.append(f"p{pg:02d} [X] 글자 크기 범위 초과: sz={v} (허용 100~400000)")
            # 도형 id 중복 — 이식 XML(지도 그룹 등)이 원본 번호를 들고 오면 생긴다.
            # PowerPoint 가 「복구」를 걸고 복구 시 도형을 통째로 버린다
            # (실기기 재현).
            ids = re.findall(r'<p:cNvPr id="(\d+)"', x)
            dups = sorted({i for i in ids if ids.count(i) > 1}, key=int)
            if dups:
                warns.append(
                    f"p{pg:02d} [X] cNvPr id 중복 {len(dups)}건 (id {', '.join(dups[:8])}"
                    f"{'…' if len(dups) > 8 else ''}) — PowerPoint 복구 유발. "
                    f"이식 XML 은 id 재번호 후 붙일 것")
            # 애니메이션이 이 장에 없는 도형을 가리킨다 — PowerPoint 가 「복구」로 연다
            # (2026-10-08 주입 실측: spTgt spid 를 없는 번호로 바꾸자 COM Open 이 예외).
            if "<p:timing" in x:
                tg = set(re.findall(r'<p:spTgt spid="(\d+)"', x))
                lost = sorted(tg - set(ids), key=int)
                if lost:
                    warns.append(f"p{pg:02d} [MOT] 애니가 없는 도형(spid {', '.join(lost[:6])})을 "
                                 f"가리킨다 — PowerPoint 복구 유발")
            bang = re.findall(r'<p:cNvPr id="\d+" name="(!![^"]*)"', x)
            twice = sorted({b for b in bang if bang.count(b) > 1})
            if twice:
                warns.append(f"p{pg:02d} [MOT] 같은 장에 모핑 이름 {', '.join(twice[:4])} 이 여러 개 "
                             f"— 짝이 엇갈린다(MS 규칙: !! 이름은 장마다 하나)")
        # 차트 axId 음수 — ISO 29500 은 xsd:unsignedInt (python-pptx 템플릿이
        # -2068027336 을 박는다). 렌더러 _add_chart 가 abs 리매핑하니 걸리면
        # 렌더러를 우회해 만든 차트다.
        for n in (n for n in z.namelist()
                  if re.fullmatch(r"ppt/charts/chart\d+\.xml", n)):
            x = z.read(n).decode("utf-8")
            neg = sorted({m for m in re.findall(
                r'<c:(?:axId|crossAx) val="(-\d+)"', x)}, key=int)
            if neg:
                warns.append(
                    f"{n.split('/')[-1]} [X] 차트 axId/crossAx 음수 {', '.join(neg)} "
                    f"— ISO 29500 unsignedInt 위반(스키마 검증 탈락)")
            # line/area 차트에 dLblPos outEnd — 막대 전용 값이라 규격 위반.
            # PowerPoint 가 복구 프롬프트를 띄운다 (2026-09-02 실기기 이분탐색 확정).
            if re.search(r"<c:(?:lineChart|areaChart)", x) and \
               re.search(r'<c:dLblPos val="outEnd"', x):
                warns.append(
                    f"{n.split('/')[-1]} [X] line/area 차트에 dLblPos=outEnd "
                    f"— 막대 전용 값, PowerPoint 복구 유발(ABOVE/t 로)")
    # 같은 원인 반복은 한 줄로
    seen, out = set(), []
    for w in warns:
        k = re.sub(r"^p\d+ ", "", w)
        if k in seen:
            continue
        seen.add(k)
        out.append(w)
    return out


def check_file(path, fonts=None, notes=False, group=True, exempt_pages=None,
               skip_codes=None):
    """pptx 하나를 검사해 경고(기본) 또는 참고 리스트를 낸다.

    같은 원인이 수백 번 반복되면 게이트가 소음이 된다 → 서명이 같은 건 한 줄로
    묶고 `×N (p03, p07…)` 로 붙인다. group=False 면 원본 그대로.

    `exempt_pages` 는 `[N]` 이 열리는 면 번호 집합이다. **묶기 전에** 뺀다 —
    묶은 뒤에 문자열로 걸러내면 `p09, p15, p20` 처럼 면제면과 비면제면이 한 줄에
    섞였을 때 줄 전체가 사라져 **비면제면의 위반까지 조용히 통과한다**
    면제는 출력 필터가 아니라 생성 지점이다.
    """
    fonts = installed_fonts() if fonts is None else fonts
    prs = Presentation(str(path))
    sw, sh_ = prs.slide_width, prs.slide_height
    multi = len(prs.slides._sldIdLst) > 1
    raw = []
    for i, slide in enumerate(prs.slides, 1):
        w, n = check_slide(slide, sw, sh_, fonts)
        for m in (n if notes else w):
            if exempt_pages and i in exempt_pages and m.startswith("[N]"):
                continue
            # 스펙 우회 규약(`gate_skip` + `gate_reason`) — 사유 없는 우회는
            # `_spec_skips` 가 애초에 담지 않는다. [N] 은 여기서 열리지 않는다
            # (넘버 서클은 사람이 손으로 못 여는 자동 판정이다).
            if skip_codes and m.startswith("["):
                if m[1:m.index("]")] in skip_codes.get(i, ()):
                    continue
            raw.append((f"p{i:02d}", m))
    # 패키지 층(수치 범위) — 눈검수·PDF 로는 안 잡히는 「PowerPoint 복구」 유발값
    pkg = [] if notes else check_package(path)
    if not group:
        return pkg + [f"{pg} {m}" if multi else m for pg, m in raw]

    order, bucket = [], {}
    for pg, m in raw:
        k = _sig(m)
        if k not in bucket:
            bucket[k] = [m, []]
            order.append(k)
        if pg not in bucket[k][1]:
            bucket[k][1].append(pg)
    out = []
    for k in order:
        first, pages = bucket[k]
        head = f"{pages[0]} {first}" if multi else first
        if len(pages) > 1:
            shown = ", ".join(pages[:6]) + ("…" if len(pages) > 6 else "")
            head += f"  ×{len(pages)}장 ({shown})"
        out.append(head)
    return pkg + out


def _nc_exempt_pages(pptx_path):
    """[N] 이 열리는 **면 번호** 집합. 사람이 켜는 플래그가 아니라 스펙이 정한다.

    `tracks/_shared/GATES.md` — "`--allow-number-circles` 플래그는 폐기됐다.
    넘버 서클 허용 여부는 사람이 손으로 켜는 게 아니라 `gate.json` 의
    `figures.number_circles` + `ref.reuse` 조합으로 자동 판정된다.
    손으로 켜는 플래그가 있으면 그게 우회로가 된다."

    그래서 pptx 옆의 스펙을 직접 읽어 **면 단위로** 연다. 덱 단위로 열면
    다음 면이 넘버 서클을 발명해도 안 잡힌다.
    """
    proj = pathlib.Path(pptx_path).resolve().parent.parent
    sys.path.insert(0, str(proj.parent.parent))
    try:
        from pptm import pages as pg_list
    except ImportError:
        return set()
    pages = pg_list.spec_files(proj)
    if not pages:
        return set()
    try:
        from pptm import track as tr
        nc = ((tr.gate_profile(tr.resolve(proj.name)[0]) or {})
              .get("figures", {}).get("number_circles", "deny"))
    except Exception:
        return set()
    if nc not in ("allow", "origin_only"):
        return set()
    out = set()
    for f in pages:
        # 면 번호는 **디렉터리 이름에서 읽는다.** 정렬 순서(enumerate)로 세면
        # p03 이 빠진 덱에서 이후 면제가 통째로 한 칸씩 밀린다.
        i = int(f.parent.name[1:])
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        reuse = ((d.get("ref") or {}).get("reuse") or [])
        if nc == "allow" or "figure" in reuse:
            out.add(i)
    return out


def _spec_skips(pptx_path):
    """면 번호 → 우회 코드 집합. 스펙의 `gate_skip` + `gate_reason` 을 그대로 읽는다.

    `tracks/_shared/GATES.md` 「우회 규약」 — "스펙에 `gate_skip`/`gate_reason`,
    **사유가 없으면 우회 자체가 실패**한다". 지금까지 이 규약은 스펙 게이트
    (`check_align.py`)에만 걸려 있었고 납품물 게이트는 무시했다. 그래서 원전 문법상
    일부러 흐리게 깔아야 하는 것(고스트 넘버·거대 인용부호·판면 밖 블리드)이
    구조적으로 0건에 도달할 수 없었다. **면 단위**로만 연다.

    **접두사 `pptx:` 가 붙은 항목만 읽는다.** 두 게이트는 같은 글자를 다른 뜻으로 쓴다
    (`[C]` = check_align 「같은 행 y 불일치」 ↔ check_pptx 「판면 밖」,
    `[L]` = 「사진 위 대비」 ↔ 「WCAG 대비」). 접두사가 없으면 스펙 게이트를 열려고 적은
    한 글자가 납품물 게이트의 다른 결함까지 조용히 덮는다.

    `[N]` 은 여기서 열리지 않는다 — 넘버 서클 면제는 `_nc_exempt_pages` 의 자동
    판정 소관이고, 손으로 켤 수 있으면 그게 우회로가 된다.
    """
    proj = pathlib.Path(pptx_path).resolve().parent.parent
    sys.path.insert(0, str(proj.parent.parent))
    try:
        from pptm import pages as pg_list
    except ImportError:
        return {}
    out = {}
    for f in pg_list.spec_files(proj):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        codes = {str(c)[5:].strip("[]").upper()
                 for c in (d.get("gate_skip") or [])
                 if str(c).lower().startswith("pptx:")}
        codes.discard("N")
        if codes and str(d.get("gate_reason") or "").strip():
            out[int(f.parent.name[1:])] = codes
    return out


def main(argv):
    show_notes = "--notes" in argv
    if "--allow-number-circles" in argv:
        print("  ! --allow-number-circles 는 폐기된 플래그다 — 무시한다. "
              "넘버 서클은 스펙의 ref.reuse 로 면 단위 판정된다"
              "(tracks/_shared/GATES.md).")
    files = [a for a in argv if not a.startswith("-")]
    if not files:
        print("사용: python3 tools/check_pptx.py <deck.pptx> [...] [--notes]")
        return 2
    fonts = installed_fonts()
    if not fonts:
        print("  ! 글꼴 스냅샷 없음 — [F] 건너뜀 (python3 tools/dump_win_fonts.py)")
    grand = 0
    for f in files:
        print(f"■ {f}")
        warns = check_file(f, fonts=fonts, exempt_pages=_nc_exempt_pages(f),
                           skip_codes=_spec_skips(f))
        for w in warns:
            print(f"    ✗ {w}")
        if not warns:
            print("    ✓")
        grand += len(warns)
        if show_notes:
            for n in check_file(f, fonts=fonts, notes=True):
                print(f"    · {n}")
    print()
    if grand:
        print(f"✗ pptx 경고 {grand}건 — 완성 선언 금지")
        return 1
    print("✓ pptx 검수 통과 (0건)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
