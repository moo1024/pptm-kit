"""네이티브 PPT 렌더러 + 테마 라이브러리 (아키텍처 v7 — Claude 직접 디자인).

Claude가 작성한 pages/pNN/layout_spec.json을 python-pptx 네이티브 개체로 그린다.
텍스트는 100% 편집 가능. AI 이미지는 image 아이템(스팟 슬롯)에만 쓴다.

layout_spec: {"page": N, "theme": "gyoan|anthropic|cupertino|stripe|figma|vercel",
              "variant": "paper|dark", "items": [...]}

테마 5종(gyoan 제외)은 공개 디자인 시스템 규칙을 한글 타이포로
번안한 것 — 팔레트 hex·서피스 문법·시그니처 무브를 원전에서 그대로 가져온다.

아이템 타입:
  glow / label / serif_label / headline / text / card / pill / chip /
  hairline / arrow / sowhat / timetable / bars / steps / image /
  mockup(기기 목업 — 폰·노트북·모니터·TV·워치 프레임 + 화면 합성) /
  map(세계·한국·서울 지도 — 네이티브 도형 이식, 지역별 단계구분도) /
  chart(네이티브 PPT 차트 — 엑셀 백업 데이터, 편집 가능) /
  table(네이티브 PPT 표 — 셀 데이터 편집 가능)
  ※ 그래프·표는 반드시 chart/table(네이티브 개체)로 — card 사각형으로 그리지 말 것.

모션(슬라이드쇼 재생): 페이지 키 transition · 아이템 키 anim / morph_id → pptm/design/motion.py

runs: [{"t": "텍스트", "color"?: hex, "weight"?: "bold", "font"?: 폰트키, "size"?: pt}]
좌표는 전부 슬라이드 비율(0~1).
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Emu, Pt

from .. import pages as pages_mod

SLIDE_W = 12192000
SLIDE_H = 6858000

# 페이지 크기 프리셋 — layout_spec의 "page_size" 키(덱 단위, 첫 페이지 기준)
PAGE_SIZES = {
    "16:9": (12192000, 6858000),
    "a4p": (6858000, 9906000),   # A4 세로 (결과보고서·수료증 등 문서형)
    # 실측 A4 세로 210×297mm — 인쇄·PDF 배포물(a4p는 540×780pt로 A4가 아니다)
    "a4": (210 * 36000, 297 * 36000),
    # 실측 A4 가로 297×210mm — 공고가 「A4 가로형 프레젠테이션」을 요구하는 서류심사 제안서.
    # 가로 A4 는 덱이다.
    "a4l": (297 * 36000, 210 * 36000),
    "3:4": (6858000, 9144000),   # 세로 포스터·키비주얼 (인스타·에타 마스터)
    # A2 세로 420×594mm — 인쇄 모집 포스터 마스터. A3(297×420) 축소도 같은 조판.
    # pt는 절대 단위이므로 이 판형에선 글자 크기가 16:9보다 3배 이상 커진다.
    "a2p": (420 * 36000, 594 * 36000),
    # 인스타 피드 4:5 (1080×1350) 파생 — 254×317.5mm
    "4:5": (9144000, 11430000),
    # B5(257×182) 2단 접지의 한 면 = 128.5×182mm. 리플렛·안내책자용.
    "b5half": (4626000, 6552000),
    # b5half 의 2배 작업용 캔버스(257×364mm). 최종 인쇄물은 50% 축소가 전제다.
    "b5half2x": (9252000, 13104000),
    # 대형 포스터 양식의 「도식화 자료 삽입」 프레임 한 칸(6.65×4.38in).
    # 남의 양식 안에 넣을 도식은 프레임 비율로 조판해야 여백이 안 남는다.
    # 폭이 478.8pt(16:9의 절반)라 TYPE SCALE 을 그만큼 키워 잡는다.
    "figframe": (6080760, 4005072),
    # 행사 현수막 4000×700mm 의 1/5 축소(800×140mm) — 실물이 PPT 최대 판형(56in)을
    # 넘어 같은 비율로 줄였다. 인쇄 발주 시 500% 확대, 5.714:1 비율 정확.
    "banner5x": (800 * 36000, 140 * 36000),
    # X배너 600×1800mm 의 1/2 축소(300×900mm) — 1:3. 인쇄 발주 시 200% 확대.
    "xbanner2x": (300 * 36000, 900 * 36000),
    # 부스 상단 현수막 1940×480mm 의 1/2 축소(970×240mm) — 4.0417:1.
    # 실물 폭이 PPT 최대 판형(56in=1422mm)을 넘어 절반으로 줄였다. 인쇄 발주 시 200% 확대.
    # 위아래 각 100mm 가 재봉선 → 정규화 y 0.2083 / 0.7917 밖으로 글자가 나가면 잘린다.
    "boothbanner2x": (970 * 36000, 240 * 36000),
    # 목걸이 명찰 내지 95×123mm — 실물 1:1(케이스 103×133mm 안에 들어간다).
    # 케이스가 상하좌우 5mm 씩 크고 **구멍은 케이스 위쪽**에 온다 → 내지 상단
    # 14mm(정규화 y 0.114) 안쪽에는 잘려도 되는 것만 둔다.
    "badge95x123": (95 * 36000, 123 * 36000),
    # 같은 명찰의 **인쇄 데이터** 판 — 재단 95×123 에 사방 3mm 도련을 더한 크기.
    # 지면이 풀블리드라 도련 없이 넘기면 재단 오차만큼 흰 실선이 난다.
    "badge101x129": (101 * 36000, 129 * 36000),
}


def _set_page_size(key: str):
    global SLIDE_W, SLIDE_H
    if key not in PAGE_SIZES:
        raise ValueError(f"unknown page_size: {key} (지원: {list(PAGE_SIZES)})")
    SLIDE_W, SLIDE_H = PAGE_SIZES[key]

# ── 교안 토큰 (하위 호환용 상수) ──────────────────────────────────────────────
PAPER = "F7F5F1"
INK = "1B1A18"
ORANGE = "E8622C"
CARD = "EFECE7"
CARD_DARKER = "E9E5DF"
GRAY = "8A857D"
HAIR = "D8D3CB"
DARK = "0B0A09"
WHITE = "F5F2ED"
DARK_SUB = "B7A99A"

FONTS = {
    "display_b": ("Paperlogy 7 Bold", False),
    "display_sb": ("Paperlogy 6 SemiBold", False),
    "body": ("Pretendard", False),
    "body_sb": ("Pretendard SemiBold", False),
    "body_b": ("Pretendard", True),
    "serif": ("Georgia", False),
}

# ── 테마 라이브러리 ───────────────────────────────────────────────────────────
# 각 테마 = 색 토큰 + 폰트 매핑 + 글로우 팔레트. 프리미티브는 전부 공유.
THEMES: dict[str, dict] = {
    # 교안 (기본): 웜페이퍼 + 챠콜 + 번트오렌지, Paperlogy/Pretendard
    "gyoan": {
        "bg": PAPER, "ink": INK, "accent": ORANGE, "on_accent": "FFFFFF",
        "card": CARD, "card2": CARD_DARKER, "gray": GRAY, "hair": HAIR,
        "arrow": "B9B2A8", "highlight": "F6DEC9", "body_ink": "4A463F",
        "dark_bg": DARK, "dark_fg": WHITE, "dark_sub": DARK_SUB,
        "fonts": dict(FONTS),
        "glow": {"core": (242, 122, 40), "deep": (120, 48, 14), "base": (11, 10, 9)},
    },
    # 앤트로픽(Claude DESIGN.md): 크림 캔버스 + 코랄 한 점, 세리프 디스플레이.
    # "굵기가 아니라 크기로 말한다" — 헤드라인은 바탕체, 라벨/본문은 산스.
    "anthropic": {
        "bg": "FAF9F5", "ink": "141413", "accent": "CC785C", "on_accent": "FFFFFF",
        "card": "EFE9DE", "card2": "E8E0D2", "gray": "6C6A64", "hair": "E6DFD8",
        "arrow": "C9C1B4", "highlight": "F5F0E8", "body_ink": "3D3D3A",
        "dark_bg": "181715", "dark_fg": "FAF9F5", "dark_sub": "A09D96",
        "fonts": {
            "display_b": ("KoPubWorld바탕체 Bold", False),
            "display_sb": ("Pretendard SemiBold", False),
            "body": ("Pretendard", False),
            "body_sb": ("Pretendard SemiBold", False),
            "body_b": ("Pretendard", True),
            "serif": ("KoPubWorld바탕체 Light", False),
        },
        "glow": {"core": (230, 140, 105), "deep": (140, 60, 35), "base": (24, 23, 21)},
    },
    # 쿠퍼티노(Apple DESIGN.md): 화이트/파치먼트/다크 타일 3서피스, 단일 블루 악센트.
    # 시스템 전체에 그림자·장식 0 — 서피스 색 전환이 곧 구획.
    "cupertino": {
        "bg": "FFFFFF", "ink": "1D1D1F", "accent": "0066CC", "on_accent": "FFFFFF",
        "card": "F5F5F7", "card2": "EBEBED", "gray": "86868B", "hair": "E0E0E0",
        "arrow": "C7C7CC", "highlight": "D9E8FA", "body_ink": "424245",
        "dark_bg": "161617", "dark_fg": "F5F5F7", "dark_sub": "98989D",
        "fonts": {
            "display_b": ("AppleSDGothicNeoB00", False),
            "display_sb": ("AppleSDGothicNeoSB00", False),
            "body": ("AppleSDGothicNeoR00", False),
            "body_sb": ("AppleSDGothicNeoSB00", False),
            "body_b": ("AppleSDGothicNeoB00", False),
            "serif": ("Georgia", False),
        },
        "glow": {"core": (238, 238, 244), "deep": (70, 70, 78), "base": (10, 10, 11)},
    },
    # 스트라이프(Stripe DESIGN.md): 딥네이비 잉크 + 인디고, 라이트웨이트 엘레강스.
    # 디스플레이는 에스코어드림 라이트 — 얇은 대형 활자가 브랜드 그 자체.
    "stripe": {
        "bg": "FFFFFF", "ink": "0D253D", "accent": "533AFD", "on_accent": "FFFFFF",
        "card": "F6F9FC", "card2": "EAF0F7", "gray": "64748D", "hair": "E3E8EE",
        "arrow": "B9C4D9", "highlight": "E7E4FE", "body_ink": "3C4257",
        "dark_bg": "1C1E54", "dark_fg": "F6F9FC", "dark_sub": "A5AEE0",
        "fonts": {
            "display_b": ("에스코어 드림 7 ExtraBold", False),
            "display_sb": ("에스코어 드림 5 Medium", False),
            "body": ("Pretendard", False),
            "body_sb": ("Pretendard SemiBold", False),
            "body_b": ("Pretendard", True),
            "serif": ("D2Coding", False),
        },
        "glow": {"core": (160, 120, 255), "deep": (80, 60, 220), "base": (28, 30, 84)},
    },
    # 피그마(Figma DESIGN.md): 흑백 크롬 + 파스텔 컬러블록, 모노 아이브로우.
    # 구조는 전부 블랙/화이트, 색은 블록(카드 fill 오버라이드)으로만.
    "figma": {
        "bg": "FFFFFF", "ink": "0A0A0A", "accent": "0A0A0A", "on_accent": "FFFFFF",
        "card": "F7F7F5", "card2": "EDEDE8", "gray": "6E6E6B", "hair": "E6E6E6",
        "arrow": "C9C9C9", "highlight": "DCEEB1", "body_ink": "1A1A1A",
        "dark_bg": "1F1D3D", "dark_fg": "FFFFFF", "dark_sub": "C9C0EE",
        "fonts": {
            "display_b": ("더잠실 5 Bold", False),
            "display_sb": ("더잠실 4 Medium", False),
            "body": ("Pretendard", False),
            "body_sb": ("Pretendard SemiBold", False),
            "body_b": ("Pretendard", True),
            "serif": ("D2Coding", False),
        },
        "glow": {"core": (197, 176, 244), "deep": (110, 85, 200), "base": (31, 29, 61)},
    },
    # 버셀(Vercel DESIGN.md): 잉크 모노크롬 + 모노스페이스 라벨, 다크 폴라리티 밴드.
    # 웨이트 상한 600, 문장형+마침표 헤드라인, 색은 링크 블루 한 점만.
    "vercel": {
        "bg": "FAFAFA", "ink": "171717", "accent": "171717", "on_accent": "FFFFFF",
        "card": "FFFFFF", "card2": "F5F5F5", "gray": "888888", "hair": "EBEBEB",
        "arrow": "BFBFBF", "highlight": "D3E5FF", "body_ink": "4D4D4D",
        "dark_bg": "0A0A0A", "dark_fg": "EDEDED", "dark_sub": "8F8F8F",
        "fonts": {
            "display_b": ("Freesentation 7 Bold", False),
            "display_sb": ("Freesentation 5 Medium", False),
            "body": ("Pretendard", False),
            "body_sb": ("Pretendard SemiBold", False),
            "body_b": ("Pretendard", True),
            "serif": ("D2Coding", False),
        },
        "glow": {"core": (0, 190, 220), "deep": (100, 40, 190), "base": (10, 10, 10)},
    },
}

# TYPE SCALE — 크기·폰트키·색 "역할" (색은 테마에서 해석)
TYPO_ROLES = {
    "label":     {"size": 10,   "font": "display_sb", "role": "accent"},
    "headline":  {"size": 23,   "font": "display_b",  "role": "ink"},
    "divider_h": {"size": 30,   "font": "display_b",  "role": "dark_fg"},
    "sub":       {"size": 13,   "font": "body",       "role": "gray"},
    "cardtitle": {"size": 13.5, "font": "display_sb", "role": "ink"},
    "body":      {"size": 11.5, "font": "body",       "role": "body_ink"},
    "sowhat":    {"size": 12,   "font": "body_sb",    "role": "ink"},
    "caption":   {"size": 8.5,  "font": "body",       "role": "gray"},
    "chip":      {"size": 10,   "font": "body_sb",    "role": "on_accent"},
    "pill":      {"size": 10.5, "font": "body_sb",    "role": "on_accent"},
    "num_xl":    {"size": 34,   "font": "display_b",  "role": "ink"},
    "meta":      {"size": 10.5, "font": "body",       "role": "dark_sub"},
    "serif":     {"size": 11,   "font": "serif",      "role": "dark_sub"},
}

_ACTIVE = THEMES["gyoan"]  # render_page가 페이지별로 설정

# JSON 테마 레지스트리 — 코드 수정 없이 테마 추가.
# 탐색 순서: 프로젝트/themes/<name>.json → pptm/design/themes/<name>.json.
# JSON은 THEMES 항목과 같은 구조(부분 정의 가능 — 빠진 키는 gyoan에서 상속).
_BUILTIN_THEME_DIR = Path(__file__).parent / "themes"


def _normalize_theme(raw: dict) -> dict:
    t = {**THEMES["gyoan"], **raw}
    t["fonts"] = {**THEMES["gyoan"]["fonts"],
                  **{k: tuple(v) for k, v in raw.get("fonts", {}).items()}}
    if "glow" in raw:
        t["glow"] = {k: tuple(v) for k, v in raw["glow"].items()}
    return t


def resolve_theme(name: str, project_dir: Path | None = None) -> dict:
    if name in THEMES:
        return THEMES[name]
    candidates = []
    if project_dir is not None:
        candidates.append(Path(project_dir) / "themes" / f"{name}.json")
    candidates.append(_BUILTIN_THEME_DIR / f"{name}.json")
    for p in candidates:
        if p.exists():
            theme = _normalize_theme(json.loads(p.read_text()))
            THEMES[name] = theme  # 캐시
            return theme
    raise KeyError(f"테마 '{name}' 없음 — THEMES 내장 또는 themes/{name}.json 필요")


def _typo(key: str) -> dict:
    r = TYPO_ROLES[key]
    return {"size": r["size"], "font": r["font"], "color": _ACTIVE[r["role"]]}


# 하위 호환: 기존 코드/스펙이 TYPO를 참조하면 gyoan 값으로 해석
TYPO = {k: {"size": v["size"], "font": v["font"], "color": THEMES["gyoan"][v["role"]]}
        for k, v in TYPO_ROLES.items()}


def _c(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr)


def _x(v: float) -> Emu:
    return Emu(int(v * SLIDE_W))


def _y(v: float) -> Emu:
    return Emu(int(v * SLIDE_H))


def _set_runs(para, runs, default_style: dict, spacing: float | None = None):
    para.alignment = PP_ALIGN.LEFT
    for r in runs:
        run = para.add_run()
        run.text = r["t"]
        fkey = r.get("font", default_style["font"])
        # 테마 fonts 의 키가 아니면 글꼴 실명으로 본다(라틴 세리프 이탤릭 등 일회성 지정).
        fname, fbold = _ACTIVE["fonts"].get(fkey, (fkey, False))
        run.font.name = fname
        # font.name은 <a:latin>만 설정 — 한글은 <a:ea>(동아시아) 슬롯을 쓰므로
        # 같이 지정하지 않으면 렌더러가 기본 고딕으로 폴백한다.
        rPr = run._r.get_or_add_rPr()
        for tag in ("a:ea", "a:cs"):
            e = rPr.find(qn(tag))
            if e is None:
                e = rPr.makeelement(qn(tag), {})
                rPr.append(e)
            e.set("typeface", fname)
        run.font.bold = fbold or (r.get("weight") == "bold")
        run.font.size = Pt(r.get("size", default_style["size"]))
        run.font.color.rgb = _c(r.get("color", default_style["color"]))
        if r.get("underline"):
            run.font.underline = True
        # 이탤릭 — 학명·in vitro 같은 라틴 표기. 한글 폰트엔 이탤릭 페이스가 없으니
        # 보통 font 를 라틴 세리프로 같이 지정한다.
        if r.get("italic"):
            run.font.italic = True
        # baseline(%) — 작게 줄인 인라인 기호(▶ 같은 진행 화살표)를 글줄 중앙까지 올린다.
        # python-pptx 미노출이라 rPr/@baseline(1/1000 %)을 직접 쓴다.
        if r.get("baseline"):
            rPr.set("baseline", str(int(round(r["baseline"] * 1000))))
        # 자간(letter-spacing) — python-pptx 미노출이라 rPr/@spc(1/100 pt)를 직접 쓴다.
        # 매거진 마스트헤드처럼 "가늘게 + 넓게" 앉히는 조판에 필수.
        track = r.get("tracking", default_style.get("tracking"))
        if track:
            rPr.set("spc", str(int(round(track * 100))))
        # 형광펜 — 성과 수치 한 줄만 짚을 때. 포트폴리오 문법에서 가져왔다.
        # python-pptx 미노출이라 rPr/a:highlight 를 직접 쓴다.
        mark = r.get("mark")
        if mark:
            hl = rPr.makeelement(qn("a:highlight"), {})
            clr = hl.makeelement(qn("a:srgbClr"), {"val": mark})
            hl.append(clr)
            rPr.append(hl)


def _textbox(slide, x, y, w, h=None):
    tb = slide.shapes.add_textbox(_x(x), _y(y), _x(w), _y(h if h else 0.06))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tb, tf


def _norm_runs(item, style):
    if "runs" in item:
        return [item["runs"]] if isinstance(item["runs"][0], dict) else item["runs"]
    lines = item.get("lines") or [item["text"]]
    color = item.get("color", style["color"])
    return [[{"t": ln, "color": color}] for ln in lines]


def _text_shadow(shp, spec):
    """글자 그림자 — 채움 없는 텍스트박스의 도형 그림자는 글자 윤곽을 따라 떨어진다.

    run 단위 `a:ln`(외곽선)·`a:effectLst` 는 LibreOffice 가 무시해 PDF·미리보기에서
    사라진다(2026-10-06 실측). 도형 단위 outerShdw 는 PowerPoint·LibreOffice 둘 다 그린다.
    LibreOffice 는 첫 그림자 한 겹만 그리니 2중 그림자는 아이템을 두 장 겹쳐 만든다.
    spec: {"color":"HEX","dist":pt,"dir":deg(기본 45=오른쪽 아래),"blur":pt,"alpha":0~1}
    """
    from lxml import etree
    spPr = shp._element.spPr
    el = spPr.find(qn("a:effectLst"))
    if el is None:
        el = etree.SubElement(spPr, qn("a:effectLst"))
    o = etree.SubElement(el, qn("a:outerShdw"), {
        "blurRad": str(int(spec.get("blur", 0) * 12700)),
        "dist": str(int(spec.get("dist", 4) * 12700)),
        "dir": str(int(spec.get("dir", 45) * 60000)),
        "algn": "tl", "rotWithShape": "0"})
    c = etree.SubElement(o, qn("a:srgbClr"), {"val": spec.get("color", "000000")})
    etree.SubElement(c, qn("a:alpha"), {"val": str(int(spec.get("alpha", 1.0) * 100000))})


def _add_text_item(slide, item, style_key):
    style = dict(_typo(style_key))
    if "size" in item:
        style["size"] = item["size"]
    if "color" in item:
        style["color"] = item["color"]
    if "tracking" in item:
        style["tracking"] = item["tracking"]
    line_groups = _norm_runs(item, style)
    tb, tf = _textbox(slide, item["x"], item["y"], item["w"], item.get("h"))
    first = True
    for runs in line_groups:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        _set_runs(para, runs, style)
        # num_xl(빅넘버)은 1.0 — 기본 1.45는 글자를 박스 아래로 밀어 밑줄/헤어라인과 겹친다.
        _def_sp = {"headline": 1.18, "divider_h": 1.18, "num_xl": 1.0}.get(style_key, 1.45)
        para.line_spacing = item.get("spacing", _def_sp)
        align = item.get("align", "left")
        para.alignment = _ALIGN[align]
    if item.get("rot"):
        tb.rotation = item["rot"]
    if item.get("shadow"):
        _text_shadow(tb, item["shadow"])
    if item.get("anchor"):
        tb.text_frame.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                                         "bottom": MSO_ANCHOR.BOTTOM}[item["anchor"]]
        # python-pptx 텍스트박스 기본값은 spAutoFit(도형을 글에 맞춰 줄임)이다.
        # 이게 켜져 있으면 h로 준 박스가 글 크기로 쪼그라들어 anchor:middle이
        # 아무 일도 안 한다(세로중앙 지시가 조용히 무시됨 — 채움 면 글자가 위로
        # 쏠리는 반복 반려의 실제 원인). h를 명시한 anchor 박스는 고정 크기로 둔다.
        if item.get("h"):
            tb.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    return tb


def _add_poly(slide, pts, fill, outline=None, outline_w=1.0):
    """프리폼 다각형 채우기 — 아이소메트릭 타일 등. pts = [(x_frac, y_frac), ...]."""
    epts = [(int(px * SLIDE_W), int(py * SLIDE_H)) for (px, py) in pts]
    fb = slide.shapes.build_freeform(epts[0][0], epts[0][1], scale=1.0)
    fb.add_line_segments(epts[1:], close=True)
    shp = fb.convert_to_shape()
    if fill == "none":
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _c(fill)
    if outline:
        shp.line.color.rgb = _c(outline)
        shp.line.width = Pt(outline_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    _strip_style(shp)
    return shp


def _pill_w(text: str, size: float) -> float:
    """필 자동 폭 — 한글/CJK는 전각, 라틴/숫자는 반각으로 추정."""
    f = 0.0
    for ch in text:
        wide = ord(ch) >= 0x1100 or ch in "·—↔→←“”"
        f += 0.0105 if wide else 0.0058
    return 0.024 + f * size / 10.5


def _strip_style(shp):
    """add_shape가 붙이는 <p:style>(테마 effectRef 그림자 참조) 제거.

    빈 <a:effectLst/>만으로는 LibreOffice가 테마 그림자를 무시하지 않는다.
    """
    el = shp._element
    st = el.find(qn("p:style"))
    if st is not None:
        el.remove(st)
    return shp


def _radius_from(item, default):
    """radius_px(절대 px, 1920×1080 기준)가 있으면 도형 크기로 역산해 adj 로 환산.

    roundRect 의 adjustments[0] 은 짧은 변 대비 비율이라 같은 값이라도 도형이
    크면 모서리가 커진다 — 덱 전체 실효 라운드를 한 px 로 통일하려면 절대값이
    필요하다 (fb021 "라운드 4px로 통일"). 없으면 기존 radius 비율 그대로.
    """
    if item.get("radius_px") is not None:
        mn = min(item.get("w", 0) * 1920.0, item.get("h", 0) * 1080.0)
        if mn > 0:
            return max(0.0, min(0.5, float(item["radius_px"]) / mn))
    return item.get("radius", default)


def _rounded(slide, x, y, w, h, fill, radius=0.5, outline=None,
             outline_w=1.2, outline_dash=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _x(x), _y(y), _x(w), _y(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    if fill == "none":
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _c(fill)
    if outline:
        shp.line.color.rgb = _c(outline)
        shp.line.width = Pt(outline_w)
        if outline_dash:  # 빈 슬롯(자리) 선언용 파선 테두리
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return _strip_style(shp)


def _hairline_shape(slide, x, y, w, color=None, dash=None, thick=None,
                    dash_seg=None, dash_gap=None):
    h = Emu(int(9525 * (thick or 1)))
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _x(x), _y(y), _x(w), h)
    ln.fill.solid()
    ln.fill.fore_color.rgb = _c(color or _ACTIVE["hair"])
    ln.line.fill.background()
    ln.shadow.inherit = False
    if dash:  # 점선: 촘촘한 사각 세그먼트로 재현 (렌더러 독립적)
        ln.fill.background()
        seg = dash_seg or 0.006
        gap = dash_gap if dash_gap is not None else seg
        cx = x
        while cx < x + w:
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, _x(cx), _y(y),
                                       _x(min(seg, x + w - cx)), h)
            s.fill.solid()
            s.fill.fore_color.rgb = _c(color or _ACTIVE["hair"])
            s.line.fill.background()
            s.shadow.inherit = False
            _strip_style(s)
            cx += seg + gap
    return _strip_style(ln)


# 범용 도형 프리미티브 — 인포그래픽 조립용 (관공서 장표 문법: 육각형·사선 밴드·파이·링)
_SHAPES = {
    "hexagon": MSO_SHAPE.HEXAGON,
    "parallelogram": MSO_SHAPE.PARALLELOGRAM,
    "pie": MSO_SHAPE.PIE,
    "arc_ring": MSO_SHAPE.BLOCK_ARC,
    "chevron": MSO_SHAPE.CHEVRON,
    "pentagon_arrow": MSO_SHAPE.PENTAGON,
    "oval": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
    "rounded": MSO_SHAPE.ROUNDED_RECTANGLE,
    # 도형 변형 모핑·강조 반복 재료(유튜브 기법 shape-morph·emphasis-loop, 2026-10-08)
    "rect": MSO_SHAPE.RECTANGLE,
    "star5": MSO_SHAPE.STAR_5_POINT,
    "star16": MSO_SHAPE.STAR_16_POINT,
    "heart": MSO_SHAPE.HEART,
    "gear": MSO_SHAPE.GEAR_9,
    "cloud": MSO_SHAPE.CLOUD,
    "donut": MSO_SHAPE.DONUT,
    "teardrop": MSO_SHAPE.TEAR,
    "moon": MSO_SHAPE.MOON,
    "lightning": MSO_SHAPE.LIGHTNING_BOLT,
}


def _add_shape_item(slide, item, T, variant="paper"):
    return _with_depth(lambda top: _add_shape_base(slide, item, T, variant, top), item)


def _add_shape_base(slide, item, T, variant="paper", with_text=True):
    shp = slide.shapes.add_shape(_SHAPES[item["shape"]], _x(item["x"]), _y(item["y"]),
                                 _x(item["w"]), _y(item["h"]))
    for i, v in enumerate(item.get("adj", [])):
        try:
            shp.adjustments[i] = v
        except Exception:
            pass
    fill = depth_fill(item, T, variant) or item.get("fill", T["accent"])
    if fill == "none":
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _c(fill)
    if item.get("outline"):
        shp.line.color.rgb = _c(item["outline"])
        shp.line.width = Pt(item.get("outline_w", 1.2))
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    if item.get("rot"):
        shp.rotation = item["rot"]
    _strip_style(shp)
    if item.get("text") and with_text:
        tf = shp.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(20000)
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        style = dict(_typo(item.get("style", "chip")))
        style["color"] = item.get("text_color", style["color"])
        style["font"] = item.get("font", style["font"])
        if "size" in item:
            style["size"] = item["size"]
        _set_runs(tf.paragraphs[0], [{"t": item["text"]}], style)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    return shp


def _apply_grad(shp, top_hex, bottom_hex, ang_deg=90.0, scaled=True,
                mids=()):
    """카드 그라데이션. 기본은 세로 2-stop(위→아래, 90°).

    `ang_deg` 는 DrawingML 각도(시계방향, 0°=왼→오른쪽, 90°=위→아래).
    `mids` 는 [[0~100, "RRGGBB"], …] 중간 스톱 — 실측 그라데이션이 한쪽에서
    빨리 꺾일 때 2-stop 으로는 안 맞는다(표지 풀블리드).
    표지 풀블리드처럼 대각 그라데이션이 필요할 때만 준다.
    """
    from pptx.oxml.ns import nsmap
    import copy
    spPr = shp._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        e = spPr.find(qn(tag))
        if e is not None:
            spPr.remove(e)
    from lxml import etree
    grad = etree.SubElement(spPr, qn("a:gradFill"))
    lst = etree.SubElement(grad, qn("a:gsLst"))
    stops = [(0, top_hex)]
    stops += [(int(round(float(p) * 1000)), h) for p, h in (mids or ())]
    stops.append((100000, bottom_hex))
    for pos, hx in sorted(stops):
        gs = etree.SubElement(lst, qn("a:gs"))
        gs.set("pos", str(pos))
        clr = etree.SubElement(gs, qn("a:srgbClr"))
        clr.set("val", hx)
    lin = etree.SubElement(grad, qn("a:lin"))
    lin.set("ang", str(int(round(float(ang_deg) % 360 * 60000))))  # 90° = 위→아래
    lin.set("scaled", "1" if scaled else "0")
    # ln(테두리)보다 앞에 와야 하므로 위치 재조정
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        spPr.remove(grad)
        ln.addprevious(grad)


def _apply_shadow(shp, alpha_pct=22):
    """소프트 드롭섀도우 (평가자료 모드 카드)."""
    from lxml import etree
    spPr = shp._element.spPr
    old = spPr.find(qn("a:effectLst"))
    if old is not None:
        spPr.remove(old)
    eff = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"))
    sh.set("blurRad", "90000")
    sh.set("dist", "28000")
    sh.set("dir", "5400000")
    sh.set("rotWithShape", "0")
    clr = etree.SubElement(sh, qn("a:srgbClr"))
    clr.set("val", "1F3864")
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(alpha_pct * 1000))


# ── 입체 표면 효과 (뉴모피즘 · 글래스모피즘 · 앰비언트 섀도우) ────────────────
# 수치는 원전 영상 실측 → DrawingML 환산. 근거·환산식은 docs/design-refs-depth-effects.md.
#   투명도 T%  → alpha = (100-T)*1000      흐리게 Xpt → blurRad = X*12700
#   간격 Xpt   → dist  = X*12700           각도 X°    → dir     = X*60000

def _pt(v: float) -> str:
    return str(int(round(v * 12700)))


def _deg(v: float) -> str:
    return str(int(round(v * 60000)))


def _op(transparency_pct: float) -> int:
    """PPT '투명도' → DrawingML alpha."""
    return int(round((100 - transparency_pct) * 1000))


# 하이라이트(좌상단 흰빛) / 그늘(우하단 검정) — 돌출 뉴모피즘은 이 둘을 겹친다.
_NEU_HI = {"kind": "outer", "color": "FFFFFF", "alpha": _op(22),
           "blur": _pt(12), "dist": _pt(10), "dir": _deg(225)}
_NEU_LO = {"kind": "outer", "color": "000000", "alpha": _op(88),
           "blur": _pt(15), "dist": _pt(10), "dir": _deg(45)}

DEPTH_PRESETS: dict[str, dict] = {
    # 돌출 — 배경과 같은 색 면이 위로 솟은 느낌. 겹 2장.
    "neu_raised": {"layers": [_NEU_HI, _NEU_LO], "fill_role": "bg"},
    # 함몰 — 눌린 홈. 안쪽 그림자 1장 + 헤어라인 링.
    # 링은 장식이 아니라 폴백이다: LibreOffice(PDF·썸네일 경로)는 a:innerShdw를
    # 렌더하지 않아 홈이 통째로 사라진다. 자세한 근거는 docs/PITFALLS.md.
    "neu_pressed": {"layers": [{"kind": "inner", "color": "000000", "alpha": _op(87),
                                "blur": _pt(11), "dist": _pt(7), "dir": _deg(220)}],
                    "fill_role": "bg", "ring": "hair"},
    # ── 부양 3종은 전부 2겹이다 ──────────────
    # 단겹·방향 없음(dist 0)은 카드 사방에 균일한 띠를 만들고 꼬리가 뚝 끊겨
    # 그림자가 아니라 '테두리'로 보인다. 실측: 단겹 10pt는 꼬리가 19px에서 끊김.
    # 자연스러운 부양 = 넓고 옅은 앰비언트 겹 + 좁고 진한 키 겹, 둘 다 아래(90°)로.
    # 판정 기준: 진하기 25~30 · 꼬리 55px+ · 인접 픽셀 최대단차 4 이하.
    # 부양 3종은 '은은한 리프트'다 — 기본이 아니라 옵트인. 일반 카드는 그림자 없이
    # fill 틴트로 나누는 게 기본이고(CLAUDE.md 규칙 14), 이건 정말 살짝 띄우고 싶을 때만.
    # 2026-07-26 사용자 피드백으로 진하기를 대폭 낮췄다: 검정 그림자가 카드에 눈에 띄게
    # 깔리면 촌스럽다. 2겹 구조(넓은 앰비언트+좁은 키)는 유지해 '테두리 띠'는 피하되,
    # 불투명도·거리를 줄여 있는 듯 없는 듯한 그늘로.
    "soft": {"layers": [                                   # 은은한 본문 카드 리프트
        {"kind": "outer", "color": "000000", "alpha": _op(97),
         "blur": _pt(26), "dist": _pt(5), "dir": _deg(90)},
        {"kind": "outer", "color": "000000", "alpha": _op(95),
         "blur": _pt(8), "dist": _pt(1), "dir": _deg(90)}]},
    # 말풍선·플로팅 칩 — 작은 요소라 더 얕게.
    "float": {"layers": [
        {"kind": "outer", "color": "000000", "alpha": _op(97),
         "blur": _pt(15), "dist": _pt(3), "dir": _deg(90)},
        {"kind": "outer", "color": "000000", "alpha": _op(95),
         "blur": _pt(6), "dist": _pt(1), "dir": _deg(90)}]},
    # 대형 패널 — 조금 더 넓게, 그래도 옅게.
    "ambient": {"layers": [
        {"kind": "outer", "color": "000000", "alpha": _op(96),
         "blur": _pt(38), "dist": _pt(8), "dir": _deg(90)},
        {"kind": "outer", "color": "000000", "alpha": _op(94),
         "blur": _pt(11), "dist": _pt(2), "dir": _deg(90)}]},
    # 하드 오프셋 — 스티커·컷아웃. 거의 안 번진다.
    "sticker": {"layers": [{"kind": "outer", "color": "000000", "alpha": _op(20),
                            "blur": _pt(1), "dist": _pt(2), "dir": _deg(45)}]},
    # 컬러 글로우 — 검정이 아니라 악센트색으로 번진다.
    "neon": {"layers": [{"kind": "outer", "color": "@accent", "alpha": _op(70),
                         "blur": _pt(15), "dist": "0", "dir": "0"}]},
    # 눌린 종이 — 방향 없는 안쪽 그림자. 표·목록 판을 살짝 가라앉힌다.
    "paper": {"layers": [{"kind": "inner", "color": "000000", "alpha": _op(80),
                          "blur": _pt(9), "dist": "0", "dir": "0"}]},
    # 유리 — 반투명 흰 면 + 헤어라인 + 안쪽 흰 그림자(가장자리 광택).
    "glass": {"layers": [{"kind": "inner", "color": "FFFFFF", "alpha": _op(45),
                          "blur": _pt(20), "dist": "0", "dir": "0"}],
              "glass": True},
}


def _resolve_depth_color(spec_color: str, item) -> str:
    """'@accent'는 테마 악센트로 해석. item.depth_color가 있으면 그게 이긴다."""
    if item.get("depth_color"):
        return item["depth_color"]
    if spec_color.startswith("@"):
        return _ACTIVE[spec_color[1:]]
    return spec_color


def _apply_effects(shp, layers, item=None):
    """effectLst 교체. DrawingML은 outerShdw/innerShdw를 각 1개만 허용하므로
    2겹이 필요한 프리셋(neu_raised)은 호출부에서 도형을 겹쳐 쌓는다."""
    from lxml import etree
    spPr = shp._element.spPr
    old = spPr.find(qn("a:effectLst"))
    if old is not None:
        spPr.remove(old)
    eff = etree.SubElement(spPr, qn("a:effectLst"))
    for spec in layers:
        tag = "a:outerShdw" if spec["kind"] == "outer" else "a:innerShdw"
        node = etree.SubElement(eff, qn(tag))
        node.set("blurRad", spec["blur"])
        node.set("dist", spec["dist"])
        node.set("dir", spec["dir"])
        if spec["kind"] == "outer":
            node.set("rotWithShape", "0")
        clr = etree.SubElement(node, qn("a:srgbClr"))
        clr.set("val", _resolve_depth_color(spec["color"], item or {}).upper())
        a = etree.SubElement(clr, qn("a:alpha"))
        a.set("val", str(spec["alpha"]))
    return shp


def _fill_alpha(shp, hex_color: str, alpha: float):
    """반투명 단색 채우기 (a:solidFill/srgbClr/alpha) — 유리 면."""
    from lxml import etree
    spPr = shp._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill"):
        e = spPr.find(qn(tag))
        if e is not None:
            spPr.remove(e)
    fill = etree.Element(qn("a:solidFill"))
    clr = etree.SubElement(fill, qn("a:srgbClr"))
    clr.set("val", hex_color.upper())
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(int(round(alpha * 100000))))
    _spPr_insert_fill(spPr, fill)


def _use_bg_fill(shp):
    """<p:sp useBgFill="1"> + noFill — 도형이 슬라이드 배경을 그대로 비춘다.

    글래스모피즘 도형의 핵심. 배경에 흐린 사진을 깔아두면
    도형이 이동해도 그 자리의 배경을 계속 따라간다.
    """
    from lxml import etree
    sp = shp._element
    sp.set("useBgFill", "1")
    spPr = sp.spPr
    for tag in ("a:solidFill", "a:gradFill", "a:blipFill", "a:noFill"):
        e = spPr.find(qn(tag))
        if e is not None:
            spPr.remove(e)
    _spPr_insert_fill(spPr, etree.Element(qn("a:noFill")))


def _spPr_insert_fill(spPr, fill_el):
    """fill은 a:ln 앞에 와야 스키마상 유효하다."""
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        ln.addprevious(fill_el)
    else:
        eff = spPr.find(qn("a:effectLst"))
        if eff is not None:
            eff.addprevious(fill_el)
        else:
            spPr.append(fill_el)


def _norm_layer(lay: dict) -> dict:
    """PPT 그림자 대화상자 단위(투명도 %·흐리게 pt·간격 pt·각도 °)를 DrawingML로.

    이미 환산된 값(문자열)이 오면 그대로 둔다 — DEPTH_PRESETS가 그 형태다.
    """
    if isinstance(lay.get("blur"), str):
        return lay
    return {"kind": lay.get("kind", "outer"),
            "color": lay.get("color", "000000"),
            "alpha": lay.get("alpha", _op(lay.get("transparency", 85))),
            "blur": _pt(lay.get("blur", 10)),
            "dist": _pt(lay.get("dist", 0)),
            "dir": _deg(lay.get("dir", 0))}


def _depth_of(item):
    """이름(프리셋) 또는 dict(인라인 기술) → 프리셋 dict. 없으면 None."""
    d = item.get("depth")
    if d is None:
        return None
    if isinstance(d, dict):
        if not d.get("layers"):
            raise KeyError("인라인 depth에는 layers가 있어야 한다")
        return {**d, "layers": [_norm_layer(l) for l in d["layers"]]}
    if d not in DEPTH_PRESETS:
        raise KeyError(f"depth '{d}' 없음 — {sorted(DEPTH_PRESETS)} 중 하나")
    return DEPTH_PRESETS[d]


def depth_fill(item, T, variant="paper"):
    """depth 프리셋이 요구하는 기본 면색 — 뉴모피즘은 배경색과 같아야 성립한다."""
    if "fill" in item:
        return item["fill"]
    d = item.get("depth")
    preset = d if isinstance(d, dict) else DEPTH_PRESETS.get(d or "", {})
    if preset.get("fill_role") == "bg":
        return T["dark_bg"] if variant == "dark" else T["bg"]
    return None


def _with_depth(draw, item):
    """draw(is_top) -> shape 콜백을 프리셋이 요구하는 겹 수만큼 호출하고 위 겹을 반환.

    아래 겹에는 텍스트를 넣지 않는다(글자가 두 번 겹쳐 두꺼워진다) —
    콜백은 is_top=False일 때 글을 쓰지 않아야 한다.
    """
    preset = _depth_of(item)
    if not preset:
        shp = draw(True)
        if item.get("shadow"):  # 하위 호환: shadow: true
            _apply_shadow(shp, item.get("shadow_alpha", 22))
        return shp
    layers = preset["layers"]
    for lay in layers[:-1]:  # 아래 겹 (뉴모피즘 하이라이트)
        _apply_effects(draw(False), [lay], item)
    top = draw(True)
    if preset.get("ring") and not item.get("outline"):
        top.line.color.rgb = _c(_ACTIVE[preset["ring"]])
        top.line.width = Pt(item.get("outline_w", 1.0))
    if preset.get("glass"):
        if item.get("bg_fill"):
            _use_bg_fill(top)
        else:
            _fill_alpha(top, item.get("fill", "FFFFFF"), item.get("fill_alpha", 0.18))
        top.line.color.rgb = _c(item.get("outline", "FFFFFF"))
        top.line.width = Pt(item.get("outline_w", 0.25))
    _apply_effects(top, [layers[-1]], item)
    return top


def _blurred_bg_png(assets_dir: Path, src: Path, blur: float, darken: float) -> Path:
    """배경용 흐림·감광 사본 (결정적 캐시). 유리 도형이 이걸 비춘다.

    blur는 **1344px 기준 픽셀**로 해석하고 원본 폭에 비례 환산한다. 안 그러면
    같은 값이 4K 소스에서는 거의 안 흐려지고 저해상 소스에서는 뭉개진다.
    """
    from PIL import ImageEnhance, ImageFilter
    key = f"bgblur_{src.stem}_{blur:.1f}_{darken:.2f}.png"
    out = assets_dir / key
    # 원본이 캐시보다 새로우면 다시 굽는다 — 이름만 보고 재사용하면 같은 이름으로 바뀐 배경이 옛 그림으로 나간다
    # (배경이 앞 빌드의 개체 박힌 그림으로 남으면 팬 전환이 교차 페이드로 보인다)
    if out.exists() and out.stat().st_mtime >= src.stat().st_mtime:
        return out
    im = Image.open(src).convert("RGB")
    if blur:
        im = im.filter(ImageFilter.GaussianBlur(blur * im.width / 1344))
    if darken:
        im = ImageEnhance.Brightness(im).enhance(1.0 - darken)
    assets_dir.mkdir(parents=True, exist_ok=True)
    im.save(out)
    return out


def _set_slide_background(slide, png: Path):
    """슬라이드 배경 자체를 그림으로 채운다 (<p:bg><p:bgPr><a:blipFill>)."""
    from lxml import etree
    cSld = slide._element.cSld
    old = cSld.find(qn("p:bg"))
    if old is not None:
        cSld.remove(old)
    _, rId = slide.part.get_or_add_image_part(str(png))
    bg = etree.Element(qn("p:bg"))
    bgPr = etree.SubElement(bg, qn("p:bgPr"))
    blip_fill = etree.SubElement(bgPr, qn("a:blipFill"))
    blip = etree.SubElement(blip_fill, qn("a:blip"))
    blip.set(qn("r:embed"), rId)
    etree.SubElement(etree.SubElement(blip_fill, qn("a:stretch")), qn("a:fillRect"))
    etree.SubElement(bgPr, qn("a:effectLst"))
    cSld.insert(0, bg)


def _glow_png(assets_dir: Path, theme_name: str, cx: float, cy: float,
              sigma: float, intensity: float, alpha: bool = False) -> Path:
    """다크 디바이더용 블룸 배경 PNG (테마 팔레트, 결정적 자체 생성).

    alpha=True — 베이스를 칠하지 않는 투명 RGBA. 글로우 아래에 풀블리드 grad
    카드가 있을 때 쓴다. 불투명 RGB(기본)는 테마 dark_bg 를 베이크해서
    아래 지면을 통째로 덮는다.
    """
    w = 2560
    h = int(round(w * SLIDE_H / SLIDE_W))
    tag = "a_" if alpha else ""
    key = f"glow_{tag}{theme_name}_{w}x{h}_{cx:.2f}_{cy:.2f}_{sigma:.2f}_{intensity:.2f}.png"
    out = assets_dir / key
    if out.exists():
        return out
    g = _ACTIVE["glow"]
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float64)
    d2 = ((xx - cx * w) ** 2 + (yy - cy * h) ** 2) / (sigma * w) ** 2
    core = np.exp(-d2 * 2.2)
    halo = np.exp(-d2 * 0.55)
    base = np.array(g["base"], dtype=np.float64)
    corec = np.array(g["core"], dtype=np.float64)
    deep = np.array(g["deep"], dtype=np.float64)
    add = deep[None, None, :] * (halo[..., None] * 0.85 * intensity) \
        + corec[None, None, :] * (core[..., None] * 0.9 * intensity)
    assets_dir.mkdir(parents=True, exist_ok=True)
    if alpha:
        # bg*(1-a) + c*a == bg + add 가 되도록 — 같은 dark_bg 위에선 기존과 동일,
        # grad 카드 위에선 카드가 비쳐 보인다.
        m = add.max(axis=2, keepdims=True)          # 채널 최대 = 알파 근거
        a = np.clip(m / 255.0, 0.0, 1.0)
        with np.errstate(divide="ignore", invalid="ignore"):
            c = np.where(m > 0.5, base[None, None, :] + add * (255.0 / np.maximum(m, 0.5)),
                         0.0)
        rgba = np.concatenate([np.clip(c, 0, 255), a * 255.0], axis=2)
        Image.fromarray(rgba.astype(np.uint8), "RGBA").save(out)
        return out
    img = base[None, None, :] + add
    rng = np.random.default_rng(0)
    img += rng.standard_normal((h, w, 1)) * 1.1
    Image.fromarray(np.clip(img, 0, 255).astype(np.uint8)).save(out)
    return out


# ── 기기 목업 ────────────────────────────────────────────────────────────────
# 프레임 PNG(구매 템플릿에서 들여옴) + 화면 마스크로 합성한다.
# 설치: `python3 tools/install_mockups.py` — 없으면 친절한 에러로 알려준다.
_MOCKUP_DIR = Path(__file__).parent / "mockups"
_MOCKUP_MANIFEST: dict | None = None


def mockup_devices() -> dict:
    global _MOCKUP_MANIFEST
    if _MOCKUP_MANIFEST is None:
        mf = _MOCKUP_DIR / "manifest.json"
        _MOCKUP_MANIFEST = json.loads(mf.read_text(encoding="utf-8")) if mf.exists() else {}
    return _MOCKUP_MANIFEST


def _screen_image(spec: dict, size: tuple[int, int], base_dir: Path,
                  fit: str) -> Image.Image:
    """목업 화면에 들어갈 그림 — 사진(image) 또는 단색/그라데이션 면."""
    sw, sh = size
    if spec.get("image"):
        src = Path(spec["image"])
        if not src.is_absolute():
            src = base_dir / spec["image"]
        im = Image.open(src).convert("RGB")
        if fit == "contain":                       # 레터박스 — 잘리면 안 되는 화면
            pad = spec.get("pad_color", "FFFFFF")
            canvas = Image.new("RGB", (sw, sh), tuple(
                int(pad[i:i + 2], 16) for i in (0, 2, 4)))
            im2 = im.copy()
            im2.thumbnail((sw, sh), Image.LANCZOS)
            canvas.paste(im2, ((sw - im2.width) // 2, (sh - im2.height) // 2))
            return canvas
        r = max(sw / im.width, sh / im.height)     # 센터크롭 (기본)
        im = im.resize((max(int(im.width * r), sw), max(int(im.height * r), sh)),
                       Image.LANCZOS)
        return im.crop(((im.width - sw) // 2, (im.height - sh) // 2,
                        (im.width - sw) // 2 + sw, (im.height - sh) // 2 + sh))
    top = spec.get("grad", [spec.get("fill", "F4F6F9")] * 2)
    if isinstance(top, str):
        top = [top, top]
    c0 = np.array([int(top[0][i:i + 2], 16) for i in (0, 2, 4)], dtype=float)
    c1 = np.array([int(top[1][i:i + 2], 16) for i in (0, 2, 4)], dtype=float)
    t = np.linspace(0, 1, sh)[:, None, None]
    arr = c0[None, None, :] * (1 - t) + c1[None, None, :] * t
    return Image.fromarray(np.repeat(np.clip(arr, 0, 255).astype(np.uint8), sw, axis=1))


def _mockup_png(assets_dir: Path, base_dir: Path, device: str, screen: dict,
                fit: str, px: int) -> tuple[Path, float]:
    """프레임 + 화면을 합성한 PNG 경로와 프레임 종횡비(w/h)를 돌려준다."""
    man = mockup_devices()
    if device not in man:
        raise ValueError(
            f"목업 '{device}' 없음 (설치된 것: {sorted(man) or '없음'}). "
            f"`python3 tools/install_mockups.py` 로 설치할 것")
    meta = man[device]
    fw, fh = meta["size"]
    ar = fw / fh
    key = json.dumps([device, screen, fit, px], sort_keys=True, ensure_ascii=False)
    out = assets_dir / f"mock_{device}_{px}_{abs(hash(key)) % (10 ** 10)}.png"
    if out.exists():
        return out, ar
    up = max(1, round(px / fw))
    frame = Image.open(_MOCKUP_DIR / f"{device}.png").convert("RGBA")
    mask = Image.open(_MOCKUP_DIR / f"{device}_mask.png").convert("L")
    if up > 1:
        frame = frame.resize((fw * up, fh * up), Image.LANCZOS)
        mask = mask.resize((fw * up, fh * up), Image.LANCZOS)
    x0, y0, x1, y1 = [v * up for v in meta["screen"]]
    shot = _screen_image(screen, (x1 - x0, y1 - y0), base_dir, fit)
    canvas = Image.new("RGBA", frame.size, (0, 0, 0, 0))
    canvas.paste(shot.convert("RGBA"), (x0, y0))
    base = Image.new("RGBA", frame.size, (0, 0, 0, 0))
    base.paste(canvas, (0, 0), mask)
    base.alpha_composite(frame)
    assets_dir.mkdir(parents=True, exist_ok=True)
    base.save(out)
    return out, ar


# ── 지도 ─────────────────────────────────────────────────────────────────────
# RedSlide 배포 지도의 프리폼 그룹을 그대로 이식한다 — 이미지가 아니라 도형이라
# 지역별 색칠(단계구분도)이 되고 받는 사람이 편집할 수도 있다.
# 설치: `python3 tools/install_maps.py`
_MAP_DIR = Path(__file__).parent / "maps"
_MAP_MANIFEST: dict | None = None


def map_regions() -> dict:
    global _MAP_MANIFEST
    if _MAP_MANIFEST is None:
        mf = _MAP_DIR / "manifest.json"
        _MAP_MANIFEST = json.loads(mf.read_text(encoding="utf-8")) if mf.exists() else {}
    return _MAP_MANIFEST


def _add_map(slide, item, T, variant):
    """지도 그룹을 이식하고 지역별로 칠한다.

    values: {"서울": 0.9, ...} 0~1 강도 → 기본면↔악센트 보간
    highlight: ["서울", 3, ...] 이름 또는 인덱스 → 악센트 단색
    """
    from pptx.oxml import parse_xml            # 커스텀 엘리먼트 클래스로 파싱해야
    from pptx.shapes.group import GroupShape   # left/width 같은 속성이 붙는다

    name = item.get("region", "korea")
    man = map_regions()
    if name not in man:
        raise ValueError(
            f"지도 '{name}' 없음 (설치된 것: {sorted(man) or '없음'}). "
            f"`python3 tools/install_maps.py` 로 설치할 것")
    el = parse_xml((_MAP_DIR / f"{name}.xml").read_bytes())
    # 이식 그룹의 cNvPr id 는 원본 pptx 의 번호(2부터)를 그대로 들고 온다 —
    # 슬라이드에 이미 있는 도형 id 와 **충돌**하면 PowerPoint 가 "복구" 프롬프트를
    # 띄우고 복구 시 도형을 통째로 버린다(실기기 재현,
    # slide47 dup 1건·slide71 dup 20건). 붙이기 전에 전부 재번호한다.
    _cnvpr = qn("p:cNvPr")
    used = {int(c.get("id")) for c in slide.shapes._spTree.iter(_cnvpr)
            if str(c.get("id", "")).isdigit()}
    nid = max(used, default=1) + 1
    for c in el.iter(_cnvpr):
        c.set("id", str(nid))
        nid += 1
    slide.shapes._spTree.append(el)
    g = GroupShape(el, slide.shapes)
    w = item["w"]
    h = item.get("h") or w * (SLIDE_W / SLIDE_H) / man[name]["aspect"]
    g.left, g.top = _x(item["x"]), _y(item["y"])
    g.width, g.height = _x(w), _y(h)

    regions = man[name].get("regions") or []
    base = item.get("base", T["card2"] if variant != "dark" else _mix_hex(
        T["dark_bg"], T["dark_fg"], 0.16))
    accent = item.get("accent", T["accent"])
    line = item.get("outline", T["bg"] if variant != "dark" else T["dark_bg"])
    values = item.get("values") or {}
    hi = item.get("highlight") or []
    hi_idx = {regions.index(k) if isinstance(k, str) and k in regions else k
              for k in hi}
    for i, sh in enumerate(g.shapes):
        key = regions[i] if i < len(regions) else None
        if i in hi_idx or (key is not None and key in hi_idx):
            fill = accent
        elif key is not None and key in values:
            fill = _mix_hex(base, accent, max(0.0, min(1.0, float(values[key]))))
        elif str(i) in values:
            fill = _mix_hex(base, accent,
                            max(0.0, min(1.0, float(values[str(i)]))))
        else:
            fill = base
        try:
            sh.fill.solid()
            sh.fill.fore_color.rgb = _c(fill)
            sh.line.color.rgb = _c(line)
            sh.line.width = Pt(item.get("outline_w", 0.75))
            sh.shadow.inherit = False
        except Exception:
            pass
    return g


def _mix_hex(a: str, b: str, t: float) -> str:
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(ca, cb))


def _scrim_png(assets_dir: Path, edge: str, color: str, strength: float,
               softness: float) -> Path:
    """풀블리드 이미지용 알파 그라데이션 스크림 PNG (투명→color). 시네마틱 자막 하단부용.

    edge: bottom/top/left/right — 그 방향 바깥으로 갈수록 진해진다. 배치 rect에 맞춰 늘어난다.
    """
    key = f"scrim_{edge}_{color}_{strength:.2f}_{softness:.2f}.png"
    out = assets_dir / key
    if out.exists():
        return out
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    vertical = edge in ("bottom", "top")
    n = 1024
    ramp = np.linspace(0.0, 1.0, n)  # 0=안쪽(투명) → 1=바깥(진함)
    if edge in ("top", "left"):
        ramp = ramp[::-1].copy()
    alpha = np.clip(ramp, 0, 1) ** softness * strength
    a8 = (alpha * 255).astype(np.uint8)
    if vertical:
        arr = np.zeros((n, 8, 4), dtype=np.uint8)
        arr[..., 3] = a8[:, None]
    else:
        arr = np.zeros((8, n, 4), dtype=np.uint8)
        arr[..., 3] = a8[None, :]
    arr[..., 0], arr[..., 1], arr[..., 2] = r, g, b
    assets_dir.mkdir(parents=True, exist_ok=True)
    Image.fromarray(arr, "RGBA").save(out)
    return out


_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT,
          "justify": PP_ALIGN.JUSTIFY, "distribute": PP_ALIGN.DISTRIBUTE}

_CHART_KIND = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,        # 가로 막대
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,  # 세로 막대
    "line": XL_CHART_TYPE.LINE,
    "line_markers": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}
# "No Style, No Grid" — 기본 밴딩/테두리 없는 표 스타일
_TABLE_NO_GRID = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def _ea_cs(rpr, fname):
    """defRPr/rPr에 동아시아(a:ea)·복합문자(a:cs) typeface 주입 — 한글 폴백 방지.

    a:latin 뒤에 오도록 append (OOXML 시퀀스 순서 유지)."""
    if rpr is None:
        return
    for tag in ("a:ea", "a:cs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = rpr.makeelement(qn(tag), {})
            rpr.append(e)
        e.set("typeface", fname)


def _font_ea(font_obj, fname, size=None, color=None, bold=None):
    font_obj.name = fname
    if size is not None:
        font_obj.size = Pt(size)
    if color is not None:
        font_obj.color.rgb = _c(color)
    if bold is not None:
        font_obj.bold = bold
    _ea_cs(font_obj._rPr, fname)


def _marker_color(ser, hex_color: str, size: int = 6):
    """선 그래프 마커도 선과 같은 색으로 (기본 마커는 오피스 팔레트를 따른다)."""
    from lxml import etree
    sp = ser._element
    marker = sp.find(qn("c:marker"))
    if marker is None:
        marker = etree.SubElement(sp, qn("c:marker"))
        spPr = sp.find(qn("c:spPr"))
        if spPr is not None:
            spPr.addnext(marker)
    for tag, val in (("c:symbol", "circle"), ("c:size", str(size))):
        el = marker.find(qn(tag))
        if el is None:
            el = etree.SubElement(marker, qn(tag))
        el.set("val", val)
    old = marker.find(qn("c:spPr"))
    if old is not None:
        marker.remove(old)
    mspPr = etree.SubElement(marker, qn("c:spPr"))
    fill = etree.SubElement(mspPr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr")).set("val", hex_color.upper())
    ln = etree.SubElement(mspPr, qn("a:ln"))
    lfill = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(lfill, qn("a:srgbClr")).set("val", hex_color.upper())


def _add_chart(slide, item):
    """네이티브 PPT 차트(엑셀 백업 데이터 — 더블클릭 편집 가능)."""
    kind = item.get("kind", "column")
    xltype = _CHART_KIND.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    size = item.get("size", 10)
    fname = _ACTIVE["fonts"][item.get("font", "body")][0]
    nfmt = item.get("number_format", "General")

    data = CategoryChartData()
    data.categories = item["categories"]
    for s in item["series"]:
        data.add_series(s.get("name", "series"), tuple(s["values"]), number_format=nfmt)

    gf = slide.shapes.add_chart(xltype, _x(item["x"]), _y(item["y"]),
                                _x(item["w"]), _y(item["h"]), data)
    chart = gf.chart
    chart.has_title = False
    _font_ea(chart.font, fname, size=size, color=item.get("text_color", _ACTIVE["body_ink"]))

    is_pie = kind in ("pie", "doughnut")
    chart.has_legend = bool(item.get("legend", len(item["series"]) > 1 or is_pie))
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        _font_ea(chart.legend.font, fname, size=size, color=_ACTIVE["body_ink"])

    plot = chart.plots[0]
    if not is_pie and item.get("gap") is not None:
        plot.gap_width = item["gap"]

    if item.get("value_labels"):
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = item.get("label_format", nfmt)
        dl.number_format_is_linked = False
        _font_ea(dl.font, fname, size=item.get("label_size", size),
                 color=item.get("value_color"), bold=True)
        try:
            # 라인차트에 OUTSIDE_END(outEnd)를 주면 OOXML 규격 위반이라
            # PowerPoint 가 "복구"를 띄운다 (2026-09-02 실기기 이분탐색으로 확정).
            # line 계열은 ABOVE, pie 는 BEST_FIT, 그 외(bar/column)만 OUTSIDE_END.
            if is_pie:
                dl.position = XL_LABEL_POSITION.BEST_FIT
            elif kind.startswith("line") or kind.startswith("area"):
                dl.position = XL_LABEL_POSITION.ABOVE
            else:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
        except (ValueError, KeyError):
            pass

    # 색상: 단일 시리즈 → 포인트별 / 다중 시리즈 → 시리즈별
    colors = item.get("colors")
    sers = list(chart.series)
    if colors and len(sers) == 1:
        for i, pt in enumerate(sers[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _c(colors[i % len(colors)])
    # 값 라벨을 막대와 같은 색으로 (`label_colors: true` 면 `colors` 를 그대로 쓴다)
    lc = item.get("label_colors")
    if lc and item.get("value_labels") and len(sers) == 1:
        lc = colors if lc is True else lc
        for i, pt in enumerate(sers[0].points):
            _font_ea(pt.data_label.font, fname,
                     size=item.get("label_size", size),
                     color=lc[i % len(lc)], bold=True)
            # 포인트별 dLbl 을 새로 만들면 서식 상속이 끊겨 숫자 포맷이
            # General 로 떨어진다("4,030" → "4030"). 여기서 다시 박는다.
            dLbl = pt.data_label._dLbl
            if dLbl is not None and dLbl.find(qn("c:numFmt")) is None:
                nf = dLbl.makeelement(qn("c:numFmt"), {
                    "formatCode": item.get("label_format", nfmt),
                    "sourceLinked": "0"})
                anchor = dLbl.find(qn("c:spPr"))
                if anchor is None:
                    anchor = dLbl.find(qn("c:txPr"))
                if anchor is not None:
                    anchor.addprevious(nf)
                else:
                    dLbl.append(nf)
            if dLbl is not None:  # 포인트별 라벨은 기본이 줄바꿈 — "4,030" 이 접힌다
                bp = dLbl.find(qn("c:txPr"))
                bp = None if bp is None else bp.find(qn("a:bodyPr"))
                if bp is not None:
                    bp.set("wrap", "none")
    else:
        is_line = kind in ("line", "line_markers")
        for i, ser in enumerate(sers):
            c = (colors[i % len(colors)] if colors else item["series"][i].get("color"))
            if not c:
                continue
            if is_line:
                # 선 그래프는 면이 아니라 '선'에 색이 걸린다. fill만 칠하면
                # 오피스 기본 팔레트(파랑·빨강)가 그대로 나와 테마가 깨진다.
                ser.format.line.color.rgb = _c(c)
                ser.format.line.width = Pt(item.get("line_w", 2.25))
                _marker_color(ser, c, item.get("marker_size", 6))
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = _c(c)

    if not is_pie:
        try:
            cat_ax, val_ax = chart.category_axis, chart.value_axis
            cat_ax.visible = item.get("cat_axis", True)
            val_ax.visible = item.get("val_axis", False)
            cat_ax.has_major_gridlines = False
            val_ax.has_major_gridlines = item.get("gridlines", False)
            if val_ax.has_major_gridlines:
                # 기본값은 오피스 회색이라 라이트 지면에서 너무 진하다.
                gl = val_ax.major_gridlines.format.line
                gl.color.rgb = _c(item.get("gridline_color", _ACTIVE["hair"]))
                gl.width = Pt(item.get("gridline_w", 0.5))
            if cat_ax.visible:
                _font_ea(cat_ax.tick_labels.font, fname, size=size, color=_ACTIVE["body_ink"])
                cat_ax.format.line.color.rgb = _c(item.get("axis_color", _ACTIVE["hair"]))
            if val_ax.visible:
                _font_ea(val_ax.tick_labels.font, fname, size=size, color=_ACTIVE["gray"])
            if item.get("value_max") is not None:
                val_ax.maximum_scale = item["value_max"]
            if item.get("value_min") is not None:
                val_ax.minimum_scale = item["value_min"]
        except (ValueError, KeyError, AttributeError):
            pass

    # python-pptx 차트 템플릿의 axId 가 음수(-2068027336·-2113994440)다.
    # ISO 29500 은 c:axId/@val 이 xsd:unsignedInt 라 **규격 위반** — xmllint 는
    # 통과해도 스키마 검증에서 걸린다(실측).
    # crossAx 참조까지 같은 값으로 바꿔야 하므로 abs 로 일괄 리매핑한다.
    cs = chart._chartSpace
    for tag in ("c:axId", "c:crossAx"):
        for e in cs.iter(qn(tag)):
            v = int(e.get("val"))
            if v < 0:
                e.set("val", str(abs(v)))
    return gf


def _tc_side_border(cell, color, w_pt, side="R"):
    """셀 좌·우 세로 룰 — `a:lnL` / `a:lnR`.

    이게 없어서 표에 **열 경계 세로 룰이 아예 그어지지 않았다.** 스펙에 `col_w`
    를 적어도 렌더러가 안 그리니 표가 격자로 안 읽히고, 게이트 `[GRAM]` 의
    `vrule_coverage` 가 커버리지 0.006 으로 잡아낸다.
    `a:lnL·a:lnR` 은 `a:lnB` 보다 **앞**에 와야 스키마가 맞는다.
    """
    tag = qn(f"a:ln{side}")
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(tag)
    if old is not None:
        tcPr.remove(old)
    ln = tcPr.makeelement(tag, {"w": str(int(w_pt * 12700)), "cap": "flat"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    fill.append(fill.makeelement(qn("a:srgbClr"), {"val": color}))
    ln.append(fill)
    # 순서: lnL → lnR → lnT → lnB → 채움
    order = ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]
    after = order[order.index(f"a:ln{side}") + 1:]
    ref = None
    for t in after + ["a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
                      "a:pattFill", "a:grpFill"]:
        found = tcPr.find(qn(t))
        if found is not None:
            ref = found
            break
    if ref is None:
        tcPr.append(ln)
    else:
        ref.addprevious(ln)


def _tc_bottom_border(cell, color, w_pt, dash=None):
    """셀 하단 헤어라인 — solidFill 앞에 a:lnB 삽입(OOXML 시퀀스 준수).

    dash = OOXML prstDash 값(`dash`/`sysDash`/`sysDot`…). 점선 구분선을 도형
    수백 개로 흉내내지 않고 선 자체 속성으로 낸다.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("a:lnB"))
    if old is not None:
        tcPr.remove(old)
    ln = tcPr.makeelement(qn("a:lnB"), {"w": str(int(w_pt * 12700)), "cap": "flat"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    clr = fill.makeelement(qn("a:srgbClr"), {"val": color})
    fill.append(clr)
    ln.append(fill)
    if dash:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    fill_tags = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
                 "a:pattFill", "a:grpFill")
    ref = None
    for tag in fill_tags:
        ref = tcPr.find(qn(tag))
        if ref is not None:
            break
    if ref is not None:
        ref.addprevious(ln)
    else:
        tcPr.append(ln)


def _cell_lines(val):
    """표 셀 값 → [[run, …], …] (줄 × 런).

    문자열은 `\\n` 으로 줄을 나눈다(python-pptx 는 run.text 안의 개행을
    `_x000A_` 로 이스케이프하므로 문단을 직접 나눠야 한다).
    리스트를 주면 각 원소가 한 줄이고, 원소가 다시 리스트/딕셔너리면 리치런이다.
    """
    lines = val if isinstance(val, list) else str(val).split("\n")
    out = []
    for ln in lines:
        if isinstance(ln, list):
            out.append(ln)
        elif isinstance(ln, dict):
            out.append([ln])
        else:
            out.append([{"t": str(ln)}])
    return out or [[{"t": ""}]]


def _add_table(slide, item):
    """네이티브 PPT 표(셀 데이터 편집 가능). 밴딩 스타일 제거 후 직접 채색."""
    rows = item["rows"]
    nrows, ncols = len(rows), len(rows[0])
    gf = slide.shapes.add_table(nrows, ncols, _x(item["x"]), _y(item["y"]),
                                _x(item["w"]), _y(item["h"]))
    table = gf.table
    table.first_row = table.last_row = table.horz_banding = table.vert_banding = False
    styleId = table._tbl.tblPr.find(qn("a:tableStyleId"))
    if styleId is None:
        styleId = table._tbl.tblPr.makeelement(qn("a:tableStyleId"), {})
        table._tbl.tblPr.append(styleId)
    styleId.text = _TABLE_NO_GRID

    col_w = item.get("col_w")
    if col_w:
        for i, cw in enumerate(col_w[:ncols]):
            table.columns[i].width = _x(item["w"] * cw)
    # row_hs = 행마다 다른 높이(정규화 y). 1:1 재현 표처럼 행 높이가 제각각일 때.
    row_hs = item.get("row_hs")
    if row_hs:
        for r, rh in zip(table.rows, row_hs):
            r.height = _y(rh)
    elif item.get("row_h"):
        for r in table.rows:
            r.height = _y(item["row_h"])
    # merges — 셀 병합. 두 형태를 받는다:
    #   [행, 시작열, 끝열]              → 가로 병합(통칸)
    #   [시작행, 시작열, 끝행, 끝열]     → 사각 병합(좌라벨 열이 여러 행을 먹는 구조)
    for m in (item.get("merges") or []):
        if len(m) == 4:
            r0, c0, r1, c1 = m
        else:
            r0, c0, c1 = m
            r1 = r0
        table.cell(r0, c0).merge(table.cell(r1, c1))

    header = item.get("header", True)
    size = item.get("size", 11)
    fontkey = item.get("font", "body")
    aligns = item.get("align")
    borders = item.get("borders", True)
    # 셀 안쪽 여백·줄간격 — 촘촘한 재현 표에서 행이 콘텐츠에 밀려 늘어나는 걸 막는다
    pad_x = item.get("pad_x", 0.014)
    pad_y = item.get("pad_y", 0.008)
    line_spacing = item.get("line_spacing")
    # 행 강조 — 타임테이블의 강연 행처럼 특정 행만 면·글자색·굵기를 바꾼다.
    # 키는 문자열/정수 인덱스 모두 허용(JSON은 정수 키를 못 쓴다).
    def _row_opt(key, ri):
        d = item.get(key) or {}
        return d.get(str(ri), d.get(ri))
    bold_rows = {str(r) for r in (item.get("row_bold") or [])}
    # col_bold = 열 단위 굵기. 보고서 표 문법의 "연분홍 첫 열 굵은 라벨" 이
    # 이것 없이는 표현이 안 돼서 세 열이 같은 굵기로 평평하게 읽혔다
    # row_bold 와 대칭으로 둔다.
    bold_cols = {str(c) for c in (item.get("col_bold") or [])}
    for ri, row in enumerate(rows):
        is_head = header and ri == 0
        row_fill = _row_opt("row_fills", ri)
        row_color = _row_opt("row_colors", ri)
        row_is_bold = str(ri) in bold_rows
        for ci in range(ncols):
            val = row[ci] if ci < len(row) else ""
            cell = table.cell(ri, ci)
            if cell.is_spanned:      # 병합에 먹힌 칸 — 원점 셀에만 쓴다
                continue
            # col_pad_x = 좌우 여백 덮어쓰기. 키는 "열" 또는 "행,열"(셀 하나).
            # 좁은 좌라벨 열이 본문 열과 같은 여백을 쓰면 라벨이 줄바꿈된다.
            _cpad = item.get("col_pad_x") or {}
            cpx = _cpad.get(f"{ri},{ci}", _cpad.get(str(ci), pad_x))
            cell.margin_left = cell.margin_right = Emu(int(cpx * SLIDE_W))
            # pad_top/pad_bottom 을 따로 주면 세로중앙이 그만큼 위·아래로 밀린다 —
            # 1:1 재현에서 글줄 상자의 어센트 여백 때문에 생기는 쏠림 보정용.
            # col_pad_top/col_pad_bottom 은 셀 단위 덮어쓰기("열" 또는 "행,열").
            # 세로 병합된 좌라벨을 원본처럼 중앙보다 위에 앉힐 때 쓴다.
            def _pv(key, dflt):
                d = item.get(key) or {}
                return d.get(f"{ri},{ci}", d.get(str(ci), dflt))
            cell.margin_top = Emu(int(
                _pv("col_pad_top", item.get("pad_top", pad_y)) * SLIDE_H))
            cell.margin_bottom = Emu(int(
                _pv("col_pad_bottom", item.get("pad_bottom", pad_y)) * SLIDE_H))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            # cell_fills = 셀 하나만 지정("행,열" 키). "none" 이면 투명 —
            # 표 뒤에 미리 깔아둔 알약·칩 같은 도형이 셀 바탕으로 비쳐 보인다.
            cell_fill = (item.get("cell_fills") or {}).get(f"{ri},{ci}")
            if cell_fill is not None:
                fill = cell_fill
            elif is_head:
                fill = item.get("header_fill", _ACTIVE["ink"])
            elif row_fill:
                fill = row_fill
            elif _row_opt("col_fills", ci):
                fill = _row_opt("col_fills", ci)
            elif item.get("alt_fill") and ri % 2 == 1:
                fill = item["alt_fill"]
            else:
                fill = item.get("body_fill", _ACTIVE["bg"])
            if fill == "none":
                cell.fill.background()
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = _c(fill)
            # cell_colors = 셀 하나만 글자색 덮어쓰기("행,열" 키). row_colors 보다
            # 우선한다 — 행 전체를 칠하면 악센트가 셀 몇 개에 걸쳐 "파란 면"처럼
            # 번져서 반려된다. 강조는 한 칸이면 족하다.
            _cell_color = (item.get("cell_colors") or {}).get(f"{ri},{ci}")
            color = (item.get("header_color", "FFFFFF") if is_head
                     else _cell_color or row_color
                     or item.get("body_color", _ACTIVE["ink"]))
            style = {**_typo("body"), "size": size,
                     "font": (item.get("header_font", "body_b") if is_head else fontkey),
                     "color": color}
            # row_aligns/row_line_spacings = 행 단위 덮어쓰기(헤더행만 가운데정렬 등)
            # cell_aligns = 셀 하나만("행,열" 키) — 열 정렬보다 우선한다.
            al = ((item.get("cell_aligns") or {}).get(f"{ri},{ci}")
                  or _row_opt("row_aligns", ri) or (
                      aligns[ci] if aligns and ci < len(aligns)
                      else item.get("cell_align", "left")))
            # cell_line_spacings = 셀 하나만("행,열" 키). 같은 행에서 좌라벨(촘촘)과
            # 본문(넉넉)의 줄간격이 다른 1:1 재현 표에 쓴다. 행 단위보다 우선한다.
            ls = ((item.get("cell_line_spacings") or {}).get(f"{ri},{ci}")
                  or _row_opt("row_line_spacings", ri) or line_spacing)
            force_bold = (is_head or row_is_bold or str(ci) in bold_cols
                          or item.get("weight") == "bold")
            tf = cell.text_frame
            for li, line_runs in enumerate(_cell_lines(val)):
                para = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
                rr = []
                for r in line_runs:
                    r = {"color": color, **r}
                    if force_bold:
                        r.setdefault("weight", "bold")
                    rr.append(r)
                _set_runs(para, rr, style)
                para.alignment = _ALIGN[al]
                if ls:
                    # 5 미만 = 배수, 이상 = 고정 pt(촘촘한 재현 표에서 줄 간격을 못박는다)
                    para.line_spacing = ls if ls < 5 else Pt(ls)
            bc = item.get("border_color", _ACTIVE["hair"])
            if isinstance(bc, list):
                bc = bc[ci] if ci < len(bc) else bc[-1]
            if borders and not is_head and ri < nrows - 1:
                # border_color 는 문자열 또는 열별 리스트. border_dash 로 점선.
                _tc_bottom_border(cell, bc, item.get("border_w", 0.75),
                                  item.get("border_dash"))
            # vborders = 열 경계 세로 룰. 표가 격자로 읽히려면 이게 있어야 한다.
            if item.get("vborders") and ci < ncols - 1:
                # 머리행은 짙은 면이라 같은 회색 룰이 묻힌다 — 그 구간만 룰이
                # 끊긴 것으로 잡힌다([GRAM] vrule_coverage 0.90). 밝게 긋는다.
                vc = (item.get("header_vborder", "FFFFFF") if is_head else bc)
                _tc_side_border(cell, vc, item.get("border_w", 0.75), "R")
            # close_rule = 마지막 행을 닫는 하단 룰. 없으면 표가 아래로 열린다.
            if item.get("close_rule") and ri == nrows - 1:
                _tc_bottom_border(
                    cell, item.get("close_color", _ACTIVE["ink"]),
                    item.get("close_w", 1.0))
            if is_head:
                _tc_bottom_border(cell, item.get("header_border", item.get("header_fill", _ACTIVE["ink"])), 1.2)
    return gf


def _clip_pics_to_canvas(slide):
    """캔버스 밖으로 나간 그림을 srcRect 크롭으로 판면 안에 눕힌다 (R-SWEEP).

    블리드 연출(워드마크·글로우·목업 걸치기)은 시각적으로 그대로 두되,
    **개체 자체는 캔버스 안**에 들어오게 한다 — 편집 화면에서 밖으로 삐져나온
    개체가 거슬린다는 반려. 회전된 그림과 이미
    srcRect 가 있는 그림은 건드리지 않는다. 도형(sp)은 크롭이 불가능해 제외.
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    eps = 1000  # EMU — 반올림 잔여는 무시
    for sh in list(slide.shapes):
        if sh.shape_type != MSO_SHAPE_TYPE.PICTURE or sh.rotation:
            continue
        # 모핑 짝(!!이름)은 화면 밖 대기·확대 파고들기가 연출이다 — 자르면 모핑이 깨진다
        if sh.name.startswith("!!"):
            continue
        l, t, w, h = sh.left, sh.top, sh.width, sh.height
        if w <= 0 or h <= 0:
            continue
        over_l = max(0, -l)
        over_t = max(0, -t)
        over_r = max(0, l + w - SLIDE_W)
        over_b = max(0, t + h - SLIDE_H)
        if max(over_l, over_t, over_r, over_b) <= eps:
            continue
        blipFill = sh._element.blipFill
        if blipFill is None or blipFill.find(qn("a:srcRect")) is not None:
            continue
        sr = blipFill.makeelement(qn("a:srcRect"), {})
        for k, v in (("l", over_l / w), ("t", over_t / h),
                     ("r", over_r / w), ("b", over_b / h)):
            if v > 0:
                sr.set(k, str(int(round(v * 100000))))
        blip = blipFill.find(qn("a:blip"))
        blip.addnext(sr)
        sh.left, sh.top = max(l, 0), max(t, 0)
        sh.width, sh.height = min(l + w, SLIDE_W) - max(l, 0), \
            min(t + h, SLIDE_H) - max(t, 0)


def render_page(slide, spec: dict, assets_dir: Path, motion_ctx: dict | None = None):
    global _ACTIVE
    theme_name = spec.get("theme", "gyoan")
    if "theme_tokens" in spec:  # 스펙 인라인 일회성 테마
        _ACTIVE = _normalize_theme(spec["theme_tokens"])
        THEMES.setdefault(theme_name, _ACTIVE)
    else:
        _ACTIVE = resolve_theme(theme_name, assets_dir.parent)
    T = _ACTIVE
    variant = spec.get("variant", "paper")

    # 배경 사진 — 슬라이드 배경 자체를 흐린 사본으로 채운다.
    # 유리 도형(depth: glass + bg_fill)이 이 배경을 그대로 비추는 구조.
    bg_spec = spec.get("background")
    if bg_spec and bg_spec.get("image"):
        src = Path(bg_spec["image"])
        if not src.is_absolute():
            src = assets_dir.parent / src
        blurred = _blurred_bg_png(assets_dir, src, bg_spec.get("blur", 18),
                                  bg_spec.get("darken", 0.15))
        _set_slide_background(slide, blurred)
        if bg_spec.get("sharp"):  # 선명한 원본을 한 겹 더 — 유리 뒤만 흐려 보인다
            slide.shapes.add_picture(str(src), 0, 0, SLIDE_W, SLIDE_H)
    elif bg_spec and bg_spec.get("color"):   # 단색 배경 지정 — useBgFill 도형은 그림 배경을 전환 중에 못 비춘다
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _c(str(bg_spec["color"]).lstrip("#"))
    else:
        # 슬라이드 **배경 속성**으로 칠한다 — 예전엔 전면 사각형 도형을 깔았는데,
        # 편집할 때 맨 뒤 판이 자꾸 잡히고 개체 수만 늘었다(2026-08-19 지적).
        # 결과는 같고, `bg_fill: true` 유리 도형이 비추는 배경도 이쪽이 정확하다.
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _c(
            T["dark_bg"] if variant == "dark" else T["bg"])

    # figure(도식 골격)는 여기서 프리미티브로 전개된다 — 아래 루프는 그대로 둔다.
    from pptm.design.figures import expand_figures
    fullbleed_below = False  # 풀블리드 card/shape 가 이미 깔렸나 — glow 알파 판정용
    # 모션 — 아이템이 만든 도형을 추적해 모핑 이름(morph_id)·애니메이션(anim)을 건다.
    from pptm.design import motion
    anim_entries, morph_used, groups = [], {}, {}
    spTree = slide.shapes._spTree
    # motion(레시피·auto)은 전개된 아이템 위에서 anim·transition 으로 풀린다 — 명시값 우선
    page_items, page_tr = motion.plan_page(
        spec, expand_figures(spec.get("items", []), T, SLIDE_W / SLIDE_H, variant), motion_ctx)
    media_nodes = []

    def _made(before):
        """아이템이 새로 만든 도형 — 3D 모델처럼 mc:AlternateContent 로 감싼 개체도 잡는다."""
        from pptm.design.media3d import _Handle
        new = list(spTree)[before:]
        shapes = {sh._element: sh for sh in slide.shapes}
        out = []
        for el in new:
            if el in shapes:
                out.append(shapes[el])
            elif el.tag.endswith("}AlternateContent"):
                nv = el.find(".//" + qn("p:cNvPr"))
                if nv is not None:
                    out.append(_Handle(el, int(nv.get("id")), nv.get("name")))
        return out

    from pptm.design.media3d import fill_media_dur
    page_items = [fill_media_dur(it, assets_dir.parent) if it.get("type") in ("video", "audio") else it
                  for it in page_items]
    deck_state = (motion_ctx or {}).setdefault("deck_state", {})   # 링크·확대/축소 — 덱 끝에서 관계를 건다
    names = {}                                                       # name → spid (anim trigger "#이름")
    for item in page_items:
        before = len(spTree)
        if item["type"] == "zoom":           # 확대/축소(Slide Zoom) — 관계는 interact.finish_deck
            from pptm.design import interact
            interact.zoom_placeholder(slide, item, _x(item["x"]), _y(item["y"]), _x(item["w"]), _y(item["h"]),
                                      assets_dir.parent, deck_state)
        else:
            _render_item(slide, item, T, variant, assets_dir, theme_name, fullbleed_below)
        if item.get("link") or item.get("fill_image") or item.get("bevel") or item.get("name"):
            from pptm.design import interact
            made = _made(before)
            els = [m._element for m in made]
            if made and item.get("name"):
                names[str(item["name"])] = made[0].shape_id
                if not item.get("morph_id"):
                    made[0].name = str(item["name"])
            if item.get("link"):             # 하이퍼링크 메뉴 — COM 견본 a:hlinkClick
                interact.note_link(deck_state, slide, els, item["link"])
            if item.get("fill_image"):       # 사진 채운 글자 — COM 견본 rPr/blipFill
                fp = Path(item["fill_image"])
                interact.text_picture_fill(slide, els, fp if fp.is_absolute() else assets_dir.parent / fp)
            if item.get("bevel"):            # 입체 서식 — sp3d
                interact.apply_bevel(els, item["bevel"])
        if item["type"] in ("video", "audio"):    # 재생 노드(p:video·p:audio) — 음량·반복·숨김
            m = item.get("media") or {}
            for sh in _made(before):
                media_nodes.append({**m, "kind": item["type"], "spid": sh.shape_id})
        if item["type"] in ("card", "shape") and item.get("w", 0) >= 0.98 \
                and item.get("h", 0) >= 0.98:
            fullbleed_below = True
        if item.get("group"):  # 한 덩어리 — 끝에서 grpSp 하나로 묶고 효과·이름은 그룹에
            made = _made(before)
            g = groups.setdefault(str(item["group"]), {"shapes": [], "anim": None,
                                                        "morph_id": None})
            g["shapes"] += made
            g["anim"] = g["anim"] or item.get("anim")
            g["morph_id"] = g["morph_id"] or item.get("morph_id")
        elif item.get("morph_id") or item.get("anim"):
            made = _made(before)
            if item.get("morph_id"):
                motion.name_morph(made, str(item["morph_id"]), morph_used)
            if item.get("anim"):
                for k, sh in enumerate(made):
                    a = item["anim"]
                    if k:  # 한 아이템이 여러 도형을 만들면 나머지는 "함께" — 지연은 첫 도형과 같게
                        a = {"effect": a} if isinstance(a, str) else a   # (그림자 2겹 카드가 계단에서 먼저 뜨던 것)
                        a = {**a, "start": "with", "delay": a.get("delay", 0)}
                    ent = (a, sh.shape_id, motion.is_build_shape(sh), motion.n_paragraphs(sh))
                    if item["type"] == "chart":    # 계열·항목별 등장
                        ent += ({"series": len(item.get("series", [])),
                                 "category": len(item.get("categories", []))},)
                    anim_entries.append(ent)

    if spec.get("clip_pics", True):
        _clip_pics_to_canvas(slide)
    for gname, g in groups.items():
        grp = motion.group_shapes(slide, g["shapes"], gname)
        if g["morph_id"]:
            motion.name_morph([grp], str(g["morph_id"]), morph_used)
        if g["anim"]:
            anim_entries.append((g["anim"], grp.shape_id, False))
    motion.apply_transition(slide, page_tr)
    anim_entries = motion.resolve_triggers(anim_entries, names)
    motion.apply_timing(slide, anim_entries, media_nodes)


def _render_item(slide, item, T, variant, assets_dir, theme_name, fullbleed_below):
    """아이템 하나를 그린다 — render_page 가 앞뒤 도형 수로 그 아이템의 도형을 잡는다."""
    t = item["type"]
    if t == "glow":
        png = _glow_png(assets_dir, theme_name, item["cx"], item["cy"],
                        item.get("sigma", 0.35), item.get("intensity", 1.0),
                        alpha=item.get("alpha", fullbleed_below))
        slide.shapes.add_picture(str(png), 0, 0, SLIDE_W, SLIDE_H)
    elif t == "scrim":
        edge = item.get("edge", "bottom")
        png = _scrim_png(assets_dir, edge, item.get("color", "000000"),
                         item.get("strength", 0.9), item.get("softness", 1.25))
        if edge in ("bottom", "top"):
            hh = item.get("h", 0.5)
            yy = 1 - hh if edge == "bottom" else 0.0
            slide.shapes.add_picture(str(png), 0, _y(yy), SLIDE_W, _y(hh))
        else:
            ww = item.get("w", 0.5)
            xx = 1 - ww if edge == "right" else 0.0
            slide.shapes.add_picture(str(png), _x(xx), 0, _x(ww), SLIDE_H)
    elif t == "label":
        _add_text_item(slide, {**item, "w": item.get("w", 0.4)}, "label")
    elif t == "serif_label":
        style = dict(_typo("serif"))
        if "color" in item:
            style["color"] = item["color"]
        if "size" in item:
            style["size"] = item["size"]
        tb, tf = _textbox(slide, item["x"], item["y"], item["w"])
        para = tf.paragraphs[0]
        _set_runs(para, [{"t": item["text"]}], style)
        for run in para.runs:
            run.font.italic = True
        para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                          "right": PP_ALIGN.RIGHT}[item.get("align", "center")]
    elif t == "headline":
        _add_text_item(slide, item, "divider_h" if variant == "dark" else "headline")
    elif t == "text":
        _add_text_item(slide, item, item.get("style", "body"))
    elif t == "card":
        fill = depth_fill(item, T, variant) or T["card"]

        def _draw_card(top=True, item=item, fill=fill):
            s = _rounded(slide, item["x"], item["y"], item["w"], item["h"],
                         fill, _radius_from(item, 0.18),
                         item.get("outline"), item.get("outline_w", 1.2),
                         item.get("outline_dash"))
            if item.get("rot"):  # 스티키 노트 등 살짝 기울이기 (도)
                s.rotation = item["rot"]
            if item.get("grad"):  # 세로 그라데이션 ["top","bottom"]
                _apply_grad(s, item["grad"][0], item["grad"][1],
                            item.get("grad_ang", 90.0),
                            item.get("grad_scaled", True),
                            item.get("grad_mids"))
            return s

        shp = _with_depth(_draw_card, item)
    elif t == "shape":
        _add_shape_item(slide, item, T, variant)
    elif t == "pill":
        h = item["h"]
        w = item.get("w", _pill_w(item["text"], item.get("size", 10.5)))
        pfill = depth_fill(item, T, variant) or T["accent"]
        shp = _with_depth(
            lambda top: _rounded(slide, item["x"], item["y"], w, h, pfill, radius=0.5),
            item)
        tf = shp.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = Emu(20000)
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        style = dict(_typo("pill"))
        style["color"] = item.get("text_color", style["color"])
        style["font"] = item.get("font", style["font"])
        if "size" in item:
            style["size"] = item["size"]
        _set_runs(para, [{"t": item["text"]}], style)
        para.alignment = PP_ALIGN.CENTER
        if item.get("grad"):  # 카드와 동일: 세로 그라데이션
            _apply_grad(shp, item["grad"][0], item["grad"][1])
    elif t == "chip":
        d = item["d"]
        shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, _x(item["x"]), _y(item["y"]),
                                     _x(d), Emu(int(d * SLIDE_W)))
        shp.fill.solid()
        shp.fill.fore_color.rgb = _c(item.get("fill", T["accent"]))
        shp.line.fill.background()
        shp.shadow.inherit = False
        _strip_style(shp)
        tf = shp.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        style = dict(_typo("chip"))
        style["color"] = item.get("text_color", style["color"])
        # text 는 선택이다. 글자 없는 작은 점은 넘버 서클의 공식 대안인데
        # (deck R-FIGURE 대안 ②), 정작 KeyError 로 렌더가 죽어서 쓸 수 없었다.
        _set_runs(para, [{"t": item.get("text", "")}], style)
        para.alignment = PP_ALIGN.CENTER
    elif t == "band":
        # 채워진 바(SO WHAT/클로저/캡션) — 라벨+본문을 세로·가로 자동 중앙 정렬.
        # card+별도 text 손좌표의 세로 어긋남·라벨↔본문 빈공백을 구조적으로 제거.
        bfill = depth_fill(item, T, variant) or T["card"]
        shp = _with_depth(
            lambda top: _rounded(slide, item["x"], item["y"], item["w"], item["h"],
                                 bfill, radius=_radius_from(item, 0.5)),
            item)
        if item.get("grad"):  # 카드와 동일 문법 — 2-stop 그라데이션
            _apply_grad(shp, item["grad"][0], item["grad"][1],
                        item.get("grad_ang", 90.0),
                        item.get("grad_scaled", True),
                        item.get("grad_mids"))
        tf = shp.text_frame
        tf.word_wrap = True
        pad = item.get("pad", 0.024)
        tf.margin_left = tf.margin_right = _x(pad)
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = tf.paragraphs[0]
        lsize = item.get("label_size", 9.5)
        runs = []
        if item.get("label"):
            runs.append({"t": item["label"], "color": item.get("label_color", T["accent"]),
                         "weight": "bold", "size": lsize})
            runs.append({"t": item.get("sep", "   "), "size": lsize})
        runs.append({"t": item.get("text", ""), "color": item.get("text_color", T["ink"]),
                     "size": item.get("size", 11)})
        _set_runs(para, runs, dict(_typo("body")))
        para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                          "right": PP_ALIGN.RIGHT}[item.get("align", "center")]
    elif t == "hairline":
        _hairline_shape(slide, item["x"], item["y"], item["w"], item.get("color"),
                        dash=item.get("dash"), thick=item.get("thick"),
                        dash_seg=item.get("dash_seg"),
                        dash_gap=item.get("dash_gap"))
    elif t == "arrow":
        ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, _x(item["x"]), _y(item["y"]),
                                    _x(item["w"]), _y(item.get("h", 0.014)))
        ar.adjustments[0] = 0.5
        ar.adjustments[1] = 0.55
        if item.get("rot"):      # 역방향 커넥터(뱀형 스텝 되돌이 행) — 180도
            ar.rotation = item["rot"]
        ar.fill.solid()
        ar.fill.fore_color.rgb = _c(item.get("color", T["arrow"]))
        ar.line.fill.background()
        ar.shadow.inherit = False
        _strip_style(ar)
    elif t == "sowhat":
        y = item.get("y", 0.885)
        # 다크 지면에서는 잉크·헤어라인도 뒤집어야 한다 — 안 그러면 검은 바탕에
        # 검은 글자가 앉는다(실측 1.0:1, 납품물 게이트 [L]).
        dark = variant == "dark"
        _hairline_shape(slide, 0.05, y, 0.90,
                        item.get("line_color",
                                 T["dark_sub"] if dark else None))
        tb, tf = _textbox(slide, 0.05, y + 0.028, 0.90)
        para = tf.paragraphs[0]
        # label:"" 이면 라벨 런을 통째로 빼고 선행 공백 3칸도 안 붙인다.
        # "SO WHAT" 은 영어 컨설팅 용어라 한국어 교육·안내 자료에서는 빼달라는
        # 요청이 나온다. 그때 이 아이템을 버리고 헤어라인+
        # 텍스트로 손조립하면 `[Q]`(본문의 하단 바 침범)·`[D]`(하단 빈칸)가
        # **하단 바를 못 찾아 통째로 침묵한다** — 두 게이트가 type in
        # ("sowhat","band") 로만 바를 찾기 때문이다. 라벨만 끄는 길을 열어
        # 게이트를 살린 채 라벨을 없앤다.
        label = item.get("label", "SO WHAT")
        runs = []
        if label:
            # label_color — 악센트를 지면당 한 곳으로 묶는 덱에서 이 고정 라벨을
            # 무채로 내린다. 생략하면 종전대로 악센트색.
            runs.append({"t": label,
                         "color": item.get("label_color", T["accent"]),
                         "font": "body_b",
                         "size": item.get("label_size",
                                          item.get("size", 12) * 0.82)})
        # 핵심 한 문장이다 — 페이지 번호처럼 작게 두면 아무도 안 읽는다.
        # size 로 키운다 (피치덱 16~20 / 문서형 12 기본).
        runs.append({"t": ("   " + item["text"]) if label else item["text"],
                     "color": item.get("color",
                                       T["dark_fg"] if dark else T["ink"]),
                     "font": "body_sb", "size": item.get("size", 12)})
        _set_runs(para, runs, _typo("sowhat"))
    elif t == "timetable":
        rows = item["rows"]
        row_h = item.get("row_h", 0.135)
        x0, y0, w0 = item["x"], item["y"], item["w"]
        time_w = item.get("time_w", 0.13)
        body_x = x0 + item.get("body_indent", 0.17)
        for i, r in enumerate(rows):
            ry = y0 + i * row_h
            cy = ry + row_h / 2
            hl = r.get("highlight", False)
            text_h = 0.036
            ty = cy - text_h / 2
            if hl:
                band_h = row_h * 0.72
                _rounded(slide, x0, cy - band_h / 2, w0, band_h,
                         item.get("highlight_fill", T["highlight"]), radius=0.5)
            tb, tf = _textbox(slide, x0 + 0.025, ty, time_w)
            _set_runs(tf.paragraphs[0], [{"t": r["time"]}],
                      {**_typo("cardtitle"), "size": 15,
                       "color": T["accent"] if hl else T["ink"]})
            tb, tf = _textbox(slide, body_x, ty + 0.002, 0.45)
            tf.word_wrap = False
            _set_runs(tf.paragraphs[0], [{"t": r["body"]}],
                      {**_typo("body"), "size": 12.5,
                       "color": T["ink"] if hl else T["body_ink"]})
            if hl and r.get("pill"):
                pill_w = item.get("pill_w", 0.055)
                px = body_x + r.get("pill_offset", 0.145)
                ph = 0.046
                shp = _rounded(slide, px, cy - ph / 2, pill_w, ph,
                               T["accent"], radius=0.5)
                ptf = shp.text_frame
                ptf.margin_left = ptf.margin_right = ptf.margin_top = ptf.margin_bottom = 0
                ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
                _set_runs(ptf.paragraphs[0], [{"t": r["pill"]}], _typo("pill"))
                ptf.paragraphs[0].alignment = PP_ALIGN.CENTER
                if r.get("note"):
                    tb, tf = _textbox(slide, px + pill_w + 0.015, ty + 0.002, 0.3)
                    tf.word_wrap = False
                    _set_runs(tf.paragraphs[0], [{"t": r["note"]}],
                              {**_typo("body"), "size": 12.5, "color": T["ink"]})
            if i < len(rows) - 1:
                _hairline_shape(slide, x0, ry + row_h - 0.001, w0)
    elif t == "bars":
        rows = item["items"]
        mx = item.get("max", max(r["value"] for r in rows))
        row_h = item["h"] / len(rows)
        bar_x = item["x"] + item.get("label_w", 0.13)
        bar_span = item["w"] - item.get("label_w", 0.13) - 0.06
        for i, r in enumerate(rows):
            yy = item["y"] + i * row_h
            tb, tf = _textbox(slide, item["x"], yy + row_h * 0.18,
                              item.get("label_w", 0.13))
            _set_runs(tf.paragraphs[0], [{"t": r["label"]}],
                      {**_typo("body"), "size": 11, "color": T["gray"]})
            bw = max(0.015, bar_span * r["value"] / mx)
            hl = r.get("highlight")
            _rounded(slide, bar_x, yy + row_h * 0.16, bw, row_h * 0.52,
                     T["accent"] if hl else T["card2"], radius=0.5)
            tb, tf = _textbox(slide, bar_x + bw + 0.012, yy + row_h * 0.18, 0.05)
            _set_runs(tf.paragraphs[0], [{"t": str(r["value"])}],
                      {**_typo("body"), "size": 11.5, "font": "body_sb",
                       "color": T["accent"] if hl else T["ink"]})
    elif t == "steps":
        names = item["items"]
        n = len(names)
        span = item["w"] / n
        d = item.get("d", 0.022)
        for i, name in enumerate(names):
            x0 = item["x"] + i * span
            shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, _x(x0), _y(item["y"]),
                                         _x(d), Emu(int(d * SLIDE_W)))
            shp.fill.solid()
            shp.fill.fore_color.rgb = _c(item.get("fill", T["accent"]))
            shp.line.fill.background()
            shp.shadow.inherit = False
            _strip_style(shp)
            tf = shp.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            _set_runs(tf.paragraphs[0], [{"t": str(i + 1)}], _typo("chip"))
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            tb, tf2 = _textbox(slide, x0 + d + 0.008, item["y"] + 0.004,
                               span - d - 0.03)
            _set_runs(tf2.paragraphs[0], [{"t": name}],
                      {**_typo("body"), "size": 11.5, "font": "body_sb",
                       "color": T["ink"]})
            if i < n - 1:
                ar = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW,
                    _x(x0 + span - 0.033), _y(item["y"] + 0.006),
                    _x(0.02), _y(0.012))
                ar.fill.solid()
                ar.fill.fore_color.rgb = _c(T["arrow"])
                ar.line.fill.background()
                ar.shadow.inherit = False
                _strip_style(ar)
    elif t == "map":
        _add_map(slide, item, T, variant)
    elif t == "mockup":
        # 기기 목업 — 프레임 안에 화면을 합성해 한 장의 PNG로 얹는다.
        # w 만 주면 높이는 프레임 비율에서 계산한다(찌그러뜨리지 않는다).
        png, ar = _mockup_png(assets_dir, assets_dir.parent,
                              item.get("device", "phone"),
                              item.get("screen", {}), item.get("fit", "cover"),
                              item.get("px", 1600))
        w = item["w"]
        h = item.get("h") or w * (SLIDE_W / SLIDE_H) / ar
        slide.shapes.add_picture(str(png), _x(item["x"]), _y(item["y"]),
                                 _x(w), _y(h))
    elif t == "image":
        p = Path(item["path"])
        if not p.is_absolute():
            p = assets_dir.parent / item["path"]
        if p.exists():
            pic = slide.shapes.add_picture(str(p), _x(item["x"]), _y(item["y"]),
                                           _x(item["w"]), _y(item["h"]))
            if item.get("rot"):      # 스티커·테이프를 손으로 붙인 듯 기울이기 (도)
                pic.rotation = item["rot"]
            if item.get("crop"):     # 자르기 {l,t,r,b} 0~1 — 모핑이 자르기 창을 보간한다
                from pptm.design.media3d import set_crop
                set_crop(pic, item["crop"])
            if item.get("flip_h") or item.get("flip_v"):
                from pptm.design.media3d import set_flip
                set_flip(pic, item.get("flip_h"), item.get("flip_v"))
            if item.get("name"):
                pic.name = str(item["name"])
            # 무채색 창고 아이콘(assets/bank/icon3d)을 덱 팔레트로 — ["어두운hex","밝은hex"].
            # PPT '다시 칠하기'와 같은 a:duotone 이라 납품 후에도 PowerPoint 에서 바꿀 수 있다.
            if item.get("duotone"):
                from lxml import etree as _et
                _ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
                _duo = _et.SubElement(pic._element.find(f".//{{{_ns}}}blip"), f"{{{_ns}}}duotone")
                for _hex in item["duotone"]:
                    _et.SubElement(_duo, f"{{{_ns}}}srgbClr", val=_hex.lstrip("#").upper())
    elif t == "model3d":     # 3D 모델(glb) — 회전·카메라는 스펙, 장 사이 모핑이 보간
        from pptm.design.media3d import add_model3d
        add_model3d(slide, item, _x(item["x"]), _y(item["y"]), _x(item["w"]), _y(item["h"]),
                    assets_dir.parent)
    elif t == "raw":         # 원본 개체 XML 그대로(글상자) — 줄마다 다른 글자 크기 모핑 재현용
        from pptm.design.media3d import add_raw
        add_raw(slide, item)
    elif t in ("video", "audio"):  # 영상·소리 — 재생은 anim(media-play·media-toggle) + media 노드
        from pptm.design.media3d import add_media
        add_media(slide, item, _x(item["x"]), _y(item["y"]), _x(item["w"]), _y(item["h"]),
                  assets_dir.parent)
    elif t == "icon":
        from pptm.design import icons as _icons_mod
        color = item.get("color", T["ink"])
        label = item.get("text")
        # 면(card) 안에 얹을 때: h + anchor:"middle" 을 주면 박스의 세로 중앙에
        # 맞춘다. 손좌표 y로 눈대중하면 줄간격 때문에 **항상** 위로 쏠린다 —
        # 반복 반려 지점이라 구조로 막는다 (절대규칙 6의 icon 판).
        box_h = item.get("h")
        centered = box_h is not None and item.get("anchor") == "middle"
        # 아이콘은 EMU 폭·높이가 같은 **정사각**이라, x분수 size 를 y분수로
        # 바꾸려면 판면 종횡비를 곱해야 한다(국기 flag 분기와 같은 공식).
        # 16:9 는 마침 16/9 라 티가 안 났지만 A4 세로에서는 2.5mm 위로 떴다.
        _ICON_AR = SLIDE_W / SLIDE_H
        # 라벨을 주면: 아이콘 크기·세로중심을 글자에 맞춰 자동 정렬(눈대중 금지).
        #   글자 광학중심 = y + 0.00207·S,  아이콘 박스높이 = size·16/9
        if label is not None:
            tsize = item.get("text_size", 10)
            size = item.get("size", round(0.00147 * tsize, 4))
            mid = (item["y"] + box_h / 2) if centered \
                else (item["y"] + 0.00207 * tsize)
            iy = round(mid - size * _ICON_AR / 2, 4)
        else:
            size = item.get("size", 0.05)
            iy = round(item["y"] + box_h / 2 - size * _ICON_AR / 2, 4) \
                if centered else item["y"]
        px = max(96, min(1024, int(size * 3000)))
        # 브랜드 로고(Claude·OpenAI·Figma 등) — "brand": true 로 켠다.
        #   기본은 풀컬러(lobehub, MIT 래퍼) — 지면 variant에 맞는 light/dark를
        #   자동으로 고른다. "mono": true 면 단색(Simple Icons, CC0)을 테마색으로
        #   리컬러한다. 로고 자체는 각 회사 상표이니 서술적 뱃지로만 쓴다.
        if item.get("brand"):
            if item.get("mono"):
                png = _icons_mod.brand_mono_png(assets_dir, item["name"], color, px)
            else:
                bvariant = item.get("brand_variant",
                                    "dark" if variant == "dark" else "light")
                png = _icons_mod.brand_png(assets_dir, item["name"], bvariant, px)
        else:
            png = _icons_mod.icon_png(assets_dir, item["name"], color, px,
                                      item.get("weight"))
        if png is not None:
            w_emu = _x(size)  # 정사각 아이콘: 폭·높이 EMU 동일 = 화면상 정사각
            slide.shapes.add_picture(str(png), _x(item["x"]), _y(iy),
                                     w_emu, w_emu)
        if label is not None:  # 아이콘 오른쪽에 라벨 (자동 x = 아이콘 끝 + gap)
            lab = {
                "text": label,
                "x": round(item["x"] + size + item.get("gap", 0.007), 4),
                "y": item["y"], "w": item.get("w", 0.4),
                "size": item.get("text_size", 10),
                "color": item.get("text_color", color),
                "align": item.get("align", "left"),
            }
            if centered:  # 박스를 그대로 물려주고 세로중앙은 PPT가 잡게 한다
                lab["h"], lab["anchor"] = box_h, "middle"
            _add_text_item(slide, lab, item.get("text_style", "caption"))
    elif t == "flag":
        # 국기(flag-icons, MIT) — 국제비교 표·지도 라벨용. `name`은 ISO
        # 3166-1 alpha-2 소문자(kr/us/jp/cn/gb/eu…). 원색이라 리컬러 안 함.
        from pptm.design import icons as _icons_mod
        shape = item.get("shape", "4x3")
        w = item["w"]
        px = max(96, min(1024, int(w * 3000)))
        png = _icons_mod.flag_png(assets_dir, item["name"], shape, px)
        if png is not None:
            # w는 x축 분수라 그대로 y축 h로 못 쓴다 — 슬라이드가 정사각이
            # 아니라서(16:9) 실제 종횡비를 지키려면 슬라이드 종횡비로 보정해야
            # 한다(기기 목업과 같은 공식). 이걸 안 줘서 국기가 눌린 타원으로
            # 나온 적이 있다(2026-08-13 실측 — 일본 국기 원이 가로로 찌그러짐).
            ar = 4 / 3 if shape != "1x1" else 1.0
            h = w * (SLIDE_W / SLIDE_H) / ar
            slide.shapes.add_picture(str(png), _x(item["x"]), _y(item["y"]),
                                     _x(w), _y(h))
    elif t == "merge":       # 도형 병합(빼기·합치기·교차·조각) — 스포트라이트·마스크·글자 창 재료
        from pptm.design.merge import add_merge
        for shp in add_merge(slide, item, SLIDE_W, SLIDE_H):
            if item.get("fill", T["accent"]) == "none":
                shp.fill.background()
            elif item.get("alpha") is not None:     # 0~1 불투명도(유리·어둠 덮개)
                _fill_alpha(shp, str(item.get("fill", T["accent"])).lstrip("#"), float(item["alpha"]))
            else:
                shp.fill.solid()
                shp.fill.fore_color.rgb = _c(item.get("fill", T["accent"]))
            if item.get("outline"):
                shp.line.color.rgb = _c(item["outline"])
                shp.line.width = Pt(item.get("outline_w", 1.0))
            else:
                shp.line.fill.background()
            shp.shadow.inherit = False
            _strip_style(shp)
    elif t == "poly":
        shp = _add_poly(slide, item["points"], item.get("fill", T["accent"]),
                        item.get("outline"), item.get("outline_w", 1.0))
        if item.get("grad"):  # card 와 같은 문법 — 세로 2-stop
            _apply_grad(shp, item["grad"][0], item["grad"][1],
                        item.get("grad_ang", 90.0),
                        item.get("grad_scaled", True),
                        item.get("grad_mids"))
    elif t == "chart":
        _add_chart(slide, item)
    elif t == "table":
        _add_table(slide, item)
    else:
        raise ValueError(f"unknown item type: {t}")


def render_deck(project_dir: Path) -> Path:
    project_dir = Path(project_dir)
    assets_dir = project_dir / "assets"
    page_dirs = pages_mod.page_dirs(project_dir)

    # 덱 페이지 크기 — 첫 스펙의 page_size 키 (없으면 16:9)
    size_key = "16:9"
    for pdir in page_dirs:
        sp = pdir / "layout_spec.json"
        if sp.exists():
            size_key = json.loads(sp.read_text()).get("page_size", "16:9")
            break
    _set_page_size(size_key)

    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    blank = prs.slide_layouts[6]
    # 덱 단위 motion — PROJECT.json "motion": "auto" | "<레시피>" (장마다 덮어쓸 수 있다)
    from pptm.design.motion import morph_ids, morph_boxes
    proj = project_dir / "PROJECT.json"
    deck_motion = json.loads(proj.read_text()).get("motion") if proj.exists() else None
    prev_ids: set = set()
    prev_boxes: dict = {}      # 자동 모드 모핑 길이 — 큰 이동이면 0.5s
    deck_state: dict = {}      # 장을 넘나드는 것(링크·확대/축소) — 끝에서 interact.finish_deck
    for pdir in page_dirs:
        spec_path = pdir / "layout_spec.json"
        if not spec_path.exists():
            continue
        spec = json.loads(spec_path.read_text())
        slide = prs.slides.add_slide(blank)
        ids = morph_ids(spec.get("items", []))
        ctx = {"deck_motion": deck_motion, "prev_morph_ids": prev_ids,
               "shared_morph": bool(ids & prev_ids), "deck_state": deck_state, "prev_boxes": prev_boxes}
        render_page(slide, spec, assets_dir, ctx)
        prev_ids = ids
        prev_boxes = morph_boxes(spec.get("items", []))

    if deck_state.get("links") or deck_state.get("zooms"):
        from pptm.design import interact
        interact.finish_deck(prs, deck_state)
    out_dir = project_dir / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "deck.pptx"
    prs.save(out)
    if deck_state.get("zooms"):               # 확대/축소 썸네일 — 대상 장을 떠서 끼운다(PowerPoint 는 안 그려 준다)
        from pptm.design import interact
        interact.refresh_zoom_thumbs(out, deck_state)
    return out
