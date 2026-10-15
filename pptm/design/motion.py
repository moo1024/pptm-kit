"""장 전환(transition) · 개체 애니메이션(anim) · 모핑 매칭(morph_id) · 모션 레시피(motion).

python-pptx 에 API 가 없어 슬라이드 XML 에 직접 쓴다. 구조는 추측이 아니라 레퍼런스
pptx 실물에서 뜯은 그대로다(2026-10-08, 141덱 전수 — `tools/motion_harvest.py`).

- 효과 XML   `tracks/deck/motion/templates/anim/<presetClass>-<presetID>-<presetSubtype>.xml`
             실물에서 뗀 `<p:par>` 조각. 여기서는 id·spid·지속·지연·nodeType·grpId 만 바꾼다.
- 전환 XML   모핑 `mc:AlternateContent` > Choice(p159) `<p159:morph>` + Fallback `<p:fade/>`,
             페이드·푸시·와이프 Choice(p14) `p14:dur`, 팬 `p14:pan`, 필오프·커튼 `p15:prstTrans`.
- 위치       전환은 `</p:clrMapOvr>` 바로 뒤, `<p:timing>` 은 그 뒤.
- 애니 골격  mainSeq 아래 클릭 그룹 par → "이전 후" 단계 par(delay=앞 단계 끝) → 효과 par.
             첫 그룹이 자동 시작이면 `cond evt="onBegin"`(tn val=2). 도형(p:sp)만
             `grpId` + `bldLst/bldP` — 그림(pic)·그룹엔 없다. 같은 도형의 두 번째 효과는 grpId=1.
- 레시피     `tracks/deck/motion/recipes/<이름>.json` — 레퍼런스 덱의 동선을 이름으로 부른다.
             `"motion": "auto"` 는 쓰임새 통계로 짠 기본 동선.

**절제가 원칙** — 기본 효과 10종 · 전환 3종(모핑·페이드·없음) · 지속 애니 0.2~1.0s / 전환 0.3~1.0s ·
장당 클릭 3회. 날아오기·회전·펄스·푸시 같은 강한 효과와 범위 밖 지속은 `"bold": true`
(사용자가 이름으로 불렀다는 표시)일 때만 열린다. 어기면 ValueError.
"""
from __future__ import annotations

import copy
import json
import re
from functools import lru_cache
from pathlib import Path

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
P = "{%s}" % NS_P

LIB = Path(__file__).resolve().parent / "motion_templates"
TPL_ANIM = LIB / "anim"
RECIPES = LIB / "recipes"   # 레시피(JSON)를 넣으면 motion: "<이름>" 으로 부른다

# ── 사전 ────────────────────────────────────────────────────────────────
# 이름 → (템플릿 키, 등급). 등급 base = 자동 모드·기본 허용, bold = 이름으로 부를 때만.
# 키는 레퍼런스에서 뗀 템플릿 파일 이름이다(없는 효과는 여기 올리지 않는다).
EFFECTS = {
    "fade-in": ("entr-10-0", "base"),         # 페이드 — 전 개체 종류 1위(1,625건)
    "rise-in": ("entr-42-0", "base"),         # 떠오르기 — 아래서 .1 올라오며 페이드
    "zoom-in": ("entr-53-16", "base"),        # 확대/축소 — 가운데서 커지며 페이드
    "appear": ("entr-1-0", "base"),           # 나타내기 — 지속 0
    "wipe-left": ("entr-22-8", "base"),       # 닦아내기 왼쪽에서 — 그래프·막대
    "wipe-right": ("entr-22-2", "base"),
    "wipe-up": ("entr-22-4", "base"),         # 아래에서 위로 — 기둥·막대 차오름
    "wipe-down": ("entr-22-1", "base"),       # 위에서 아래로 — 세로선·목록
    "fade-out": ("exit-10-0", "base"),
    "disappear": ("exit-1-0", "base"),
    # ── 강한 효과(bold) ──
    "fly-in-left": ("entr-2-8", "bold"),
    "fly-in-right": ("entr-2-2", "bold"),
    "fly-in-bottom": ("entr-2-4", "bold"),
    "fly-in-top-right": ("entr-2-3", "bold"),
    "fly-in-bottom-left": ("entr-2-12", "bold"),
    "zoom-pop": ("entr-23-16", "bold"),       # 확대/축소(기본) — 페이드 없이 0→원래
    "zoom-fade": ("entr-53-0", "bold"),
    "grow-turn": ("entr-31-0", "bold"),       # 회전하며 커지기
    "wheel": ("entr-21-1", "bold"),
    "split": ("entr-16-21", "bold"),
    "stretch-top": ("entr-17-1", "bold"),
    "stretch-right": ("entr-17-2", "bold"),
    "stretch-bottom": ("entr-17-4", "bold"),
    "stretch-left": ("entr-17-8", "bold"),
    "stretch-across": ("entr-17-10", "bold"),
    "strips-up-right": ("entr-18-3", "bold"),
    "strips-down-right": ("entr-18-6", "bold"),
    "strips-up-left": ("entr-18-9", "bold"),
    "strips-down-left": ("entr-18-12", "bold"),
    "box-in": ("entr-4-16", "bold"),
    "box-out": ("entr-4-32", "bold"),
    "circle-in": ("entr-6-16", "bold"),
    "dissolve": ("entr-9-0", "bold"),
    "random-bars": ("entr-14-10", "bold"),
    "pulse": ("emph-26-0", "bold"),           # 강조 펄스 — 105% 왕복
    "grow": ("emph-6-0", "bold"),             # 강조 크게(150%)
    "transparency": ("emph-9-0", "bold"),     # 강조 투명도 50%
    "fill-color": ("emph-1-2", "bold"),       # 강조 채우기 색 — `color` 로 바꾼다
    "font-color": ("emph-3-2", "bold"),       # 강조 글꼴 색 — `color`
    "move": ("path", "bold"),                 # 이동 경로 직선 — `dx`·`dy`(슬라이드 비율)
    # ── 경로·미디어(2026-10-08 5덱 재현) ──
    "path-curve": ("path-1-0", "bold"),       # 곡선 경로 — `path`("M 0 0 C …", 슬라이드 비율)
    "media-play": ("mediacall-1-0", "bold"),  # 영상·소리 재생(playFrom 0) — `dur` = 재생 길이
    "media-toggle": ("mediacall-2-0", "bold"),  # 일시정지/재개 — 보통 `trigger: "self"`(그 영상을 클릭)
    "turntable-3d": ("emph-100-1", "bold"),   # 3D 턴테이블 — model3d 아이템 전용
    "swing-3d": ("emph-101-1", "bold"),       # 3D 스윙 — model3d 아이템 전용
    "path-infinity": ("path-26-0", "bold"),   # 무한대 경로 — 레퍼런스 실물 1건
    # ── 유튜브 강좌 기법(2026-10-08) — 레퍼런스에 없어 PowerPoint COM 으로 만든 실물에서 뗐다 ──
    "spin": ("emph-8-0", "bold"),             # 강조 회전 — `deg`(기본 360) · 톱니·휠 반복은 repeat:"indefinite"
}
MEDIA_KEYS = ("mediacall-1-0", "mediacall-2-0")
BASE_EFFECTS = tuple(k for k, (_, t) in EFFECTS.items() if t == "base")
STARTS = ("click", "after", "with")
BYS = ("object", "word", "letter", "paragraph", "series", "category")
CHART_BYS = ("series", "category")       # 그래프 계열별·항목별

TRANSITIONS = {"morph": "base", "fade": "base", "none": "base",
               "push": "bold", "wipe": "bold", "strips": "bold", "pan": "bold",
               "peel": "bold", "curtains": "bold"}
MORPH_OPTIONS = ("byObject", "byWord", "byChar")
MORPH_ALIAS = {"object": "byObject", "word": "byWord", "words": "byWord", "char": "byChar", "chars": "byChar",
               "letter": "byChar", "letters": "byChar"}   # 유튜브 강좌 이름(Words·Characters) — COM 생성 실물과 같은 XML
DIRS = {"push": ("l", "r", "u", "d"), "wipe": ("l", "r", "u", "d"),   # 첫 값 = OOXML 기본
        "strips": ("lu", "ru", "ld", "rd")}

ANIM_DUR = (0.2, 1.0)        # base
TR_DUR = (0.3, 1.0)          # base
BOLD_ANIM_DUR = (0.0, 60.0)
MEDIA_DUR = (0.0, 7200.0)    # 미디어 재생 길이
BOLD_TR_DUR = (0.1, 5.0)
MAX_CLICKS, BOLD_MAX_CLICKS = 3, 40   # bold 상한(재현용)


def _dur_ms(v, what, rng) -> int:
    v = float(v)
    if not rng[0] <= v <= rng[1]:
        raise ValueError(f"{what} 지속시간 {v}s — 허용 {rng[0]}~{rng[1]}s "
                         f"(범위 밖은 \"bold\": true — 사용자가 이름으로 부른 경우만)")
    return int(round(v * 1000))


def _spd(ms: int) -> str:
    return "fast" if ms <= 500 else "med" if ms <= 750 else "slow"


# ── 장 전환 ──────────────────────────────────────────────────────────────
def _ac(requires: str, ns_decl: str, choice: str, fallback: str) -> str:
    return (f'<mc:AlternateContent xmlns:mc="{NS_MC}" {ns_decl}>'
            f'<mc:Choice Requires="{requires}">{choice}</mc:Choice>'
            f'<mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>')


def transition_xml(tr) -> str | None:
    if tr is None:
        return None
    if isinstance(tr, str):
        tr = {"type": tr}
    kind = tr.get("type", "none")
    if kind not in TRANSITIONS:
        raise ValueError(f"transition {kind!r} — 허용: {tuple(TRANSITIONS)}")
    bold = bool(tr.get("bold"))
    if TRANSITIONS[kind] == "bold" and not bold:
        raise ValueError(f"transition {kind!r} 는 강한 전환 — 사용자가 이름으로 불렀을 때만 "
                         f"\"bold\": true 로 쓴다(절제 원칙: 기본은 morph·fade)")
    adv = ""
    if tr.get("advance_after") is not None:      # 자동 넘김(초) — 키오스크·연속 장표용
        adv = f' advTm="{int(round(float(tr["advance_after"]) * 1000))}"'
    if tr.get("advance_click") is False:
        adv += ' advClick="0"'
    # spd 만 쓰고 p14:dur 가 없는 꼴(`<p:transition spd="slow"><p:wipe/>`)
    # 은 "dur": null 로 부른다. 그때 spd 는 명시값(없으면 속성 자체를 안 쓴다).
    no_dur = "dur" in tr and tr["dur"] is None
    spd_attr = f' spd="{tr["spd"]}"' if tr.get("spd") else ""
    if kind == "none":
        if not adv and tr.get("dur") is None:
            return None
        # 전환 없이 자동 넘김만 — Choice p14 dur 0(또는 2000 + spd slow)
        ms = int(round(float(tr.get("dur") or 0) * 1000))
        return _ac("p14", f'xmlns:p14="{NS_P14}"',
                   f'<p:transition xmlns:p="{NS_P}"{spd_attr} p14:dur="{ms}"{adv}/>',
                   f'<p:transition xmlns:p="{NS_P}"{spd_attr}{adv}/>')
    default = {"morph": 0.8, "fade": 0.7, "pan": 1.3, "peel": 1.25, "curtains": 3.0}.get(kind, 1.0)
    rng = BOLD_TR_DUR if bold else TR_DUR
    if no_dur:
        spd = tr.get("spd")
        head = f'<p:transition xmlns:p="{NS_P}"{spd_attr}{adv}>'
        fade_fb = f'<p:transition xmlns:p="{NS_P}"{spd_attr}{adv}><p:fade/></p:transition>'
    else:
        ms = _dur_ms(tr.get("dur", default), "transition", rng)
        spd = tr.get("spd") or _spd(ms)
        if tr.get("spd") == "":
            spd = None
        sa = f' spd="{spd}"' if spd else ""
        fade_fb = f'<p:transition xmlns:p="{NS_P}"{sa}{adv}><p:fade/></p:transition>'
        head = f'<p:transition xmlns:p="{NS_P}"{sa} p14:dur="{ms}"{adv}>'
    if kind == "morph":
        opt = MORPH_ALIAS.get(tr.get("option", "byObject"), tr.get("option", "byObject"))
        if opt not in MORPH_OPTIONS:
            raise ValueError(f"morph option {opt!r} — 허용: {MORPH_OPTIONS}")
        return _ac("p159", f'xmlns:p159="{NS_P159}" xmlns:p14="{NS_P14}"',
                   f'{head}<p159:morph option="{opt}"/></p:transition>', fade_fb)
    if kind == "pan":                           
        return _ac("p14", f'xmlns:p14="{NS_P14}"', f'{head}<p14:pan/></p:transition>', fade_fb)
    if kind in ("peel", "curtains"):            
        prst = "peelOff" if kind == "peel" else "curtains"
        return _ac("p15", f'xmlns:p15="{NS_P15}" xmlns:p14="{NS_P14}"',
                   f'{head}<p15:prstTrans prst="{prst}"/></p:transition>', fade_fb)
    child = "<p:fade/>"
    if kind in DIRS:
        d = tr.get("dir")
        if d is None and not no_dur:
            d = DIRS[kind][0]
        if d is not None and d not in DIRS[kind]:
            raise ValueError(f"{kind} dir {d!r} — 허용: {DIRS[kind]}")
        child = f'<p:{kind} dir="{d}"/>' if d else f"<p:{kind}/>"
    if no_dur:                                    # AlternateContent 없이 기본 전환 하나
        return f"{head}{child}</p:transition>"
    sa = f' spd="{spd}"' if spd else ""
    return _ac("p14", f'xmlns:p14="{NS_P14}"', f"{head}{child}</p:transition>",
               f'<p:transition xmlns:p="{NS_P}"{sa}{adv}>{child}</p:transition>')


def apply_transition(slide, tr):
    xml = transition_xml(tr)
    if xml is None:
        return
    sld = slide._element
    el = etree.fromstring(xml)
    clr = sld.find(f"{P}clrMapOvr")
    if clr is not None:
        clr.addnext(el)
    else:
        sld.find(f"{P}cSld").addnext(el)


# ── 모핑 이름 ────────────────────────────────────────────────────────────
def name_morph(shapes, morph_id: str, used: dict):
    """한 아이템이 만든 도형들에 `!!id`, `!!id#2`… — 앞뒤 장에서 같은 순서면 같은 이름."""
    for sh in shapes:
        n = used.get(morph_id, 0) + 1
        used[morph_id] = n
        sh.name = f"!!{morph_id}" if n == 1 else f"!!{morph_id}#{n}"


# ── 묶음 ────────────────────────────────────────────────────────────────
def group_shapes(slide, shapes, name):
    """도형들을 grpSp 하나로 묶는다 — 첫 구성원 자리(z 순서)에 앉힌다.

    메모지 한 장(종이·압정·글·스티커)에 효과를 조각마다 걸면 zoom 은 조각이 각자 중심에서
    커지고 rise 는 압정이 종이보다 먼저 뜬다(qa-r1 반려). 묶어서 효과를 그룹 하나에 건다.
    """
    tree = slide.shapes._spTree
    els = [sh._element for sh in shapes]
    anchor = min(list(tree).index(e) for e in els)
    grp = slide.shapes.add_group_shape(shapes)       # 맨 끝에 붙고 범위를 다시 잰다
    tree.remove(grp._element)
    tree.insert(anchor, grp._element)
    grp.name = str(name)
    return grp


# ── 효과 정규화 ─────────────────────────────────────────────────────────
@lru_cache(maxsize=None)
def _tpl_text(key: str) -> str:
    f = TPL_ANIM / f"{key}.xml"
    if not f.exists():
        raise ValueError(f"효과 템플릿 {key} 없음 — 레퍼런스에서 뗀 것만 쓴다({TPL_ANIM})")
    lines = f.read_text(encoding="utf-8").split("\n", 1)
    return lines[1] if lines[0].startswith("<!--") else lines[0]


def _resolve(eff: str, bold: bool) -> tuple[str, str]:
    if eff in EFFECTS:
        key, tier = EFFECTS[eff]
    elif re.fullmatch(r"(entr|exit|emph|path)-\d+-\d+(-it(wd|lt))?", eff or ""):
        key, tier = eff, "bold"                 # 사전 키 그대로 — 레퍼런스 템플릿이 있어야 한다
        _tpl_text(key)
    else:
        raise ValueError(f"anim effect {eff!r} — 기본 {BASE_EFFECTS}, 강한 효과는 "
                         f"{tuple(k for k in EFFECTS if k not in BASE_EFFECTS)} (bold)")
    if tier == "bold" and not bold:
        raise ValueError(f"anim effect {eff!r} 는 강한 효과 — 사용자가 이름으로 불렀을 때만 "
                         f"\"bold\": true 로 쓴다(바운스·회전·날아들기 절제)")
    return key, tier


def _norm_anim(a) -> dict:
    if isinstance(a, str):
        a = {"effect": a}
    eff = a.get("effect", "fade-in")
    bold = bool(a.get("bold"))
    key, tier = _resolve(eff, bold)
    start = a.get("start", "after")
    if start not in STARTS:
        raise ValueError(f"anim start {start!r} — 허용: {STARTS}")
    delay = float(a.get("delay", 0))
    if not 0 <= delay <= (10 if bold else 3):
        raise ValueError(f"anim delay {delay}s — 0~3s")
    by = a.get("by", "object")
    if by not in BYS:
        raise ValueError(f"anim by {by!r} — 허용: {BYS}")
    dflt = 0 if eff in ("appear", "disappear") else 0.5
    dur = a.get("dur", dflt)
    rng = MEDIA_DUR if key in MEDIA_KEYS else BOLD_ANIM_DUR if bold else ANIM_DUR
    if key in MEDIA_KEYS and float(dur) == 0:
        dms = 0
    elif (eff in ("appear", "disappear") or key.startswith("entr-1-0")) and float(dur) == 0:
        dms = 0
    else:
        dms = _dur_ms(dur, "anim", rng)
    out = {"effect": eff, "key": key, "bold": bold, "start": start, "dur": dms,
           "delay": int(round(delay * 1000)), "order": a.get("order"), "by": by,
           "step": a.get("step"), "after": a.get("after"), "repeat": a.get("repeat"),
           "color": a.get("color"), "dx": float(a.get("dx", 0)), "dy": float(a.get("dy", 0)),
           "scale": a.get("scale"),
           # 실물 그대로 옮길 때 쓰는 세부값(곡선 경로: accel·decel 50% · spd -100% · rAng)
           "path": a.get("path"), "path_edit": a.get("path_edit"), "rang": a.get("rang"),
           "pts_types": a.get("pts_types"), "rctr": a.get("rctr"), "origin": a.get("origin"),
           "accel": a.get("accel"), "decel": a.get("decel"), "speed": a.get("speed"),
           "auto_rev": a.get("auto_rev"), "trigger": a.get("trigger"), "fill": a.get("fill"),
           "cmd": a.get("cmd"), "deg": a.get("deg"), "each": a.get("each")}
    if a.get("trigger") not in (None, "self") and not str(a.get("trigger")).isdigit():
        t = str(a.get("trigger"))
        raise ValueError(f"anim trigger {t!r} — \"self\"(그 개체를 클릭) · \"#이름\"(같은 장의 그 name 아이템을 클릭)"
                         + (" — 이 장에 그 이름의 아이템이 없다" if t.startswith("#") else "") + " · spid")
    if key == "path-1-0" and not out["path"]:
        raise ValueError("path-curve 효과는 path(\"M 0 0 C …\", 슬라이드 비율)가 있어야 한다")
    if key == "path" and not (out["dx"] or out["dy"]):
        raise ValueError("move 효과는 dx·dy(슬라이드 비율) 중 하나가 있어야 한다")
    return out


# ── 효과 XML ────────────────────────────────────────────────────────────
_PATH_ID = {"right": "63", "left": "35", "up": "64", "down": "0"}   # 실물 프리셋(0=사용자 지정)


def _effect_xml(e, ids) -> str:
    """레퍼런스 템플릿 → id·spid·지속·지연·nodeType·grpId 만 바꾼 효과 par."""
    key = e["key"]
    src = _tpl_text("path-64-0" if key == "path" else key)
    src = src.replace("{SPID}", e["spid"])
    if e.get("para") is not None:                      # 단락 단위
        i = e["para"]
        src = src.replace(f'<p:spTgt spid="{e["spid"]}"/>',
                          f'<p:spTgt spid="{e["spid"]}"><p:txEl><p:pRg st="{i}" end="{i}"/>'
                          f'</p:txEl></p:spTgt>')
    el = etree.fromstring(src)
    ctn = el.find(f"{P}cTn")
    my_id = next(ids)
    ctn.set("id", str(my_id))
    ctn.set("nodeType", e["node"])
    if e["grp"] is None:
        ctn.attrib.pop("grpId", None)
    else:
        ctn.set("grpId", str(e["grp"]))
    if e.get("repeat") is not None:
        r = e["repeat"]
        ctn.set("repeatCount", "indefinite" if r == "indefinite" else str(int(float(r) * 1000)))
    else:
        ctn.attrib.pop("repeatCount", None)
    ctn.find(f"{P}stCondLst/{P}cond").set("delay", str(e["delay"]))
    if key == "path":
        dx, dy = e["dx"], e["dy"]
        d = ("right" if dx > 0 else "left") if abs(dx) >= abs(dy) else ("down" if dy > 0 else "up")
        ctn.set("presetID", _PATH_ID[d])
        mot = ctn.find(f".//{P}animMotion")
        mot.set("path", f"M 0 0 L {dx:.5f} {dy:.5f} ")
        rc = mot.find(f"{P}rCtr")
        if rc is not None:
            rc.set("x", str(int(round(dx * 50000))))
            rc.set("y", str(int(round(dy * 50000))))
    if key == "path-1-0" and e.get("speed") is None:
        e = dict(e, speed=1)                       # 템플릿의 역재생 spd=-100% 를 기본으로 물려받지 않는다
    mot = ctn.find(f".//{P}animMotion")
    if mot is not None and e.get("path"):          # 경로를 실물 그대로(곡선 C · 닫힌 Z 포함)
        mot.set("path", e["path"])
        for k, attr in (("path_edit", "pathEditMode"), ("rang", "rAng"), ("pts_types", "ptsTypes"),
                        ("origin", "origin")):
            if e.get(k) is not None:
                mot.set(attr, str(e[k]))
        rc = mot.find(f"{P}rCtr")
        if e.get("rctr") is not None and rc is not None:
            rc.set("x", str(int(e["rctr"][0])))
            rc.set("y", str(int(e["rctr"][1])))
    for k, attr in (("accel", "accel"), ("decel", "decel")):
        if e.get(k) is not None:
            if float(e[k]):
                ctn.set(attr, str(int(round(float(e[k]) * 100000))))
            else:
                ctn.attrib.pop(attr, None)
    if e.get("auto_rev"):
        ctn.set("autoRev", "1")
    if e.get("fill"):
        ctn.set("fill", e["fill"])
    if e.get("cmd"):
        for c in ctn.iter(f"{P}cmd"):
            c.set("cmd", e["cmd"])
    # 지속 — 템플릿 안쪽 cTn 의 dur 를 비율로 늘이고 줄인다(dur=1 은 '보이기' 스위치라 그대로)
    inner = [c for c in ctn.iter(f"{P}cTn") if c is not ctn]
    base = max([int(c.get("dur")) for c in inner
                if (c.get("dur") or "").isdigit() and c.get("dur") != "1"] or [0])
    for c in inner:
        c.set("id", str(next(ids)))
        dv = c.get("dur")
        if base and dv and dv.isdigit() and dv != "1":
            c.set("dur", str(max(1, round(int(dv) * e["dur"] / base))))
        if e.get("speed") is not None and c.get("dur") not in (None, "1"):
            if float(e["speed"]) == 1:
                c.attrib.pop("spd", None)
            else:
                c.set("spd", str(int(round(float(e["speed"]) * 100000))))
        if base:
            for cond in c.findall(f"{P}stCondLst/{P}cond"):
                dl = cond.get("delay")
                if dl and dl.isdigit() and int(dl) > 0:
                    cond.set("delay", str(round(int(dl) * e["dur"] / base)))
    # 이 효과 끝난 뒤 숨기기(subTnLst) — 기본은 떼고, after:"hide" 일 때만 남긴다
    sub = ctn.find(f"{P}subTnLst")
    if sub is not None:
        if e.get("after") == "hide":
            for tn in sub.iter(f"{P}tn"):
                tn.set("val", str(my_id))
        else:
            ctn.remove(sub)
    # 글자 단위 — 나타내기 단어 10ms·페이드 단어 10% 실물 형식
    it = ctn.find(f"{P}iterate")
    if it is not None:
        ctn.remove(it)
    if e["by"] in ("word", "letter"):
        typ = "wd" if e["by"] == "word" else "lt"
        if e["dur"] == 0:
            step = int(round(float(e.get("step") or 0.01) * 1000))
            tm = f'<p:tmAbs xmlns:p="{NS_P}" val="{step}"/>'
        else:
            tm = f'<p:tmPct xmlns:p="{NS_P}" val="{int(round(float(e.get("step") or 0.1) * 100000))}"/>'
        new = etree.fromstring(f'<p:iterate xmlns:p="{NS_P}" type="{typ}">{tm}</p:iterate>')
        ctn.find(f"{P}childTnLst").addprevious(new)
    # 색 강조 — 템플릿의 목표색만 바꾼다
    if e.get("color") and key in ("emph-1-2", "emph-3-2"):
        for v in ctn.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr"):
            v.set("val", str(e["color"]).lstrip("#").upper())
    if key == "emph-8-0" and e.get("deg") is not None:   # 회전 각도(도) — 실물 by=21600000(360°)
        ctn.find(f".//{P}animRot").set("by", str(int(round(float(e["deg"]) * 60000))))
    if e.get("chart_el") is not None:                     # 그래프 한 계열·항목만 — 실물 graphicEl
        si, ci, step = e["chart_el"]
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        for t in ctn.iter(f"{P}spTgt"):
            for old in list(t):
                t.remove(old)
            ge = etree.SubElement(t, f"{P}graphicEl")
            etree.SubElement(ge, f"{{{A_NS}}}chart", seriesIdx=str(si), categoryIdx=str(ci), bldStep=step)
    if e.get("scale") and key == "emph-6-0":
        by = ctn.find(f".//{P}animScale/{P}by")
        s = int(round(float(e["scale"]) * 100000))
        by.set("x", str(s))
        by.set("y", str(s))
    return etree.tostring(el, encoding="unicode")


def _groups_xml(flat, ctr, interactive=False) -> list[str]:
    """효과 목록 → 클릭 그룹 par 들. interactive 면 개체 클릭 시퀀스 안쪽(첫 그룹 cond delay 0)."""
    groups = []   # [{"auto":bool, "steps":[[eff…]]}]
    for e in flat:
        if e["start"] == "click" or not groups:
            groups.append({"auto": e["start"] != "click" and not interactive, "steps": [[e]]})
        elif e["start"] == "with":
            groups[-1]["steps"][-1].append(e)
        else:
            groups[-1]["steps"].append([e])
    out = []
    for gi, g in enumerate(groups):
        gid = next(ctr)
        if interactive:
            cond = '<p:cond delay="0"/>' if gi == 0 else '<p:cond delay="indefinite"/>'
        else:
            cond = '<p:cond delay="indefinite"/>'
            if gi == 0 and g["auto"]:
                cond += '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>'
        steps_xml, t0 = [], 0
        for si, step in enumerate(g["steps"]):
            sid = next(ctr)
            eff_xml, end = [], 0
            for ei, e in enumerate(step):
                if si == 0 and ei == 0:
                    e["node"] = "clickEffect" if (not g["auto"]) else (
                        "afterEffect" if e["start"] == "after" else "withEffect")
                elif ei == 0:
                    e["node"] = "afterEffect"
                else:
                    e["node"] = "withEffect"
                eff_xml.append(_effect_xml(e, ctr))
                if e.get("repeat") is None:
                    end = max(end, e["delay"] + e["dur"])
            steps_xml.append(
                f'<p:par><p:cTn id="{sid}" fill="hold"><p:stCondLst><p:cond delay="{t0}"/>'
                f'</p:stCondLst><p:childTnLst>{"".join(eff_xml)}</p:childTnLst></p:cTn></p:par>')
            t0 += end
        out.append(f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst>{cond}</p:stCondLst>'
                   f'<p:childTnLst>{"".join(steps_xml)}</p:childTnLst></p:cTn></p:par>')
    return out


def _media_node_xml(m, ctr) -> str:
    """`p:video`/`p:audio` 재생 노드 (음량 80% · 반복 · 멈추면 숨김 display=0)."""
    tag = "audio" if m.get("kind") == "audio" else "video"
    attrs = ""
    if m.get("vol") is not None:
        attrs += f' vol="{int(round(float(m["vol"]) * 100000))}"'
    if m.get("mute"):
        attrs += ' mute="1"'
    if m.get("show_when_stopped") is False:
        attrs += ' showWhenStopped="0"'
    ctn = f'<p:cTn id="{next(ctr)}"'
    if m.get("loop"):
        ctn += ' repeatCount="indefinite"'
    ctn += f' fill="{m.get("fill", "hold")}"'
    if m.get("hide", True):
        ctn += ' display="0"'
    full = ' fullScrn="1"' if m.get("fullscreen") else ""
    return (f'<p:{tag}{full}><p:cMediaNode{attrs}>{ctn}><p:stCondLst><p:cond delay="indefinite"/>'
            f'</p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{m["spid"]}"/></p:tgtEl></p:cMediaNode></p:{tag}>')


def timing_xml(entries, media=None) -> str | None:
    """entries: [(anim, spid, is_sp[, n_paragraphs]), …] (아이템 순서).

    anim 은 dict·문자열 또는 그 목록(한 도형에 등장+강조처럼 효과 여러 개).
    `trigger: "self"`(또는 spid) 인 효과는 mainSeq 가 아니라 그 개체를 클릭하면 도는
    interactiveSeq 로 간다(영상 클릭 = 일시정지/재개). media = 재생 노드 목록.
    """
    media = media or []
    if not entries and not media:
        return None
    effs, idx = [], 0
    for ent in entries:
        a, spid, txt = ent[0], ent[1], ent[2]
        npara = ent[3] if len(ent) > 3 else 1
        nchart = ent[4] if len(ent) > 4 else None       # 그래프 {"series": n, "category": m}
        for one in (a if isinstance(a, list) else [a]):
            n = _norm_anim(one)
            n.update(spid=str(spid), txt=txt, idx=idx, npara=npara, nchart=nchart)
            if n["by"] in CHART_BYS and not nchart:
                raise ValueError(f"anim by {n['by']!r} 는 chart 아이템에만 — 계열·항목 수를 알아야 한다")
            effs.append(n)
            idx += 1
    effs.sort(key=lambda e: (e["order"] if e["order"] is not None else 1e9, e["idx"]))
    # 같은 도형에 효과가 여러 번이면 grpId 0,1,2…(등장 0 · 강조 1)
    seen = {}
    for e in effs:
        g = seen.get(e["spid"], -1) + 1
        seen[e["spid"]] = g
        e["grp"] = g if (e["txt"] or e["by"] in CHART_BYS) else None

    # 단락 단위는 여기서 펼친다 — 첫 단락은 원래 시작, 나머지는 '함께' + 계단 지연
    flat = []
    for e in effs:
        if e["by"] in CHART_BYS:
            # 계열·항목마다 효과 하나(계열=클릭마다 · 항목=이전 후 250ms)
            cnt = int(e["nchart"][e["by"]])
            each = e.get("each") or "after"
            if each not in STARTS:
                raise ValueError(f"anim each {each!r} — 허용: {STARTS}")
            for i in range(cnt):
                el = (i, -4, "series") if e["by"] == "series" else (-4, i, "category")
                f = dict(e, para=None, chart_el=el)
                if i:
                    f.update(start=each, delay=0)
                flat.append(f)
            continue
        if e["by"] == "paragraph" and e["txt"] and e["npara"] > 1:
            stg = int(round(float(e.get("step") or 0.25) * 1000))
            for i in range(e["npara"]):
                f = dict(e, para=i)
                if i:
                    f.update(start="with", delay=e["delay"] + i * stg)
                flat.append(f)
        else:
            flat.append(dict(e, para=None))
    main = [e for e in flat if e.get("trigger") is None]
    inter = {}
    for e in flat:
        if e.get("trigger") is not None:
            t = e["spid"] if e["trigger"] == "self" else str(e["trigger"])
            inter.setdefault(t, []).append(e)
    clicks = sum(1 for k, e in enumerate(main) if e["start"] == "click" and k)
    clicks += 1 if main and main[0]["start"] == "click" else 0
    cap = BOLD_MAX_CLICKS if any(e["bold"] for e in flat) else MAX_CLICKS
    if clicks > cap:
        raise ValueError(f"장당 클릭 {clicks}회 — 절제 원칙상 {cap}회까지. "
                         f"나머지는 start:'after'(이전 후 자동)")

    ctr = iter(range(3, 100000))
    body = "".join(_groups_xml(main, ctr)) if main else ""
    seqs = ""
    if main:
        seqs += (f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
                 f'<p:childTnLst>{body}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" '
                 f'delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                 f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
                 f'</p:cond></p:nextCondLst></p:seq>')
    for m in media:                               # 실물 순서: mainSeq → 재생 노드 → 클릭 시퀀스
        seqs += _media_node_xml(m, ctr)
    for spid, lst in inter.items():
        sid = next(ctr)
        trg = f'<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cond>'
        seqs += (f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="{sid}" restart="whenNotActive" '
                 f'fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst>{trg}'
                 f'</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>'
                 f'<p:childTnLst>{"".join(_groups_xml(lst, ctr, interactive=True))}</p:childTnLst></p:cTn>'
                 f'<p:nextCondLst>{trg}</p:nextCondLst></p:seq>')

    bld, done = [], set()
    for e in flat:
        if e.get("chart_el") is not None and (e["spid"], e["grp"]) not in done:
            done.add((e["spid"], e["grp"]))
            bld.append(f'<p:bldGraphic spid="{e["spid"]}" grpId="{e["grp"]}"><p:bldSub>'
                       f'<a:bldChart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                       f'bld="{e["by"]}" animBg="0"/></p:bldSub></p:bldGraphic>')
            continue
        if not e["txt"] or (e["spid"], e["grp"]) in done:
            continue
        done.add((e["spid"], e["grp"]))
        extra = ' build="p"' if e["para"] is not None else ""
        bld.append(f'<p:bldP spid="{e["spid"]}" grpId="{e["grp"]}"{extra}/>')
    bld = f"<p:bldLst>{''.join(bld)}</p:bldLst>" if bld else ""
    # 효과 조각이 제 xmlns 를 들고 오므로 루트에서 한 번 선언해 두고 그대로 잇는다
    return (f'<p:timing xmlns:p="{NS_P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
            f'restart="never" nodeType="tmRoot"><p:childTnLst>{seqs}</p:childTnLst></p:cTn></p:par></p:tnLst>'
            f'{bld}</p:timing>')


def resolve_triggers(entries, names: dict):
    """anim `trigger: "#이름"` → 그 이름 아이템의 spid. 유튜브 강좌의 트리거(버튼을 누르면 다른 개체가 나온다)."""
    out = []
    for ent in entries:
        a = ent[0]
        lst = a if isinstance(a, list) else [a]
        new = []
        for one in lst:
            if isinstance(one, dict) and str(one.get("trigger", "")).startswith("#"):
                nm = str(one["trigger"])[1:]
                if nm in names:
                    one = {**one, "trigger": str(names[nm])}
            new.append(one)
        out.append((new if isinstance(a, list) else new[0], *ent[1:]))
    return out


def _clean_ns(el):
    etree.cleanup_namespaces(el)
    return el


def apply_timing(slide, entries, media=None):
    xml = timing_xml(entries, media)
    if xml is None:
        return
    sld = slide._element
    old = sld.find(f"{P}timing")
    if old is not None:
        sld.remove(old)
    el = _clean_ns(etree.fromstring(xml))
    # 순서: … clrMapOvr → transition(AC) → timing → extLst
    ext = sld.find(f"{P}extLst")
    if ext is not None:
        ext.addprevious(el)
    else:
        sld.append(el)


def is_build_shape(shape) -> bool:
    """bldP·grpId 대상 = p:sp 전부(글자 유무 무관). 그림·그룹·연결선은 없다 — 레퍼런스 3벌
    실측(sp 는 전부 bldP, pic·grpSp·cxnSp 0건)."""
    return etree.QName(shape._element).localname == "sp"


def n_paragraphs(shape) -> int:
    try:
        return max(1, len([p for p in shape.text_frame.paragraphs]))
    except Exception:
        return 1


# ── 레시피 · 자동 모드 ──────────────────────────────────────────────────
@lru_cache(maxsize=None)
def load_recipe(name: str) -> dict:
    f = RECIPES / f"{name}.json"
    if not f.exists():
        have = sorted(p.stem for p in RECIPES.glob("*.json"))
        raise ValueError(f"motion 레시피 {name!r} 없음 — 있는 것: {have}")
    return json.loads(f.read_text(encoding="utf-8"))


def recipe_names() -> list[str]:
    return sorted(p.stem for p in RECIPES.glob("*.json"))


# 레시피 등급 — present(발표용 기본) / show(쇼용, 이름으로 부를 때만).
# 연출이 과한 레시피 이름을 여기에 넣으면 자동 모드가 고르지 않는다.
SHOW_RECIPES: frozenset = frozenset()


def recipe_grade(name: str) -> str:
    return "show" if name in SHOW_RECIPES else "present"


def auto_pool() -> list[str]:
    """이름 없이 맡겼을 때 고를 수 있는 레시피 — 발표용만. 쇼용은 사용자가 이름을 말할 때만."""
    return [n for n in recipe_names() if recipe_grade(n) == "present"]


# 자동 모드 — 실물 덱 쓰임새 통계에서 뽑은 기본 동선.
# 모든 개체 종류에서 페이드 0.5s 가 1위 · '함께'+250ms 계단이 최빈 간격 · 대부분의 덱이
# 전환 2종 이하. 큰 숫자는 확대, 그래프는 왼쪽에서 닦기.
# 모핑 길이는 0.3s, 큰 이동(화면 밖 진입·크기 1.8배 이상·중심 30% 이상 이동)은 0.5s.
AUTO_MORPH_SMALL, AUTO_MORPH_BIG = 0.3, 0.5
AUTO = {
    "name": "auto",
    "transition": {"type": "fade", "dur": 0.5},
    "morph_transition": {"type": "morph", "dur": AUTO_MORPH_SMALL},
    "page": {
        "title": {"effect": "fade-in", "dur": 0.5},
        "body": {"effect": "fade-in", "dur": 0.5, "mode": "with", "stagger": 0.25},
        "roles": {"number": {"effect": "zoom-in", "dur": 0.5},
                  "chart": {"effect": "wipe-left", "dur": 0.6}},
        "conclusion": {"effect": "fade-in", "dur": 0.5, "start": "after"},
        "max_steps": 6,
        "new_only_on_morph": True,
    },
}

_TITLE = {"headline", "label", "serif_label"}
_CARD = {"card", "pill", "chip", "band", "steps", "timetable", "figure"}
_IMAGE = {"image", "mockup", "map", "flag"}
_CHART = {"chart", "bars", "table"}
_DECOR = {"hairline", "arrow", "poly", "shape", "icon"}
_BG = {"glow", "scrim"}


def role_of(item) -> str:
    t = item.get("type")
    w, h = float(item.get("w", 0) or 0), float(item.get("h", 0) or 0)
    if t in _BG or (w >= 0.98 and h >= 0.98):
        return "bg"
    if t in _TITLE:
        return "title"
    if t == "sowhat" or (t == "band" and float(item.get("y", 0)) >= 0.72):
        return "conclusion"
    if t == "text":
        st = str(item.get("style", ""))
        if st.startswith("num"):
            return "number"
        if st in ("caption", "meta"):
            return "caption"
        return "text"
    if t in _CARD:
        return "card"
    if t in _IMAGE:
        return "image"
    if t in _CHART:
        return "chart"
    if t in _DECOR:
        return "decor"
    return "text"


def _box(item):
    x, y = float(item.get("x", 0) or 0), float(item.get("y", 0) or 0)
    w, h = float(item.get("w", 0) or 0), float(item.get("h", 0) or 0)
    return x, y, w, h


def _inside(child, parent) -> bool:
    cx, cy, cw, ch = _box(child)
    px, py, pw, ph = _box(parent)
    if pw <= 0 or ph <= 0:
        return False
    ax = cx + (min(cw, pw) * 0.5 if cw else 0)
    ay = cy + (ch * 0.5 if ch else 0.01)
    return px <= ax <= px + pw and py <= ay <= py + ph


def _units(items, roles):
    """아이템 → 동선 단위. group 키는 한 단위, 카드·그림 위에 얹힌 글·장식은 그 카드 단위로."""
    units, by_group, host_of = [], {}, {}
    hosts = [i for i, it in enumerate(items) if roles[i] in ("card", "image", "chart", "conclusion")]
    for i, it in enumerate(items):
        if roles[i] == "bg":
            continue
        g = it.get("group")
        if g is not None:
            if g in by_group:
                by_group[g]["members"].append(i)
                continue
            u = {"lead": i, "members": [i], "role": roles[i] if roles[i] != "decor" else "card",
                 "grouped": True}
            by_group[g] = u
            units.append(u)
            continue
        if roles[i] in ("text", "caption", "decor", "number"):
            host = next((h for h in reversed(hosts) if h < i and _inside(it, items[h])), None)
            if host is not None:
                host_of[i] = host
                continue
        units.append({"lead": i, "members": [i], "role": roles[i], "grouped": False})
    lead_unit = {u["lead"]: u for u in units}
    for i, h in host_of.items():
        u = lead_unit.get(h)
        if u is None:     # 호스트가 그룹 구성원 — 그룹 단위에 붙인다
            u = next((x for x in units if h in x["members"]), None)
        if u is not None:
            u["members"].append(i)
    return units


def _reading_key(items, u):
    x, y, _, _ = _box(items[u["lead"]])
    return (round(y / 0.08), x)


def _anim(base: dict, **kw) -> dict:
    a = {k: v for k, v in base.items() if k not in ("mode", "stagger")}
    a.update({k: v for k, v in kw.items() if v is not None})
    return a


def plan_page(spec: dict, items: list, ctx: dict | None = None):
    """페이지의 motion(레시피/auto)을 아이템 anim·페이지 transition 으로 푼다.

    명시한 `anim`·`transition` 은 건드리지 않는다. 반환 (items, transition).
    """
    ctx = ctx or {}
    mode = spec.get("motion", ctx.get("deck_motion"))
    items = [dict(it) for it in items]
    # 개체 레시피 — "motion": "<이름>" 이 아이템에 붙은 경우
    for it in items:
        m = it.get("motion")
        if m and not it.get("anim"):
            r = load_recipe(m) if isinstance(m, str) else m
            if "item" not in r:
                raise ValueError(f"레시피 {r.get('name')} 은 페이지용 — 아이템엔 item 블록이 있는 레시피만")
            it["anim"] = r["item"]
            if r.get("media") and it.get("type") in ("video", "audio") and not it.get("media"):
                it["media"] = r["media"]          # 영상 레시피 — 음량·반복·숨김
    tr = spec.get("transition")
    if not mode or mode == "none":
        return items, tr
    recipe = AUTO if mode == "auto" else (load_recipe(mode) if isinstance(mode, str) else mode)
    shared = bool(ctx.get("shared_morph"))
    if tr is None:
        tr = recipe.get("transition")
        if shared and recipe.get("morph_transition"):
            tr = dict(recipe["morph_transition"])
            if recipe is AUTO and _big_move(items, ctx.get("prev_boxes", {})):
                tr["dur"] = AUTO_MORPH_BIG
    pg = recipe.get("page")
    if not pg:
        return items, tr
    prev_ids = ctx.get("prev_morph_ids", set())
    roles = [role_of(it) for it in items]
    on_morph = isinstance(tr, dict) and tr.get("type") == "morph" or tr == "morph"
    skip = set()
    for i, it in enumerate(items):
        if it.get("anim"):
            skip.add(i)
        if (pg.get("new_only_on_morph") or pg.get("new_only")) and on_morph and \
                it.get("morph_id") and str(it["morph_id"]) in prev_ids:
            skip.add(i)       # 모핑으로 움직이는 개체 — 등장 효과를 겹치지 않는다
    units = [u for u in _units(items, roles) if not (set(u["members"]) & skip)]
    tcfg = pg.get("title")
    ccfg = pg.get("conclusion")
    titles = [u for u in units if u["role"] == "title"] if not pg.get("title_in_body") else []
    concl = [u for u in units if u["role"] == "conclusion"] if ccfg != "chain" else []
    body = [u for u in units if u not in titles and u not in concl]
    body.sort(key=lambda u: _reading_key(items, u))
    if pg.get("rank"):                            # 역할 순서(표식 → 그래프 → 설명)
        rank = pg["rank"]
        body.sort(key=lambda u: rank.get(u["role"], max(rank.values()) + 1))
    if ccfg == "chain":                           # 결론도 계단의 마지막 칸
        body.sort(key=lambda u: u["role"] == "conclusion")

    seq = [0]

    def give(u, anim):
        # 재생 순서 = 계획한 순서(z 순서가 아니다) — order 를 박아 timing_xml 이 이 순서로 줄 세운다
        seq[0] += 1
        o = seq[0] * 10
        if isinstance(anim, list):
            anim = [{**a, "order": o} for a in anim]
        else:
            anim = {**anim, "order": o}
        lead = items[u["lead"]]
        first_eff = anim[0] if isinstance(anim, list) else anim
        nonfade = first_eff.get("effect") not in ("fade-in", "appear", "fade-out", "disappear")
        if len(u["members"]) > 1 and not u["grouped"] and nonfade and \
                not any(items[m].get("morph_id") for m in u["members"]):
            gname = f"_mo{u['lead']}"             # 조각마다 따로 커지지 않게 한 덩어리로
            for m in u["members"]:
                items[m]["group"] = gname
            u["grouped"] = True
        lead["anim"] = anim
        if u["grouped"]:
            return
        for m in u["members"][1:]:                # 얹힌 글·장식 — 같은 때 함께
            items[m]["anim"] = {**first_eff, "start": "with"}

    def build(cfg, start, delay):
        """cfg 가 목록이면 한 도형에 효과 여러 개 — 둘째부터 '함께' + 제 지연."""
        if isinstance(cfg, list):
            out = [_anim(cfg[0], start=start, delay=round(delay + cfg[0].get("delay", 0), 3))]
            for x in cfg[1:]:
                out.append(_anim(x, start="with", delay=round(delay + x.get("delay", 0), 3)))
            return out
        return _anim(cfg, start=start, delay=round(delay, 3))

    first = True
    if tcfg:
        t0 = tcfg[0] if isinstance(tcfg, list) else tcfg
        tc = tcfg if isinstance(tcfg, list) else {k2: v for k2, v in tcfg.items() if k2 != "delay"}
        for k, u in enumerate(titles):
            st = t0.get("start", "after") if k == 0 else "with"
            dl = (0 if isinstance(tcfg, list) else t0.get("delay", 0)) if k == 0 \
                else t0.get("stagger", 0.25) * k
            give(u, build(tc, st, dl))
            first = False
    bcfg = pg.get("body")
    if bcfg and body:
        mode_b = bcfg.get("mode", "with")
        stg = float(bcfg.get("stagger", 0.25))
        maxs = int(pg.get("max_steps", 8))
        roles_cfg = pg.get("roles", {})
        cycle = bcfg.get("cycle")
        base_delay = float(bcfg.get("delay", 0))
        offset = 1 if (titles and tcfg) else 0      # 제목 다음 칸부터 계단
        step = 0
        for k, u in enumerate(body):
            rc = roles_cfg.get(u["role"], {})
            if isinstance(rc, list):
                cfg = [{**{kk: vv for kk, vv in bcfg.items() if kk not in ("delay", "cycle")},
                        **x} for x in rc]
            else:
                cfg = {kk: vv for kk, vv in {**bcfg, **rc}.items() if kk not in ("delay", "cycle")}
                if cycle:
                    cfg["effect"] = cycle[k % len(cycle)]
            lead_first = first and k == 0
            if u["role"] == "decor" and pg.get("decor") == "with_prev" and k:
                give(u, build(cfg, "with", last_delay if mode_b == "with" else 0))
                continue
            if mode_b == "with":
                st = (bcfg.get("first_start", "after")) if lead_first else "with"
                dl = base_delay if lead_first else base_delay + stg * min(step + offset, maxs)
            elif mode_b == "after":
                st = bcfg.get("first_start", "after") if lead_first else (
                    "after" if step < maxs else "with")
                dl = base_delay if lead_first else stg
            else:  # click
                st = "click" if step < int(pg.get("max_clicks", MAX_CLICKS)) else "after"
                dl = 0
            last_delay = dl
            give(u, build(cfg, st, dl))
            step += 1
            first = False
    if isinstance(ccfg, dict):
        for k, u in enumerate(concl):
            give(u, build({kk: v for kk, v in ccfg.items() if kk != "delay"},
                          ccfg.get("start", "after") if k == 0 else "with", ccfg.get("delay", 0)))
    elif concl and bcfg:                           # 결론 설정이 없으면 본문 끝에 이어 붙인다
        for u in concl:
            give(u, build({k: v for k, v in bcfg.items() if k not in ("delay", "cycle")},
                          "after", 0))
    return items, tr


def _big_move(items, prev_boxes: dict) -> bool:
    """앞 장과 짝인 개체 중 하나라도 크게 움직이나 — 화면 밖에서 들어옴 · 크기 1.8배 이상 · 중심 0.3 이상 이동."""
    for it in items:
        mid = it.get("morph_id") if isinstance(it, dict) else None
        if not mid or str(mid) not in prev_boxes:
            continue
        x0, y0, w0, h0 = prev_boxes[str(mid)]
        x1, y1, w1, h1 = _box(it)
        off = lambda x, y, w, h: x + w <= 0 or x >= 1 or y + h <= 0 or y >= 1     # noqa: E731
        if off(x0, y0, w0, h0) != off(x1, y1, w1, h1):
            return True
        a0, a1 = max(w0 * h0, 1e-6), max(w1 * h1, 1e-6)
        if max(a0, a1) / min(a0, a1) >= 1.8 ** 2:
            return True
        if ((x0 + w0 / 2 - x1 - w1 / 2) ** 2 + (y0 + h0 / 2 - y1 - h1 / 2) ** 2) ** .5 >= 0.3:
            return True
    return False


def morph_boxes(items) -> dict:
    """morph_id → (x,y,w,h) — 다음 장의 자동 모드가 이동량을 잰다."""
    out = {}
    for it in items:
        if isinstance(it, dict) and it.get("morph_id") and "x" in it:
            out[str(it["morph_id"])] = _box(it)
    return out


def morph_ids(items) -> set:
    out = set()
    for it in items:
        if not isinstance(it, dict):          # steps·bars 의 items 는 문자열일 수 있다
            continue
        if it.get("morph_id"):
            out.add(str(it["morph_id"]))
        for k in ("items", "children"):
            if isinstance(it.get(k), list):
                out |= morph_ids(it[k])
    return out


# ── 모핑 짝 검사 ────────────────────────────────────────────────────────
def lint_morph(path) -> list[str]:
    """저장된 pptx 를 다시 읽어 모핑 짝을 검사한다 — 오류 아닌 경고.

    외부 조사(docs/motion-external-research.md §2-②: ppt-master · OfficeCLI morph-check ·
    MS morph-lint)에서 가져온 것 중 레퍼런스와 부딪히지 않는 것만:
    짝 없는 `!!` · 종류 다른 짝(sp↔pic 은 보간 안 되고 교차 페이드) · 차트 `!!`(차트는 모핑
    안 됨, MS 공식) · 모핑 아닌 장의 `!!` 짝 · 모핑 짝 도형에 등장 효과.
    '이동량 ≥5cm' 규칙은 넣지 않았다 — 실물 덱은 2% 만 움직이는 게 연출이다.
    등장 효과 경고도 오류로 올리지 않는다 — 모핑 짝 그래프를 다시 닦아 그린다.
    """
    import zipfile
    out = []
    with zipfile.ZipFile(str(path)) as z:
        names = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                       key=lambda n: int(re.findall(r"\d+", n)[0]))
        prev = None
        for i, n in enumerate(names, 1):
            sld = etree.fromstring(z.read(n))
            bang = {}
            for nv in sld.iter(f"{P}cNvPr"):
                nm = nv.get("name") or ""
                if nm.startswith("!!"):
                    tag = etree.QName(nv.getparent().getparent()).localname
                    bang.setdefault(nm, []).append((tag, nv.get("id")))
            is_morph = sld.find(f".//{{{NS_P159}}}morph") is not None
            for nm, lst in bang.items():
                if len(lst) > 1:
                    out.append(f"p{i:02d} 같은 장에 {nm} 이 {len(lst)}개 — 짝이 엇갈린다")
                if lst[0][0] == "graphicFrame":
                    out.append(f"p{i:02d} {nm} 는 차트·표 — 모핑 안 되고 교차 페이드만 된다")
            if prev is not None:
                shared = set(bang) & set(prev)
                if is_morph and bang and not shared:
                    out.append(f"p{i:02d} 모핑 전환인데 앞 장과 짝인 !! 이름이 없다")
                if shared and not is_morph:
                    out.append(f"p{i:02d} 앞 장과 !! 짝 {len(shared)}개인데 전환이 모핑이 아니다")
                for nm in shared:
                    if bang[nm][0][0] != prev[nm][0][0]:
                        out.append(f"p{i:02d} {nm} 종류가 {prev[nm][0][0]}→{bang[nm][0][0]} — "
                                   f"보간 안 되고 교차 페이드로 떨어진다")
                if is_morph and shared:
                    ids = {bang[nm][0][1] for nm in shared}
                    hit = {t.get("spid") for c in sld.iter(f"{P}cTn")
                           if c.get("presetClass") == "entr" for t in c.iter(f"{P}spTgt")}
                    if ids & hit:
                        out.append(f"p{i:02d} 모핑 짝 도형 {len(ids & hit)}개에 등장 효과 — "
                                   f"모핑 도중 사라졌다 다시 나온다(의도면 무시)")
            prev = bang
        out += _lint_tuning(z, names)
    return out


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _bang_boxes(sld):
    """!! 이름 → (z 순서, 종류, 상자 EMU, 자식 수). 최상위 개체만."""
    out = {}
    tree = sld.find(f"{P}cSld/{P}spTree")
    for zi, el in enumerate(tree):
        nv = el.find(f".//{P}cNvPr")
        if nv is None or not (nv.get("name") or "").startswith("!!"):
            continue
        xf = el.find(f".//{{{A_NS}}}xfrm")
        if xf is None:
            xf = el.find(f".//{P}xfrm")
        box = None
        if xf is not None and xf.find(f"{{{A_NS}}}off") is not None:
            o, e = xf.find(f"{{{A_NS}}}off"), xf.find(f"{{{A_NS}}}ext")
            box = (int(o.get("x")), int(o.get("y")), int(e.get("cx")), int(e.get("cy")))
        kids = len([c for c in el if etree.QName(c).localname in ("sp", "pic", "grpSp", "cxnSp", "graphicFrame")])
        out[nv.get("name")] = (zi, etree.QName(el).localname, box, kids)
    return out


def _lint_tuning(z, names) -> list[str]:
    """유튜브 강좌들이 짚은 모핑 함정을 경고로(2026-10-08 · docs/motion-youtube 기법 `morph-tuning`).

    ① 이어지는 모핑 사슬 안에서 지속시간이 다르다(끊겨 보인다 — 한 사슬은 한 길이로)
    ② 화면 밖 `!!` 개체가 앞뒤 장 어디와도 짝이 없다(유령 개체 — 지워도 되고, 남기면 파일만 무겁다)
    ③ 겹친 두 짝의 앞뒤(겹 순서)가 장 사이에 뒤집힌다(모핑 중 겹침이 한순간 튄다)
    ④ 그룹 짝의 구성원 수가 바뀐다(그룹째 교차 페이드 — 그룹을 풀어 개체마다 이름을)
    ⑤ 하이퍼링크로 앞 장에 돌아가는데 그 장에 모핑 전환이 없다(돌아갈 때 모핑 없이 점프)
    """
    out = []
    sw = sh = None
    try:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        sz = pres.find(f"{P}sldSz")
        sw, sh = int(sz.get("cx")), int(sz.get("cy"))
    except Exception:
        pass
    slds = [etree.fromstring(z.read(n)) for n in names]
    boxes = [_bang_boxes(x) for x in slds]
    morph = [x.find(f".//{{{NS_P159}}}morph") is not None for x in slds]
    durs = []
    for x in slds:
        tr = x.find(f".//{{{NS_MC}}}Choice/{P}transition")
        durs.append(tr.get(f"{{{NS_P14}}}dur") if tr is not None else None)
    # ① 사슬 — 모핑 전환이 연달아 붙은 장들
    i = 0
    while i < len(slds):
        if not morph[i]:
            i += 1
            continue
        j = i
        while j + 1 < len(slds) and morph[j + 1]:
            j += 1
        ds = [durs[k] for k in range(i, j + 1)]
        if j > i and len(set(ds)) > 1:
            pretty = "·".join(f"{int(d) / 1000:g}s" if d else "기본" for d in ds)
            out.append(f"p{i + 1:02d}~p{j + 1:02d} 모핑 사슬 지속시간이 {pretty} 로 다르다 — 이어지는 모핑은 한 길이로"
                       f"(달리 쓰면 장마다 속도가 튄다)")
        i = j + 1
    for k, bx in enumerate(boxes):
        nb = set(boxes[k - 1]) if k else set()
        na = set(boxes[k + 1]) if k + 1 < len(boxes) else set()
        for nm, (zi, tag, box, kids) in bx.items():
            # ② 유령 개체
            if box and sw and nm not in nb and nm not in na:
                x, y, w, h = box
                if x >= sw or y >= sh or x + w <= 0 or y + h <= 0:
                    out.append(f"p{k + 1:02d} {nm} 이 화면 밖에 있는데 앞뒤 장에 짝이 없다 — 유령 개체(지워도 된다)")
            # ④ 그룹 구성 변경
            if k and nm in boxes[k - 1] and tag == "grpSp" and boxes[k - 1][nm][1] == "grpSp" \
                    and boxes[k - 1][nm][3] != kids and morph[k]:
                out.append(f"p{k + 1:02d} 그룹 {nm} 구성원 {boxes[k - 1][nm][3]}→{kids}개 — 그룹째 교차 페이드된다"
                           f"(그룹을 풀어 개체마다 morph_id 를)")
        # ③ 겹 순서
        if k and morph[k]:
            prev = boxes[k - 1]
            shared = [nm for nm in bx if nm in prev and bx[nm][2] and prev[nm][2]]
            for a_i, a in enumerate(shared):
                for b in shared[a_i + 1:]:
                    if (bx[a][0] < bx[b][0]) == (prev[a][0] < prev[b][0]):
                        continue

                    def ov(p_, q_):
                        return p_[0] < q_[0] + q_[2] and q_[0] < p_[0] + p_[2] and \
                            p_[1] < q_[1] + q_[3] and q_[1] < p_[1] + p_[3]
                    if ov(bx[a][2], bx[b][2]) or ov(prev[a][2], prev[b][2]):
                        out.append(f"p{k + 1:02d} 겹친 짝 {a}·{b} 의 앞뒤 순서가 앞 장과 반대 — 모핑 중 겹침이 튄다"
                                   f"(두 장의 아이템 순서를 맞춰라)")
    # ⑤ 돌아가는 링크
    for k, n in enumerate(names):
        rel_n = n.replace("slides/", "slides/_rels/") + ".rels"
        if rel_n not in z.namelist():
            continue
        rels = etree.fromstring(z.read(rel_n))
        tgt = {r.get("Id"): r.get("Target") for r in rels if r.get("Type", "").endswith("/slide")}
        R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        for hl in slds[k].iter(f"{{{A_NS}}}hlinkClick"):
            t = tgt.get(hl.get(R_ID))
            if not t:
                continue
            j = next((q for q, m in enumerate(names) if m.endswith("/" + t.split("/")[-1])), None)
            if j is not None and j < k and set(boxes[j]) & set(boxes[k]) and not morph[j]:
                out.append(f"p{k + 1:02d} → p{j + 1:02d} 로 돌아가는 링크인데 p{j + 1:02d} 에 모핑 전환이 없다 — "
                           f"돌아갈 때 모핑 없이 점프한다(메뉴 장에도 transition morph)")
                break
    return out
