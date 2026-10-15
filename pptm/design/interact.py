"""장 사이를 잇는 개체 — 하이퍼링크(`link`) · 확대/축소(`zoom`) · 사진 채운 글자(`fill_image`) · 입체 서식(`bevel`).

유튜브 강좌 기법(2026-10-08 `tracks/deck/motion/YOUTUBE.md`)을 피마2 문법으로 옮긴 것. XML 은 손으로 짓지 않았다:
- 링크 `a:hlinkClick` — PowerPoint COM 으로 만든 견본(`templates/interact/samples.json` · shape_links/run_link)
- 확대/축소 `pslz:sldZm`
- 사진 채운 글자 `a:rPr/a:blipFill` — COM 견본(text_picture_fill)
- 입체 서식 `a:scene3d`+`a:sp3d`

링크·확대/축소는 대상 장이 다 만들어진 뒤에야 관계(rId)를 걸 수 있어 `finish_deck` 이 덱 끝에서 한 번에 푼다.
"""
from __future__ import annotations

import copy
import random
import uuid
from pathlib import Path

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_PSLZ = "http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"
NS_P166 = "http://schemas.microsoft.com/office/powerpoint/2016/6/main"
P, A, R = f"{{{NS_P}}}", f"{{{NS_A}}}", f"{{{NS_R}}}"
RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"

# COM 견본 그대로 — 다음·이전·처음·마지막·쇼 끝
SHOW_JUMP = {"next": "nextslide", "prev": "previousslide", "first": "firstslide", "last": "lastslide",
             "end": "endshow", "back": "lastslideviewed"}


def _link_target(link):
    """`link` 값 → ("jump", 이름) | ("page", n) | ("url", 주소)."""
    if isinstance(link, int):
        return "page", link
    if isinstance(link, str):
        if link in SHOW_JUMP:
            return "jump", SHOW_JUMP[link]
        if link.isdigit():
            return "page", int(link)
        if link.startswith(("http://", "https://", "mailto:")):
            return "url", link
        raise ValueError(f"link {link!r} — 장 번호(1부터) · {tuple(SHOW_JUMP)} · URL")
    if isinstance(link, dict):
        if "page" in link:
            return "page", int(link["page"])
        if "url" in link:
            return "url", link["url"]
    raise ValueError(f"link {link!r} — 장 번호 · {tuple(SHOW_JUMP)} · {{\"page\": n}} · {{\"url\": …}}")


def _cnvpr(el):
    return el.find(f".//{P}cNvPr")


def _put_hlink(nv, rid: str, action: str | None):
    """cNvPr 의 첫 자식으로(CT_NonVisualDrawingProps: hlinkClick → hlinkHover → extLst)."""
    for old in nv.findall(f"{A}hlinkClick"):
        nv.remove(old)
    h = etree.Element(f"{A}hlinkClick")
    h.set(R + "id", rid)
    if action:
        h.set("action", action)
    nv.insert(0, h)


def note_link(state: dict, slide, elements, link):
    """렌더 중 — 링크를 걸 개체를 적어 둔다(대상 장이 아직 없을 수 있다)."""
    kind, val = _link_target(link)
    state.setdefault("links", []).append((slide, list(elements), kind, val))


def text_picture_fill(slide, elements, path: Path):
    """글자를 사진으로 채운다 — COM 견본 `<a:rPr><a:blipFill><a:blip r:embed/>…</a:blipFill></a:rPr>`."""
    _, rid = slide.part.get_or_add_image_part(str(path))
    order = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]
    for el in elements:
        for tag in ("rPr", "endParaRPr"):
            for rp in el.iter(f"{A}{tag}"):
                for f in [c for c in rp if etree.QName(c).localname in order[1:]]:
                    rp.remove(f)
                bf = etree.fromstring(f'<a:blipFill xmlns:a="{NS_A}" xmlns:r="{NS_R}"><a:blip r:embed="{rid}"/>'
                                      f'<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
                pos = 0
                for i, c in enumerate(rp):
                    if etree.QName(c).localname == "ln":
                        pos = i + 1
                rp.insert(pos, bf)


BEVEL_DEFAULT = {"w": 30, "h": 9, "prst": "relaxedInset", "material": "plastic", "light": "contrasting"}


def apply_bevel(elements, spec):
    """입체 서식 (scene3d orthographicFront · lightRig contrasting · sp3d plastic bevelT)."""
    cfg = dict(BEVEL_DEFAULT)
    if isinstance(spec, dict):
        cfg.update(spec)
    for el in elements:
        sp = el.find(f"{P}spPr")
        if sp is None:
            continue
        for old in sp.findall(f"{A}scene3d") + sp.findall(f"{A}sp3d"):
            sp.remove(old)
        sc = etree.fromstring(
            f'<a:scene3d xmlns:a="{NS_A}"><a:camera prst="orthographicFront"/>'
            f'<a:lightRig rig="{cfg["light"]}" dir="t"><a:rot lat="0" lon="0" rev="4200000"/></a:lightRig></a:scene3d>')
        s3 = etree.fromstring(
            f'<a:sp3d xmlns:a="{NS_A}" prstMaterial="{cfg["material"]}"><a:bevelT w="{int(cfg["w"] * 12700)}" '
            f'h="{int(cfg["h"] * 12700)}" prst="{cfg["prst"]}"/></a:sp3d>')
        # spPr 순서: xfrm → 기하 → 채움 → ln → effectLst → scene3d → sp3d → extLst
        ext = sp.find(f"{A}extLst")
        if ext is not None:
            ext.addprevious(sc)
            ext.addprevious(s3)
        else:
            sp.append(sc)
            sp.append(s3)


# ── 확대/축소(Slide Zoom) ───────────────────────────────────────────────
def zoom_placeholder(slide, item, x, y, w, h, assets_root: Path, state: dict):
    """렌더 중 — 썸네일 그림을 먼저 놓고(Fallback 이 될 p:pic), 대상 장 관계는 finish_deck 에서."""
    page = int(item["page"])
    thumb = item.get("thumb")
    if thumb:
        p = Path(thumb)
        png = p if p.is_absolute() else assets_root / thumb
    else:
        png = _thumb_placeholder(assets_root / "assets", page)
    pic = slide.shapes.add_picture(str(png), x, y, w, h)
    pic.name = str(item.get("name") or f"확대/축소 {page}")
    state.setdefault("zooms", []).append({"slide": slide, "pic": pic._element, "page": page,
                                          "return": bool(item.get("return", True)),
                                          "dur": float(item.get("dur", 1.0)), "thumb_given": bool(thumb)})
    return pic


def _thumb_placeholder(assets: Path, page: int) -> Path:
    """썸네일을 안 주면 회색 판 — 실제 장 모습은 PowerPoint 가 열 때 대상 장으로 다시 그린다(검증: COM 영상)."""
    from PIL import Image, ImageDraw
    assets.mkdir(parents=True, exist_ok=True)
    out = assets / f"zoom_thumb_{page:02d}.png"
    if not out.exists():
        im = Image.new("RGB", (480, 270), (236, 236, 236))
        ImageDraw.Draw(im).text((20, 20), f"{page}", fill=(120, 120, 120))
        im.save(out)
    return out


def _creation_id(slide) -> str:
    """대상 장의 p14:creationId — 없으면 붙인다(Zoom 의 cId 가 이 값을 가리킨다, 실물 cId=2446311471)."""
    sld = slide._element
    for c in sld.iter(f"{{{NS_P14}}}creationId"):
        return c.get("val")
    val = str(random.Random(slide.slide_id).randint(10 ** 9, 4 * 10 ** 9))
    # 자리는 p:sld 가 아니라 p:cSld 의 extLst(spTree 뒤) — PowerPoint 가 저장한 견본(com_made.pptx) 꼴.
    # p:sld 끝에 두면 PowerPoint 가 못 읽어 확대/축소가 '끊어진 링크' 아이콘으로 그려진다(2026-10-08 영상 실측)
    csld = sld.find(f"{P}cSld")
    ext_lst = csld.find(f"{P}extLst")
    if ext_lst is None:
        ext_lst = etree.SubElement(csld, f"{P}extLst")
    ext = etree.SubElement(ext_lst, f"{P}ext", uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}")
    # 실물 꼴 그대로 p14 접두어 — lxml 이 지어 붙이는 ns2: 로 쓰면 PowerPoint 가 못 읽어 확대/축소가 '끊어진 링크'가 된다
    etree.SubElement(ext, f"{{{NS_P14}}}creationId", nsmap={"p14": NS_P14}, val=val)
    return val


def _wrap_zoom(z, target, rid_slide):
    pic = z["pic"]
    nv = _cnvpr(pic)
    sid, name = nv.get("id"), nv.get("name")
    blip = pic.find(f".//{A}blip")
    rid_img = blip.get(R + "embed")
    off = pic.find(f"{P}spPr/{A}xfrm/{A}off")
    ext = pic.find(f"{P}spPr/{A}xfrm/{A}ext")
    x, y, cx, cy = off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")
    gid = "{" + str(uuid.UUID(int=random.Random(f"{sid}-{z['page']}").getrandbits(128))).upper() + "}"
    ret = "1" if z["return"] else "0"
    dur = int(round(z["dur"] * 1000))
    # 실물 꼴 그대로 — Choice = graphicFrame(pslz:sldZm), Fallback = 링크 건 그림
    xml = (f'<mc:AlternateContent xmlns:mc="{NS_MC}" xmlns:p="{NS_P}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
           f'<mc:Choice xmlns:pslz="{NS_PSLZ}" Requires="pslz"><p:graphicFrame><p:nvGraphicFramePr>'
           f'<p:cNvPr id="{sid}" name="{name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
           f'</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/>'
           f'<a:ext cx="{cx}" cy="{cy}"/></p:xfrm><a:graphic><a:graphicData uri="{NS_PSLZ}"><pslz:sldZm>'
           f'<pslz:sldZmObj sldId="{target.slide_id}" cId="{_creation_id(target)}"><pslz:zmPr id="{gid}" '
           f'returnToParent="{ret}" transitionDur="{dur}"><p166:blipFill xmlns:p166="{NS_P166}">'
           f'<a:blip r:embed="{rid_img}"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>'
           f'<p166:spPr xmlns:p166="{NS_P166}"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
           f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="3175"><a:solidFill><a:prstClr val="ltGray"/>'
           f'</a:solidFill></a:ln></p166:spPr></pslz:zmPr></pslz:sldZmObj></pslz:sldZm></a:graphicData></a:graphic>'
           f'</p:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent>')
    ac = etree.fromstring(xml)
    fb = ac.find(f"{{{NS_MC}}}Fallback")
    parent = pic.getparent()
    idx = list(parent).index(pic)
    parent.remove(pic)
    _put_hlink(nv, rid_slide, "ppaction://hlinksldjump")
    ln = pic.find(f"{P}spPr/{A}ln")
    if ln is None:
        sp = pic.find(f"{P}spPr")
        sp.append(etree.fromstring(f'<a:ln xmlns:a="{NS_A}" w="3175"><a:solidFill><a:prstClr val="ltGray"/>'
                                   f'</a:solidFill></a:ln>'))
    fb.append(pic)
    parent.insert(idx, ac)


def finish_deck(prs, state: dict):
    """덱 렌더가 끝난 뒤 — 링크·확대/축소의 대상 장 관계를 건다."""
    slides = list(prs.slides)

    def target_of(n, where):
        if not 1 <= n <= len(slides):
            raise ValueError(f"{where} 대상 장 {n} — 덱은 1~{len(slides)}장")
        return slides[n - 1]

    for slide, els, kind, val in state.get("links", []):
        for el in els:
            nv = _cnvpr(el)
            if nv is None:
                continue
            if kind == "jump":
                _put_hlink(nv, "", f"ppaction://hlinkshowjump?jump={val}")
            elif kind == "page":
                rid = slide.part.relate_to(target_of(val, "link").part, RT_SLIDE)
                _put_hlink(nv, rid, "ppaction://hlinksldjump")
            else:
                rid = slide.part.relate_to(val, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                                           is_external=True)
                _put_hlink(nv, rid, None)
    for z in state.get("zooms", []):
        target = target_of(z["page"], "zoom")
        rid = z["slide"].part.relate_to(target.part, RT_SLIDE)
        z["img_part"] = str(z["slide"].part.related_part(
            z["pic"].find(f".//{A}blip").get(R + "embed")).partname).lstrip("/")
        _wrap_zoom(z, target, rid)


def refresh_zoom_thumbs(pptx: Path, state: dict) -> int:
    """썸네일을 안 준 확대/축소 개체 — 저장된 덱에서 대상 장을 그림으로 떠서 그림 파트를 바꿔 끼운다.

    PowerPoint 는 열 때 확대/축소 썸네일을 다시 그리지 않는다(2026-10-08 COM 영상 실측: 회색 판 그대로).
    그래서 렌더 직후 LibreOffice → PDF → pdftoppm 으로 대상 장을 떠서 넣는다(pptm export 와 같은 경로).
    """
    import shutil
    import subprocess
    import tempfile
    import zipfile
    need = {z["img_part"]: z["page"] for z in state.get("zooms", []) if not z["thumb_given"] and z.get("img_part")}
    if not need or shutil.which("soffice") is None or shutil.which("pdftoppm") is None:
        return 0
    from pptm.stages.export import export_pdf
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf = export_pdf(Path(pptx), tmp)
        pngs = {}
        for part, page in need.items():
            stem = tmp / f"z{page}"
            subprocess.run(["pdftoppm", "-png", "-r", "60", "-f", str(page), "-l", str(page), "-singlefile",
                            str(pdf), str(stem)], capture_output=True, timeout=120)
            if stem.with_suffix(".png").exists():
                pngs[part] = stem.with_suffix(".png").read_bytes()
        if not pngs:
            return 0
        src = Path(pptx)
        out = tmp / "re.pptx"
        with zipfile.ZipFile(src) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
            for it in zi.infolist():
                zo.writestr(it, pngs.get(it.filename, zi.read(it.filename)))
        shutil.copy(out, src)
    return len(pngs)
