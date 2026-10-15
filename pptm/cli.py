"""pptm CLI — layout_spec.json → 네이티브 PPTX 렌더 + PDF·미리보기 내보내기."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import typer

app = typer.Typer(help="pptm — layout_spec 을 편집 가능한 PPTX 로 렌더한다")

_ROOT = Path(__file__).parent.parent


def _proj(slug: str) -> Path:
    p = _ROOT / "projects" / slug
    if not p.exists():
        typer.echo(f"프로젝트 '{slug}' 없음. `pptm new {slug}` 먼저 실행하세요.", err=True)
        raise typer.Exit(1)
    return p


# ── 커맨드 ────────────────────────────────────────────────────────────────

@app.command()
def new(
    slug: str = typer.Argument(..., help="프로젝트 식별자 (영문·숫자·하이픈)"),
    source: Optional[Path] = typer.Option(None, help="자료 파일 (.md / .txt)"),
    request: Optional[str] = typer.Option(None, help="제작 요청 문자열"),
):
    """새 프로젝트 폴더를 만듭니다."""
    import datetime as _dt
    import shutil

    proj = _ROOT / "projects" / slug
    (proj / "input").mkdir(parents=True, exist_ok=True)
    (proj / "pages").mkdir(parents=True, exist_ok=True)
    (proj / "assets" / "spots").mkdir(parents=True, exist_ok=True)
    if source:
        shutil.copy(source, proj / "input" / "source.md")
    if request:
        (proj / "input" / "request.md").write_text(request, encoding="utf-8")
    pj = proj / "PROJECT.json"
    if not pj.exists():
        pj.write_text(json.dumps({
            "slug": slug,
            "created_at": _dt.date.today().isoformat(),
        }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    typer.echo(f"✓ 프로젝트 생성: {proj}")
    typer.echo(f"  그다음: {proj}/pages/p01/layout_spec.json 작성 → pptm rendernative {slug}")


@app.command()
def rendernative(slug: str = typer.Argument(...)):
    """layout_spec.json → 네이티브 PPTX 렌더."""
    from pptm.design.gyoan import render_deck
    out = render_deck(_proj(slug))
    typer.echo(f"✓ 네이티브 렌더: {out}")


@app.command()
def export(
    slug: str = typer.Argument(...),
    dpi: int = typer.Option(150, help="썸네일 해상도 (DPI)"),
    force: bool = typer.Option(False, "--force", help="캐시 무시 재생성"),
):
    """deck.pptx → PDF + 썸네일 + 미리보기 PNG (LibreOffice + pdftoppm)."""
    from pptm.stages.export import run_export
    run_export(project_dir=_proj(slug), dpi=dpi, force=force)
    typer.echo(f"✓ export 완료: {_proj(slug) / 'output'}")


@app.command()
def check(slug: str = typer.Argument(...)):
    """완성된 deck.pptx 를 기계 검사합니다 (넘침·겹침·정렬 등)."""
    import subprocess
    import sys
    pptx = _proj(slug) / "output" / "deck.pptx"
    if not pptx.exists():
        typer.echo("deck.pptx 가 없습니다. `pptm rendernative` 먼저.", err=True)
        raise typer.Exit(1)
    r = subprocess.run([sys.executable, str(_ROOT / "tools" / "check_pptx.py"), str(pptx)])
    raise typer.Exit(r.returncode)


@app.command()
def status(slug: str = typer.Argument(...)):
    """프로젝트 진행 상황을 출력합니다."""
    proj = _proj(slug)
    typer.echo(f"\n프로젝트: {slug}")
    typer.echo("─" * 48)

    pages_root = proj / "pages"
    specs = sorted(pages_root.glob("p*/layout_spec.json")) if pages_root.exists() else []
    typer.echo(f"  스펙  : {'✓ ' + str(len(specs)) + '장' if specs else '─ 없음'}")

    spots = list((proj / "assets" / "spots").glob("*.png")) if (proj / "assets" / "spots").exists() else []
    if spots:
        typer.echo(f"  이미지: ✓ {len(spots)}장")

    out_dir = proj / "output"
    pptx, pdf = out_dir / "deck.pptx", out_dir / "deck.pdf"
    thumbs = list((out_dir / "thumbnails").glob("slide-*.jpg")) if (out_dir / "thumbnails").exists() else []

    typer.echo(f"  PPTX : {'✓ ' + str(round(pptx.stat().st_size / 1024)) + ' KB' if pptx.exists() else '─ 미생성'}")
    typer.echo(f"  PDF  : {'✓ ' + str(round(pdf.stat().st_size / 1024)) + ' KB' if pdf.exists() else '─ 미생성'}")
    typer.echo(f"  썸네일: {'✓ ' + str(len(thumbs)) + '장' if thumbs else '─ 미생성'}")
    typer.echo("")


@app.command()
def doctor():
    """환경 점검 (LibreOffice, poppler, 폰트)."""
    import shutil
    import subprocess
    ok = True

    for cmd, why in [("soffice", "PDF·썸네일 내보내기"), ("pdftoppm", "썸네일 변환")]:
        if shutil.which(cmd):
            typer.echo(f"✓ {cmd}")
        else:
            typer.echo(f"✗ {cmd} 없음 — {why} 불가", err=True)
            ok = False

    if shutil.which("fc-list"):
        fams = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True).stdout
        for font in ("Pretendard", "Paperlogy"):
            if font in fams:
                typer.echo(f"✓ 폰트 {font}")
            else:
                typer.echo(f"! 폰트 {font} 없음 — 미리보기에서 다른 글꼴로 대체된다 (README 참고)")

    raise typer.Exit(0 if ok else 1)


@app.command()
def icons(
    query: str = typer.Argument("", help="아이콘 이름 검색어 (부분일치). 비우면 개수만"),
    limit: int = typer.Option(60, help="최대 표시 개수"),
):
    """Lucide 라인 아이콘 이름 검색. layout_spec `icon` 아이템용."""
    from pptm.design import icons as ic

    all_names = ic.names()
    if not query:
        typer.echo(f"총 {len(all_names)}종. 예: pptm icons chart / pptm icons bike")
        return
    q = query.lower()
    hits = [n for n in all_names if q in n]
    if not hits:
        typer.echo(f"'{query}' 일치 없음.")
        raise typer.Exit(1)
    typer.echo(f"'{query}' {len(hits)}종:")
    for n in hits[:limit]:
        typer.echo(f"  {n}")
    if len(hits) > limit:
        typer.echo(f"  … 외 {len(hits) - limit}종 (--limit 로 확대)")


if __name__ == "__main__":
    app()
