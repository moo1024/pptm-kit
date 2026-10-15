# pptm-kit

AI 에이전트(Codex · Claude Code)가 **편집 가능한 PPTX** 를 만들게 하는 엔진.

에이전트가 장마다 `layout_spec.json`(좌표·글자·도식을 적은 JSON)을 쓰면,
`pptm` 이 그것을 python-pptx 네이티브 개체로 렌더한다. 결과물은 통짜 이미지가 아니라
글상자·도형·차트·표라서 PowerPoint 에서 바로 고칠 수 있다.

```
내용 → (에이전트) 컷 분할·디자인 → pages/pNN/layout_spec.json
     → pptm rendernative → deck.pptx
     → pptm export       → deck.pdf + 미리보기 PNG
     → pptm check        → 넘침·겹침·정렬 기계 검사
```

## 설치

필요한 것: Python 3.11+, LibreOffice(PDF 변환), poppler(`pdftoppm`).

```bash
# Ubuntu / WSL
sudo apt install libreoffice poppler-utils
# macOS
brew install --cask libreoffice && brew install poppler

git clone <이 저장소 주소> pptm-kit && cd pptm-kit
python3 -m venv .venv && source .venv/bin/activate
pip install -e .
pptm doctor
```

**폰트** — 테마가 무료 글꼴 [Pretendard](https://github.com/orioncactus/pretendard)와
[Paperlogy](https://freesentation.blog/paperlogyfont) 를 쓴다. 설치하지 않으면 미리보기에서 다른 글꼴로 바뀐다.

## 바로 해보기

```bash
pptm rendernative sample
pptm export sample
# projects/sample/output/ 에 deck.pptx · deck.pdf · preview/*.png
```

## 에이전트에게 시키기

이 폴더에서 Codex(또는 Claude Code)를 켜면 `AGENTS.md` 를 읽고 그 순서대로 일한다. 예:

```
AGENTS.md 읽고 PPT 만들어줘.

- 용도/보는 사람: 동아리 정기총회, 부원 30명, 발표 5분
- 핵심 한 문장: 올해 예산은 행사 수를 줄이고 한 번을 크게
- 분량: 8장 내외
- 내용 원본: input/총회자료.md
- 실제 자료: assets/ 폴더의 작년 행사 사진 6장
- 원하는 느낌: 밝고 깔끔하게, 다크 배경 싫음
```

## 구조

```
AGENTS.md            에이전트 작업 지침 (Claude Code 는 CLAUDE.md 로 복사해 쓰면 된다)
docs/SPEC.md         layout_spec 아이템 문법
pptm/
  cli.py             pptm 명령
  design/gyoan.py    렌더러 본체
  design/figures.py  도식 골격 (timeline·funnel·grid·quadrant …)
  design/motion.py   전환·애니메이션·모핑
  design/themes/     테마 JSON
  design/icons/      Lucide·Tabler 아이콘
  stages/export.py   PDF·미리보기 내보내기
tools/check_pptx.py  납품물 기계 검사
projects/<slug>/     덱마다 폴더 하나 (sample 만 들어 있다)
```

## 라이선스 메모

아이콘 데이터는 [Lucide](https://lucide.dev)(ISC)와 [Tabler Icons](https://tabler.io/icons)(MIT)에서 왔다.
