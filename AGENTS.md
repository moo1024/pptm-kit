# pptm-kit — 에이전트 작업 지침

이 저장소는 **편집 가능한 PPTX 를 만드는 엔진**이다. 에이전트(Codex 등)가 내용을 정리해
장마다 `layout_spec.json` 을 쓰면, `pptm` 이 python-pptx 네이티브 개체(글상자·도형·차트·표)로 렌더한다.
결과 pptx 는 PowerPoint 에서 글자 하나까지 고칠 수 있다. 통짜 이미지 슬라이드를 만들지 않는다.

## 작업 순서

1. **브리프부터 받는다.** 만들기 전에 아래가 빠졌으면 묻는다.
   - 용도와 보는 사람 (화면에 띄우는 발표인가, 읽는 문서인가)
   - 핵심 한 문장 · 분량 · 마감
   - 내용 원본 (기획서·메모·데이터)
   - 실제 자료 (로고·사진·화면 캡처) — 없으면 아이콘으로 때우지 말고 요청한다
   - 원하는 느낌 · 피할 것
2. **컷 분할** — 한 장 = 한 메시지. 표지 → 본문 N장 → 마무리. 장마다 제목 한 줄을 먼저 적어 사용자에게 보인다.
3. **프로젝트 생성** — `pptm new <slug>`
4. **1~2장 먼저** — 스펙을 쓰고 렌더해 미리보기를 보여 방향을 확인받은 뒤 나머지를 만든다.
5. **스펙 작성** — `docs/SPEC.md` 문법대로 `projects/<slug>/pages/pNN/layout_spec.json`.
6. **렌더 → 내보내기 → 검사 → 눈으로 확인**
   ```bash
   pptm rendernative <slug>     # → output/deck.pptx
   pptm export <slug>           # → output/deck.pdf, output/preview/slide-NN.png
   pptm check <slug>            # 넘침·겹침·정렬 기계 검사
   ```
   `output/preview/*.png` 를 **한 장도 빠짐없이 직접 열어 본다.** 검사 0건은 통과가 아니라 "기계가 못 찾았다"일 수 있다.
7. **납품** — `deck.pptx` + `deck.pdf` + 미리보기 PNG.

## 디자인 원칙

- **도식이 주인공이다.** 라벨:값 목록과 표만 반복하면 심심하다. 먼저 `figure` 골격(timeline·funnel·grid·quadrant…)을 고른다.
- **악센트는 한 색.** 면을 악센트로 칠하면 강조가 죽는다. 면은 무채·연한 톤, 악센트는 숫자·핵심어 한 곳.
- **화면 글자는 명사구로 끊는다.** 제목·카드·캡션에 `~습니다` 를 쓰지 않는다. (`처리 시간 40% 단축` ○ / `처리 시간을 40% 줄였습니다` ✗)
- **한 장에 같은 모양을 두 번 쓰지 않는다.** 카드 3개 + 또 카드 3개 대신 하나는 차트·표·도식으로.
- **세로 정렬은 계산에 맡긴다.** 면 안 글자·아이콘은 `h` + `anchor: "middle"`. 눈대중 y 좌표는 늘 위로 쏠린다.
- **차트 축은 0부터.** `value_max` 를 주면 `value_min: 0` 도 준다.
- **동그라미 안 숫자 배지 금지.** 순서는 `indexed_card` 나 타임라인으로.
- 다크 지면에서 `figure` 를 쓰면 `fill` 을 꼭 지정한다.

## 이미지

본문은 네이티브 도형이 원칙이다. 사진·일러스트는 표지나 포인트 자리에만 쓴다.

- 실제 사진·로고가 있으면 그것을 `projects/<slug>/assets/` 에 넣어 쓴다(가장 좋음).
- 생성이 필요하면 에이전트의 이미지 생성 기능(Codex 는 `$imagegen`)으로 만들어
  `projects/<slug>/assets/spots/` 에 저장하고 `image` 아이템으로 참조한다.
  프롬프트에 **슬라이드 배경색을 명시**하고(`on solid warm paper #FBF9F5 background`), 이미지 안에 글자를 넣지 말라고 쓴다.
- `image` 는 자르지 않고 박스에 늘린다. 넣기 전에 목표 비율로 잘라 둔다.

## 테마

`pptm/design/themes/` — `warm-ivory`(기본 추천) · `premium-white` · `premium-navy` · `graphite-slate` ·
`charcoal-sage` · `swiss` · `magazine` · `hanji` · `neumorph`. 새 테마는 아무 JSON 을 복사해
색(`bg` `ink` `accent` `card` …)만 바꾸면 된다. 프로젝트 전용이면 `projects/<slug>/themes/` 에 둔다.

## 문서

- `docs/SPEC.md` — 아이템 문법 전체
- `projects/sample/` — 3장짜리 예시 스펙 (`pptm rendernative sample` 로 바로 렌더된다)
