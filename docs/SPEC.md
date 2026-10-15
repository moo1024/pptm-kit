# layout_spec.json 문법

경로: `projects/<slug>/pages/pNN/layout_spec.json` — 한 장에 파일 하나. `pNN` 숫자 순서가 장 순서다.

## 기본 형태

```json
{"page": 1, "theme": "warm-ivory", "variant": "paper", "items": [ ... ]}
```

- 좌표는 전부 **슬라이드 비율 0~1**(16:9 기준). 좌우 여백은 x = 0.05~0.07 이 무난하다.
- `variant: "dark"` = 테마의 `dark_bg` 를 배경으로 쓴다. 표지·간지용.
- `theme` = `pptm/design/themes/<이름>.json` 또는 `projects/<slug>/themes/<이름>.json`.
  일회성 테마는 spec 에 `theme_tokens: {...}` 를 인라인으로 넣는다.
- `page_size` — 생략 시 16:9.
  - `"a4"` = 실측 A4 세로(210×297mm). 인쇄·PDF 배포물.
  - `"3:4"` = 세로 포스터 · `"4:5"` = 인스타 피드 · `"a2p"` = A2 세로 포스터.
  - **pt 는 절대 단위다.** 판형이 바뀌면 글자 크기를 판형 폭에 맞춰 다시 잡는다
    (16:9 폭 960pt → 헤드라인 23pt 안팎).
- `background` (페이지 키) — `{"color": "FFFFFF"}` 단색, 또는
  `{"image": "assets/x.png", "blur": 18, "darken": 0.15, "sharp": true}` 사진 배경.

## 텍스트

- `label` — 좌상단 아이브로우(작은 머리글)
- `headline` — `lines`(줄 배열) 또는 `runs`(리치텍스트), `size`, `spacing`
- `text` — `text` 또는 `runs`, `style: sub|body|cardtitle|caption|meta|num_xl …`,
  `x` `y` `w` (`h` + `anchor: "middle"` 이면 그 높이 안에서 세로 가운데)
- 공통 키 `tracking`(자간, pt). 음수면 조인다(큰 제목 -0.1~-0.3).
- `text` 의 `shadow` `{color, dist, dir, blur, alpha}` — 도형 단위 그림자.

### 리치텍스트

```json
"runs": [[{"t": "앞부분 ", "color": "534C44"}, {"t": "강조", "weight": "bold", "mark": "F1E4C6"}]]
```

바깥 리스트가 줄, 안쪽이 런. `font` 는 테마 `fonts` 의 키(`display_b`·`body`·`serif` 등),
`mark` = 형광펜(런 배경색). 강조는 한 장에 한 곳만.

## 면·선

- `card` — `x y w h`, `fill`, `outline`, `radius`, `rot`, `grad`[위색, 아래색], `depth`
- `depth` — 입체 프리셋. `card`·`pill`·`band`·`shape` 공통.
  `soft`(본문 카드 기본) · `ambient`(큰 패널) · `float`(칩·말풍선) · `sticker`(컷아웃) ·
  `neon`(컬러 글로우, `depth_color`) · `paper`(눌린 판) · `neu_raised`/`neu_pressed`(테마 `neumorph` 와 함께) ·
  `glass`(유리 — `fill_alpha`, 배경 비춤은 `bg_fill: true`). **오타는 KeyError 로 멈춘다.**
- `pill` — 자동 폭 알약. `fill`, `text_color`
- `chip` — 원형
- `hairline` — 가는 선. `color`, `dash`, `thick`
- `arrow`
- `band` — 채워진 바. 라벨+본문을 자동으로 가운데 정렬. `fill` `label` `text` `align`
- `scrim` — 풀블리드 사진 위 알파 그라데이션. `edge`(bottom/top/left/right) `h|w` `color` `strength`
- `sowhat` — 하단 결론 바
- `shape` — `shape: hexagon|parallelogram|pie|arc_ring|chevron|pentagon_arrow|oval|triangle|diamond|rounded|rect|star5|heart|gear|cloud|donut …`,
  `fill`("none" 가능) `outline` `adj` `rot` 내부 `text`

## 도식 — `figure`

손으로 좌표를 찍기 전에 먼저 본다. **골격 × 스킨** 을 고르면 렌더러가 좌표를 계산한다.
전개 결과는 card/text/chip/arrow/hairline 이라 PowerPoint 에서 그대로 편집된다.

```json
{"type": "figure", "figure": "timeline", "item": "compact_card",
 "x": 0.07, "y": 0.28, "w": 0.86, "h": 0.36,
 "nodes": [{"title": "진단", "value": "2주", "caption": "지금 무엇을 얼마나 하는지 센다"},
           {"title": "설계", "value": "3주", "caption": "바꿀 지점 세 곳을 고른다"}]}
```

- 골격: `grid` `timeline` `steps_up` `waterfall` `funnel` `pyramid` `zigzag` `roadmap` `snake`
  `quadrant`(축 라벨 `axes`) `circular` `hub` `tree` `gallery`(이미지 격자, `feature` 로 대표 1장 크게)
- 스킨: `plain` `compact_card` `outline_card` `indexed_card` `underline_text` `pill_badge` `stat`(라벨→큰 숫자→룰→캡션)
- 노드: `title` / `value` / `caption` / `icon`(Lucide 이름) / `image` / `fill`
- 옵션: `num_size`(stat 숫자 pt 고정) `colors` `fills` `accent` `gap` `radius` `fill` `align` `cols`
- **다크 지면에서는 `fill`(또는 `colors`)을 꼭 준다.** 없으면 렌더가 멈춘다.
- 동그라미 안 숫자 배지(`badge_card`·`badge`)는 막혀 있다. 번호가 필요하면 `indexed_card`.
- `h` 는 내용이 정한다. 너무 크게 주면 위아래가 빈다.

## 아이콘 · 이미지

- `icon` — Lucide 라인 아이콘(약 2,000종)을 테마색으로 칠해 PNG 로 넣는다.
  `name`(예 `bike` `gauge` `map-pin` `database`) `x` `y` `size` `color` `weight`.
  검색: `pptm icons <검색어>`. **이모지 대신 이걸 쓴다.**
  - 라벨 붙은 아이콘: `{"type":"icon","name":"lightbulb","x":0.055,"y":0.687,"text":"진단 관점","text_size":10}` —
    아이콘 크기·세로중심·라벨 x 를 렌더러가 계산한다. 아이콘과 글자를 따로 찍지 않는다.
  - 카드 안 아이콘 행: icon 에 카드의 `y` 와 `h`, `anchor: "middle"` 을 준다.
- `image` — `path`(프로젝트 기준 상대경로) `x y w h`, `rot`, `crop` {l,t,r,b}, `flip_h`.
  **이미지는 자르지 않고 박스에 늘린다** — 목표 종횡비(`w×960 / (h×540)`)로 미리 잘라 둔다.

## 네이티브 개체 (편집 가능)

- `chart` — 엑셀 데이터 내장. `kind`(bar/column/line/line_markers/pie/doughnut) `categories`
  `series`[{name, values, color}] `value_labels` `label_format` `legend` `gridlines` `value_max`
  - **`value_max` 를 주면 `value_min: 0` 도 같이 준다.** 안 주면 축이 잘려 막대 길이가 거짓말을 한다.
- `table` — `rows`[[…]] `header` `col_w` `align` `header_fill` `body_fill` `alt_fill` `border_color`
  `size` `row_h` `row_hs` `pad_x` `pad_y` `line_spacing`, 강조 `row_fills` `col_fills` `row_bold`,
  병합 `merges` `[[행, 시작열, 끝열]]`. 셀 값에 `\n` 을 넣으면 여러 줄.
  `row_h` 는 최솟값이다 — 실제 행은 셀 여백+줄높이만큼 커진다.

## 모션 (PowerPoint 슬라이드쇼에서만 보인다)

```json
{"page": 5, "transition": {"type": "morph", "dur": 0.8}, "items": [
  {"type": "text", "text": "STEP 2", "x": 0.6, "y": 0.7, "w": 0.2,
   "anim": {"effect": "rise-in", "start": "after", "dur": 0.6}}]}
```

- `transition`(페이지) — `"morph"` · `"fade"` · `"none"`, `dur` 초, `advance_after` 초(자동 넘김).
  강한 전환(`push` `wipe` `pan` `peel` `curtains` …)은 `"bold": true` 일 때만.
- `anim`(아이템) — `effect`: `fade-in` `rise-in` `zoom-in` `appear` `wipe-left|right|up|down` `fade-out` `disappear`,
  `start`: `click` · `after` · `with`, `dur`, `delay`, `by`: `word`·`letter`·`paragraph`.
  강한 효과(`fly-in-*` `zoom-pop` `spin` `pulse` …)는 `"bold": true` 일 때만.
- `morph_id`(아이템) — 앞뒤 장에서 같은 id 끼리 모핑 짝이 된다.
- `motion: "auto"`(페이지 키 또는 `PROJECT.json`) — 페이드 위주의 무난한 기본 동선을 자동으로 건다.
- `link`(아이템) — 장 번호 · `"next"` · `"prev"` · `{"url": …}`. 목차·돌아가기 버튼.

## 자주 쓰는 배치

- **3카드 가로**: x = 0.07 / 0.37 / 0.67, w = 0.26
- **표지**: label(y0.09) → headline(y0.2~0.33, 28~35pt) → sub → pill → 하단 meta
- **하단 결론 바**: `sowhat` 또는 `band` 를 y ≈ 0.88 에 하나만

## 이 키트에 없는 것

기기 목업(`mockup`), 지도(`map`), 국기(`flag`), 브랜드 로고(`icon` + `brand`)는 빠져 있다.
쓰면 "설치 필요" 오류가 난다. 필요하면 각자 자산을 구해 넣는다.
