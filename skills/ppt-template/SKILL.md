---
name: ppt-template
description: Use when generating the final PPTX research report deck. Loads the company background template (template.pptx = 보고서양식/POSCO IH 가로형) and enforces head-message placement, slide layouts (표지/들어가며/목차/간지/본문/엔딩), and runs build_deck.py to produce the deck. Trigger whenever a user or the ppt-writer agent needs to turn synthesized research into a company-formatted .pptx.
---

# 사내 PPT 작성 규칙 (보고서양식 / POSCO IH)

이 스킬은 **정해진 사내 템플릿(`보고서양식.pptx`)을 베이스로** 보고서 deck을 만든다.
자유 디자인을 새로 그리지 않는다. 모든 슬라이드는 `template.pptx` 의 레이아웃을 상속한다.

## 절대 규칙

1. **빈 deck에서 새로 만들지 말 것.** 반드시 `build_deck.py` 가 `template.pptx`
   를 열어 레이아웃(배경 디자인 포함)을 상속한 뒤 정해진 좌표에 텍스트만 올린다.
2. deck 생성은 직접 python-pptx 코드를 새로 짜지 말고 **`build_deck.py` 를 실행**한다.
   콘텐츠는 JSON spec 파일로 전달한다 (`spec.schema.json` 참조).
3. 폰트는 사내 폰트(`HY견고딕`, `맑은 고딕`, 서문은 `바탕`)를 사용한다.
   스크립트에 이미 박혀 있으니 건드리지 않는다.
4. **본문은 항상 '본문1' 레이아웃을 베이스로** 한다. 본문2/3/4 레이아웃에는 빨간
   가이드 박스가 배경에 박혀 있어 쓰면 안 된다. 단/표/이미지 배치는 코드가 자동 구성한다.

## 슬라이드 구성 순서 (고정)

| 순서 | 레이아웃 | 용도 | spec 키 |
|---|---|---|---|
| 1 | 표지 | 제목/배지/부제/날짜/기관 | `title`, `report_type`, `subtitle`, `date`, `org` |
| 2 | 들어가며 | 발주처 수신·제목 + 서문 | `foreword` (객체) |
| 3 | 간지(목차) | 장 목록 | `toc` (없으면 sections에서 자동 추출) |
| 4+ | 간지 | 장 구분(섹션 표지) | sections[].`divider` |
| 4+ | 본문1 | 본문 | sections[] 본문 객체 |
| 끝 | 간지(엔딩) | 마무리("감사합니다.") | `closing` (선택) |

> 신규 양식은 목차·간지·엔딩이 모두 **'간지' 레이아웃을 재사용**한다. 이전 버전의
> Executive Summary는 이 양식의 **'들어가며'(서문)** 으로 대체되었다.

## 표지 / 들어가며

- **표지**: `report_type`(배지, 기본 "완료보고서") · `title`(42pt 네이비) · `subtitle`(선택)
  · `date` · `org`(기본 "POSCO IH"). 모두 가운데 정렬, 좌표 고정.
- **들어가며(서문)**: `foreword.recipient`(수신), `foreword.title`(제목),
  `foreword.body`(서문 본문 — 문자열이면 `\n` 으로 단락 분리). 발주 용역 보고서의 격식 있는 서문.

## 본문 슬라이드 작성 규칙 — 가장 중요

각 본문 슬라이드는 **head message 한 줄이 핵심**이다. 본문을 요약한 "결론형 문장"을
head에 넣는다 (제목이 아니라 주장/발견을 담은 완전한 문장).

- `section_title` (권장): 상단 장 라벨 (예: "Ⅱ. 비용 절감 동향") — HY견고딕 20pt
- `head` (필수): 대제목 자리에 들어가는 핵심 한 줄. **가운데 정렬 22pt 네이비 + 밑줄**.
  너무 길면 줄바꿈되니 한 줄(한국어 ~40자 내)로 압축. 예) "자율주행 인식은 카메라
  단독에서 다중센서 융합 구조로 빠르게 이동하고 있음"
- `subtitle`: ■ 소제목 (18pt, 본문 블록 위)
- `body`: 글머리표 리스트(2~5개 권장). 한 슬라이드에 과하게 넣지 말 것.
- `columns`: 2단 비교 `[[좌측 항목들], [우측 항목들]]`. 가운데 구분선 자동.
- `col_titles`: 2단일 때 좌/우 컬럼 소제목(선택).
- `table`: 표. `{"headers": [...], "rows": [[...], ...]}` 형식. 네이비 헤더 자동.
- `image`: 이미지 경로(spec 파일 기준 상대경로 가능). 비율 자동 유지.
- `source`: 하단 출처 (출처가 있으면 반드시 기입 — 리서치 보고서이므로)

한 슬라이드에 메시지는 하나. body가 6개를 넘으면 슬라이드를 나눈다.

### 콘텐츠 배치 자동 선택 (내용따라)
- `columns` → **2단 비교**(좌/우 글머리표 + 가운데 구분선).
- `body` + `table`(또는 `image`) → **분할**: 본문 좌측(폭 4.61"), 시각자료 우측.
- `table` 또는 `image` 단독 → **전체 폭**(0.45", 9.92") 배치.
- `body` 단독 → 전체 폭 글머리표.
- 표는 헤더 행을 주면 네이비 배경 흰 글씨로 강조된다. 행이 많으면(>8행) 슬라이드를 나눈다.

## 실행 방법

```bash
# 1) 콘텐츠 spec(JSON)을 작성한다 (spec.schema.json 형식)
# 2) 스크립트 실행
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt-template/build_deck.py" spec.json -o report.pptx
```

`${CLAUDE_PLUGIN_ROOT}` 는 플러그인 설치 경로로 자동 치환된다. spec.json 경로와
출력 경로(`-o`)는 사용자의 작업 폴더 기준으로 지정한다.

> Windows에서 한글 spec를 읽을 때 인코딩 오류가 나면
> `PYTHONUTF8=1` 환경변수를 주고 실행한다.

## 시각 검수 (선택, 권장)

LibreOffice가 설치돼 있으면 deck을 만든 뒤 페이지별 이미지로 렌더해 **head message
위치·줄바꿈·겹침을 눈으로** 확인할 수 있다. 텍스트 좌표가 어긋났는지 코드로는 잘
안 보이므로 중요한 검수 단계다.

```bash
bash "${CLAUDE_PLUGIN_ROOT}/skills/ppt-template/render_preview.sh" report.pptx
# -> ./preview/slide01.png, slide02.png, ... 생성
```

렌더한 이미지를 Read 도구로 확인해 배지 줄바꿈/텍스트 겹침/빈 슬라이드가 없는지 본다.

## 산출물 검수 체크리스트

- [ ] 표지 제목/날짜/기관이 채워졌는가
- [ ] 들어가며(서문)가 수신·제목·서문 본문을 담고 있는가
- [ ] 모든 본문에 head message가 완전한 문장으로 들어갔는가 (한 줄 분량)
- [ ] 모든 본문에 출처가 달렸는가
- [ ] 슬라이드당 메시지 1개 원칙을 지켰는가
