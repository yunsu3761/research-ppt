---
name: ppt-template
description: Use when generating the final PPTX research report deck. Loads the company background template (template.pptx) and enforces head-message placement, slide layouts (표지/목차/Executive Summary/간지/내지/엔딩표지), and runs build_deck.py to produce the deck. Trigger whenever a user or the ppt-writer agent needs to turn synthesized research into a company-formatted .pptx.
---

# 사내 PPT 작성 규칙

이 스킬은 **정해진 사내 템플릿을 베이스로** 보고서 deck을 만든다. 자유 디자인을
새로 그리지 않는다. 모든 슬라이드는 `template.pptx` 의 레이아웃을 상속한다.

## 절대 규칙

1. **빈 deck에서 새로 만들지 말 것.** 반드시 `build_deck.py` 가 `template.pptx`
   를 열어 레이아웃(배경 디자인 포함)을 상속한 뒤 정해진 좌표에 텍스트만 올린다.
2. deck 생성은 직접 python-pptx 코드를 새로 짜지 말고 **`build_deck.py` 를 실행**한다.
   콘텐츠는 JSON spec 파일로 전달한다 (`spec.schema.json` 참조).
3. 폰트는 사내 폰트(`Hyundai Sans Text Pro`, `Hyundai Sans Head Pro`)를 사용한다.
   스크립트에 이미 박혀 있으니 건드리지 않는다.

## 슬라이드 구성 순서 (고정)

| 순서 | 레이아웃 | 용도 | spec 키 |
|---|---|---|---|
| 1 | 표지 | 제목/부제/작성자/일자 | `title`, `subtitle`, `author`, `date` |
| 2 | 목차 | 장 목록 | `toc` (없으면 sections에서 자동 추출) |
| 3 | Executive Summary | 핵심 결론 3가지 | `exec_summary` (문자열 3개) |
| 4+ | 간지 | 장 구분(섹션 표지) | sections[].`divider` |
| 4+ | 내지 | 본문 | sections[] 본문 객체 |
| 끝 | 엔딩표지 | 마무리 | (자동) |

## 내지(본문) 슬라이드 작성 규칙 — 가장 중요

각 내지 슬라이드는 **head message 한 줄이 핵심**이다. 본문을 요약한 "결론형 문장"을
head에 넣는다 (제목이 아니라 주장/발견을 담은 완전한 문장).

- `head` (필수): 슬라이드 상단(0.34", 1.84")에 들어가는 핵심 한 줄. 예) "자율주행
  인식은 카메라 단독에서 다중센서 융합 구조로 빠르게 이동하고 있음"
- `section_title`: 좌상단 장 제목 (예: "Ⅱ. 비용 절감 동향")
- `no` + `tab`: 번호 배지 + 소주제 탭 라벨 (예: "01" + "센서")
- `subtitle`: 본문 블록 위 소제목(밑줄 자동)
- `body`: 글머리표 리스트(2~5개 권장). 한 슬라이드에 과하게 넣지 말 것.
- `table`: 표. `{"headers": [...], "rows": [[...], ...]}` 형식. 네이비 헤더 자동.
- `image`: 이미지 경로(spec 파일 기준 상대경로 가능). 비율 자동 유지.
- `source`: 하단 출처 (출처가 있으면 반드시 기입 — 리서치 보고서이므로)
- `breadcrumb`: 우상단 경로 표시(선택)

한 슬라이드에 메시지는 하나. body가 6개를 넘으면 슬라이드를 나눈다.

### 표·이미지 배치 규칙
- `body` + `table`(또는 `image`) → **분할 레이아웃**: 본문 좌측(폭 4.97"), 시각자료 우측(5.53"~).
- `table` 또는 `image` 단독 → **전체 폭**(0.34", 10.16") 배치.
- 표는 헤더 행을 주면 네이비 배경 흰 글씨로 강조된다. 행이 많으면(>8행) 슬라이드를 나눈다.

## 실행 방법

```bash
# 1) 콘텐츠 spec(JSON)을 작성한다 (spec.schema.json 형식)
# 2) 스크립트 실행
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt-template/build_deck.py" spec.json -o report.pptx
```

`${CLAUDE_PLUGIN_ROOT}` 는 플러그인 설치 경로로 자동 치환된다. spec.json 경로와
출력 경로(`-o`)는 사용자의 작업 폴더 기준으로 지정한다.

## 시각 검수 (선택, 권장)

LibreOffice가 설치돼 있으면 deck을 만든 뒤 페이지별 PNG로 렌더해 **head message
위치·줄바꿈·겹침을 눈으로** 확인할 수 있다. 텍스트 좌표가 어긋났는지 코드로는 잘
안 보이므로 중요한 검수 단계다.

```bash
bash "${CLAUDE_PLUGIN_ROOT}/skills/ppt-template/render_preview.sh" report.pptx
# -> ./preview/slide01.png, slide02.png, ... 생성
```

렌더한 이미지를 Read 도구로 확인해 배지 줄바꿈/텍스트 겹침/빈 슬라이드가 없는지 본다.

## 산출물 검수 체크리스트

- [ ] 표지 제목/작성자/일자가 채워졌는가
- [ ] Executive Summary가 결론을 담고 있는가 (배경 나열 X)
- [ ] 모든 내지에 head message가 완전한 문장으로 들어갔는가
- [ ] 모든 본문에 출처가 달렸는가
- [ ] 슬라이드당 메시지 1개 원칙을 지켰는가
