# 다른 PC / 다른 템플릿으로 옮기기 (PORTING)

이 플러그인을 다른 PC에서, 다른 PPT 템플릿·다른 리서치 자료·다른 주제로 쓰는 방법.

## 먼저 알아둘 것: 무엇이 바뀌고 무엇이 그대로인가

| 구성요소 | 새 환경에서 | 이유 |
|---|---|---|
| **리서치 자료 폴더** | `RESEARCH_DIR` 환경변수만 새 경로로 | 경로는 코드에 없음. 자료가 바뀌어도 코드 변경 0 |
| **보고서 주제/방향** | `/research-report [새 주제]` 인자만 | 주제는 실행 시 전달. 코드 변경 0 |
| **에이전트 4종 · 커맨드 · 웹검색** | 그대로 복사 | 템플릿·자료에 독립적 |
| **PPT 템플릿** | ⚠️ **교체 + 좌표 재조정 필요** | `build_deck.py` 의 좌표/폰트/레이아웃 이름이 *현재* 템플릿에서 추출된 값이라 새 템플릿에 맞춰야 함 |

즉 **유일하게 손이 가는 부분은 PPT 템플릿**이다. 나머지는 복사 + 환경변수면 끝.

---

## A. 새 PC에 옮기고 설치하는 절차

1. **플러그인 폴더 전체를 복사**한다 (USB / git / 압축).
   - `sample-research/` 의 PDF는 테스트용이므로 빼도 된다. (용량 큼)
   - `report.pptx`, `deck_spec.json` 도 결과물이라 빼도 된다.

2. **사전 준비물 설치** (새 PC):
   ```bash
   npm install -g @anthropic-ai/claude-code     # Claude Code
   pip3 install python-pptx PyMuPDF             # deck 생성 + 미리보기
   # (선택) 미리보기 렌더용 LibreOffice
   #   mac:    brew install --cask libreoffice
   #   win:    https://www.libreoffice.org 에서 설치
   ```

3. **리서치 자료 폴더 경로 지정**:
   ```bash
   export RESEARCH_DIR="/새 PC/내 리서치자료 폴더"     # mac/linux
   # Windows(PowerShell):  setx RESEARCH_DIR "C:\research"
   ```

4. **플러그인 설치** (Claude Code 안에서):
   ```
   /plugin marketplace add /복사한/경로/research-ppt-plugin
   /plugin install research-ppt-plugin
   ```

5. **실행**:
   ```
   /research-report 새 보고서 주제
   ```

> 템플릿을 아직 안 바꿨다면 결과 deck은 *이전(현대엔지비) 디자인*으로 나온다.
> 디자인을 새 것으로 바꾸려면 아래 B를 먼저 한다.

---

## B. PPT 템플릿 교체 절차 (가장 중요)

새 템플릿을 `skills/ppt-template/template.pptx` 로 넣는 것만으로는 부족하다.
레이아웃 이름·좌표·폰트가 다르므로 `build_deck.py` 를 새 템플릿에 맞춰야 한다.

### B-1. 새 템플릿 넣기
```bash
cp "새_사내템플릿.pptx" skills/ppt-template/template.pptx
```

### B-2. 새 템플릿 분석 (자동)
```bash
python3 skills/ppt-template/inspect_template.py skills/ppt-template/template.pptx
```
출력에서 **레이아웃 이름 목록**을 확인한다. 예를 들어 새 템플릿이
`Title / Section / Contents / Body / Summary / Closing` 처럼 영어 이름이면,
본문 슬라이드 좌표는 아래로 상세 확인:
```bash
# 본문(내지에 해당) 샘플 슬라이드 번호를 골라 좌표/폰트 덤프
python3 skills/ppt-template/inspect_template.py skills/ppt-template/template.pptx --slide 8
```

### B-3. build_deck.py 상단 블록 수정
`build_deck.py` 맨 위 `==== 새 템플릿으로 교체할 때 ====` 블록만 바꾼다:

- **(1) LAYOUTS** — 새 템플릿의 레이아웃 이름으로 매핑
  ```python
  LAYOUTS = {
      "cover":   "Title",        # 새 표지 레이아웃 이름
      "toc":     "Contents",
      "exec":    "Summary",
      "divider": "Section",
      "body":    "Body",
      "ending":  "Closing",
  }
  ```
  (해당 역할의 레이아웃이 없으면, 가장 가까운 레이아웃 이름을 넣는다.)

- **(2) 폰트** `F_TITLE/F_HEAD/F_BODY/F_SRC/F_DIV` — 새 사내 폰트 이름으로
- **(3) 색상** `NAVY/DARK/...` — 새 브랜드 컬러로

### B-4. 좌표 맞추기 (눈으로 확인하며)
요소 위치(head message, 제목, 출처 등)는 `build_*` 함수 안의 `Inches(...)` 값이다.
B-2 에서 본 새 템플릿 본문 슬라이드의 좌표를 참고해 조정한다. 특히:
- `build_body()` 의 head message 좌표 `(0.34, 1.84)` 와 본문영역 `body_top=3.05`
- `build_cover()` 의 제목/작성자 좌표
- `build_divider()` 의 번호/제목 좌표·색상 (밝은 배경이면 어두운 글씨, 어두운 배경이면 흰 글씨)

### B-5. 렌더해서 검증 (반복)
```bash
# 테스트 spec 하나로 빌드 후 페이지별 PNG 렌더
python3 skills/ppt-template/build_deck.py deck_spec.json -o test.pptx
bash skills/ppt-template/render_preview.sh test.pptx ./preview
# ./preview/slide01.png ... 를 열어 글자 겹침/위치를 눈으로 확인하고 B-4 로 돌아가 미세조정
```
겹침·줄바꿈·빈 공간이 없을 때까지 **B-4 ↔ B-5 를 반복**한다.
(이번 프로젝트에서도 이 렌더-확인 루프로 배지 줄바꿈·로고 겹침·간지 흰글씨를 잡았다.)

### B-6. (선택) 새 레이아웃 구조가 많이 다르면
새 템플릿에 표 슬라이드·차트 슬라이드 등 고유 레이아웃이 있으면 `build_body` 처럼
`build_xxx()` 함수를 추가하고 `build()` 에서 호출, `spec.schema.json`·`SKILL.md` 에
새 spec 키를 문서화한다.

---

## C. 빠른 체크리스트

- [ ] 플러그인 폴더 복사 (sample-research 제외 가능)
- [ ] python-pptx / PyMuPDF (필요시 LibreOffice) 설치
- [ ] `RESEARCH_DIR` 새 경로로 설정
- [ ] 새 `template.pptx` 넣기
- [ ] `inspect_template.py` 로 레이아웃 이름/좌표 확인
- [ ] `build_deck.py` 상단 LAYOUTS/폰트/색상 수정
- [ ] `render_preview.sh` 로 렌더 → 좌표 미세조정 (반복)
- [ ] `/plugin install` → `/research-report [주제]` 실행
