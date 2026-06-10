# research-ppt-plugin

내부 자료 폴더와 웹 검색으로 **자동 리서치**한 뒤, **사내 PPT 템플릿**에 맞춰 보고서
deck(.pptx)을 생성하는 Claude Code 멀티 에이전트 플러그인.

## 무엇을 하나

`/research-report [주제]` 한 번이면:

1. 주제를 하위 질문으로 분해하고
2. **내부 자료 조사**(internal-researcher)와 **웹 조사**(web-researcher)를 병렬 수행
3. 결과를 **종합**(synthesizer)해 결론·목차·섹션 구조를 만들고
4. **사내 템플릿**(`template.pptx`)을 베이스로 **PPTX deck**(ppt-writer + ppt-template 스킬)을 생성

## 구성

```
research-ppt-plugin/
├── .claude-plugin/
│   ├── plugin.json          # 매니페스트
│   └── marketplace.json     # 공유용 마켓플레이스 선언
├── agents/                  # 멀티 에이전트
│   ├── internal-researcher.md
│   ├── web-researcher.md
│   ├── synthesizer.md
│   └── ppt-writer.md
├── skills/
│   └── ppt-template/        # PPT 템플릿 준수 스킬 (핵심)
│       ├── SKILL.md
│       ├── template.pptx    # 사내 배경 템플릿
│       ├── build_deck.py    # python-pptx 생성 스크립트
│       └── spec.schema.json # deck 콘텐츠 JSON 스키마
├── commands/
│   └── research-report.md   # 오케스트레이터 슬래시 커맨드
├── .mcp.json                # 내부 폴더 접근 (filesystem MCP)
└── README.md
```

## 사전 준비

- **Node.js** 와 **Claude Code** 설치 (`npm install -g @anthropic-ai/claude-code`)
- **Python 3 + python-pptx**: `pip3 install python-pptx`
- **리서치 자료 폴더 경로**를 환경변수로 지정:
  ```bash
  export RESEARCH_DIR="/경로/내/리서치자료폴더"
  ```
  (배포 대상마다 폴더가 다르므로 경로는 플러그인에 넣지 않고 환경변수로 둔다.)

## 설치

### 로컬에서 바로
```
/plugin marketplace add /Users/yunsu/Desktop/research-ppt-plugin
/plugin install research-ppt-plugin
```

### GitHub 마켓플레이스로 공유
저장소를 push 한 뒤, 사용하는 사람은:
```
/plugin marketplace add <owner>/<repo>
/plugin install research-ppt-plugin
```
설치 한 번으로 에이전트·스킬·MCP 설정이 모두 구성된다.

## 사용

```
/research-report 산업계 자율주행 최신 기술 동향
```

생성된 deck은 현재 작업 폴더에 `report.pptx` 로 저장된다.

## 템플릿 직접 호출(에이전트 없이)

콘텐츠 JSON만 있으면 스크립트를 직접 돌릴 수도 있다:
```bash
python3 skills/ppt-template/build_deck.py deck_spec.json -o report.pptx
```
`deck_spec.json` 형식은 `skills/ppt-template/spec.schema.json` 참조.

## 템플릿 레이아웃 (추출 결과)

`template.pptx` 는 6종 레이아웃을 제공한다: `표지`, `목차`, `Executive Summary`,
`간지`(장 구분), `내지`(본문), `엔딩표지`. 본문(내지)의 핵심은 상단 **head message**
(0.34", 1.84" 위치, 결론형 한 문장)이다. 자세한 규칙은
`skills/ppt-template/SKILL.md` 참조.

## 다른 PC / 다른 템플릿으로 옮기기

리서치 자료와 주제는 `RESEARCH_DIR` 환경변수와 `/research-report` 인자로만 바뀌므로
코드 변경이 없다. **유일하게 손이 가는 건 PPT 템플릿 교체**이며, 절차는
[`PORTING.md`](./PORTING.md) 에 단계별로 정리돼 있다. 새 템플릿 분석은
`skills/ppt-template/inspect_template.py` 로 자동화돼 있다.

## 보안 메모

`research-files` MCP 서버는 `RESEARCH_DIR` 한 폴더에만 **읽기 범위**로 접근하도록
최소 권한을 권장한다. MCP는 폴더 접근용 하나로 최소화하고, 나머지 기능은 스킬·에이전트로 구성해 초기 컨텍스트 비용을 줄였다.
