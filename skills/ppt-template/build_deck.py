#!/usr/bin/env python3
"""
build_deck.py — 사내 보고서 템플릿(가로형)을 베이스로 PPTX 보고서를 생성한다.

핵심 원칙
  - 절대 빈 deck에서 시작하지 않는다. template.pptx 를 열어 디자인(배경 그래픽이
    들어있는 슬라이드 레이아웃)을 그대로 상속한 뒤, 정해진 좌표에 텍스트박스만 올린다.
  - 좌표/폰트/크기는 실제 템플릿 내지 슬라이드에서 추출한 값을 따른다.

사용법
  python3 build_deck.py spec.json -o output.pptx
  python3 build_deck.py spec.json            # -> output.pptx

spec.json 스키마는 같은 폴더의 SKILL.md 와 spec.schema.json 을 참조.
"""

import argparse
import copy
import json
import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.pptx")

# spec 파일이 있는 폴더 (이미지 상대경로 해석 기준). build()에서 설정.
SPEC_DIR = os.getcwd()


def _resolve(path, sec=None):
    """이미지 경로 해석: 절대경로면 그대로, 상대경로면 spec 폴더 기준."""
    if os.path.isabs(path) or os.path.exists(path):
        return path
    cand = os.path.join(SPEC_DIR, path)
    if not os.path.exists(cand):
        raise FileNotFoundError(f"이미지를 찾을 수 없습니다: {path} (기준: {SPEC_DIR})")
    return cand

# ====================================================================
# 새 템플릿으로 교체할 때는 아래 블록만 수정하면 된다 (PORTING.md 참조).
# inspect_template.py 로 새 template.pptx 의 레이아웃 이름/좌표/폰트를 먼저 확인할 것.
# --------------------------------------------------------------------
# (1) 레이아웃 이름 — 새 템플릿의 슬라이드 레이아웃 이름과 1:1로 맞춘다.
LAYOUTS = {
    "cover":   "표지",
    "toc":     "목차",
    "exec":    "Executive Summary",
    "divider": "간지",
    "body":    "내지",
    "ending":  "엔딩표지",
}

# (2) 폰트 패밀리 — 새 템플릿의 사내 폰트로 바꾼다.
F_TITLE = "Hyundai Sans Text Pro Bold"     # 장 제목 26pt
F_HEAD = "Hyundai Sans Head Pro Medium"    # head message / 소제목
F_BODY = "Hyundai Sans Text Pro Medium"    # 본문 11pt
F_SRC = "Hyundai Sans Head Pro Light"      # 출처 8pt
F_DIV = "Hyundai Sans Head Pro Bold"       # 간지 제목 40pt

# (3) 색상 — 새 템플릿의 브랜드 컬러로 바꾼다.
NAVY = RGBColor(0x00, 0x3C, 0x83)
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE = RGBColor(0x00, 0x6D, 0xB7)
# (4) 각 슬라이드 요소의 좌표(Inches)는 아래 build_* 함수 안에 있다.
#     좌표가 어긋나면 render_preview.sh 로 렌더해 눈으로 보며 조정한다.
# ====================================================================


# ---------------------------------------------------------------- helpers
def layout_by_name(prs, name):
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if lay.name == name:
                return lay
    raise KeyError(f"레이아웃 '{name}' 을(를) 템플릿에서 찾을 수 없습니다.")


def wipe_slides(prs):
    """템플릿에 포함된 샘플 슬라이드를 모두 제거(레이아웃/마스터는 유지).

    슬라이드 ID 목록뿐 아니라 부모 파트의 관계(rId)까지 끊어 슬라이드 파트가
    저장 시 패키지에서 빠지도록 한다. (안 하면 slideN.xml 이름이 충돌함)
    """
    xml_slides = prs.slides._sldIdLst
    for sldId in list(xml_slides):
        prs.part.drop_rel(sldId.rId)
        xml_slides.remove(sldId)


def add_slide(prs, layout_name):
    return prs.slides.add_slide(layout_by_name(prs, layout_name))


def textbox(slide, left, top, width, height, text,
            font=F_BODY, size=11, bold=False, color=GRAY,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.0,
            wrap=True):
    """단일 문자열 텍스트박스. wrap=False면 짧은 라벨이 줄바꿈되지 않음."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    f = run.font
    f.name = font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    return tb


def bullets(slide, left, top, width, height, items,
            font=F_BODY, size=11, color=GRAY, bullet_char="• ",
            bold=False, line_spacing=1.15, space_after=6):
    """여러 줄 글머리표 텍스트박스. items 는 문자열 리스트."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = f"{bullet_char}{item}"
        f = run.font
        f.name = font
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = color
    return tb


def hline(slide, left, top, width, color=NAVY, weight=1.25):
    from pptx.enum.shapes import MSO_CONNECTOR
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    Inches(left), Inches(top),
                                    Inches(left + width), Inches(top))
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def _set_cell(cell, text, size=10, bold=False, color=GRAY,
              fill=None, align=PP_ALIGN.LEFT):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.06)
    cell.margin_right = Inches(0.06)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    f = run.font
    f.name = F_BODY
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color


def add_table(slide, left, top, width, height, headers, rows,
              size=10):
    """네이비 헤더 + 흰 본문의 깔끔한 표. headers 없으면 헤더행 생략."""
    has_header = bool(headers)
    ncols = len(headers) if has_header else (len(rows[0]) if rows else 1)
    nrows = len(rows) + (1 if has_header else 0)
    nrows = max(nrows, 1)
    gframe = slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top),
                                    Inches(width), Inches(height))
    tbl = gframe.table
    # 기본 줄무늬 스타일 끄기 (셀 채움을 직접 제어)
    tbl.first_row = has_header
    tbl.horz_banding = False
    r0 = 0
    if has_header:
        for c, h in enumerate(headers):
            _set_cell(tbl.cell(0, c), h, size=size, bold=True,
                      color=WHITE, fill=NAVY, align=PP_ALIGN.CENTER)
        r0 = 1
    for ri, row in enumerate(rows):
        for c in range(ncols):
            val = row[c] if c < len(row) else ""
            _set_cell(tbl.cell(r0 + ri, c), val, size=size, color=GRAY,
                      fill=WHITE)
    return gframe


def add_image_fit(slide, path, left, top, max_w, max_h):
    """비율 유지하며 (max_w x max_h) 박스에 맞춰 이미지 배치(좌상단 기준)."""
    pic = slide.shapes.add_picture(path, Inches(left), Inches(top),
                                   width=Inches(max_w))
    if pic.height > Inches(max_h):
        ratio = Inches(max_h) / pic.height
        pic.width = int(pic.width * ratio)
        pic.height = int(pic.height * ratio)
    return pic


# ---------------------------------------------------------------- slide builders
def build_cover(prs, spec):
    s = add_slide(prs, LAYOUTS["cover"])
    title = spec.get("title", "제목 없음")
    textbox(s, 0.6, 2.6, 9.6, 1.6, title,
            font=F_TITLE, size=32, bold=True, color=NAVY,
            anchor=MSO_ANCHOR.TOP, line_spacing=1.05)
    sub = spec.get("subtitle")
    if sub:
        textbox(s, 0.62, 4.2, 9.6, 0.6, sub,
                font=F_HEAD, size=15, color=GRAY)
    meta = "   |   ".join(
        x for x in [spec.get("author"), spec.get("date")] if x)
    if meta:
        # 하단 좌측 로고와 겹치지 않도록 우측 정렬
        textbox(s, 5.5, 6.55, 5.0, 0.4, meta,
                font=F_BODY, size=11, color=GRAY, align=PP_ALIGN.RIGHT)


def build_toc(prs, spec):
    items = spec.get("toc") or [sec.get("divider") or sec.get("section_title")
                                for sec in spec.get("sections", [])
                                if sec.get("divider") or sec.get("section_title")]
    if not items:
        return
    s = add_slide(prs, LAYOUTS["toc"])
    textbox(s, 0.6, 0.8, 6, 0.8, "목차", font=F_TITLE, size=30,
            bold=True, color=NAVY)
    top = 1.9
    for i, it in enumerate(items, 1):
        textbox(s, 1.0, top, 1.0, 0.5, f"{i:02d}",
                font=F_TITLE, size=20, bold=True, color=BLUE)
        textbox(s, 1.9, top + 0.03, 8.0, 0.5, it,
                font=F_HEAD, size=16, color=NAVY)
        top += 0.75


def build_exec_summary(prs, spec):
    pts = spec.get("exec_summary")
    if not pts:
        return
    s = add_slide(prs, LAYOUTS["exec"])
    # 'Executive Summary' 타이틀은 레이아웃에 이미 있음. 본문 3블록 영역에 배치.
    top = 1.5
    for i, pt in enumerate(pts[:3], 1):
        textbox(s, 0.34, top, 0.6, 0.5, f"{i:02d}",
                font=F_TITLE, size=18, bold=True, color=BLUE)
        textbox(s, 1.0, top, 9.5, 1.3, pt,
                font=F_BODY, size=12, bold=True, color=GRAY,
                line_spacing=1.2)
        top += 1.9


def build_divider(prs, text):
    """간지(장 구분). 템플릿은 밝은 배경 → 네이비 번호 + 어두운 제목, 좌측 정렬."""
    s = add_slide(prs, LAYOUTS["divider"])
    # "Ⅰ. 제목" 형태면 번호와 제목을 분리해 번호는 네이비로 강조
    num, _, rest = text.partition(".")
    if rest.strip():
        textbox(s, 0.75, 1.87, 1.5, 0.7, f"{num.strip()}.",
                font=F_DIV, size=40, bold=True, color=NAVY, wrap=False)
        hline(s, 1.44, 2.33, 1.75, color=NAVY, weight=1.5)
        textbox(s, 0.75, 2.6, 8.5, 1.35, rest.strip(),
                font=F_DIV, size=40, bold=True, color=DARK,
                anchor=MSO_ANCHOR.TOP, line_spacing=1.05)
    else:
        textbox(s, 0.75, 2.6, 8.5, 1.35, text,
                font=F_DIV, size=40, bold=True, color=DARK)


def build_body(prs, sec):
    """내지(본문) 슬라이드. 템플릿 좌표 그대로."""
    s = add_slide(prs, LAYOUTS["body"])
    section_title = sec.get("section_title", "")
    if section_title:
        textbox(s, 0.2, 0.55, 5.53, 0.5, section_title,
                font=F_TITLE, size=26, color=NAVY)
    # 상단 우측 breadcrumb
    if sec.get("breadcrumb"):
        textbox(s, 7.51, 0.12, 3.07, 0.3, sec["breadcrumb"],
                font=F_BODY, size=10, color=GRAY, align=PP_ALIGN.RIGHT)
    # 번호 배지 + 탭 라벨
    no = sec.get("no")
    if no:
        textbox(s, 0.34, 1.39, 0.5, 0.3, str(no),
                font=F_TITLE, size=16, bold=True, color=NAVY, wrap=False)
    if sec.get("tab"):
        textbox(s, 0.86, 1.39, 3.0, 0.3, sec["tab"],
                font=F_TITLE, size=16, bold=True, color=NAVY, wrap=False)
    # head message — 핵심 한 줄 (필수)
    head = sec.get("head", "")
    if head:
        textbox(s, 0.34, 1.84, 10.23, 0.6, head,
                font=F_HEAD, size=15, color=NAVY, line_spacing=1.1)
    # 소제목 + 밑줄
    body_top = 3.05
    if sec.get("subtitle"):
        textbox(s, 0.6, 2.58, 9.0, 0.3, sec["subtitle"],
                font=F_HEAD, size=13.5, color=GRAY)
        hline(s, 0.34, 2.88, 10.16, color=NAVY, weight=1.25)
    # 콘텐츠 영역: 본문(글머리표) / 표 / 이미지 조합 배치
    body = sec.get("body") or []
    table = sec.get("table")
    image = sec.get("image")
    area_h = 6.8 - body_top  # 출처선 위까지
    if (table or image) and body:
        # 분할 레이아웃: 본문 좌측, 시각자료 우측
        bullets(s, 0.34, body_top, 4.97, area_h, body,
                font=F_BODY, size=11, color=GRAY)
        if table:
            add_table(s, 5.53, body_top, 4.96, min(area_h, 2.92),
                      table.get("headers"), table.get("rows", []))
        elif image:
            add_image_fit(s, _resolve(image, sec), 5.53, body_top, 4.96, area_h)
    elif table:
        add_table(s, 0.34, body_top, 10.16, min(area_h, 3.0),
                  table.get("headers"), table.get("rows", []))
    elif image:
        add_image_fit(s, _resolve(image, sec), 0.34, body_top, 10.16, area_h)
    elif body:
        bullets(s, 0.34, body_top, 10.16, area_h, body,
                font=F_BODY, size=11, color=GRAY)
    # 출처
    if sec.get("source"):
        textbox(s, 0.34, 6.97, 9.0, 0.25, f"출처  I  {sec['source']}",
                font=F_SRC, size=8, color=GRAY)


def build_ending(prs, spec):
    add_slide(prs, LAYOUTS["ending"])


# ---------------------------------------------------------------- main
def build(spec, out_path):
    prs = Presentation(TEMPLATE)
    wipe_slides(prs)

    build_cover(prs, spec)
    build_toc(prs, spec)
    build_exec_summary(prs, spec)

    for sec in spec.get("sections", []):
        if sec.get("divider"):
            build_divider(prs, sec["divider"])
        else:
            build_body(prs, sec)

    build_ending(prs, spec)
    prs.save(out_path)
    return out_path, len(prs.slides._sldIdLst)


def main():
    ap = argparse.ArgumentParser(description="사내 템플릿 기반 PPTX 보고서 생성")
    ap.add_argument("spec", help="deck 콘텐츠 JSON 파일 경로")
    ap.add_argument("-o", "--output", default="output.pptx", help="출력 PPTX 경로")
    args = ap.parse_args()

    if not os.path.exists(TEMPLATE):
        sys.exit(f"[오류] 템플릿을 찾을 수 없습니다: {TEMPLATE}")
    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)

    global SPEC_DIR
    SPEC_DIR = os.path.dirname(os.path.abspath(args.spec))
    out, n = build(spec, args.output)
    print(f"[완료] {n}장 생성 -> {os.path.abspath(out)}")


if __name__ == "__main__":
    main()
