#!/usr/bin/env python3
"""extract_figure.py — 원문 PDF에서 그림/도표를 잘라 PNG로 저장한다.

deck에 '원문 참고 이미지'를 넣을 때 사용한다. 페이지 전체를 렌더하거나,
페이지 내 영역(분수 좌표)을 잘라낸다. 워터마크·이미지형 PDF에도 동작한다.

사용법
  # 페이지 전체를 PNG로 (1-based 페이지 번호)
  python3 extract_figure.py --pdf "report.pdf" --page 5 -o fig.png

  # 페이지 내 영역만 크롭 (x0,y0,x1,y1 = 0~1 분수 좌표)
  python3 extract_figure.py --pdf "report.pdf" --page 5 --box 0.07,0.55,0.96,0.80 -o fig.png

요령
  - 먼저 --page 로 전체를 떠서 그림 위치를 눈으로 확인한 뒤, --box 로 좁혀 다시 뜬다.
  - dpi 는 200 권장(슬라이드에서 선명). 본문 텍스트가 섞이면 box 를 더 좁힌다.
"""
import argparse
import os
import sys

import fitz  # PyMuPDF


def main():
    ap = argparse.ArgumentParser(description="원문 PDF에서 그림 크롭")
    ap.add_argument("--pdf", required=True, help="원본 PDF 경로")
    ap.add_argument("--page", type=int, required=True, help="페이지 번호(1-based)")
    ap.add_argument("--box", help="크롭 영역 'x0,y0,x1,y1' (0~1 분수). 생략 시 전체 페이지")
    ap.add_argument("--dpi", type=int, default=200, help="렌더 해상도(기본 200)")
    ap.add_argument("-o", "--output", required=True, help="출력 PNG 경로")
    args = ap.parse_args()

    if not os.path.exists(args.pdf):
        sys.exit(f"[오류] PDF를 찾을 수 없습니다: {args.pdf}")
    doc = fitz.open(args.pdf)
    if not (1 <= args.page <= doc.page_count):
        sys.exit(f"[오류] 페이지 범위 초과: {args.page} (총 {doc.page_count}쪽)")
    page = doc[args.page - 1]

    clip = None
    if args.box:
        try:
            x0, y0, x1, y1 = [float(v) for v in args.box.split(",")]
        except Exception:
            sys.exit("[오류] --box 형식은 'x0,y0,x1,y1' (0~1 분수) 입니다.")
        r = page.rect
        clip = fitz.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height,
                         r.x0 + x1 * r.width, r.y0 + y1 * r.height)

    pix = page.get_pixmap(dpi=args.dpi, clip=clip)
    out_dir = os.path.dirname(os.path.abspath(args.output))
    os.makedirs(out_dir, exist_ok=True)
    pix.save(args.output)
    print(f"[완료] {pix.width}x{pix.height}px -> {os.path.abspath(args.output)}")


if __name__ == "__main__":
    main()
