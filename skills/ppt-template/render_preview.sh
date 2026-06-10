#!/usr/bin/env bash
# render_preview.sh — 생성한 PPTX를 페이지별 PNG로 렌더해 시각 검수용 미리보기를 만든다.
# 사용법: render_preview.sh report.pptx [출력폴더]
# 요구사항: LibreOffice(soffice) + python3 + PyMuPDF(fitz)
set -euo pipefail

PPTX="${1:?사용법: render_preview.sh <pptx> [outdir]}"
OUTDIR="${2:-$(dirname "$PPTX")/preview}"
SOFFICE="${SOFFICE_BIN:-soffice}"
command -v "$SOFFICE" >/dev/null 2>&1 || SOFFICE="/opt/homebrew/bin/soffice"

mkdir -p "$OUTDIR"
"$SOFFICE" --headless --convert-to pdf --outdir "$OUTDIR" "$PPTX" >/dev/null 2>&1
PDF="$OUTDIR/$(basename "${PPTX%.*}").pdf"

python3 - "$PDF" "$OUTDIR" <<'PY'
import sys, fitz
pdf, outdir = sys.argv[1], sys.argv[2]
doc = fitz.open(pdf)
for i, page in enumerate(doc):
    page.get_pixmap(dpi=90).save(f"{outdir}/slide{i+1:02d}.png")
print(f"[완료] {len(doc)}장 미리보기 -> {outdir}/slideNN.png")
PY
