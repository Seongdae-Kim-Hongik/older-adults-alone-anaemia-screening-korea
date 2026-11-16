#!/bin/bash
# 30호 ANEM Maturitas 전체 빌드(2026-11-16): 그림 → 원고 → 보충 → 커버 → sanitize → PDF → 검사기
set -e; cd "$(dirname "$0")"; O=../out
python3 figs_anem.py
python3 build_mat_anem.py
python3 esm_anem.py
python3 cover_anem.py "$@"
python3 sanitize_docx.py $O/Manuscript_MAT_ANEM.docx $O/Supplementary_material.docx $O/Cover_Letter_MAT_ANEM.docx $O/Highlights.docx > /dev/null
for f in Manuscript_MAT_ANEM Supplementary_material Cover_Letter_MAT_ANEM; do /opt/homebrew/bin/soffice --headless -env:UserInstallation=file:///tmp/lo_anem_all --convert-to pdf --outdir $O $O/$f.docx > /dev/null 2>&1; done
python3 check_mat_anem.py
