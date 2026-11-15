#!/bin/bash
# 30호 ANEM European Journal of Nutrition 전체 빌드(2026-11-16): 그림 → 원고 → 보충 → 커버 → sanitize → PDF → 검사기
set -e; cd "$(dirname "$0")"; O=../out
python3 figs_anem.py
python3 build_ejn_anem.py
python3 esm_anem.py
python3 cover_anem.py "$@"
python3 sanitize_docx.py $O/Manuscript_EJN_ANEM.docx $O/ESM_1.docx $O/Cover_Letter_EJN_ANEM.docx > /dev/null
for f in Manuscript_EJN_ANEM ESM_1 Cover_Letter_EJN_ANEM; do /opt/homebrew/bin/soffice --headless -env:UserInstallation=file:///tmp/lo_anem_all --convert-to pdf --outdir $O $O/$f.docx > /dev/null 2>&1; done
python3 check_ejn_anem.py
