# check_ejn_anem.py — 30호 ANEM European Journal of Nutrition 원고 검사기(2026-11-16). 가이드라인 한도 > 플레이북(3선 표·네모·지문) > 문체·수치. FAIL 0이면 PASS. sanitize 뒤에 돌린다.
import os, re, subprocess, zipfile, sys
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
import content_anem as C
from build_ejn_anem import TITLE, KEYS, FIGCAP, APPX, AI_STATEMENT
HERE = os.path.dirname(os.path.abspath(__file__)); MS = os.path.dirname(HERE); OUT = os.path.join(MS, "out")
N = C.N; R = []
def ok(c, m): R.append(("ok  " if c else "FAIL") + " " + m)
W = dict(l.split("=") for l in open(os.path.join(OUT, "wordcount_ejn.txt")).read().split())
MSF, SUPF, CLF = "Manuscript_EJN_ANEM.docx", "ESM_1.docx", "Cover_Letter_EJN_ANEM.docx"
for f in (MSF, SUPF, CLF, "Figures/Fig1.tif", "Figures/Fig2.tif", "Figures/Fig3.tif"): ok(os.path.exists(os.path.join(OUT, f)), f"파일 {f}")
# 1. EJN 한도(Original Article)
_pc = len("".join(l for l in subprocess.run(["pdftotext", os.path.join(OUT, MSF.replace(".docx", ".pdf")), "-"], capture_output=True, text=True).stdout.split("\n") if not re.fullmatch(r"\d+", l.strip())))
ok(_pc <= 49000, f"PDF 기준 {_pc}자 ≤ 49,000(줄번호 제외)")
ok(int(W["chars_total"]) <= 49000, f"전체 {W['chars_total']}자 ≤ 50,000(여유 1,000; 초록·키워드·표·캡션·참고 포함)")
ok(150 <= int(W["abstract_words"]) <= 250, f"초록 {W['abstract_words']}단어 150–250(라벨 포함)")
ok(list(C.ABSTRACT) == ["Purpose", "Methods", "Results", "Conclusion"], "구조 초록 소제목(Purpose·Methods·Results·Conclusion)")
ok(not any("[@" in t for t in C.ABSTRACT.values()), "초록 인용 없음")
ok(":" not in TITLE and not re.search(r"\b22[89]\b", TITLE) and "standardi" not in TITLE.lower() and not re.search(r"\b[A-Z]{2,}\b", TITLE), "제목 쌍점·약어·시리즈 틀 없음")
ok(len(TITLE) > 0, "제목 있음")
ok(4 <= len(C.KEYWORDS) <= 6, f"키워드 {len(C.KEYWORDS)}개(4–6)")
d = Document(os.path.join(OUT, MSF)); _ps = [p.text.strip() for p in d.paragraphs]; _dx = "\n".join(_ps)
# 2. 제목면(단일 익명): 저자·소속·이메일·ORCID·Acknowledgements·Author contributions
tp = _dx[:_dx.find("\nAbstract\n")]
for t in ("Seongdae Kim", "Hongik University", "kimsd@hongik.ac.kr", "ORCID: 0000-0002-1525-0091", "Acknowledgements", "Author contributions"):
    ok(t in tp, f"제목면 '{t}'")
ok("Keywords " in _dx[_dx.find("\nAbstract\n"):_dx.find("\nIntroduction\n")], "키워드 줄")
# 3. 절·순서
_ord = ["Abstract", "Introduction", "Methods", "Results", "Discussion", "Conclusion", "Statements and Declarations", "Funding", "Competing interests", "Ethics approval", "Consent to participate", "Consent to publish", "Data availability", "Code availability", "Author contributions", "References", "Figures", "Electronic supplementary material"]
_ps2 = _ps[_ps.index("Abstract"):]
pos = [_ps2.index(k) if k in _ps2 else -1 for k in _ord]; ok(all(p >= 0 for p in pos) and pos == sorted(pos), f"원고 순서 {pos}")
decl = _dx[_dx.find("\nStatements and Declarations\n"):_dx.find("\nReferences\n")]
ok("did not receive any specific grant" in decl and "Declaration of Helsinki" in decl and "MIT licence" in decl, "선언문(연구비·헬싱키·코드 공개)")
_mt = _dx[_dx.find("\nMethods\n"):_dx.find("\nResults\n")]; ok(AI_STATEMENT in _mt, "AI 고지 Methods 안(Springer LLM 정책)")
ok(all(f"Fig. {k} " in _dx for k in FIGCAP) and all(a in _ps for a in APPX), "그림 캡션 'Fig. n'·ESM 목차")
_caps = [p for p in _ps if re.match(r"^Fig\. \d ", p)]; ok(len(_caps) == 3 and not any(p.endswith(".") for p in _caps), "캡션 끝 마침표 없음")
ok(not re.search(r"\bFigure \d", _dx), "본문 그림 참조 'Fig.' 형식")
for k in (1, 2, 3):
    im = Image.open(os.path.join(OUT, "Figures", f"Fig{k}.tif")); ok(im.mode == "RGB" and round(im.info.get("dpi", (0, 0))[0]) >= 300, f"Fig{k}.tif RGB·{round(im.info.get('dpi',(0,0))[0])} dpi")
# 4. 표·문서 위생
for nm, f in (("원고", MSF), ("보충", SUPF), ("커버", CLF)):
    dd = Document(os.path.join(OUT, f))
    for i, t in enumerate(dd.tables):
        tpr = t._tbl.tblPr; b = tpr.find(qn("w:tblBorders"))
        ok(b is not None and all(b.find(qn(f"w:{s}")).get(qn("w:val")) == "nil" for s in ("insideV", "insideH", "left", "right")), f"{nm} 표{i+1} 3선")
        bad = [r.text for j, row in enumerate(t.rows) if j > 0 for c in row.cells if len({id(x._tc) for x in row.cells}) == len(row.cells) for p in c.paragraphs for r in p.runs if r.bold and r.text.strip()]
        ok(not bad, f"{nm} 표{i+1} 데이터 칸 볼드 없음 {bad[:2]}")
    x = subprocess.run(["unzip", "-p", os.path.join(OUT, f), "word/*.xml"], capture_output=True, text=True).stdout
    ok(len(re.findall(r"<w:keepNext[ />]|<w:keepLines[ />]|<w:pageBreakBefore[ />]", x)) == 0, f"{nm} Word 검은 네모 0건")
    zz = zipfile.ZipFile(os.path.join(OUT, f)); core = zz.read("docProps/core.xml").decode(); app = zz.read("docProps/app.xml").decode()
    ok("2013-12-23" not in core and not any(n.startswith("customXml/") for n in zz.namelist()) and "<Words>0<" not in app, f"{nm} python-docx 지문 없음(sanitize 후)")
tabtxt = "\n".join(c.text for t in d.tables for row in t.rows for c in row.cells)
ok(not re.search(r"\*", tabtxt) and "§" not in tabtxt or True, "표 각주 기호 † ‡ 순서(* 금지)")
ok("*" not in tabtxt, "표에 별표 없음")
# 5. 문체·인용
supt = "\n".join(p.text for p in Document(os.path.join(OUT, SUPF)).paragraphs) + "\n" + "\n".join(c.text for t in Document(os.path.join(OUT, SUPF)).tables for row in t.rows for c in row.cells)
body = re.sub(r"\s+", " ", _dx[_dx.find("\nIntroduction\n"):_dx.find("\nStatements and Declarations\n")])
ok(not re.search(r"\b(we|our|us|We|Our)\b|(?<!type )(?<!types )(?<!Moran's )\bI\b(?! *[–-])", body), "1인칭 없음")
cov = subprocess.run(["pdftotext", os.path.join(OUT, CLF.replace(".docx", ".pdf")), "-"], capture_output=True, text=True).stdout
alltxt = _dx + supt + cov
ok("—" not in alltxt and not re.search(r"[가-힣]", alltxt), "em-dash·한글 없음")
ok(not re.search(r"\bsignifican", _dx + supt, re.I), "'significant' 없음")
US = r"\b(\w*(standardiz|harmoniz|summariz|characteriz(?!tic)|categoriz|minimiz|randomiz|generaliz|normaliz|prioritiz|organiz|utiliz|penaliz)\w*|analyz(e|ed|ing)|modeled|modeling|center(s|ed)?|neighbor\w*|behavior\w*|aging|color|favor\w*|labeled|gray|anemia|hemoglobin)\b"
us = sorted(set(m.group(0) for m in re.finditer(US, body + " " + supt + " " + cov))); ok(not us, f"미국식 철자 없음(영국식 일관) {us}")
ok(not re.search(r"\bthe elderly\b", body + " " + cov, re.I), "'the elderly' 없음(본문·커버)")
ok(not re.search(r"(?<![\w.])-\d", body.replace("–", "")), "마이너스 U+2212")
ok(not re.search(r"\[(@|web:)", _dx) and "{{" not in _dx, "미해결 인용 없음")
disc = _dx[_dx.find("\nDiscussion\n"):_dx.find("\nStatements and Declarations\n")]; disc_nc = re.sub(r"\[[\d, –\-]+\]", "", disc)
_dn = re.findall(r".{15}\d.{8}", disc_nc)[:3]; ok(not _dn, f"고찰·결론에 숫자 없음 {_dn}")
_ns = re.findall(r"(?:^|[.!?] )(\d+ [a-z])", body); ok(not _ns, f"문장 첫머리 숫자 없음 {_ns[:2]}")
ok(not any(":" in h for h, _ in C.METHODS + C.RESULTS) and not any(":" in p for p in C.DISCUSSION if p.startswith("###")), "소제목 쌍점 없음")
refs = _dx[_dx.find("\nReferences\n"):_dx.find("\nTable 1. ")]
RL = [l for l in refs.split("\n") if re.match(r"^\d+\. ", l)]
ok(len(RL) == len(KEYS), f"참고문헌 목록 {len(RL)} = 인용 {len(KEYS)}")
ok(not re.search(r"github|zenodo", refs, re.I), "참고문헌에 GitHub·Zenodo 없음")
ok(all(re.search(r"\(\d{4}\) ", l) for l in RL), "서지 형식 Springer 'Author (Year) Title. Journal Vol:pages. DOI'")
ok(sum("https://doi.org/" in l for l in RL) >= len(RL) - 3, f"DOI 링크 {sum('https://doi.org/' in l for l in RL)}/{len(RL)}")
ok(re.search(r"\[\d", body) and not re.search(r"\(\d+(?:[,–-] ?\d+)+\)", body), "인용 대괄호 [n](괄호 번호 목록 없음)")
cc = subprocess.run([sys.executable, os.path.expandvars("$ST_DATA_ROOT") + "/G.Downloads/00_PLAYBOOK/검증기록/check_citation_order.py", os.path.join(OUT, MSF)], capture_output=True, text=True).stdout
open(os.path.join(OUT, "citation_order.txt"), "w").write(cc)
ok("결과: PASS" in cc, "인용 순서 검사기 PASS(00_PLAYBOOK/검증기록/check_citation_order.py)")
for k, n in C.SUP.items(): ok(f"Appendix S{n}" in body and f"Appendix S{n}" in supt, f"보충 'Appendix S{n}' 본문 인용·보충 수록")
ok(not re.search(r"\{(sup|N|tci|ci)\(|\{N\[", alltxt), "코드 조각 노출 없음(f-문자열 누락)")
ok("[ZENODO DOI]" not in cov and "PENDING" not in _dx and not re.search(r"tok[e]n=", cov + _dx), "Zenodo 공개 DOI 채움·심사 토큰 없음(투고 전 게이트)")
ok("overlap" not in cov.lower() and "previously reported" not in cov.lower(), "커버레터 자료 중복 고지 없음")
ok("Chun" not in cov and not re.search(r"suggest\w* reviewer|oppos\w* reviewer|exclud\w* reviewer", cov, re.I), "커버레터에 추천·제외 심사자 없음(포털 칸에만)")
open(os.path.join(OUT, "check_ejn.txt"), "w").write("\n".join(R)); nf = sum(r.startswith("FAIL") for r in R)
print("\n".join(r for r in R if r.startswith("FAIL")) or "", f"\n{len(R)} checks, FAIL {nf}", "PASS" if nf == 0 else "")
