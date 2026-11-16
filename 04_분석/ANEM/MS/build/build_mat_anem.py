# build_mat_anem.py — Maturitas(Elsevier EM) Original article 원고 조립(2026-11-16, 30호 ANEM).
#  규정(sciencedirect.com/journal/maturitas/publish/guide-for-authors, 2026-11-16 열람): 본문 2,000-3,000단어·참고 ≤30, 구조 초록 ≤250(Objectives·Study design·Main outcome measures·Results·Conclusions),
#   키워드 1-6, Highlights 3-5(≤85자, 별도 파일), 번호 절, 인용 [n], Elsevier 번호 참고문헌, 그림 별도 파일, 지도 경계 문구, Acknowledgements는 참고문헌 바로 앞, GenAI 선언 절은 참고문헌 앞.
import os, re, json
from jkms_engine import JKMS
from british import brit
import content_anem as C
HERE = os.path.dirname(os.path.abspath(__file__)); MS = os.path.dirname(HERE); OUT = os.path.join(MS, "out"); os.makedirs(OUT, exist_ok=True)
N = C.N; REFS = json.load(open(os.path.join(HERE, "refs_anem.json"), encoding="utf-8"))
rdf = lambda f, d: open(os.path.join(OUT, f)).read().strip() if os.path.exists(os.path.join(OUT, f)) else d
ZENODO = rdf("zenodo_doi.txt", "10.5281/zenodo.PENDING"); GITHUB = rdf("github_url.txt", "https://github.com/PENDING")
TAG = "MAT_ANEM"; TITLE = C.TITLE
ORCID = "0000-0002-1525-0091"
AFFIL = "Department of Biohealth Convergence Open Sharing System, Hongik University, Seoul, Republic of Korea"
ADDR = "Department of Biohealth Convergence Open Sharing System, Hongik University, 94 Wausan-ro, Mapo-gu, Seoul 04066, Republic of Korea"
FUNDING = "This research did not receive any specific grant from funding agencies in the public, commercial, or not-for-profit sectors."
AI_STATEMENT = "The author used Claude (Anthropic) and ChatGPT (OpenAI) only for translation and English-language editing of the manuscript."
COI = "The author has no relevant financial or non-financial interests to disclose."
ETHICS = "The Institutional Review Board of Hongik University approved the study on 4 November 2025 (7002340-202511-HR-012) and its continuing review on 3 November 2026 (7002340-202611-HR-040)."
CONSENT = "Informed consent was waived by the Institutional Review Board because only published aggregate statistics were used and no individual was identifiable."
PERMISSION = "No material reproduced from other sources."
TRIAL = "Not applicable."
def data_statement():
    return ("The district statistics analysed are publicly available from the Korean Statistical Information Service (https://kosis.kr) and the other official sources cited in the Methods. "
            "The processed district-year data are available from the author on reasonable request. "
            )
def code_statement():
    return (f"The analysis code is openly available under the MIT licence on Zenodo (https://doi.org/{ZENODO}) and GitHub ({GITHUB}). It contains code only; no individual-level data were used.")

def pandoc_cites(t): return re.sub(r"\s*\[(@[^\]]+)\]", lambda m: " {{" + ",".join(k.strip().lstrip("@") for k in m.group(1).split(";")) + "}}", t)
def strip(t): return re.sub(r"\s*\[@[^\]]+\]", "", t)
def all_keys():
    src = open(os.path.join(HERE, "content_anem.py"), encoding="utf-8").read()
    src = re.sub(r"(?m)^\s*#.*$", "", src)
    return sorted(k for k in {k.strip().lstrip("@") for g in re.findall(r"\[(@[^\]]+)\]", src) for k in g.split(";")} if re.fullmatch(r"(web:)?[a-z][a-z0-9]*", k))
KEYS = all_keys()

def _ini(n):   # "Besora-Moreno M" → "M. Besora-Moreno", "Kim HS" → "H.S. Kim"
    p = n.rsplit(" ", 1)
    if len(p) == 2 and p[1].isupper() and len(p[1]) <= 3: return ".".join(p[1]) + ". " + p[0]
    return n
def ref(k):   # Elsevier 번호: H.S. Ballard, Title, J Abbrev 13 (1989) 706–720. https://doi.org/...
    r = REFS[k]
    if r["type"] == "web": return f"{r['authors']}, {r['title']}, {r['year']}. {r['url']} (accessed {r['accessed']})."
    au = [_ini(x) for x in r["authors"]]; a = ", ".join(au) if len(au) <= 6 else ", ".join(au[:6]) + ", et al."
    t = r["title"].rstrip(".")
    s = f"{a}, {t}, {r['journal']}"
    if r.get("vol"): s += f" {r['vol']}"
    s += f" ({r['year']})"
    if r.get("pages"): s += f" {r['pages'].replace('-', chr(8211))}"
    s += "."
    if r.get("doi"): s += f" https://doi.org/{r['doi']}."
    return s

class TX(JKMS):
    raw = False
    def _emit(self, p, text, sz=12, bold=False, italic=False):
        if not self.raw:
            text = brit(text); text = re.sub(r"(^|[\s(=])-(?=\d)", "\\1\u2212", text)
        return super()._emit(p, text, sz, bold, italic)
    def head(self, text, level=1):
        p = super().head(text, level)
        if p is not None: p.paragraph_format.keep_with_next = True
        return p

def foot_row(J, text, sz=9, ls=15):
    from docx.enum.text import WD_LINE_SPACING
    from docx.shared import Pt
    t = J.doc.tables[-1]; r = t.add_row(); m = r.cells[0].merge(r.cells[-1]); m.text = ""; p = m.paragraphs[0]
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY; p.paragraph_format.line_spacing = Pt(ls); p.paragraph_format.space_before = Pt(4)
    c = J.counting; J.counting = False; J._emit(p, text, sz); J.counting = c
def put_table(J, t, fs=9):
    c = J.counting; J.counting = False; J.table_fit(brit(t[0]), [brit(x) for x in t[1]], [[brit(x) for x in r] for r in t[2]], "", fs_max=fs, title_bold=True); foot_row(J, brit(t[3]), 8); J.counting = c

LABEL = {"basic_livelihood_rate": "Basic livelihood security receipt", "tertiary_educated_share": "Tertiary education", "ext__nutrition_label_literacy": "Nutrition-label literacy",
         "older_single_household_share": "Older adults living alone", "multicultural_marriage_share": "Marriages with a foreign spouse", "ext__total_fertility_rate": "Total fertility rate",
         "ext__dx_dm_share": "Diagnosed diabetes among attendees", "current_smoking_rate": "Current smoking", "high_risk_drinking_rate": "High-risk drinking",
         "walking_practice_rate": "Walking", "moderate_pa_rate": "Moderate or vigorous physical activity", "im_specialists_per_100k": "Internal medicine specialists",
         "screening_uptake": "Screening uptake", "urban_population_share": "Urban population share", "physicians_per_1000": "Physicians",
         "sex_ratio": "Sex ratio", "fiscal_independence": "Fiscal autonomy", "pop80_share": "Population aged 80 years or older"}
GROUPS = [("Deprivation, household and nutrition-label literacy", ["basic_livelihood_rate", "tertiary_educated_share", "ext__nutrition_label_literacy", "older_single_household_share"]),
          ("Women and childbearing", ["multicultural_marriage_share", "ext__total_fertility_rate"]),
          ("Chronic disease", ["ext__dx_dm_share"]),
          ("Lifestyle and physical activity", ["current_smoking_rate", "high_risk_drinking_rate", "walking_practice_rate", "moderate_pa_rate"]),
          ("Access and selection", ["im_specialists_per_100k", "screening_uptake"]),
          ("Forced controls", ["urban_population_share", "physicians_per_1000", "sex_ratio", "fiscal_independence", "pop80_share"])]
EXPECT = {"basic_livelihood_rate": "+", "tertiary_educated_share": "−", "ext__nutrition_label_literacy": "−", "older_single_household_share": "+", "multicultural_marriage_share": "+",
          "ext__total_fertility_rate": "+", "ext__dx_dm_share": "+", "current_smoking_rate": "−", "high_risk_drinking_rate": "+", "walking_practice_rate": "−", "moderate_pa_rate": "−",
          "im_specialists_per_100k": "−", "screening_uptake": "−"}
UNIT = {"basic_livelihood_rate": "%", "tertiary_educated_share": "% of adults", "ext__nutrition_label_literacy": "%", "older_single_household_share": "% of households",
        "multicultural_marriage_share": "% of marriages", "ext__total_fertility_rate": "births per woman", "ext__dx_dm_share": "% of attendees", "current_smoking_rate": "%",
        "high_risk_drinking_rate": "%", "walking_practice_rate": "%", "moderate_pa_rate": "%", "im_specialists_per_100k": "per 100,000", "screening_uptake": "%",
        "urban_population_share": "%", "physicians_per_1000": "per 1,000", "sex_ratio": "men per 100 women", "fiscal_independence": "%", "pop80_share": "%"}
assert set(LABEL) == set(N["main"]) and sum(len(v) for g, v in GROUPS[:-1]) == N["n_hyp"]

def table1():
    T = N["T1"]; cols = ["W1_2011-16", "W2_2017-24", "2011-2024"]
    rows = [["__H__", "Screening"],
            ["  District-years, n"] + [f"{T[c]['n']:,}" for c in cols],
            ["  Screening records, million"] + [f"{T[c]['exam']/1e6:.1f}" for c in cols],
            ["  Records flagged with suspected anaemia, n"] + [f"{T[c]['cases']:,}" for c in cols],
            ["  Crude positivity, %"] + [f"{T[c]['pos']:.2f}" for c in cols],
            ["  Standardised anaemia ratio, median†"] + [f"{T[c]['sir_med']:.2f}" for c in cols],
            ["  Standardised anaemia ratio, P10–P90†"] + [f"{T[c]['sir_p10']:.2f}–{T[c]['sir_p90']:.2f}" for c in cols]]
    for g, vs in GROUPS:
        rows.append(["__H__", f"{g}, mean (SD)"])
        for v in vs:
            mx = max(abs(T[c]['cov'][v][0]) for c in cols); dp = 0 if mx >= 100 else (2 if mx < 3 else 1)
            rows.append([f"  {LABEL[v]}, {UNIT[v]}"] + [f"{T[c]['cov'][v][0]:,.{dp}f} ({T[c]['cov'][v][1]:,.{dp}f})" for c in cols])
    foot = ("†Observed over expected number flagged, summed over the window by district; expected numbers from national positivity by year, sex and 5-year age group applied to the estimated age structure of each district's attendees. "
            "Covariates are shown before standardisation. P10–P90, 10th–90th percentile; SD, standard deviation.")
    return ("Table 1. Anaemia flagged at national general health screening and district characteristics by calendar window, Korea, 2011–2024", ["Characteristic", "2011–2016", "2017–2024", "2011–2024"], rows, foot)

def table2():
    rows = []
    for g, vs in GROUPS:
        rows.append(["__H__", g])
        for v in vs:
            m = N["main"][v]; mv = N["mv"].get(v); k2 = N["s43"][("KH_typeII", v)]; k4 = N["s43"][("KH_typeIV", v)]
            rows.append([f"  {LABEL[v]}", EXPECT.get(v, "NA"), f"{m['IRR']:.3f} ({m['lo']:.3f}–{m['hi']:.3f})", f"{100*mv['cred_frac']:.0f}" if mv else "NA", f"{k2[0]:.3f} ({k2[1]:.3f}–{k2[2]:.3f})", f"{k4[0]:.3f} ({k4[1]:.3f}–{k4[2]:.3f})"])
    foot = (f"RR, rate ratio per 1 standard deviation; CrI, credible interval; NA, not applicable (forced control); expected: + more, − less anaemia. Principal model: negative binomial, Besag–York–Mollié 2 district effect, random-walk and independent year effects (n={N['n_rows']:,} district-years). "
            "†Share of 200 covariate-subset fits containing the covariate with a CrI excluding 1. "
            f"‡Knorr-Held type II (best fitting; ΔWAIC {C.dW('KH_typeII'):,.0f}) or type IV ({C.dW('KH_typeIV'):,.0f}) space–time interaction added; WAIC, widely applicable information criterion. Bounds shown as 1.000 were classified from unrounded values.")
    return ("Table 2. District characteristics and the age- and sex-standardised anaemia ratio in the principal and space–time interaction models", ["Characteristic", "Expected", "Principal RR (95% CrI)", "Subsets credible, %†", "Type II RR (95% CrI)‡", "Type IV RR (95% CrI)‡"], rows, foot)

def tight(tb, fs):   # 가로 표 한 쪽(2026-11-12): 셀 줄간격 고정
    from docx.enum.text import WD_LINE_SPACING
    from docx.shared import Pt
    for row in tb.rows:
        for c in row.cells:
            for p in c.paragraphs: p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY; p.paragraph_format.line_spacing = Pt(round(1.6 * fs, 1)); p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(0)
FIGCAP = {
 1: "National positivity for suspected anaemia at general health screening by 5-year age group in men (A) and women (B), Korea, selected years 2011–2024. Positivity is the number of screening records flagged divided by the number of screening records with a recorded judgement in each sex and age group, shown on a log scale. Data from the national statistics portal.",
 2: "Geography of the age- and sex-standardised anaemia ratio, Korea, 2011–2024. (A) Residual spatial rate ratio (RR) by district, the posterior mean of the exponentiated Besag–York–Mollié 2 district effect in the principal model. Solid outlines, hot spots (posterior probability of a residual spatial RR above 1 greater than 0.95); dashed outlines, cold spots (less than 0.05); thick lines, province boundaries. (B) Number of calendar windows (2011–2016, 2017–2024) in which each district was a hot or cold spot in the principal model refitted in each window; numbers of districts in parentheses. Sejong (hatched) was excluded. District boundaries as of 2024. Map lines delineate study areas and do not necessarily depict accepted national boundaries.",
 3: "Rate ratios per 1 standard deviation with 95% credible intervals for the four subset-stable district characteristics across models: the principal model, models with Knorr-Held type II (best fitting) and type IV space–time interactions, the principal model fitted to 2011–2016 and to 2017–2024 only, and the principal model with expected numbers computed from the resident population instead of attendees. The dotted vertical line marks a rate ratio of 1; log scale.",
}
APPX = ["Appendix S1. Supplementary methods: data sources, district harmonisation, standardisation and model specification",
        "Appendix S2. Covariate definitions, sources, expected directions and filled values",
        "Appendix S3. Panel checks before modelling",
        "Appendix S4. Emerging hot spot categories and persistence (Figure S1)",
        "Appendix S5. Window-specific rate ratios and interactions with the window and the COVID-19 period",
        "Appendix S6. Robustness of the subset-stable characteristics, comparison outcomes, negative controls and model checks",
        "Appendix S7. Post hoc examination of associations against the expected direction",
        "Appendix S8. STROBE checklist"]

def counts():
    aw = sum(len(t.split()) for t in C.ABSTRACT.values()) + sum(len(h.split()) for h in C.ABSTRACT)
    secs = [("Introduction", C.INTRO), ("Methods", [h for h, _ in C.METHODS] + [t for _, t in C.METHODS]), ("Results", [h for h, _ in C.RESULTS] + [t for _, t in C.RESULTS]), ("Discussion", [x.replace("### ", "") for x in C.DISCUSSION])]
    bw = sum(len(h.split()) + sum(len(strip(t).split()) for t in ts) for h, ts in secs)
    return aw, bw

AUTHCONTRIB = "Seongdae Kim conceived and designed the study, assembled and checked the data, performed the analysis, interpreted the results, wrote the manuscript and approved the final version."
GENAI = "During the preparation of this work the author used Claude (Anthropic) and ChatGPT (OpenAI) only for translation and English-language editing. After using these tools, the author reviewed and edited the content as needed and takes full responsibility for the content of the published article."
HIGHLIGHTS = ["Anaemia flagged at Korean health screening clustered in the south, 2011–2024",
              "Districts with more older adults living alone had more anaemia flagged",
              "Diabetes among attendees was associated, less robustly, with anaemia",
              "Physical activity showed no credible district-level association",
              "Ecological findings call for individual studies of diet and iron status"]
assert 3 <= len(HIGHLIGHTS) <= 5 and all(len(h) <= 85 for h in HIGHLIGHTS), [len(h) for h in HIGHLIGHTS]
def build():
    aw, bw = counts(); nref = len(KEYS)
    J = TX(cite="bracket")
    J.para(f"<b>{brit(TITLE)}</b>")
    J.para("Seongdae Kim", before=12); J.para(ADDR)
    J.para(f"<b>Corresponding author:</b> Seongdae Kim, {ADDR}. Email: kimsd@hongik.ac.kr. ORCID: {ORCID}")
    J.para(f"<b>Word count:</b> abstract {aw}; main text {bw:,}. <b>References:</b> {nref}. <b>Tables:</b> 2. <b>Figures:</b> 3.", before=6)
    J.page_break(); J.head("Abstract", 1)
    for h, t in C.ABSTRACT.items(): J.para(f"<b>{h}:</b> {t}")
    J.raw = True; J.para(f"<b>Keywords:</b> {'; '.join(C.KEYWORDS)}", before=6); J.raw = False
    J.page_break(); J.counting = True
    J.head("1. Introduction", 1)
    for p in C.INTRO: J.para(pandoc_cites(p), indent=True)
    J.head("2. Methods", 1)
    for i, (h, p) in enumerate(C.METHODS, 1): J.head(f"2.{i}. {h}", 2); J.para(pandoc_cites(p), indent=True)
    J.head("3. Results", 1)
    for i, (h, p) in enumerate(C.RESULTS, 1): J.head(f"3.{i}. {h}", 2); J.para(pandoc_cites(p), indent=True)
    J.head("4. Discussion", 1); _d = 0
    for p in C.DISCUSSION:
        if p.startswith("### Conclusion"): J.head("5. Conclusion", 1); continue
        if p.startswith("### "): _d += 1; J.head(f"4.{_d}. {p[4:].strip()}", 2); continue
        J.para(pandoc_cites(p), indent=True)
    J.counting = False
    for h, t in (("Contributors", AUTHCONTRIB), ("Conflict of interest", COI), ("Funding", FUNDING), ("Ethical approval", ETHICS + " " + CONSENT + " The study was conducted in accordance with the Declaration of Helsinki."),
                 ("Provenance and peer review", "This article was not commissioned and was externally peer reviewed."), ("Research data (data sharing statement)", data_statement() + " " + code_statement()),
                 ("Declaration of generative AI and AI-assisted technologies in the manuscript preparation process", GENAI), ("Acknowledgements", "None.")):
        J.head(h, 1); J.para(t)
    assert sorted(J.cites) == KEYS, (set(J.cites) ^ set(KEYS))
    J.page_break(); J.head("References", 1); J.raw = True
    for i, k in enumerate(J.cites, 1): J.para(f"[{i}] " + ref(k))
    J.raw = False
    J.page_break(); put_table(J, table1()); J.new_page("landscape"); put_table(J, table2(), fs=8); tight(J.doc.tables[-1], 8)
    J.new_page("portrait"); J.head("Figure captions", 1)
    for k in sorted(FIGCAP): J.para(f"<b>Fig. {k}.</b> {FIGCAP[k]}")
    J.head("Supplementary material", 1)
    J.para("Supplementary material (Supplementary_material.pdf) contains the following appendices:")
    for a_ in APPX: J.para(a_)
    return J, aw, bw, nref

def highlights():
    from docx import Document
    d = Document(); d.add_paragraph("Highlights")
    for h in HIGHLIGHTS: d.add_paragraph(h, style="List Bullet")
    d.core_properties.author = "Seongdae Kim"; d.save(os.path.join(OUT, "Highlights.docx"))

def charcount(path):
    from docx import Document
    d = Document(path); ps = [p.text for p in d.paragraphs]
    seen = set(); cells = []
    for tb in d.tables:
        for row in tb.rows:
            for c in row.cells:
                if id(c._tc) in seen: continue
                seen.add(id(c._tc)); cells.append(c.text)
    t = "\n".join(ps) + "\n" + "\n".join(cells)   # 병합 칸은 한 번만
    return len(t)

if __name__ == "__main__":
    J, aw, bw, nref = build()
    assert aw <= 250, aw; assert 2000 <= bw <= 3000, bw; assert nref <= 30, nref
    p = os.path.join(OUT, f"Manuscript_{TAG}.docx"); J.doc_title = TITLE; J.save(p); highlights()
    open(os.path.join(OUT, "wordcount_mat.txt"), "w").write(f"body_words={bw}\nabstract_words={aw}\nrefs={nref}\ntables=2\nfigures=3\ntitle_chars={len(TITLE)}\n")
    print(f"body {bw} | abstract {aw} | refs {nref} | title chars {len(TITLE)}")
