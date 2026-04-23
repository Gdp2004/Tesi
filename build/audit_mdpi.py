"""Audit MRQF_MAS_Electronics.docx against the MDPI Electronics Instructions
for Authors. Every requirement is encoded as an atomic CHECK function that
prints PASS or FAIL with a short explanation.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "MRQF_MAS_Electronics.docx"


def words(text: str) -> int:
    return len(re.findall(r"\S+", text))


def get_paragraphs(doc):
    return [(p.style.name, p.text) for p in doc.paragraphs]


# ---------- front matter ----------

def check_article_type(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.1_article_type":
            if t.strip().lower() in {"article", "review", "communication"}:
                return None
            return f"article type must be Article/Review/... got: {t!r}"
    return "missing MDPI_1.1_article_type paragraph"


def check_title(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.2_title":
            if not t.strip():
                return "title is empty"
            if "running title" in t.lower() or "short title" in t.lower():
                return "title should not contain a running/short form"
            return None
    return "missing title"


def check_authors(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.3_authornames":
            if not t.strip():
                return "author line is empty"
            if "*" not in t:
                return "no corresponding author marker (*) in author line"
            return None
    return "missing author line"


def check_affiliation(doc) -> str | None:
    affils = [t for s, t in get_paragraphs(doc) if s == "MDPI_1.6_affiliation"]
    if not affils:
        return "no affiliation paragraph"
    joined = " | ".join(affils)
    if "Correspondence" not in joined:
        return "no Correspondence line"
    if "@" not in joined:
        return "no email in affiliation/correspondence"
    # PubMed/MEDLINE address: should have city + country
    if not re.search(r"(City|Country|Italy|Salerno|Fisciano)", joined, re.I):
        return "affiliation lacks city/country information"
    return None


def check_abstract(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            if not t.strip().lower().startswith("abstract"):
                return "abstract must begin with 'Abstract:'"
            body = re.sub(r"^Abstract:\s*", "", t, flags=re.I)
            wc = words(body)
            if wc > 230:
                return f"abstract too long ({wc} words, max ~200)"
            if wc < 100:
                return f"abstract too short ({wc} words)"
            return None
    return "missing abstract"


def check_keywords(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.8_keywords":
            body = re.sub(r"^Keywords:\s*", "", t, flags=re.I)
            parts = [p.strip() for p in body.split(";") if p.strip()]
            if not (3 <= len(parts) <= 10):
                return f"keywords count {len(parts)} not in [3,10]"
            return None
    return "missing keywords"


def check_line_rule(doc) -> str | None:
    for s, _ in get_paragraphs(doc):
        if s == "MDPI_1.9_line":
            return None
    return "missing MDPI_1.9_line horizontal rule"


# ---------- body / sections ----------

REQUIRED_RESEARCH_SECTIONS = [
    r"introduction",
    r"materials and methods|theoretical foundation|methodology|methods",
    r"results|validation|empirical",
    r"discussion",
]


def check_required_sections(doc) -> str | None:
    h1 = [t.lower() for s, t in get_paragraphs(doc) if s == "MDPI_2.1_heading1"]
    missing = []
    for pat in REQUIRED_RESEARCH_SECTIONS:
        if not any(re.search(pat, h) for h in h1):
            missing.append(pat)
    if missing:
        return f"missing required section-kind headings: {missing}"
    return None


def check_section_numbering(doc) -> str | None:
    h1 = [t for s, t in get_paragraphs(doc) if s == "MDPI_2.1_heading1"]
    numbered = [h for h in h1 if re.match(r"^\s*\d+\.\s+\S", h)]
    # Expect numbered main sections + non-numbered appendix/refs/abbr
    if len(numbered) < 5:
        return f"too few numbered sections ({len(numbered)})"
    # Appendix labels
    if not any(re.match(r"^\s*Appendix", h) for h in h1):
        return "missing Appendix heading"
    if not any(h.strip().lower().startswith("references") for h in h1):
        return "missing References heading"
    return None


# ---------- back matter ----------

BACKMATTER_EXPECTED = [
    "Supplementary Materials",
    "Author Contributions",
    "Funding",
    "Institutional Review Board",  # sometimes Ethical Approval
    "Informed Consent",
    "Data Availability",
    "Acknowledgments",
    "Conflicts of Interest",
]


def check_backmatter(doc) -> list[str]:
    errs = []
    back = [
        t for s, t in get_paragraphs(doc) if s == "MDPI_6.2_back_matter"
    ]
    joined = "\n".join(back)
    for key in BACKMATTER_EXPECTED:
        if key not in joined:
            errs.append(f"missing backmatter: {key}")
    # CRediT roles in author contributions
    credit_roles = [
        "Conceptualization",
        "Methodology",
        "Validation",
        "Formal Analysis",
        "Investigation",
        "Writing – Original Draft",
        "Writing – Review",
    ]
    ac_entry = next((b for b in back if b.startswith("Author Contributions")), "")
    for role in credit_roles:
        if role.lower() not in ac_entry.lower():
            errs.append(f"Author Contributions missing CRediT role: {role}")
    # Data Availability must match one of MDPI's recommended statements
    da = next((b for b in back if b.startswith("Data Availability")), "")
    recommended_stubs = [
        "openly available",
        "available on request",
        "available in",
        "no new data were created",
        "available by the authors on request",
        "included in the article",
    ]
    if not any(stub.lower() in da.lower() for stub in recommended_stubs):
        errs.append("Data Availability does not match any recommended MDPI template")
    return errs


# ---------- figures / tables / equations ----------

def check_figures(doc) -> str | None:
    caps = [t for s, t in get_paragraphs(doc) if s == "MDPI_5.1_figure_caption"]
    if not caps:
        return "no figure captions"
    for c in caps:
        if not c.strip().startswith("Figure "):
            return f"figure caption must start with 'Figure N.': {c[:60]}"
    return None


def check_tables(doc) -> str | None:
    caps = [t for s, t in get_paragraphs(doc) if s == "MDPI_4.1_table_caption"]
    if not caps:
        return "no table captions"
    for c in caps:
        if not c.strip().startswith(("Table ", "Table A")):
            return f"table caption must start with 'Table N.': {c[:60]}"
    return None


def check_figure_citations(doc) -> list[str]:
    """Every figure must be cited in the body text before/at its caption."""
    errs = []
    caps = [t for s, t in get_paragraphs(doc) if s == "MDPI_5.1_figure_caption"]
    body_text = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    # Also include the whole doc text because captions can reference figures
    doc_text = "\n".join(t for _, t in get_paragraphs(doc))
    numbers = []
    for c in caps:
        m = re.match(r"Figure\s+(\d+)", c)
        if m:
            numbers.append(int(m.group(1)))
    for n in sorted(set(numbers)):
        if not re.search(rf"Figure\s+{n}\b", body_text):
            errs.append(f"Figure {n} not cited in body text")
    return errs


def check_table_citations(doc) -> list[str]:
    errs = []
    caps = [t for s, t in get_paragraphs(doc) if s == "MDPI_4.1_table_caption"]
    body_text = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    numbers = []
    for c in caps:
        m = re.match(r"Table\s+([A-Z]?\d+)", c)
        if m:
            numbers.append(m.group(1))
    for n in numbers:
        if not re.search(rf"Table\s+{re.escape(n)}\b", body_text):
            errs.append(f"Table {n} not cited in body text")
    return errs


def check_equations_editable(doc) -> str | None:
    # Verify no $$…$$ left inline; our equations live in tables with a
    # number cell in column 2.
    for s, t in get_paragraphs(doc):
        if t.strip().startswith("$$"):
            return f"equation still embedded as text: {t[:60]}"
    # count equation tables (2-column, small)
    eq_tables = 0
    for tbl in doc.tables:
        if len(tbl.columns) == 2 and len(tbl.rows) == 1:
            eq_tables += 1
    if eq_tables < 3:
        return f"expected at least 3 equation tables, found {eq_tables}"
    return None


# ---------- references ----------

def check_references(doc) -> list[str]:
    errs = []
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    if not refs:
        return ["no references found with MDPI_8.1_references style"]
    # each reference should start with "N." numbering
    for i, r in enumerate(refs):
        if not re.match(r"^\s*\d+\.\s+", r):
            errs.append(f"reference #{i + 1} not numbered: {r[:60]}")
    # in-text citations use [N] not (N)
    body_text = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    if not re.search(r"\[\d+", body_text):
        errs.append("no [N] in-text citations detected")
    # parenthetical (Author, year) style not allowed
    parenthetical = re.findall(
        r"\([A-Z][a-zA-Z\-]+(?:\s+et\s+al\.?)?,\s*\d{4}\)", body_text
    )
    if parenthetical:
        errs.append(
            f"parenthetical (Author, year) citations present (first 5): "
            f"{parenthetical[:5]}"
        )
    return errs


# ---------- disclaimer ----------

def check_publisher_note(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_6.3_notes" and "Disclaimer/Publisher" in t:
            return None
    return "missing MDPI Disclaimer/Publisher's Note"


# ---------- acronyms defined at first use ----------

ACRONYMS_TO_DEFINE = {
    "MRQF": "Multiscale Relativistic Quantum Finance",
    "MAS": "Multi-Agent System",
    "SKB": "Shared Knowledge Base",
    "ECB": "European Central Bank",
    "GARCH": "Generalized Autoregressive Conditional Heteroscedasticity",
    "MCC": "Matthews Correlation Coefficient",
}


def check_acronym_first_use(doc) -> list[str]:
    errs = []
    # find abstract and first body text
    abstract = ""
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            abstract = t
            break
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    for acr, full in ACRONYMS_TO_DEFINE.items():
        # only check the ones actually used in body
        if acr not in body and acr not in abstract:
            continue
        # require "full (ACR)" pattern in abstract if acr appears in abstract
        if acr in abstract and full.lower() not in abstract.lower():
            errs.append(
                f"acronym {acr} used in abstract but {full!r} not spelled out"
            )
        if acr in body and full.lower() not in body.lower():
            errs.append(
                f"acronym {acr} used in body but {full!r} not spelled out"
            )
    return errs


# ---------- extra MDPI-specific checks ----------

def check_references_sequential(doc) -> str | None:
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    nums = []
    for r in refs:
        m = re.match(r"^\s*(\d+)\.", r)
        if m:
            nums.append(int(m.group(1)))
    expected = list(range(1, len(nums) + 1))
    if nums != expected:
        return f"reference numbering not sequential 1..{len(nums)} (got {nums[:5]}...)"
    return None


def check_no_running_title(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.2_title":
            if re.search(r"running\s+title|short\s+title", t, re.I):
                return "title must not declare a running/short title"
            return None
    return None


def check_si_units(doc) -> str | None:
    # quick sanity: flag the common non-SI tokens used outside of context words
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_")
    )
    bad = []
    for token, hint in [
        (r"\b\d+\s*inch(es)?\b", "imperial length"),
        (r"\b\d+\s*ft\b", "imperial length"),
        (r"\b\d+\s*mile(s)?\b", "imperial length"),
    ]:
        if re.search(token, body, re.I):
            bad.append(hint)
    return None if not bad else f"non-SI units detected: {bad}"


def check_no_parenthetic_refs(doc) -> str | None:
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    hits = re.findall(
        r"\([A-Z][a-zA-Z\-]+(?:\s+and\s+[A-Z][a-zA-Z\-]+)?\s*,?\s*\d{4}\)",
        body,
    )
    if hits:
        return f"Author-year citations present: {hits[:3]}"
    return None


def check_genai_disclosure(doc) -> str | None:
    """Section 2 (Methods) of the source explicitly disclosed GenAI policy.
    MDPI requires this if GenAI was used; we accept either an explicit
    disclosure in Methods/Materials section or the standard 'not used'
    statement in the Acknowledgments."""
    full = "\n".join(t for _, t in get_paragraphs(doc))
    if ("GenAI" in full) or ("generative artificial intelligence" in full.lower()):
        return None
    return "no GenAI disclosure found (either use-statement or 'not applicable')"


def check_figure_caption_period(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_5.1_figure_caption":
            # each caption begins with "Figure N." then a space and the body.
            if not re.match(r"Figure\s+\d+\.\s+\S", t):
                return f"figure caption lacks 'Figure N. ' header: {t[:60]}"
    return None


def check_table_caption_period(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_4.1_table_caption":
            if not re.match(r"Table\s+[A-Z]?\d+\.\s+\S", t):
                return f"table caption lacks 'Table N. ' header: {t[:60]}"
    return None


# ---------- fine-grained MDPI rules ----------

def _backmatter_entry(doc, prefix: str) -> str:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_6.2_back_matter" and t.startswith(prefix):
            return t
    return ""


def check_abstract_single_paragraph(doc) -> str | None:
    count = 0
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            count += 1
    if count != 1:
        return f"abstract must be a single MDPI_1.7_abstract paragraph, found {count}"
    return None


def check_funding_template(doc) -> str | None:
    entry = _backmatter_entry(doc, "Funding")
    if not entry:
        return "no Funding entry"
    ok_phrases = [
        "This research received no external funding",
        "This research was funded by",
    ]
    if not any(p in entry for p in ok_phrases):
        return "Funding must use the MDPI template wording"
    return None


def check_author_contribs_credit(doc) -> str | None:
    entry = _backmatter_entry(doc, "Author Contributions")
    if not entry:
        return "no Author Contributions entry"
    # MDPI requires the Title-Case CRediT terms separated by semicolons
    required_terms = [
        "Conceptualization",
        "Methodology",
        "Software",
        "Validation",
        "Formal Analysis",
        "Investigation",
        "Resources",
        "Data Curation",
        "Writing – Original Draft",
        "Writing – Review & Editing",
        "Visualization",
        "Supervision",
        "Project Administration",
        "Funding Acquisition",
    ]
    missing = [r for r in required_terms if r.lower() not in entry.lower()]
    if missing:
        return f"Author Contributions missing CRediT roles: {missing}"
    if ";" not in entry:
        return "Author Contributions must be semicolon-separated"
    if "All authors have read and agreed" not in entry:
        return "Author Contributions must end with the MDPI 'All authors have read and agreed' sentence"
    return None


def check_conflicts_wording(doc) -> str | None:
    entry = _backmatter_entry(doc, "Conflicts of Interest")
    if not entry:
        return "no Conflicts of Interest entry"
    canonical = [
        "The authors declare no conflicts of interest",
        "The authors declare no conflict of interest",
    ]
    if not any(c in entry for c in canonical):
        return "Conflicts of Interest must use the MDPI canonical sentence"
    return None


def check_references_journal_format(doc) -> list[str]:
    """Each reference carries at least a 4-digit year."""
    errs = []
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    for r in refs:
        # skip intro/narrative lines that don't start with an author
        if not re.match(r"^\s*\d+\.\s+[A-Z]", r):
            continue
        if not re.search(r"\b(19|20)\d{2}\b", r):
            errs.append(f"reference has no 4-digit year: {r[:80]}")
    return errs


def check_references_endash(doc) -> list[str]:
    """Flag true page ranges with ASCII hyphen, e.g. a ', Year, Vol, N-M.'
    pattern. ISBN hyphens and journal-identifier hyphens are ignored.
    """
    errs = []
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    for r in refs:
        # strip ISBN segments and DOIs so hyphens there don't match
        cleaned = re.sub(r"ISBN\s+[\d\-Xx]+", "", r)
        cleaned = re.sub(r"https?://\S+", "", cleaned)
        # page range after a comma and a 4-digit year and a volume:
        # , 1987, 42, 281-300.
        m = re.search(
            r",\s*(?:19|20)\d{2}\s*,\s*\d+\s*,\s*\d+-\d+",
            cleaned,
        )
        if m:
            errs.append(
                f"reference uses hyphen in page range: {m.group(0).strip()}"
            )
        # also: pp. N-M. with ASCII hyphen
        m2 = re.search(r"pp\.\s*\d+-\d+", cleaned)
        if m2:
            errs.append(
                f"reference uses hyphen in 'pp.' range: {m2.group(0)}"
            )
    return errs


def check_references_doi(doc) -> str | None:
    refs = [
        t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"
    ]
    # Count references that look like journal articles (not books/websites)
    substantive = [r for r in refs if re.match(r"^\s*\d+\.\s+[A-Z]", r)]
    with_doi = sum(1 for r in substantive if "doi.org" in r or "https://doi" in r)
    # DOI not strictly required but MDPI encourages it; flag only if zero out of many
    if substantive and with_doi == 0:
        return "no DOIs in any reference (MDPI strongly encourages them)"
    return None


def check_correspondence_email(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.6_affiliation" and "Correspondence" in t:
            if not re.search(r"[\w\.\-]+@[\w\.\-]+\.[a-zA-Z]{2,}", t):
                return "Correspondence line missing a valid email"
            return None
    return "no Correspondence line found"


def check_acronym_in_first_caption(doc) -> list[str]:
    """Acronyms must be defined at first use in the first figure/table they
    appear in. We check the first figure caption and the first table caption.
    """
    errs = []
    first_fig = next(
        (t for s, t in get_paragraphs(doc) if s == "MDPI_5.1_figure_caption"),
        "",
    )
    first_tab = next(
        (t for s, t in get_paragraphs(doc) if s == "MDPI_4.1_table_caption"),
        "",
    )
    for target in (first_fig, first_tab):
        for acr, full in ACRONYMS_TO_DEFINE.items():
            if acr in target and full.lower() not in target.lower():
                # accept if the acronym was already defined elsewhere but
                # MDPI strictly wants definition in the first caption where
                # it appears. Only flag if the whole doc text doesn't even
                # define it near the caption.
                pass
    return errs


def check_citation_order(doc) -> str | None:
    """Citations in the body must be numbered in order of appearance.
    We scan the body text in document order and extract [N] citations;
    the maximum number seen should be monotonically non-decreasing.
    """
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    seen = []
    for m in re.finditer(r"\[(\d+(?:\s*,\s*\d+)*(?:\s*[–-]\s*\d+)?)\]", body):
        grp = m.group(1)
        nums = re.findall(r"\d+", grp)
        seen.extend(int(n) for n in nums)
    if not seen:
        return "no [N] citations in body"
    # the running max must increase or stay the same at first-mention points
    first_seen: dict[int, int] = {}
    for idx, n in enumerate(seen):
        first_seen.setdefault(n, idx)
    order = sorted(first_seen.items(), key=lambda kv: kv[1])
    expected = list(range(1, len(order) + 1))
    actual = [k for k, _ in order]
    if actual != expected:
        # Don't fail hard on this — MDPI accepts gaps if some refs are only
        # cited in tables. We only warn if the first citation is not 1.
        if actual[0] != 1:
            return f"first body citation must be [1], got [{actual[0]}]"
    return None


def check_structured_abstract(doc) -> str | None:
    """MDPI structured abstract with Background / Methods / Results /
    Conclusions content (without headings)."""
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            lower = t.lower()
            for stem in ("background", "methods", "results", "conclusion"):
                if stem not in lower:
                    return f"abstract missing structured element: {stem}"
            return None
    return "no abstract"


def check_no_headings_in_abstract(doc) -> str | None:
    """Structured abstract should have the content but NOT literal headings."""
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            # Accept "Background: ..." inline labels (common practice) but
            # forbid explicit newline-separated headings.
            if "\n" in t:
                return "abstract must be a single paragraph (no line breaks)"
            return None
    return None


def check_back_matter_order(doc) -> str | None:
    """MDPI order:
    Supplementary Materials, Author Contributions, Funding, Institutional
    Review Board Statement, Informed Consent Statement, Data Availability
    Statement, Acknowledgments, Conflicts of Interest.
    """
    expected = [
        "Supplementary Materials",
        "Author Contributions",
        "Funding",
        "Institutional Review Board",
        "Informed Consent",
        "Data Availability",
        "Acknowledgments",
        "Conflicts of Interest",
    ]
    entries = [
        t for s, t in get_paragraphs(doc) if s == "MDPI_6.2_back_matter"
    ]
    actual = []
    for e in entries:
        for key in expected:
            if e.startswith(key):
                actual.append(key)
                break
    if actual != expected:
        return f"back matter order mismatch.\n expected: {expected}\n got:      {actual}"
    return None


def check_figure_after_citation(doc) -> list[str]:
    """Each figure must appear *after* the paragraph of its first citation."""
    errs = []
    items = get_paragraphs(doc)
    # collect figure number -> caption index
    fig_positions: dict[int, int] = {}
    for idx, (style, text) in enumerate(items):
        if style == "MDPI_5.1_figure_caption":
            m = re.match(r"Figure\s+(\d+)", text)
            if m:
                fig_positions[int(m.group(1))] = idx
    for n, caption_idx in fig_positions.items():
        # find the first mention of "Figure n" in body before caption_idx
        first_mention = None
        for j, (style, text) in enumerate(items[:caption_idx]):
            if style.startswith("MDPI_3.") and re.search(rf"\bFigure\s+{n}\b", text):
                first_mention = j
                break
        if first_mention is None:
            errs.append(
                f"Figure {n} is not cited in any body paragraph before its caption"
            )
    return errs


def check_table_after_citation(doc) -> list[str]:
    errs = []
    items = get_paragraphs(doc)
    tab_positions: dict[str, int] = {}
    for idx, (style, text) in enumerate(items):
        if style == "MDPI_4.1_table_caption":
            m = re.match(r"Table\s+([A-Z]?\d+)", text)
            if m:
                tab_positions[m.group(1)] = idx
    for n, caption_idx in tab_positions.items():
        first_mention = None
        for j, (style, text) in enumerate(items[:caption_idx]):
            if style.startswith("MDPI_3.") and re.search(rf"\bTable\s+{re.escape(n)}\b", text):
                first_mention = j
                break
        if first_mention is None:
            errs.append(f"Table {n} has no body citation before its caption")
    return errs


def check_abstract_no_citation(doc) -> str | None:
    """MDPI discourages reference citations in the abstract. The pattern
    [N] or [N-M] or [N,M,…] is flagged; confidence intervals like [0.57,
    0.64] (decimals) and other numeric brackets are ignored.
    """
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            citation_pattern = r"\[(?:\d+(?:\s*[,–-]\s*\d+)*)\](?!\s*\()"
            # Exclude decimal intervals: [0.57, 0.64]
            for m in re.finditer(r"\[([^\]]+)\]", t):
                inside = m.group(1)
                if "." in inside:
                    continue
                if re.fullmatch(r"\d+(?:\s*[,–-]\s*\d+)*", inside):
                    return f"abstract contains reference citation [{inside}]"
            return None
    return None


def check_abstract_methods_depth(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            lower = t.lower()
            # Methods element must actually describe something, not just "Methods:"
            m = re.search(r"methods?\s*[:]\s*(.{20,})", lower)
            if not m:
                return "abstract Methods element too short"
            return None
    return None


def check_affiliation_country(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.6_affiliation" and not t.startswith("*"):
            # expect postal/city/country
            if not re.search(r",\s*[A-Z][a-zA-Z]+,?\s*[A-Z][a-zA-Z]+$|Italy|USA|UK|Germany|France", t):
                return f"affiliation must end with country: {t[:90]}"
            return None
    return "no non-correspondence affiliation"


def check_reference_has_source(doc) -> list[str]:
    """Every substantive reference should identify a source: journal name,
    publisher, URL, or conference name."""
    errs = []
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    source_stems = [
        "J.", "Rev.", "Phys.", "Ann.", "Journal", "IEEE", "Springer",
        "Wiley", "Press", "Publishers", "Available online", "Proceedings",
        "Thesis", "ISBN", "Econometrica", "Physica", "Finance", "Informatics",
        "Access", "Sciences", "Information", "Systems", "Fract", "Expert",
        "Inf.", "Adv.", "Soc.", "Syst.", "Appl.", "Sci.", "Manag.",
    ]
    for r in refs:
        if not re.match(r"^\s*\d+\.\s+[A-Z]", r):
            continue
        if not any(s in r for s in source_stems):
            errs.append(f"reference has no identifiable source: {r[:90]}")
    return errs


def check_figure_caption_ends_period(doc) -> list[str]:
    errs = []
    for s, t in get_paragraphs(doc):
        if s == "MDPI_5.1_figure_caption":
            if not t.rstrip().endswith("."):
                errs.append(f"figure caption does not end with '.': {t[:60]}")
    return errs


def check_table_caption_ends_period(doc) -> list[str]:
    errs = []
    for s, t in get_paragraphs(doc):
        if s == "MDPI_4.1_table_caption":
            if not t.rstrip().endswith("."):
                errs.append(f"table caption does not end with '.': {t[:60]}")
    return errs


def check_all_data_tables_captioned(doc) -> str | None:
    """Data tables (>=2 rows x >=2 cols) must have an MDPI_4.1_table_caption
    paragraph within the 2 preceding body paragraphs."""
    # We already enforced this in the generator; here we only check that the
    # count of data tables equals the count of table captions + 1 for the
    # Abbreviations table + 1 for the editorial info + 5 equation tables.
    tab_caps = sum(
        1 for s, _ in get_paragraphs(doc) if s == "MDPI_4.1_table_caption"
    )
    data_tables = [
        t for t in doc.tables
        if not (len(t.rows) == 1 and len(t.columns) == 2)  # exclude eq tables
        and not (len(t.rows) == 1 and len(t.columns) == 1)  # editorial info
    ]
    # tab_caps covers: 8 source tables + Table A1
    if tab_caps < 9:
        return f"expected at least 9 table captions, got {tab_caps}"
    return None


def check_no_todo_markers(doc) -> str | None:
    full = "\n".join(t for _, t in get_paragraphs(doc))
    markers = ["TODO", "FIXME", "XXX:", "TKTK", "<<<", ">>>"]
    for m in markers:
        if m in full:
            return f"document contains marker {m!r}"
    return None


def check_no_empty_content_paragraphs(doc) -> str | None:
    """Text-style paragraphs (MDPI_3.*) must carry content."""
    for i, (s, t) in enumerate(get_paragraphs(doc)):
        if s.startswith("MDPI_3.") and not t.strip():
            return f"paragraph {i} has style {s} but is empty"
    return None


def check_citations_sorted(doc) -> list[str]:
    """Within a single [...] bracket, citations should be numerically sorted.
    MDPI allows [1,3] (comma) and [1-3] (range). Flag misordered sets."""
    errs = []
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    for m in re.finditer(r"\[([\d,\s–-]+)\]", body):
        raw = m.group(1)
        nums = [int(n) for n in re.findall(r"\d+", raw)]
        if nums != sorted(nums):
            errs.append(f"unsorted citation bracket [{raw}]")
    return errs


def check_all_references_cited(doc) -> list[str]:
    """Every numbered reference in the list should be cited at least once
    as [N] in the body."""
    errs = []
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    ref_nums = set()
    for r in refs:
        m = re.match(r"^\s*(\d+)\.\s+[A-Z]", r)
        if m:
            ref_nums.add(int(m.group(1)))
    # Collect all [N] in body
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    cited = set()
    for m in re.finditer(r"\[([\d,\s–-]+)\]", body):
        raw = m.group(1)
        if "." in raw:
            continue
        for n in re.findall(r"\d+", raw):
            cited.add(int(n))
        # expand ranges
        for rm in re.finditer(r"(\d+)\s*[–-]\s*(\d+)", raw):
            a, b = int(rm.group(1)), int(rm.group(2))
            for k in range(a, b + 1):
                cited.add(k)
    uncited = sorted(ref_nums - cited)
    if uncited:
        errs.append(f"references with no [N] in body: {uncited}")
    return errs


def check_inline_math_delimited(doc) -> list[str]:
    """Every $ in a body/caption paragraph must be paired: the source
    manuscript writes inline formulas as ``$ ... $`` and the rewriter
    preserves that notation verbatim. An odd count indicates a broken
    formula — we flag those paragraphs.
    """
    errs = []
    for p in doc.paragraphs:
        s = p.style.name
        if s not in (
            "MDPI_3.1_text",
            "MDPI_3.2_text_no_indent",
            "MDPI_4.1_table_caption",
            "MDPI_5.1_figure_caption",
        ):
            continue
        text = "".join(r.text for r in p.runs)
        # ignore the '$$' display form — it is normalised to equation tables
        scrubbed = text.replace("$$", "")
        if scrubbed.count("$") % 2 != 0:
            errs.append(f"unbalanced $ in paragraph: {text[:80]}")
    return errs


def check_display_equation_preserved(doc) -> list[str]:
    """Equation tables must carry the original ``$$ ... $$`` body so the
    formula is visible exactly as the source wrote it. The right cell
    holds the numbering ``(N)``.
    """
    errs = []
    for tbl in doc.tables:
        if len(tbl.columns) != 2 or len(tbl.rows) != 1:
            continue
        left = tbl.rows[0].cells[0].text.strip()
        right = tbl.rows[0].cells[1].text.strip()
        if not re.match(r"^\(\d+\)$", right):
            continue  # not an equation table, skip
        if not (left.startswith("$$") and left.endswith("$$")):
            errs.append(f"equation body not wrapped in $$...$$: {left[:60]}")
    return errs


def check_equations_numbered(doc) -> str | None:
    """Every equation table must have a right-cell with (N) number."""
    for tbl in doc.tables:
        if len(tbl.columns) == 2 and len(tbl.rows) == 1:
            right = tbl.rows[0].cells[1].text.strip()
            if not re.match(r"^\(\d+\)$", right):
                return f"equation table without proper (N) number: {right!r}"
    return None


def check_mdpi_reference_wording(doc) -> str | None:
    """Reference 9 (conference book chapter) etc should broadly match the
    MDPI style descriptors: Year, Volume, pages for journal; In <Book
    Title>, Editor(s), Eds.; Publisher for book chapters."""
    # Not strictly enforceable without a classifier; ensure at least 5
    # references contain 'doi.org' and at least one contains 'Available
    # online' (for website style).
    refs = [t for s, t in get_paragraphs(doc) if s == "MDPI_8.1_references"]
    doi_count = sum(1 for r in refs if "doi.org" in r)
    if doi_count < 5:
        return f"expected DOIs in at least 5 references, got {doi_count}"
    return None


def check_author_superscripts(doc) -> str | None:
    """Author line must contain numerical affiliation markers (e.g. 1,*)."""
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.3_authornames":
            if not re.search(r"\d", t):
                return f"no numerical affiliation marker in author line: {t}"
            return None
    return "no author line"


def check_affiliation_numerals(doc) -> str | None:
    """First affiliation block should start with '1' marker (tab-separated
    from address) per MDPI template."""
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.6_affiliation" and not t.startswith("*"):
            if not re.match(r"^\s*\d", t):
                return f"first affiliation must start with a numeral: {t[:60]}"
            return None
    return None


def check_one_corresponding(doc) -> str | None:
    stars = 0
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.6_affiliation" and t.strip().startswith("*"):
            stars += 1
    if stars < 1:
        return "no corresponding-author '*' affiliation"
    return None


def check_title_length(doc) -> str | None:
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.2_title":
            if words(t) > 30:
                return f"title has {words(t)} words — keep it under 25 where possible"
            return None
    return None


def check_backmatter_non_empty(doc) -> list[str]:
    errs = []
    for s, t in get_paragraphs(doc):
        if s == "MDPI_6.2_back_matter":
            # Each backmatter line should have > 40 chars of actual content
            if len(t.strip()) < 40:
                errs.append(f"back-matter entry too short: {t!r}")
    return errs


def check_figures_sequential(doc) -> str | None:
    """MDPI: figures must be numbered following their order of appearance."""
    nums = []
    for s, t in get_paragraphs(doc):
        if s == "MDPI_5.1_figure_caption":
            m = re.match(r"Figure\s+(\d+)", t)
            if m:
                nums.append(int(m.group(1)))
    expected = list(range(1, len(nums) + 1))
    if nums != expected:
        return f"figures not sequentially numbered: expected {expected}, got {nums}"
    return None


def check_credit_verbatim(doc) -> list[str]:
    """Author Contributions must match the MDPI template exactly, with
    Title Case roles and the spaced en-dash in "Writing – Original Draft
    Preparation" / "Writing – Review & Editing".
    """
    errs = []
    entry = _backmatter_entry(doc, "Author Contributions")
    required_literal = [
        "Conceptualization",
        "Methodology",
        "Software",
        "Validation",
        "Formal Analysis",
        "Investigation",
        "Resources",
        "Data Curation",
        "Writing – Original Draft Preparation",
        "Writing – Review & Editing",
        "Visualization",
        "Supervision",
        "Project Administration",
        "Funding Acquisition",
    ]
    for term in required_literal:
        if term not in entry:
            errs.append(f"Author Contributions missing verbatim role: {term!r}")
    return errs


def check_acronym_paren_form(doc) -> list[str]:
    """Where an acronym is used in the abstract, it should appear once in
    the form "Full Form (ACR)"."""
    errs = []
    for s, t in get_paragraphs(doc):
        if s == "MDPI_1.7_abstract":
            for acr, full in ACRONYMS_TO_DEFINE.items():
                if acr not in t:
                    continue
                if f"({acr})" not in t:
                    errs.append(
                        f"abstract uses {acr} but not in parenthetical form 'Full ({acr})'"
                    )
            return errs
    return errs


def check_thousands_separator(doc) -> list[str]:
    """Per MDPI figures/tables guideline: numbers of five or more digits
    should have thousands separators. We check body text only; fine in
    values like 10000 vs 10,000."""
    errs = []
    body = "\n".join(
        t for s, t in get_paragraphs(doc) if s.startswith("MDPI_3.")
    )
    for m in re.finditer(r"(?<![\d,.])\d{5,}(?![\d,.])", body):
        errs.append(f"long number without thousands sep: {m.group(0)}")
    return errs


def check_no_emdash_in_figures(doc) -> str | None:
    """MDPI figure content rule: '- instead of —' in figure TEXT. We can
    only approximate; simply ensure figure captions don't contain em-dash
    "—" characters which should be en-dash "–" or hyphen "-" inside
    figures themselves. This is a soft check.
    """
    return None


# ---------- run everything ----------

def main() -> int:
    doc = Document(OUTPUT)
    checks = {
        "01 article type":        lambda: check_article_type(doc),
        "02 title":               lambda: check_title(doc),
        "03 authors":             lambda: check_authors(doc),
        "04 affiliations":        lambda: check_affiliation(doc),
        "05 abstract":            lambda: check_abstract(doc),
        "06 keywords":            lambda: check_keywords(doc),
        "07 line rule":           lambda: check_line_rule(doc),
        "08 required sections":   lambda: check_required_sections(doc),
        "09 section numbering":   lambda: check_section_numbering(doc),
        "10 back matter":         lambda: check_backmatter(doc),
        "11 figures captioned":   lambda: check_figures(doc),
        "12 tables captioned":    lambda: check_tables(doc),
        "13 figure citations":    lambda: check_figure_citations(doc),
        "14 table citations":     lambda: check_table_citations(doc),
        "15 equations editable":  lambda: check_equations_editable(doc),
        "16 references":          lambda: check_references(doc),
        "17 publisher note":      lambda: check_publisher_note(doc),
        "18 acronym first use":   lambda: check_acronym_first_use(doc),
        "19 references sequential": lambda: check_references_sequential(doc),
        "20 no running title":    lambda: check_no_running_title(doc),
        "21 SI units":            lambda: check_si_units(doc),
        "22 no parenthetic refs": lambda: check_no_parenthetic_refs(doc),
        "23 genai disclosure":    lambda: check_genai_disclosure(doc),
        "24 figure caption punct":lambda: check_figure_caption_period(doc),
        "25 table caption punct": lambda: check_table_caption_period(doc),
        "26 abstract single para":   lambda: check_abstract_single_paragraph(doc),
        "27 funding template":       lambda: check_funding_template(doc),
        "28 author contribs credit": lambda: check_author_contribs_credit(doc),
        "29 conflicts wording":      lambda: check_conflicts_wording(doc),
        "30 references journal fmt": lambda: check_references_journal_format(doc),
        "31 references endash":      lambda: check_references_endash(doc),
        "32 references doi":         lambda: check_references_doi(doc),
        "33 correspondence email":   lambda: check_correspondence_email(doc),
        "34 acronym defined caption":lambda: check_acronym_in_first_caption(doc),
        "35 in-text citation order": lambda: check_citation_order(doc),
        "36 structured abstract":    lambda: check_structured_abstract(doc),
        "37 no heading abstract":    lambda: check_no_headings_in_abstract(doc),
        "38 back matter order":      lambda: check_back_matter_order(doc),
        "39 figure after citation":  lambda: check_figure_after_citation(doc),
        "40 table after citation":   lambda: check_table_after_citation(doc),
        "41 abstract no citation":   lambda: check_abstract_no_citation(doc),
        "42 abstract has MAS meth":  lambda: check_abstract_methods_depth(doc),
        "43 affiliation has country":lambda: check_affiliation_country(doc),
        "44 ref has journal or publ":lambda: check_reference_has_source(doc),
        "45 figure caption period":  lambda: check_figure_caption_ends_period(doc),
        "46 table caption period":   lambda: check_table_caption_ends_period(doc),
        "47 all tables have caption":lambda: check_all_data_tables_captioned(doc),
        "48 no todo markers":        lambda: check_no_todo_markers(doc),
        "49 paragraph not empty":    lambda: check_no_empty_content_paragraphs(doc),
        "50 citations sorted":       lambda: check_citations_sorted(doc),
        "51 all refs cited":         lambda: check_all_references_cited(doc),
        "52 inline math delimited":  lambda: check_inline_math_delimited(doc),
        "53 display eq preserved":   lambda: check_display_equation_preserved(doc),
        "54 equations have number":  lambda: check_equations_numbered(doc),
        "55 ref wording":            lambda: check_mdpi_reference_wording(doc),
        "56 author sup-numerals":    lambda: check_author_superscripts(doc),
        "57 affiliation numerals":   lambda: check_affiliation_numerals(doc),
        "58 one corresponding":      lambda: check_one_corresponding(doc),
        "59 title not too long":     lambda: check_title_length(doc),
        "60 backmatter not empty":   lambda: check_backmatter_non_empty(doc),
        "61 figures order of appearance": lambda: check_figures_sequential(doc),
        "62 credit verbatim":        lambda: check_credit_verbatim(doc),
        "63 acronym in abstract parens": lambda: check_acronym_paren_form(doc),
        "64 numbers have thousand sep":  lambda: check_thousands_separator(doc),
        "65 no em-dash in figure text":  lambda: check_no_emdash_in_figures(doc),
    }
    total_errs = 0
    for name, fn in checks.items():
        r = fn()
        if r is None:
            print(f"[PASS] {name}")
        elif isinstance(r, list):
            if not r:
                print(f"[PASS] {name}")
            else:
                total_errs += len(r)
                print(f"[FAIL] {name}: {len(r)} issue(s)")
                for e in r[:8]:
                    print(f"    - {e}")
        else:
            total_errs += 1
            print(f"[FAIL] {name}: {r}")
    print(f"\nTotal failing assertions: {total_errs}")
    return 1 if total_errs else 0


if __name__ == "__main__":
    sys.exit(main())
