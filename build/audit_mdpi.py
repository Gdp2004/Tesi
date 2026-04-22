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
        "Formal analysis",
        "Investigation",
        "Writing—original draft",
        "Writing—review",
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
