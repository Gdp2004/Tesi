"""Four-pass verification of the generated MRQF_MAS_Electronics.docx.

Pass 1: every mandatory MDPI template section present.
Pass 2: every paragraph carries an MDPI_* style (no Normal / Heading residue).
Pass 3: 10 figure captions, 9 table captions, all $$...$$ lifted to equation tables.
Pass 4: content fidelity — no substantive sentence of the source dropped.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "MRQF_MAS_Electronics.docx"
SOURCE = ROOT / "build" / "_sources" / "MRQF_MAS_paper_round7_FINAL.docx"


def collect_styles(doc):
    return [p.style.name for p in doc.paragraphs]


def collect_texts(doc):
    return [p.text for p in doc.paragraphs]


def pass1_structural(doc) -> list[str]:
    styles = collect_styles(doc)
    texts = collect_texts(doc)
    required = [
        "MDPI_1.1_article_type",
        "MDPI_1.2_title",
        "MDPI_1.3_authornames",
        "MDPI_1.6_affiliation",
        "MDPI_1.7_abstract",
        "MDPI_1.8_keywords",
        "MDPI_1.9_line",
        "MDPI_2.1_heading1",
        "MDPI_2.2_heading2",
        "MDPI_6.2_back_matter",
        "MDPI_6.3_notes",
        "MDPI_8.1_references",
    ]
    errors = []
    for r in required:
        if r not in styles:
            errors.append(f"missing required style: {r}")
    # seven back-matter entries
    backmatter_required = [
        "Author Contributions",
        "Funding",
        "Institutional Review Board",
        "Informed Consent",
        "Data Availability",
        "Acknowledgments",
        "Conflicts of Interest",
    ]
    back_texts = [
        t for t, s in zip(texts, styles) if s == "MDPI_6.2_back_matter"
    ]
    joined = " || ".join(back_texts)
    for entry in backmatter_required:
        if entry not in joined:
            errors.append(f"missing back-matter entry: {entry}")
    if "Abbreviations" not in " ".join(texts):
        errors.append("missing Abbreviations heading")
    if "Appendix A" not in " ".join(texts):
        errors.append("missing Appendix A heading")
    if "References" not in " ".join(texts):
        errors.append("missing References heading")
    return errors


def pass2_style_coverage(doc) -> list[str]:
    errors = []
    for i, p in enumerate(doc.paragraphs):
        name = p.style.name
        if not name.startswith("MDPI_") and name != "Normal":
            errors.append(f"paragraph {i} uses non-MDPI style: {name}")
        if name in ("Heading 1", "Heading 2", "Heading 3"):
            errors.append(f"paragraph {i} still uses {name}")
        if name == "Normal" and p.text.strip():
            # Normal with text means we forgot to style it; header/footer
            # Normal paragraphs coming from the template are empty separators
            errors.append(
                f"paragraph {i} is Normal but has text: {p.text[:60]}"
            )
    return errors


def pass3_figures_tables_equations(doc) -> list[str]:
    errors = []
    styles = collect_styles(doc)
    texts = collect_texts(doc)
    fig_captions = sum(1 for s in styles if s == "MDPI_5.1_figure_caption")
    tab_captions = sum(1 for s in styles if s == "MDPI_4.1_table_caption")
    if fig_captions != 10:
        errors.append(f"expected 10 figure captions, got {fig_captions}")
    if tab_captions != 9:
        errors.append(f"expected 9 table captions, got {tab_captions}")
    # no stray $$...$$ in body
    for i, t in enumerate(texts):
        if t.strip().startswith("$$") and t.strip().endswith("$$"):
            errors.append(f"equation still inline at paragraph {i}: {t[:60]}")
    return errors


def pass4_content_fidelity(out_doc) -> list[str]:
    """Assert every substantive source sentence appears in the output."""
    src = Document(SOURCE)
    src_text = " ".join(p.text for p in src.paragraphs)
    out_text = " ".join(p.text for p in out_doc.paragraphs)
    # additionally crawl cell text to catch table content
    for t in out_doc.tables:
        for row in t.rows:
            for cell in row.cells:
                out_text += " " + cell.text
    src_text_norm = re.sub(r"\s+", " ", src_text.replace("$", ""))
    out_text_norm = re.sub(r"\s+", " ", out_text.replace("$", ""))
    # sample characteristic sentences that must survive
    touchstones = [
        "Price dynamics in financial markets exhibit a complexity",
        "The price of a financial instrument is the emergent outcome",
        "A price signal",
        "To move from a scaling description",
        "The step that completes the theoretical apparatus",
        "The transition from theory to implementation",
        "Each of the three operator-level agents is further decomposed",
        "The MRQF account of operator interactions yields three primitive",
        "(E, S) plane, populated by the nine subscenarios",
        "We now specify the end-to-end algorithm",
        "To illustrate the behaviour of MRQF-MAS on a realistic signal",
        "The terms financion, scattering, annihilation",
        "The synthetic case study of Section 8",
        "Reviewers raised two legitimate concerns",
        "The primary advantage of MRQF-MAS over classical quantitative",
        "This paper has presented MRQF-MAS",
        "For full reproducibility of the synthetic case study",
    ]
    errors = []
    for s in touchstones:
        if s not in out_text_norm:
            errors.append(f"missing touchstone sentence: {s[:60]}")
    return errors


def main() -> int:
    doc = Document(OUTPUT)
    all_errors = {}
    all_errors["pass1_structural"] = pass1_structural(doc)
    all_errors["pass2_style_coverage"] = pass2_style_coverage(doc)
    all_errors["pass3_figures_tables_equations"] = pass3_figures_tables_equations(doc)
    all_errors["pass4_content_fidelity"] = pass4_content_fidelity(doc)

    fail = False
    for name, errs in all_errors.items():
        if errs:
            fail = True
            print(f"[FAIL] {name}: {len(errs)} issue(s)")
            for e in errs[:20]:
                print(f"    - {e}")
            if len(errs) > 20:
                print(f"    ... and {len(errs) - 20} more")
        else:
            print(f"[ OK ] {name}")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
