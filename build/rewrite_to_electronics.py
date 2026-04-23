"""Rewrite MRQF-MAS manuscript into the Electronics (MDPI) template.

Input  : ../_sources/electronicstemplate.dot  (MDPI Electronics template)
         ../_sources/MRQF_MAS_paper_round7_FINAL.docx (source manuscript)
Output : ../MRQF_MAS_Electronics.docx (submission-ready copy)

The template ships with a complete set of named MDPI_* paragraph styles and a
mandatory structural skeleton (article-type, title, authors, affiliations,
editorial info table, abstract, keywords, horizontal rule, numbered sections,
figure/table captions, back-matter, Abbreviations, Appendix, References and
Publisher's Note). We preserve that skeleton by *cloning* the template file
itself, wiping only its illustrative body while keeping the style tree, theme,
headers/footers and editorial info table, and then populating the body with
MRQF content in the correct MDPI styles.
"""

from __future__ import annotations

import copy
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parent.parent
BUILD = ROOT / "build"
SOURCES = BUILD / "_sources"
MEDIA = BUILD / "media"

TEMPLATE_DOT = SOURCES / "electronicstemplate.dot"
MRQF_DOCX = SOURCES / "MRQF_MAS_paper_round7_FINAL.docx"
OUTPUT = ROOT / "MRQF_MAS_Electronics.docx"
WORK_TEMPLATE = BUILD / "_template_fixed.docx"


# ---------------------------------------------------------------------------
# Step 1 — patch the template so python-docx will open it
# ---------------------------------------------------------------------------

def patch_template_content_type(src: Path, dst: Path) -> None:
    """Convert ``.dot`` (template content-type) into ``.docx`` (document).

    python-docx rejects files whose content type is the Word template form;
    swapping the string inside ``[Content_Types].xml`` is enough to make the
    file behave as a regular ``.docx`` while preserving every style,
    relationship and embedded resource bit-for-bit.
    """
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument."
                    b"wordprocessingml.template.main+xml",
                    b"application/vnd.openxmlformats-officedocument."
                    b"wordprocessingml.document.main+xml",
                )
            zout.writestr(item, data)


# ---------------------------------------------------------------------------
# Step 2 — iterate body elements of a docx in document order
# ---------------------------------------------------------------------------

def iter_block_items(element):
    for child in element.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag in ("p", "tbl"):
            yield child


def para_text(p) -> str:
    texts = p.findall(".//" + qn("w:t"))
    return "".join(t.text or "" for t in texts)


def para_style(p) -> str:
    s = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return s.get(qn("w:val")) if s is not None else "Normal"


def has_drawing(p) -> bool:
    return bool(p.findall(".//" + qn("w:drawing")))


# ---------------------------------------------------------------------------
# Step 3 — source parsing helpers
# ---------------------------------------------------------------------------

def load_source_stream():
    """Return a list of typed blocks from the MRQF manuscript.

    Block kinds: ``("para", idx, style, text, has_img)`` and
    ``("table", idx, Table)``.
    """
    doc = Document(MRQF_DOCX)
    blocks = []
    tables_iter = iter(doc.tables)
    para_idx = 0
    table_idx = 0
    body = doc.element.body
    for el in iter_block_items(body):
        tag = el.tag.split("}")[-1]
        if tag == "p":
            blocks.append(
                (
                    "para",
                    para_idx,
                    para_style(el),
                    para_text(el),
                    has_drawing(el),
                )
            )
            para_idx += 1
        else:
            tbl = next(tables_iter)
            blocks.append(("table", table_idx, tbl))
            table_idx += 1
    return blocks


# ---------------------------------------------------------------------------
# Step 4 — prepare output document
# ---------------------------------------------------------------------------

def prepare_output_document() -> Document:
    """Load the patched template and wipe its placeholder content."""
    patch_template_content_type(TEMPLATE_DOT, WORK_TEMPLATE)
    doc = Document(WORK_TEMPLATE)
    body = doc.element.body

    sect_pr = None
    editorial_table = None  # first table in the template = editorial info

    to_remove = []
    for idx, child in enumerate(list(body)):
        tag = child.tag.split("}")[-1]
        if tag == "sectPr":
            sect_pr = child
            continue
        if tag == "tbl" and editorial_table is None:
            editorial_table = child
            continue
        to_remove.append(child)
    for child in to_remove:
        body.remove(child)

    return doc


# ---------------------------------------------------------------------------
# Step 5 — writers
# ---------------------------------------------------------------------------

def _strip_inline_maths(text: str) -> str:
    """Remove $…$ delimiters and convert LaTeX inside to Unicode. Used for
    captions and other single-run contexts where italic runs are not set.
    """
    out = []
    in_math = False
    buf: list[str] = []
    for ch in text:
        if ch == "$":
            chunk = "".join(buf)
            if in_math:
                chunk = latex_to_unicode(chunk)
            out.append(chunk)
            buf = []
            in_math = not in_math
            continue
        buf.append(ch)
    tail = "".join(buf)
    if in_math:
        tail = latex_to_unicode(tail)
    out.append(tail)
    return "".join(out)


def add_para(doc: Document, style: str, text: str = ""):
    # Callers pass plain text with optional $…$ inline maths; we normalise
    # so neither the $ delimiter nor the LaTeX macros survive in the docx.
    p = doc.add_paragraph(_strip_inline_maths(text) if text else "")
    p.style = doc.styles[style]
    return p


def add_inline_math_para(doc: Document, style: str, text: str):
    """Body text with $...$ inline math rendered as italic runs."""
    p = doc.add_paragraph()
    p.style = doc.styles[style]
    _emit_with_inline_math(p, text)
    return p


def _emit_with_inline_math(p, text: str):
    """Split on $...$ boundaries and emit italic runs for the maths.

    LaTeX macros inside math segments are converted to Unicode.
    """
    in_math = False
    i = 0
    current_run_text = []
    while i < len(text):
        ch = text[i]
        if ch == "$":
            if current_run_text:
                chunk = "".join(current_run_text)
                if in_math:
                    chunk = latex_to_unicode(chunk)
                run = p.add_run(chunk)
                run.italic = in_math
                current_run_text = []
            in_math = not in_math
            i += 1
            continue
        current_run_text.append(ch)
        i += 1
    if current_run_text:
        chunk = "".join(current_run_text)
        if in_math:
            chunk = latex_to_unicode(chunk)
        run = p.add_run(chunk)
        run.italic = in_math


def add_figure(doc: Document, image_path: Path, caption: str):
    p = doc.add_paragraph()
    p.style = doc.styles["MDPI_5.2_figure"]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(str(image_path), width=Cm(14))
    cap = doc.add_paragraph(renumber_figures(_strip_inline_maths(caption)))
    cap.style = doc.styles["MDPI_5.1_figure_caption"]
    return cap


def add_equation(doc: Document, latex_body: str, number: int):
    """Insert a two-column equation table (equation | number)."""
    table = doc.add_table(rows=1, cols=2)
    table.autofit = True
    left, right = table.rows[0].cells
    p_left = left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_left.add_run(latex_to_unicode(latex_body.strip()))
    run.italic = True
    p_right = right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run(f"({number})")
    # make left cell wider
    left.width = Cm(13)
    right.width = Cm(2)


def add_table_caption(doc: Document, caption: str):
    p = doc.add_paragraph(_strip_inline_maths(caption))
    p.style = doc.styles["MDPI_4.1_table_caption"]


def add_data_table(doc: Document, header: list[str], rows: list[list[str]]):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(header))
    try:
        tbl.style = doc.styles["Table Grid"]
    except KeyError:
        pass
    for j, h in enumerate(header):
        c = tbl.rows[0].cells[j]
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.bold = True
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            tbl.rows[i + 1].cells[j].text = val
    return tbl


def copy_source_table(doc: Document, src_table):
    """Clone a source docx table into the destination document.

    Normalise every text node in the clone so LaTeX macros are converted to
    Unicode and $…$ delimiters are removed.
    """
    tbl_xml = copy.deepcopy(src_table._tbl)
    doc.element.body.append(tbl_xml)
    for pStyle in tbl_xml.iter(qn("w:tblStyle")):
        pStyle.set(qn("w:val"), "Table Grid")
    for t_el in tbl_xml.iter(qn("w:t")):
        if t_el.text:
            t_el.text = renumber_figures(_strip_inline_maths(t_el.text))


# ---------------------------------------------------------------------------
# Step 6 — the rewrite driver
# ---------------------------------------------------------------------------

STYLE_MAP = {
    "Heading1": "MDPI_2.1_heading1",
    "Heading2": "MDPI_2.2_heading2",
    "Heading3": "MDPI_2.3_heading3",
}


def is_equation_para(text: str) -> bool:
    t = text.strip()
    return t.startswith("$$") and t.endswith("$$") and len(t) > 4


def is_figure_caption(text: str) -> bool:
    return text.strip().startswith("Figure ")


def is_table_caption(text: str) -> bool:
    return text.strip().startswith("Table ")


REFERENCE_RE = re.compile(r"^\s*\d{1,3}\.\s+\S")


# MDPI requires editable maths rather than raw LaTeX text in body paragraphs.
# We normalise the LaTeX macros produced by the source manuscript to their
# Unicode equivalents.
LATEX_GREEK = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
    r"\epsilon": "ε", r"\varepsilon": "ε", r"\zeta": "ζ", r"\eta": "η",
    r"\theta": "θ", r"\vartheta": "ϑ", r"\iota": "ι", r"\kappa": "κ",
    r"\lambda": "λ", r"\mu": "μ", r"\nu": "ν", r"\xi": "ξ",
    r"\pi": "π", r"\varpi": "ϖ", r"\rho": "ρ", r"\varrho": "ϱ",
    r"\sigma": "σ", r"\varsigma": "ς", r"\tau": "τ", r"\upsilon": "υ",
    r"\phi": "φ", r"\varphi": "φ", r"\chi": "χ", r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ", r"\Lambda": "Λ",
    r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ", r"\Upsilon": "Υ",
    r"\Phi": "Φ", r"\Psi": "Ψ", r"\Omega": "Ω",
}
LATEX_OPERATORS = {
    r"\to": "→", r"\Rightarrow": "⇒", r"\Leftrightarrow": "⇔",
    r"\le": "≤", r"\leq": "≤", r"\ge": "≥", r"\geq": "≥",
    r"\ne": "≠", r"\neq": "≠", r"\approx": "≈", r"\equiv": "≡",
    r"\pm": "±", r"\mp": "∓", r"\times": "×", r"\cdot": "·",
    r"\infty": "∞", r"\partial": "∂", r"\nabla": "∇",
    r"\sum": "∑", r"\prod": "∏", r"\int": "∫",
    r"\langle": "⟨", r"\rangle": "⟩", r"\star": "⋆", r"\ast": "∗",
    r"\leftrightarrow": "↔",
}
LATEX_STRIP = [
    r"\,", r"\;", r"\:", r"\!", r"\\", r"\displaystyle",
    r"\bigl", r"\bigr", r"\Bigl", r"\Bigr", r"\big", r"\Big",
    r"\left", r"\right",
]
LATEX_MATHCAL = {
    r"\mathcal{A}": "𝒜", r"\mathcal{B}": "ℬ", r"\mathcal{C}": "𝒞",
    r"\mathcal{D}": "𝒟", r"\mathcal{E}": "ℰ", r"\mathcal{F}": "ℱ",
    r"\mathcal{N}": "𝒩", r"\mathcal{O}": "𝒪", r"\mathcal{R}": "ℛ",
    r"\mathcal{S}": "𝒮", r"\mathcal{U}": "𝒰",
}


def latex_to_unicode(s: str) -> str:
    """Best-effort LaTeX → Unicode conversion for inline maths."""
    if not s:
        return s
    # 1) calligraphic mathcal macros first (they contain braces)
    for k, v in LATEX_MATHCAL.items():
        s = s.replace(k, v)
    # 2) \text{…} → plain text
    s = re.sub(r"\\text\{([^}]*)\}", r"\1", s)
    s = re.sub(r"\\mathrm\{([^}]*)\}", r"\1", s)
    s = re.sub(r"\\mathbf\{([^}]*)\}", r"\1", s)
    # 3) \frac{a}{b} → a/b
    s = re.sub(r"\\frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
    # 4) Greek letters and operators
    for k, v in LATEX_GREEK.items():
        s = s.replace(k, v)
    for k, v in LATEX_OPERATORS.items():
        s = s.replace(k, v)
    # 5) strip spacing commands
    for tok in LATEX_STRIP:
        s = s.replace(tok, "")
    # 6) collapse residual braces and backslash spaces
    s = re.sub(r"\\ ", " ", s)
    s = re.sub(r"\\([a-zA-Z]+)", r"\1", s)  # any unknown \macro → macro name
    s = s.replace("{", "").replace("}", "")
    # 7) tidy double spaces created by strips
    s = re.sub(r"  +", " ", s)
    return s


# MDPI requires every table to be cited in the body as "Table N". The source
# manuscript only has an explicit "Table N" reference for a subset of tables;
# we inject the missing citations at the natural end of the paragraph that
# precedes the table, without altering the author's argument. Keys are source
# paragraph indices, values are the sentence to append.
# MDPI requires acronyms to be defined at first use in each of: abstract,
# main text, and first figure/table. The abstract is rewritten with the
# expansions; for the body, we perform a single-pass replacement at the
# first occurrence of the acronym.
ACRONYM_FIRST_USE_REPLACEMENTS = {
    # find -> replace-with (only first occurrence)
    "GARCH-like": "Generalized Autoregressive Conditional Heteroscedasticity (GARCH)-like",
}


TABLE_CITATION_INJECTIONS = {
    58: " Table 1 summarises the resulting mapping from physics (QED) "
        "through MRQF finance to the MRQF-MAS software layer.",
    62: " Table 2 summarises the inputs, outputs and primary role of each "
        "sub-agent.",
    79: " Table 3 summarises the activation region, the dominant agent and "
        "the size policy associated with each strategy class.",
    130: " Table 4 reports the metrics produced by MRQF-MAS on this "
         "synthetic realisation.",
    146: " Table 8 decomposes these metrics across the pre-split and "
         "post-split sub-samples.",
    147: " The per-episode detection rates are reported in Table 7.",
    192: " Table A1 summarises the notation used throughout the paper.",
}


# MDPI requires each figure to be cited in the body *before* its caption.
# Where the source places a figure caption before its first narrative
# mention, we append a forward reference to the paragraph that introduces
# the figure.
FIGURE_CITATION_INJECTIONS = {
    125: " An overview of the resulting tick series is shown in Figure 4.",
}


# Ordered list of image files as they appear in the MRQF document. The source
# manuscript numbers them inconsistently with document order; MDPI rule
# (section "Preparing Figures, Schemes and Tables") requires figures to be
# numbered following their order of appearance. We renumber both the caption
# and every body-text reference using FIGURE_RENUMBER below.
IMAGE_ORDER = [
    "86dbf103603466e89edd3bb8f86d0ec50c76bfba.png",  # source Figure 1
    "9f69b37eab66096cea01f44f7ffe53df1b253963.png",  # source Figure 2
    "77d7691f31d0656d51a02e6ab2029d3f06ab9164.png",  # source Figure 3
    "04ad1bf59b0b7a48f944f24d8219df4080c18cd5.png",  # source Figure 5
    "63f6b8d4790b659471f2264181753247c162d360.png",  # source Figure 4
    "5f959ce10d672d7bd140eeedf1d0ac432b454c2b.png",  # source Figure 7
    "5c3b2910fa051cef9634a1af41ddbb39fbfc2e45.png",  # source Figure 10
    "0dde4cad5cf01590766ae782390fc23ab711029d.png",  # source Figure 8
    "d4fd55400f8fc85be9fd0bddc6b5af062d7e2196.png",  # source Figure 9
    "01ec5859f3b859a750561738bdbd2c9fb62107cf.png",  # source Figure 6
]

# old figure number -> new figure number (order-of-appearance)
FIGURE_RENUMBER = {
    1: 1, 2: 2, 3: 3,
    4: 5,  # case-study figure appears after the pipeline figure
    5: 4,  # pipeline figure appears first
    6: 10,
    7: 6,
    8: 8,
    9: 9,
    10: 7,
}


_FIG_RE = re.compile(r"\bFigure\s+(\d+)")


def renumber_figures(text: str) -> str:
    if not text:
        return text
    return _FIG_RE.sub(
        lambda m: f"Figure {FIGURE_RENUMBER.get(int(m.group(1)), int(m.group(1)))}",
        text,
    )

# Paragraphs that make up Algorithm 1 (pseudocode) in the source
ALG_START = 94
ALG_END = 123  # inclusive


def rewrite() -> None:
    blocks = load_source_stream()
    doc = prepare_output_document()

    # ---------- front matter ----------
    add_para(doc, "MDPI_1.1_article_type", "Article")

    # title
    title_block = next(b for b in blocks if b[0] == "para" and b[1] == 0)
    add_para(doc, "MDPI_1.2_title", title_block[3])

    # authors
    authors_block = next(b for b in blocks if b[0] == "para" and b[1] == 1)
    add_para(doc, "MDPI_1.3_authornames", authors_block[3])

    # affiliations — split the single source line into the mandatory 2+ lines
    affil_text = next(b for b in blocks if b[0] == "para" and b[1] == 2)[3]
    add_para(
        doc,
        "MDPI_1.6_affiliation",
        "1\tDepartment of Computer Science, University of Salerno, "
        "Via Giovanni Paolo II 132, 84084 Fisciano (SA), Italy; "
        "author@unisa.it",
    )
    add_para(
        doc,
        "MDPI_1.6_affiliation",
        "*\tCorrespondence: author@unisa.it",
    )

    # abstract — compressed to <=200 words, structured (Background / Methods /
    # Results / Conclusions) per MDPI Electronics Instructions, with the
    # acronyms MAS and GARCH spelled out at first use.
    abstract_body = (
        "Background: price dynamics in financial markets exhibit "
        "scale-invariant volatility, quantised liquidity and collective "
        "behaviour that resist single-paradigm models; Multiscale "
        "Relativistic Quantum Finance (MRQF) reconciles these facets on an "
        "energy-entropy (E,S) plane, but its translation into a deployable "
        "decision system has remained open. Methods: we propose MRQF-MAS, a "
        "cooperative Multi-Agent System (MAS) in which institutional, "
        "commercial and retail operators become first-class agents, each "
        "decomposed into signal, energy, entropy, risk and execution "
        "sub-agents that share beliefs through a horizontal cooperation "
        "layer and a Shared Knowledge Base (SKB) of (E,S) trajectories. The "
        "framework is benchmarked as a high-volatility regime classifier "
        "on 3840 daily EUR/USD reference rates published by the European "
        "Central Bank (ECB) over 1999-2026 against four baselines including "
        "Generalized Autoregressive Conditional Heteroscedasticity "
        "(GARCH)(1,1). Results: MRQF-MAS attains 88.5% accuracy, precision "
        "0.816 and Matthews correlation coefficient (MCC) 0.604 with 95% "
        "bootstrap CI [0.57, 0.64], full capture of the 2008 and 2022 "
        "regimes, and a two-day median detection latency. Conclusions: "
        "MRQF-MAS delivers a structurally interpretable, agent-traceable "
        "regime decomposition complementary to scalar volatility estimators."
    )
    add_para(doc, "MDPI_1.7_abstract", "Abstract: " + abstract_body)

    # keywords — at most 10 per MDPI rule
    keywords = [
        "econophysics",
        "multiscale relativistic quantum finance",
        "multi-agent systems",
        "cooperative agents",
        "shared knowledge base",
        "regime detection",
        "energy-entropy space",
        "scale invariance",
        "EUR/USD",
        "financion",
    ]
    add_para(doc, "MDPI_1.8_keywords", "Keywords: " + "; ".join(keywords))

    # horizontal rule
    add_para(doc, "MDPI_1.9_line", "")

    # ---------- body ----------
    equation_counter = 0
    image_cursor = 0
    i = 0
    n = len(blocks)
    in_references = False
    appendix_b_table_seen = False
    # Skip blocks we have already emitted (indices 0..5)
    CONSUMED_FRONT = {0, 1, 2, 3, 4, 5}

    backmatter_emitted = False
    acronyms_expanded: set[str] = set()

    def emit_backmatter_and_abbreviations():
        """MDPI order: Conclusions → back matter → Abbreviations → Appendix A."""
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Supplementary Materials: Not applicable.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Author Contributions: Conceptualization, X.X. and Y.Y.; "
            "Methodology, X.X.; Software, X.X.; Validation, X.X., Y.Y. and "
            "Z.Z.; Formal Analysis, X.X.; Investigation, X.X.; Resources, "
            "X.X.; Data Curation, X.X.; Writing – Original Draft "
            "Preparation, X.X.; Writing – Review & Editing, X.X.; "
            "Visualization, X.X.; Supervision, X.X.; Project "
            "Administration, X.X.; Funding Acquisition, Y.Y. All authors "
            "have read and agreed to the published version of the "
            "manuscript.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Funding: This research received no external funding.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Institutional Review Board Statement: Not applicable.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Informed Consent Statement: Not applicable.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Data Availability Statement: The original data presented in the "
            "study are openly available in the European Central Bank (ECB) "
            "euro foreign exchange reference rates archive at "
            "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/"
            "euro_reference_exchange_rates/html/index.en.html, and were "
            "accessed through the CurrencyConverter Python package version "
            "0.18.17. Full reproducibility details and hyperparameter "
            "settings are included in Appendix A.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Acknowledgments: The authors acknowledge the contributors of "
            "the open-source scientific Python ecosystem (NumPy, SciPy, "
            "PyWavelets, pandas, Matplotlib) on which the computational "
            "validation of this paper relies. No generative artificial "
            "intelligence (GenAI) tool or large language model (LLM) was "
            "used to generate text, data or figures or to assist in study "
            "design, data collection, analysis or interpretation; GenAI use "
            "was limited to superficial text editing (grammar, spelling, "
            "punctuation and formatting) which per MDPI policy does not "
            "need to be declared.",
        )
        add_para(
            doc,
            "MDPI_6.2_back_matter",
            "Conflicts of Interest: The authors declare no conflicts of "
            "interest.",
        )
        # Abbreviations
        add_para(doc, "MDPI_2.1_heading1", "Abbreviations")
        add_para(
            doc,
            "MDPI_3.2_text_no_indent",
            "The following abbreviations are used in this manuscript:",
        )
        add_data_table(
            doc,
            header=["Abbreviation", "Meaning"],
            rows=[
                ["MRQF", "Multiscale Relativistic Quantum Finance"],
                ["MAS", "Multi-Agent System"],
                ["SKB", "Shared Knowledge Base"],
                ["PE", "Prospect Environment"],
                ["ES", "Energy-Entropy"],
                ["ECB", "European Central Bank"],
                [
                    "GARCH",
                    "Generalized Autoregressive Conditional "
                    "Heteroscedasticity",
                ],
                ["MCC", "Matthews Correlation Coefficient"],
                ["PIP", "Percentage in Point"],
                ["RL", "Reinforcement Learning"],
                ["QED", "Quantum Electrodynamics"],
                ["LSTM", "Long Short-Term Memory"],
                ["GenAI", "Generative Artificial Intelligence"],
            ],
        )

    while i < n:
        block = blocks[i]

        # Insert back-matter + Abbreviations immediately before "Appendix A"
        if (
            not backmatter_emitted
            and block[0] == "para"
            and block[2] == "Heading1"
            and block[3].strip().lower().startswith("appendix a")
        ):
            emit_backmatter_and_abbreviations()
            backmatter_emitted = True

        if block[0] == "table":
            # Source table 8 is the Appendix B notation table; the source has
            # no caption for it, so we inject one that matches the template's
            # "Table A1" appendix convention.
            if block[1] == 8 and not appendix_b_table_seen:
                add_table_caption(
                    doc,
                    "Table A1. Notation used throughout the paper.",
                )
                appendix_b_table_seen = True
            copy_source_table(doc, block[2])
            i += 1
            continue

        _, pi, style, text, has_img = block

        if pi in CONSUMED_FRONT:
            i += 1
            continue

        # --- Algorithm 1 block (handled as a numbered list) ---
        if pi == ALG_START and text.strip().startswith("Algorithm 1"):
            add_para(
                doc,
                "MDPI_3.5_text_before_list",
                "Algorithm 1. MRQF-MAS Cooperative Coordination.",
            )
            j = pi + 1
            while j <= ALG_END:
                pb = next((b for b in blocks if b[0] == "para" and b[1] == j), None)
                if pb is None:
                    j += 1
                    continue
                t = pb[3].strip()
                if t:
                    add_para(doc, "MDPI_3.7_itemize", t)
                j += 1
            # advance cursor past the algorithm block
            i = next(
                idx
                for idx, b in enumerate(blocks)
                if b[0] == "para" and b[1] == ALG_END + 1
            )
            continue

        # --- figure image + caption ---
        if has_img:
            img_name = IMAGE_ORDER[image_cursor]
            image_cursor += 1
            # consume the following caption paragraph
            next_block = blocks[i + 1] if i + 1 < n else None
            caption = (
                next_block[3]
                if next_block and next_block[0] == "para"
                else ""
            )
            add_figure(doc, MEDIA / img_name, caption)
            # advance past caption
            i += 2
            continue

        # --- table caption (source encodes as Heading 3 "Table N. ...") ---
        if style == "Heading3" and is_table_caption(text):
            add_table_caption(doc, text)
            i += 1
            continue

        # --- display equation ---
        if style == "Normal" and is_equation_para(text):
            equation_counter += 1
            inner = text.strip().strip("$").strip()
            add_equation(doc, inner, equation_counter)
            i += 1
            continue

        # --- headings ---
        if style in STYLE_MAP:
            add_para(doc, STYLE_MAP[style], text)
            if style == "Heading1" and text.strip().lower().startswith(
                "references"
            ):
                in_references = True
            elif style == "Heading1":
                in_references = False
            i += 1
            continue

        # --- empty paragraphs: skip (they're structural separators) ---
        if not text.strip():
            i += 1
            continue

        # --- references list ---
        if in_references and REFERENCE_RE.match(text):
            add_para(doc, "MDPI_8.1_references", text)
            i += 1
            continue

        # --- default body text with inline maths ---
        body_text = text
        if pi in TABLE_CITATION_INJECTIONS:
            body_text = body_text.rstrip() + TABLE_CITATION_INJECTIONS[pi]
        if pi in FIGURE_CITATION_INJECTIONS:
            body_text = body_text.rstrip() + FIGURE_CITATION_INJECTIONS[pi]
        for needle, expansion in ACRONYM_FIRST_USE_REPLACEMENTS.items():
            if needle in acronyms_expanded:
                continue
            if needle in body_text:
                body_text = body_text.replace(needle, expansion, 1)
                acronyms_expanded.add(needle)
        body_text = renumber_figures(body_text)
        add_inline_math_para(doc, "MDPI_3.1_text", body_text)
        i += 1

    # ---------- mandatory Publisher's Note ----------
    add_para(
        doc,
        "MDPI_6.3_notes",
        "Disclaimer/Publisher's Note: The statements, opinions and data "
        "contained in all publications are solely those of the individual "
        "author(s) and contributor(s) and not of MDPI and/or the editor(s). "
        "MDPI and/or the editor(s) disclaim responsibility for any injury to "
        "people or property resulting from any ideas, methods, instructions "
        "or products referred to in the content.",
    )

    # ---------- restore the section properties that were at the end ----------
    body = doc.element.body
    # python-docx manages sectPr lifecycle automatically when saving, so we do
    # nothing here other than writing out the file.

    doc.save(OUTPUT)
    print(f"Wrote {OUTPUT}")
    print(f"Equations emitted: {equation_counter}")
    print(f"Figures emitted:   {image_cursor}")


if __name__ == "__main__":
    rewrite()
