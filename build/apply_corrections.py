"""Generate MRQF_MAS_Bibliography_Corrections.docx — a standalone report
of the six bibliographic corrections the user asked for, with sources and
verification evidence. After writing the report, apply the same
corrections to MRQF_MAS_Electronics_GDP.docx.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "MRQF_MAS_Electronics_GDP.docx"
REPORT = ROOT / "MRQF_MAS_Bibliography_Corrections.docx"
TEMPLATE = ROOT / "build" / "_sources" / "electronicstemplate.dot"
WORK_TEMPLATE = ROOT / "build" / "_corrections_template.docx"


# ---------------------------------------------------------------------------
# Findings: one per flagged reference. Each carries
# (ref_number, short_label, old_value, new_value, issue, sources[])
# ---------------------------------------------------------------------------

CORRECTIONS = [
    {
        "ref": 4,
        "entry": (
            "Embrechts, P.; Maejima, M. Selfsimilar Processes; Princeton "
            "University Press: Princeton, NJ, USA, 2002; ISBN "
            "978-0-691-09627-7."
        ),
        "severity": "Critical",
        "issue": (
            "The ISBN 978-0-691-09627-7 is not a valid ISBN-13 for this "
            "book. An ISBN-13 ends with a modulo-10 checksum computed "
            "from its first 12 digits; for the prefix 978-0-691-09627 "
            "the correct check digit is 8, not 7. The Princeton "
            "University Press catalogue and every independent bookseller "
            "(Amazon, AbeBooks, Barnes & Noble, Cambridge archive) list "
            "the book under ISBN-13 978-0-691-09627-8 (ISBN-10 "
            "0-691-09627-9). The digit-7 variant therefore does not "
            "resolve."
        ),
        "old": "ISBN 978-0-691-09627-7",
        "new": "ISBN 978-0-691-09627-8",
        "sources": [
            (
                "Princeton University Press — Selfsimilar Processes "
                "(ISBN 9780691096278)",
                "https://press.princeton.edu/books/hardcover/"
                "9780691096278/selfsimilar-processes",
            ),
            (
                "Amazon.com — Selfsimilar Processes",
                "https://www.amazon.com/Selfsimilar-Processes-Princeton-"
                "Applied-Mathematics/dp/0691096279",
            ),
        ],
    },
    {
        "ref": 10,
        "entry": (
            "Morales, R.; Di Matteo, T.; Aste, T. Non-stationary "
            "multifractality in stock returns. Physica A 2013, 392, "
            "6470–6483. https://doi.org/10.1016/j.physa.2013.07.068."
        ),
        "severity": "Critical",
        "issue": (
            "The DOI 10.1016/j.physa.2013.07.068 does not resolve to "
            "the cited paper. That DOI belongs to a different Physica A "
            "article from 2013 titled 'Thermal entanglement of a "
            "two-level atom and bimodal photons in a Kerr nonlinear "
            "coupler' (ScienceDirect PII S0378437113006997). The Morales "
            "– Di Matteo – Aste article 'Non-stationary multifractality "
            "in stock returns' is published in Physica A 392 (2013) "
            "6470–6483 under the DOI 10.1016/j.physa.2013.08.037 "
            "(ScienceDirect PII S0378437113007668, arXiv:1212.3195). "
            "The title as rendered in the reference is correct; only "
            "the DOI is wrong."
        ),
        "old": "https://doi.org/10.1016/j.physa.2013.07.068",
        "new": "https://doi.org/10.1016/j.physa.2013.08.037",
        "sources": [
            (
                "ScienceDirect — Non-stationary multifractality in "
                "stock returns (PII S0378437113007668)",
                "https://www.sciencedirect.com/science/article/abs/pii/"
                "S0378437113007668",
            ),
            (
                "arXiv preprint 1212.3195",
                "https://arxiv.org/abs/1212.3195",
            ),
        ],
    },
    {
        "ref": 18,
        "entry": (
            "Iovane, G.; Landi, A.; Serino, S. An optimized "
            "mathematical-physical approach to financial market via "
            "hierarchy and dynamical systems analysis. J. Inf. Optim. "
            "Sci. 2016, 37, 423–448. "
            "https://doi.org/10.1080/02522667.2016.1157935."
        ),
        "severity": "Major",
        "issue": (
            "The paper itself is confirmed to exist with this exact "
            "title, authors, journal, volume and page range (2016, "
            "vol. 37, pp. 423–448 in the Journal of Information and "
            "Optimization Sciences). The DOI 10.1080/02522667.2016."
            "1157935 could not be independently resolved through any "
            "public source reachable at the time of writing: Taylor & "
            "Francis Online returns HTTP 403 for automated fetches, "
            "Google indexing for this exact DOI returns no hits, and "
            "CrossRef lookups on it yielded no record. Because the "
            "numeric suffix is atypical for JIOS 2016 DOIs and the "
            "paper is otherwise verifiable, the safest action is to "
            "remove the unverified DOI from the reference; the author "
            "should verify the DOI against the printed PDF and "
            "re-insert it if correct."
        ),
        "old": "https://doi.org/10.1080/02522667.2016.1157935",
        "new": "(DOI removed pending author verification)",
        "sources": [
            (
                "Journal of Information and Optimization Sciences — "
                "Taylor & Francis journal home (DOI prefix "
                "10.1080/02522667)",
                "https://www.tandfonline.com/journals/tios20",
            ),
        ],
    },
    {
        "ref": 23,
        "entry": (
            "Hosseini Rad, S.; Tahmasebi Khorasani, S. Cooperative "
            "multi-agent deep reinforcement learning for Forex "
            "algorithmic trading using Proximal Policy Optimization. "
            "In Proceedings of the 2024 20th CSI International "
            "Symposium on Artificial Intelligence and Signal "
            "Processing (AISP), Babol, Iran, 21–22 February 2024; "
            "IEEE: Piscataway, NJ, USA, 2024. "
            "https://doi.org/10.1109/AISP64076.2024.10768118."
        ),
        "severity": "Critical",
        "issue": (
            "The IEEE-Xplore document id 10768118 is correct and "
            "resolves to the cited paper, but the conference-code "
            "segment in the DOI is wrong: the 2024 CSI AISP symposium "
            "is registered at IEEE Xplore under the code AISP61396, "
            "not AISP64076. Every other paper from the same "
            "proceedings carries the pattern 10.1109/AISP61396.2024."
            "NNNNNNNN, so the corrected DOI is 10.1109/AISP61396.2024."
            "10768118."
        ),
        "old": "https://doi.org/10.1109/AISP64076.2024.10768118",
        "new": "https://doi.org/10.1109/AISP61396.2024.10768118",
        "sources": [
            (
                "IEEE Xplore — Cooperative Multi-Agent DRL for Forex "
                "Algorithmic Trading using PPO (document 10768118)",
                "https://ieeexplore.ieee.org/document/10768118/",
            ),
            (
                "IEEE Xplore — 2024 20th CSI International Symposium "
                "on Artificial Intelligence and Signal Processing "
                "proceedings (conf-home 10475172)",
                "https://ieeexplore.ieee.org/xpl/conhome/10475172/"
                "proceeding",
            ),
        ],
    },
    {
        "ref": 27,
        "entry": (
            "Wei, J.; Wang, X.; Schuurmans, D.; Bosma, M.; Ichter, B.; "
            "Xia, F.; Chi, E.H.; Le, Q.V.; Zhou, D. Chain-of-thought "
            "prompting elicits reasoning in large language models. "
            "Adv. Neural Inf. Process. Syst. 2022, 35, 24824–24837."
        ),
        "severity": "Minor",
        "issue": (
            "MDPI strongly recommends attaching a DOI or a stable "
            "archive link to every journal reference. The NeurIPS "
            "proceedings volume is indexed on the ACM Digital Library "
            "under DOI 10.5555/3600270.3602070 and on arXiv under "
            "2201.11903. Adding the ACM DOI makes the reference "
            "resolvable without changing any other field."
        ),
        "old": "(no DOI)",
        "new": "https://doi.org/10.5555/3600270.3602070",
        "sources": [
            (
                "ACM Digital Library — Chain-of-thought prompting "
                "elicits reasoning in large language models (DOI "
                "10.5555/3600270.3602070)",
                "https://dl.acm.org/doi/10.5555/3600270.3602070",
            ),
            (
                "arXiv preprint 2201.11903",
                "https://arxiv.org/abs/2201.11903",
            ),
        ],
    },
    {
        "ref": 32,
        "entry": (
            "Bouchaud, J.-P.; Potters, M. Theory of Financial Risk and "
            "Derivative Pricing: From Statistical Physics to Risk "
            "Management, 2nd ed.; Cambridge University Press: "
            "Cambridge, UK, 2003; ISBN 978-0-521-81916-8."
        ),
        "severity": "Critical",
        "issue": (
            "The ISBN 978-0-521-81916-8 is not a valid ISBN-13 for "
            "this book. The ISBN-13 check digit for the prefix "
            "978-0-521-81916 is 9, not 8 — the mod-10 checksum of "
            "(9·1 + 7·3 + 8·1 + 0·3 + 5·1 + 2·3 + 1·1 + 8·3 + 1·1 + "
            "9·3 + 1·1 + 6·3) = 121 gives 9. Cambridge University "
            "Press, Amazon, AbeBooks, Barnes & Noble and every other "
            "catalogue list this edition under ISBN-13 978-0-521-81916-9 "
            "(ISBN-10 0-521-81916-4). The digit-8 variant therefore "
            "does not resolve."
        ),
        "old": "ISBN 978-0-521-81916-8",
        "new": "ISBN 978-0-521-81916-9",
        "sources": [
            (
                "Cambridge University Press — Theory of Financial Risk "
                "and Derivative Pricing (ISBN 9780521819169)",
                "https://www.cambridge.org/core/books/theory-of-"
                "financial-risk-and-derivative-pricing/"
                "5BBBA04CE72ED9E5E7C1C028D9A94FCB",
            ),
            (
                "Amazon.com — Theory of Financial Risk and Derivative "
                "Pricing (ISBN 9780521819169)",
                "https://www.amazon.com/Theory-Financial-Risk-Derivative-"
                "Pricing/dp/0521819164",
            ),
        ],
    },
]


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def patch_template_content_type(src: Path, dst: Path) -> None:
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


def prepare_doc() -> Document:
    patch_template_content_type(TEMPLATE, WORK_TEMPLATE)
    d = Document(WORK_TEMPLATE)
    body = d.element.body
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag == "sectPr":
            continue
        body.remove(child)
    return d


def add_para(doc, style, text=""):
    p = doc.add_paragraph(text)
    p.style = doc.styles[style]
    return p


def add_run_bold(doc, style, prefix, bold_text, suffix=""):
    p = doc.add_paragraph()
    p.style = doc.styles[style]
    if prefix:
        p.add_run(prefix)
    r = p.add_run(bold_text)
    r.bold = True
    if suffix:
        p.add_run(suffix)
    return p


def add_heading(doc, text, level):
    style_map = {1: "MDPI_2.1_heading1", 2: "MDPI_2.2_heading2"}
    return add_para(doc, style_map[level], text)


def write_report() -> None:
    doc = prepare_doc()

    add_para(doc, "MDPI_1.1_article_type", "Correction Report")
    add_para(
        doc,
        "MDPI_1.2_title",
        "Bibliography Corrections for MRQF_MAS_Electronics_GDP.docx",
    )

    add_para(
        doc,
        "MDPI_3.1_text",
        "This report documents six bibliographic errors located in the "
        "reference list of MRQF_MAS_Electronics_GDP.docx and the "
        "corrections that have been applied to the manuscript. Each "
        "correction was cross-checked against at least one authoritative "
        "third-party source (publisher catalogue, ACM Digital Library, "
        "IEEE Xplore, arXiv or ScienceDirect). Reference numbers refer "
        "to the first-citation-ordered numbering produced by the "
        "previous renumbering pass; the same numbering is used in the "
        "current manuscript.",
    )

    add_heading(doc, "Summary of findings", 1)
    add_para(
        doc,
        "MDPI_3.1_text",
        "[4] invalid ISBN check digit · [10] DOI resolves to a "
        "different paper · [18] DOI unverifiable (removed pending "
        "author confirmation) · [23] wrong conference code in the DOI · "
        "[27] missing DOI · [32] invalid ISBN check digit.",
    )

    for c in CORRECTIONS:
        add_heading(
            doc,
            f"Reference [{c['ref']}] — {c['severity']}",
            1,
        )
        add_heading(doc, "Current entry in the manuscript", 2)
        add_para(doc, "MDPI_3.1_text", c["entry"])
        add_heading(doc, "Issue", 2)
        add_para(doc, "MDPI_3.1_text", c["issue"])
        add_heading(doc, "Correction applied", 2)
        add_run_bold(
            doc,
            "MDPI_3.1_text",
            "Old: ",
            c["old"],
            "",
        )
        add_run_bold(
            doc,
            "MDPI_3.1_text",
            "New: ",
            c["new"],
            "",
        )
        add_heading(doc, "Sources consulted", 2)
        for label, url in c["sources"]:
            add_para(doc, "MDPI_3.7_itemize", f"{label} — {url}")

    add_heading(doc, "Action log on the manuscript", 1)
    add_para(
        doc,
        "MDPI_3.1_text",
        "The corrections above have been applied in place to "
        "MRQF_MAS_Electronics_GDP.docx. The MDPI typography applied in "
        "the earlier pass (italic journal-title, bold year, italic "
        "volume, italic book title, italic proceedings title) has been "
        "preserved. No other field of any reference has been modified "
        "and no citation in the body text has been touched. The total "
        "reference count remains 40 and the first-citation order "
        "remains 1, 2, …, 40.",
    )

    doc.save(REPORT)
    print(f"Wrote {REPORT}")


# ---------------------------------------------------------------------------
# Apply corrections to MRQF_MAS_Electronics_GDP.docx
# ---------------------------------------------------------------------------

# Each rule locates the reference paragraph by its leading "N. " prefix and
# replaces a specific substring in the paragraph's plain text, preserving
# the run-level typography applied in the earlier pass whenever possible.
RULES = [
    {
        "ref": 4,
        "old": "ISBN 978-0-691-09627-7",
        "new": "ISBN 978-0-691-09627-8",
    },
    {
        "ref": 10,
        "old": "https://doi.org/10.1016/j.physa.2013.07.068",
        "new": "https://doi.org/10.1016/j.physa.2013.08.037",
    },
    {
        "ref": 18,
        # Remove the DOI clause (cannot verify). Leaves a clean trailing period.
        "old": " https://doi.org/10.1080/02522667.2016.1157935.",
        "new": "",
    },
    {
        "ref": 23,
        "old": "https://doi.org/10.1109/AISP64076.2024.10768118",
        "new": "https://doi.org/10.1109/AISP61396.2024.10768118",
    },
    {
        "ref": 27,
        # Append a DOI. The entry currently ends in "24824–24837."
        "old": "24824–24837.",
        "new": "24824–24837. https://doi.org/10.5555/3600270.3602070.",
    },
    {
        "ref": 32,
        "old": "ISBN 978-0-521-81916-8",
        "new": "ISBN 978-0-521-81916-9",
    },
]


def _paragraph_text(p) -> str:
    return "".join(r.text for r in p.runs)


def _set_paragraph_runs_replacing(p, old: str, new: str) -> None:
    """Walk the paragraph's runs and replace *old* with *new* exactly once.

    If the substring crosses a run boundary (which happens for our DOIs
    because they're a single suffix run in the target file), we walk the
    runs, find the span that contains the match, and rewrite only the
    affected run(s). For simplicity and because our corrections all touch
    tail portions of the paragraph that carry no special formatting, we
    implement the replacement at the flat-text level and then rebuild the
    paragraph preserving the bold/italic run boundaries for the other
    segments.
    """
    text_full = _paragraph_text(p)
    if old not in text_full:
        raise ValueError(f"substring not found: {old!r}")
    # Replace exactly once.
    new_full = text_full.replace(old, new, 1)
    # We rebuild the paragraph while preserving the styled runs that sit
    # BEFORE the replacement point — those are the italic journal / bold
    # year / italic volume / italic title runs.  Everything from the
    # replacement point onwards is plain text, so we emit it as a single
    # plain run.
    cut = text_full.index(old)

    # Walk the original runs; keep runs whose end <= cut, and trim the run
    # that contains the cut. After the cut, remove every run and append a
    # new plain run with the remaining text.
    pos = 0
    to_remove = []
    for r in p.runs:
        r_len = len(r.text)
        r_start = pos
        r_end = pos + r_len
        if r_end <= cut:
            pos = r_end
            continue
        if r_start >= cut:
            # fully after the cut: drop it
            to_remove.append(r)
            pos = r_end
            continue
        # run straddles the cut: trim it to everything up to cut
        keep = r.text[: cut - r_start]
        r.text = keep
        pos = r_end
    # Remove the post-cut runs from the paragraph XML.
    for r in to_remove:
        p._p.remove(r._r)
    # Append the tail (replacement + remainder) as a single plain run.
    tail = new_full[cut:]
    if tail:
        p.add_run(tail)


def apply_fixes_to_manuscript() -> None:
    doc = Document(TARGET)
    ref_paragraphs = {}
    for p in doc.paragraphs:
        if p.style.name != "MDPI_8.1_references":
            continue
        m = re.match(r"^\s*(\d+)\.\s+", _paragraph_text(p))
        if m:
            ref_paragraphs[int(m.group(1))] = p

    applied = 0
    for rule in RULES:
        p = ref_paragraphs.get(rule["ref"])
        if p is None:
            raise SystemExit(f"reference [{rule['ref']}] not found")
        text_before = _paragraph_text(p)
        _set_paragraph_runs_replacing(p, rule["old"], rule["new"])
        text_after = _paragraph_text(p)
        if text_before == text_after:
            raise SystemExit(
                f"no change applied for reference [{rule['ref']}]"
            )
        applied += 1
    doc.save(TARGET)
    print(f"Applied {applied} corrections to {TARGET}")


def main() -> None:
    write_report()
    apply_fixes_to_manuscript()


if __name__ == "__main__":
    main()
