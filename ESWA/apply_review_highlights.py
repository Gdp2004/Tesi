"""
Open Predictive_MRQF_MS_ESWA_Elsevier.docx, apply yellow highlights to every
flagged passage from the reviewer audit, and save as
Predictive_MRQF_MS_ESWA_Elsevier_REVIEW.docx.

Each highlight has a numeric tag (R1, R2, ...) that matches an entry in
REVIEW_REPORT.md, so the user can cross-reference document and report.
"""

import sys
from copy import deepcopy
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = "/home/user/Tesi/ESWA/Predictive_MRQF_MS_ESWA_Elsevier.docx"
DST = "/home/user/Tesi/ESWA/Predictive_MRQF_MS_ESWA_Elsevier_REVIEW.docx"


def _make_run_clone(template_run_elem, new_text, *, highlight=False):
    """Clone a w:r element, set its text, optionally apply yellow highlight."""
    new_r = deepcopy(template_run_elem)
    # Strip existing w:t and w:highlight
    for t in new_r.findall(qn("w:t")):
        new_r.remove(t)
    rPr = new_r.find(qn("w:rPr"))
    if rPr is not None:
        hl = rPr.find(qn("w:highlight"))
        if hl is not None:
            rPr.remove(hl)
    # Add highlight if requested
    if highlight:
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            new_r.insert(0, rPr)
        hl = OxmlElement("w:highlight")
        hl.set(qn("w:val"), "yellow")
        rPr.append(hl)
    # Add fresh text element
    t_el = OxmlElement("w:t")
    t_el.text = new_text
    t_el.set(qn("xml:space"), "preserve")
    new_r.append(t_el)
    return new_r


def highlight_substring(paragraph, substring):
    """If substring appears in any single run of the paragraph, split that run
    so the matched span is in its own run, then yellow-highlight that run.
    Returns True on success, False if not found in any single run."""
    for run in paragraph.runs:
        if substring in run.text:
            full = run.text
            i = full.index(substring)
            before = full[:i]
            match = full[i : i + len(substring)]
            after = full[i + len(substring) :]
            r_elem = run._element
            parent = r_elem.getparent()
            pos = list(parent).index(r_elem)
            # Mutate current run -> "before" (no highlight)
            for t in r_elem.findall(qn("w:t")):
                r_elem.remove(t)
            t_el = OxmlElement("w:t")
            t_el.text = before
            t_el.set(qn("xml:space"), "preserve")
            r_elem.append(t_el)
            # Insert highlighted "match" run
            new_match = _make_run_clone(r_elem, match, highlight=True)
            parent.insert(pos + 1, new_match)
            # Insert "after" run if non-empty
            if after:
                new_after = _make_run_clone(r_elem, after, highlight=False)
                parent.insert(pos + 2, new_after)
            return True
    # Fallback: substring crosses run boundaries (rare in our build) — highlight
    # all runs whose combined text spans the match.
    full_text = paragraph.text
    if substring not in full_text:
        return False
    # In the multi-run case, just highlight all runs whose own text overlaps the match.
    start = full_text.index(substring)
    end = start + len(substring)
    cursor = 0
    for run in paragraph.runs:
        rtext = run.text
        rstart = cursor
        rend = cursor + len(rtext)
        cursor = rend
        if rend > start and rstart < end:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return True


def highlight_paragraph(paragraph):
    """Highlight every run in a paragraph (whole-paragraph yellow)."""
    for run in paragraph.runs:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW


# Each item: (tag, mode, locator, target_substring_or_None, note)
#   mode = "phrase" -> highlight just that substring inside paragraph(s) matching locator
#   mode = "para"   -> highlight the entire paragraph that contains locator
HIGHLIGHTS = [
    # ---- Numeric / claim mismatches ----
    ("R1", "phrase",
     "We propose P-MRQF-MS, a predictive machine learning framework",
     "P-MRQF-MS attains cross-pair mean $\\mathrm{MCC} = 0.326$",
     "Abstract MCC=0.326 conflicts with Table 4 (0.325 for H=5)."),
    ("R2", "phrase",
     "Cross-pair mean MCC 0.326 beats tuned LSTM",
     "Cross-pair mean MCC 0.326 beats tuned LSTM, XGBoost-classical and logistic.",
     "Highlights: 0.326 vs Table 4 0.325 — pick one."),
    ("R3", "phrase",
     "We execute the full pipeline on real ECB daily data",
     "tuned LSTM ensemble",
     '"Tuned LSTM" label covers only 3/6 pairs (USD,JPY,CAD). Mislabel.'),
    ("R4", "phrase",
     "We execute the full pipeline on real ECB daily data",
     "EUR/USD, EUR/JPY and EUR/CAD",
     "Tuning was applied only to these 3 pairs — fairness issue."),
    ("R5", "phrase",
     "Table 2 reports the cross-pair Wilcoxon",
     "wins in 6 of 6 pairs",
     "Contradicts Table 1: Logistic 5/6, p=0.063 (not 6/6 / 0.031)."),
    ("R6", "phrase",
     "Table 3 reports per-pair DM tests",
     "on 4 of the 6 pairs",
     "USD has DM=+0.96 p=0.34 vs Persistence — does not satisfy 'all four'."),
    ("R7", "phrase",
     "Three contiguous segments per pair",
     "(2,236 observations after the 500-day warm-up)",
     "Sample arithmetic does not match: 500+2236+1254+3465=7455 ≠ 7115."),

    # ---- Wrong / weak citations ----
    ("R8", "phrase",
     "An EGARCH(1,1) model",
     "every 250 trading days [12, 18]",
     "Engle 1982 [18] is ARCH not rolling-refit. Replace with Brownlees & Gallo."),
    ("R9", "phrase",
     "Energy reflects directional potential",
     "pre-registered in the protocol of [29]",
     "[29] Iovane 2026a is unpublished — unverifiable load-bearing claim."),
    ("R10", "phrase",
     "The architecture inherits the operator taxonomy",
     "operator taxonomy and cooperation primitive of [29]",
     "[29] Iovane 2026a unpublished — foundational reference must be archived."),
    ("R11", "phrase",
     "Third, these qualitative patterns can be interpreted",
     "two-dimensional complex-time perspective introduced by Iovane and Iovane [31]",
     "[31] is on LLM hallucination, not FX. Speculative transplant — demote to footnote."),
    ("R12", "phrase",
     "Section 6.7 reports a finding that is scientifically interesting",
     "complex-time",
     "Same speculative complex-time interpretation — see R11."),

    # ---- Self-citation density (reference list) ----
    ("R13", "para_at",
     195,
     None,
     "Iovane 2026a [29] unpublished but cited 7× in body. Archive on Zenodo or remove."),
    ("R14", "para_at",
     196,
     None,
     "Iovane 2026b [30] unpublished. Same problem as R13."),

    # ---- Sections that reviewers will ask to cut / move to appendix ----
    ("R15", "para_at",
     104,
     None,
     "§6.6 COVID trajectories: qualitative only — move to Appendix B."),
    ("R16", "para_at",
     108,
     None,
     "§6.7 multi-horizon: sample-dependent (n=6); cut or shrink."),
    ("R17", "para_at",
     124,
     None,
     "§7.2 Stacking on EUR/USD only — single pair, not informative."),
    ("R18", "para_at",
     142,
     None,
     "§8.5 Complex-time discussion: out of ESWA scope; cut or appendicise."),

    # ---- Defensive / hedging paragraphs ----
    ("R19", "phrase",
     "On the nature of the contribution",
     "On the nature of the contribution",
     "Defensive rebuttal — shorten by 60% or move to cover letter."),
    ("R20", "phrase",
     "On the economic meaning of the target",
     "On the economic meaning of the target",
     "Defensive rebuttal — shorten or move to cover letter."),
    ("R21", "phrase",
     "This is an honest negative result",
     "honest negative result",
     'Hedging adverbs ("honest", "frankly") weaken contribution. Cut.'),

    # ---- Missing citations / unsupported claims ----
    ("R22", "phrase",
     "Regime-transition forecasting in financial time series",
     "persistence-dominated baselines that inflate apparent predictive accuracy",
     'Strong empirical claim — cite (e.g., Patton 2011, Hansen & Lunde 2005).'),

    # ---- LSTM tuned mean discussion ----
    ("R23", "phrase",
     "The LSTM hyperparameter grid was designed",
     "tuned LSTM",
     "Counter-intuitive that tuning *degraded* MCC (0.305→0.284). Audit early-stop / lr."),

    # ---- 0.326 vs 0.325 mismatch in conclusions ----
    ("R24", "phrase",
     "We presented P-MRQF-MS",
     "0.326",
     "Conclusions echo 0.326 — must reconcile with Table 4 (0.325)."),
]


def main():
    doc = Document(SRC)
    paragraphs = list(doc.paragraphs)

    counts = {"phrase": 0, "para_at": 0, "missed": 0}

    for tag, mode, locator, target, note in HIGHLIGHTS:
        if mode == "phrase":
            # Find all paragraphs whose text starts with or contains 'locator',
            # then highlight 'target' inside them.
            found_any = False
            for p in paragraphs:
                if locator in p.text:
                    if target and target in p.text:
                        if highlight_substring(p, target):
                            found_any = True
                    else:
                        # locator IS the target
                        if target is None or target == locator:
                            if highlight_substring(p, locator):
                                found_any = True
            if found_any:
                counts["phrase"] += 1
                print(f"  [{tag}] phrase OK :: {note}")
            else:
                counts["missed"] += 1
                print(f"  [{tag}] MISSED :: {note}")
        elif mode == "para_at":
            idx = locator
            if 0 <= idx < len(paragraphs):
                highlight_paragraph(paragraphs[idx])
                counts["para_at"] += 1
                print(f"  [{tag}] para#{idx} OK :: {note}")
            else:
                counts["missed"] += 1
                print(f"  [{tag}] index out of range :: {note}")

    doc.save(DST)
    print()
    print(f"Saved: {DST}")
    print(f"Phrase highlights: {counts['phrase']}, "
          f"Paragraph highlights: {counts['para_at']}, "
          f"Missed: {counts['missed']}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
