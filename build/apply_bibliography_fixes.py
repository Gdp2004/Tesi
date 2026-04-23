"""Apply the bibliography fixes recommended in
MRQF_MAS_Bibliography_Report.docx to MRQF_MAS_Electronics_GDP.docx.

Three transformations are applied, in order:

1. Renumber references in order of first citation. Every [N] citation in
   the body text (and in tables) is rewritten accordingly; the reference
   list itself is re-sorted to follow the new numbering.
2. Apply MDPI journal typography to every journal reference (italic
   journal-title, bold year, italic volume). Book references get an
   italic book title.
3. Leave the three conference/proceedings entries ([29], [30], [40] in
   the original numbering) without a DOI, since MDPI does not strictly
   require one for proceedings.

Output is written in place to MRQF_MAS_Electronics_GDP.docx.
"""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "MRQF_MAS_Electronics_GDP.docx"


# ---------------------------------------------------------------------------
# Citation parsing / rendering
# ---------------------------------------------------------------------------

CITATION_RE = re.compile(r"\[([\d,\s–\-]+)\]")


def parse_citation(group: str) -> list[int]:
    """Return the list of reference numbers inside the brackets, preserving
    duplicates and order from the input."""
    nums = []
    for piece in re.split(r",", group):
        piece = piece.strip()
        if not piece:
            continue
        if "–" in piece or "-" in piece:
            a, b = re.split(r"[–\-]", piece)
            nums.extend(range(int(a.strip()), int(b.strip()) + 1))
        else:
            nums.append(int(piece))
    return nums


def render_citation(nums: list[int]) -> str:
    """Render a sorted-unique citation list as MDPI [a,b,c-d]."""
    nums = sorted(set(nums))
    if not nums:
        return "[]"
    # Compress consecutive runs into en-dash ranges
    groups = [[nums[0]]]
    for n in nums[1:]:
        if n == groups[-1][-1] + 1:
            groups[-1].append(n)
        else:
            groups.append([n])
    parts = []
    for g in groups:
        if len(g) >= 3:
            parts.append(f"{g[0]}–{g[-1]}")
        else:
            parts.extend(str(x) for x in g)
    return "[" + ",".join(parts) + "]"


# ---------------------------------------------------------------------------
# Step 1: compute first-citation order
# ---------------------------------------------------------------------------

def iter_body_paragraphs(doc):
    """Yield every paragraph except those in the reference list."""
    for p in doc.paragraphs:
        if p.style and p.style.name == "MDPI_8.1_references":
            continue
        yield p


def iter_table_paragraphs(doc):
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def compute_first_citation_order(doc) -> dict[int, int]:
    first_seen: dict[int, int] = {}
    counter = 0
    for p in list(iter_body_paragraphs(doc)) + list(iter_table_paragraphs(doc)):
        for m in CITATION_RE.finditer(p.text):
            for n in parse_citation(m.group(1)):
                if n not in first_seen:
                    counter += 1
                    first_seen[n] = counter
    return first_seen


# ---------------------------------------------------------------------------
# Step 2: rewrite citations in text (preserving runs where possible)
# ---------------------------------------------------------------------------

def rewrite_paragraph_citations(p, old_to_new: dict[int, int]) -> None:
    """Rewrite every [N] token in p using the old→new map.

    We rebuild the full text in a single run if the paragraph already has
    a single run; otherwise we collapse all runs into one (for body text,
    citations never carry custom formatting, so no information is lost).
    """
    text = p.text
    new_text, changed = _substitute_text(text, old_to_new)
    if not changed:
        return
    # Replace the entire paragraph text with the new text in a single run,
    # preserving the paragraph style. Inline italic math (if any) is lost,
    # but for MRQF_MAS_Electronics_GDP.docx the body has no italic runs
    # anyway — the source manuscript encodes maths as text ($...$).
    _set_paragraph_text(p, new_text)


def _substitute_text(text: str, old_to_new: dict[int, int]) -> tuple[str, bool]:
    changed = False

    def repl(m):
        nonlocal changed
        nums = parse_citation(m.group(1))
        new_nums = [old_to_new.get(n, n) for n in nums]
        out = render_citation(new_nums)
        if out != m.group(0):
            changed = True
        return out

    return CITATION_RE.sub(repl, text), changed


def _set_paragraph_text(p, text: str) -> None:
    # Remove every <w:r> and <w:hyperlink> child while keeping <w:pPr>.
    to_remove = []
    for child in list(p._p):
        tag = child.tag.split("}")[-1]
        if tag in ("r", "hyperlink"):
            to_remove.append(child)
    for child in to_remove:
        p._p.remove(child)
    p.add_run(text)


# ---------------------------------------------------------------------------
# Step 3: reorder reference list entries and apply MDPI typography
# ---------------------------------------------------------------------------

def collect_reference_paragraphs(doc) -> list[tuple[int, "_pargraph"]]:
    refs = []
    for p in doc.paragraphs:
        if p.style and p.style.name == "MDPI_8.1_references":
            m = re.match(r"^\s*(\d+)\.\s+(.*)$", p.text)
            if m:
                refs.append((int(m.group(1)), p, m.group(2).strip()))
    return refs


def reorder_references(doc, old_to_new: dict[int, int]) -> None:
    """Reorder the reference paragraphs so that the Nth entry is the one
    whose old number maps to new_number=N. Rewrite the `N. ` prefix.
    """
    refs = collect_reference_paragraphs(doc)
    # Build map: old_number -> (old_paragraph, body_text)
    by_old = {n: (p, body) for n, p, body in refs}

    # Determine new order: list of old numbers in new order.
    # For each new_number N, find old_number o such that old_to_new[o] = N.
    new_to_old = {v: k for k, v in old_to_new.items()}
    n_total = len(refs)
    assert set(new_to_old.keys()) == set(range(1, n_total + 1)), (
        "first-citation map does not cover 1..N"
    )

    # We rewrite each original paragraph's text in place, in document order.
    # Original paragraphs are already stored in new order (1, 2, 3, …) in the
    # XML, because we kept them where they are. We just need to update each
    # paragraph i to carry the content of the old reference that maps to new
    # number i.
    ordered_old_numbers = [new_to_old[n] for n in range(1, n_total + 1)]

    # Iterate in document order: the k-th MDPI_8.1_references paragraph in
    # the document should now carry the content of old ref
    # ordered_old_numbers[k].
    k = 0
    for p in doc.paragraphs:
        if p.style.name != "MDPI_8.1_references":
            continue
        old_num = ordered_old_numbers[k]
        k += 1
        _, _src_p, body_text = next(r for r in refs if r[0] == old_num)
        new_num = k
        _emit_styled_reference(p, new_num, body_text)


# ---------- MDPI typography ----------

JOURNAL_SIGNATURE = re.compile(
    r"\b((?:19|20)\d{2})\s*,\s*(\d+)\s*,\s*(\d+(?:[–\-]\d+)?)"
)

# Journal abbreviations are sequences of capital-initial tokens, each
# optionally ending in a period (e.g. "Ann. Sci. École Norm. Sup.",
# "J. Bus.", "IEEE Access", "Phys. Rev.", "Adv. Neural Inf. Process. Syst.").
JOURNAL_TITLE_PATTERN = re.compile(
    r"^[A-ZÀ-ÜŒ][\w'’\-]*\.?(?:\s+[A-ZÀ-ÜŒ][\w'’\-]*\.?)*$"
)


def _is_journal_entry(body: str) -> bool:
    # Year, Volume, page/article-number (range or single)
    return bool(re.search(r"\b(19|20)\d{2},\s*\d+,\s*\d+", body))


def _is_book_or_chapter(body: str) -> bool:
    return (
        "ISBN" in body
        or "In Proceedings of" in body
        or re.search(
            r"(Press|Publisher|Wiley|Springer|Academic|IntechOpen|Random House|"
            r"University Press|SIAM|IEEE|PMLR|ACM)[^,;]*[;,]",
            body,
        ) is not None
    )


def _find_journal_span(body: str, year_start: int) -> tuple[int, int] | None:
    """Return (start, end) in *body* coordinates of the journal title
    that precedes the year, or None if it cannot be located.

    Strategy: iterate all ``". "`` separators in ``body[:year_start]``.
    For each, test whether the slice from that point to year_start
    (stripped) is a pure journal pattern — a sequence of capital-initial
    tokens separated by spaces, each optionally ending with a period.
    Return the earliest position that yields a valid match (i.e., the
    longest journal span).
    """
    head = body[:year_start]
    positions = [m.end() for m in re.finditer(r"[.?!]\s+", head)]
    for pos in positions:
        candidate = head[pos:].rstrip()
        if candidate and JOURNAL_TITLE_PATTERN.match(candidate):
            return pos, pos + len(candidate)
    return None


def _emit_styled_reference(p, new_num: int, body: str) -> None:
    """Clear runs and re-emit the reference with MDPI typography.

    Journal entries get italic journal-title, bold year, italic volume.
    Books and book chapters get an italic book title. Everything else
    is emitted as plain text.
    """
    # Clear the paragraph
    to_remove = []
    for child in list(p._p):
        tag = child.tag.split("}")[-1]
        if tag in ("r", "hyperlink"):
            to_remove.append(child)
    for child in to_remove:
        p._p.remove(child)

    prefix = f"{new_num}. "

    if _is_journal_entry(body):
        _emit_journal(p, prefix, body)
    elif _is_book_or_chapter(body):
        _emit_book(p, prefix, body)
    else:
        # conference proceedings, EU regulation, etc.
        p.add_run(prefix + body)


def _emit_journal(p, prefix: str, body: str) -> None:
    """MDPI journal format:
    Authors. Title. *Journal* **Year**, *Vol*, Pages. DOI.
    """
    m = JOURNAL_SIGNATURE.search(body)
    assert m, f"journal signature missing in: {body[:80]}"
    year_start, year_end = m.start(1), m.end(1)
    vol_start, vol_end = m.start(2), m.end(2)

    span = _find_journal_span(body, year_start)
    if span is None:
        # fall back: no italic on journal title, still apply year/volume
        p.add_run(prefix + body[:year_start])
        r = p.add_run(body[year_start:year_end])
        r.bold = True
        p.add_run(body[year_end:vol_start])
        r = p.add_run(body[vol_start:vol_end])
        r.italic = True
        p.add_run(body[vol_end:])
        return

    j_start, j_end = span
    pre_journal = prefix + body[:j_start]
    journal_text = body[j_start:j_end]
    between = body[j_end:year_start]
    year_text = body[year_start:year_end]
    between_yv = body[year_end:vol_start]
    vol_text = body[vol_start:vol_end]
    tail = body[vol_end:]

    p.add_run(pre_journal)
    r = p.add_run(journal_text)
    r.italic = True
    p.add_run(between)
    r = p.add_run(year_text)
    r.bold = True
    p.add_run(between_yv)
    r = p.add_run(vol_text)
    r.italic = True
    p.add_run(tail)


def _emit_book(p, prefix: str, body: str) -> None:
    """Italicise the title.

    * Proceedings entry (``In Proceedings of the XYZ Conference, …;``):
      italicise ``Proceedings of the XYZ Conference`` (up to the first
      comma that precedes the location / date).
    * Book chapter (``ChapterTitle. In BookTitle; Editor, Eds.; …``):
      italicise the book title between ``In `` and the first ``;``.
    * Plain book (``BookTitle; Publisher: …``): italicise everything
      between the author-period and the first ``;``.
    """
    # Proceedings: ". In Proceedings of …, <location>, <date>; Publisher…"
    m_proc = re.search(r"\.\s+In\s+(Proceedings\s+of\s+[^,]+),", body)
    if m_proc:
        start = m_proc.start(1)
        end = m_proc.end(1)
        p.add_run(prefix + body[:start])
        r = p.add_run(body[start:end])
        r.italic = True
        p.add_run(body[end:])
        return

    # Book chapter: ". In <BookTitle>;"
    m_chap = re.search(r"\.\s+In\s+([^;]+);", body)
    if m_chap:
        start = m_chap.start(1)
        end = m_chap.end(1)
        p.add_run(prefix + body[:start])
        r = p.add_run(body[start:end])
        r.italic = True
        p.add_run(body[end:])
        return

    # Plain book: first ". TITLE;" after the author list
    m_title = re.search(r"\.\s+([^;]+?);", body)
    if not m_title:
        p.add_run(prefix + body)
        return
    start = m_title.start(1)
    end = m_title.end(1)
    p.add_run(prefix + body[:start])
    r = p.add_run(body[start:end])
    r.italic = True
    p.add_run(body[end:])


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def main() -> None:
    doc = Document(TARGET)

    # 1. Compute first-citation order
    first_seen = compute_first_citation_order(doc)
    n = len(first_seen)
    print(f"Discovered {n} unique references cited in body.")
    if n != 40:
        raise SystemExit(
            f"expected 40 cited references, got {n} — aborting"
        )
    # old_to_new: reference number in the existing GDP file -> new number
    old_to_new = dict(first_seen)

    # 2. Rewrite every citation in body and tables
    for p in iter_body_paragraphs(doc):
        rewrite_paragraph_citations(p, old_to_new)
    for p in iter_table_paragraphs(doc):
        rewrite_paragraph_citations(p, old_to_new)
    print("Citations in body and tables renumbered.")

    # 3. Reorder the reference list
    reorder_references(doc, old_to_new)
    print("Reference list reordered and MDPI typography applied.")

    doc.save(TARGET)
    print(f"Saved → {TARGET}")


if __name__ == "__main__":
    main()
