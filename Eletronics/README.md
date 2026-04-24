# Tesi

`MRQF_MAS_Electronics.docx` is the MRQF-MAS manuscript rewritten into the MDPI
*Electronics* template. To regenerate it:

```
python3 build/rewrite_to_electronics.py
python3 build/verify.py
```

- `build/rewrite_to_electronics.py` clones the Electronics template, wipes its
  placeholder body and re-emits the MRQF content using the template's named
  `MDPI_*` styles (article type, title, authors, affiliations, abstract,
  keywords, horizontal rule, numbered sections, figure/table captions,
  back-matter, Abbreviations, Appendices, References, Publisher's Note).
- `build/verify.py` runs four structural/content checks: mandatory template
  sections, style coverage, figure/table/equation counts, and content fidelity
  against the source manuscript.
- `build/_sources/` holds the upstream template and the MRQF source docx.
- `build/media/` holds the ten figure PNGs extracted from the source.
