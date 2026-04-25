# Review report — `Predictive_MRQF_MS_ESWA_Elsevier.docx`

Companion to `Predictive_MRQF_MS_ESWA_Elsevier_REVIEW.docx`, in which every
item below is highlighted in **yellow** with the matching tag (R1, R2, …).
Open the `_REVIEW.docx`, search for "0.326" / "wins in 6 of 6 pairs" / etc.
to jump to each highlight; this file describes the issue and the fix.

**Overall verdict:** *Major revision* — the empirical work, statistics and
reproducibility appendix are solid, but the manuscript has (a) several
internal numeric mismatches, (b) load-bearing citations to unpublished
self-references, (c) an unfair "tuned LSTM" label, and (d) ESWA-fit drift
from §6.7 / §8.5 (complex-time / sophimatics). All of these are
addressable in revision without new experiments, except for one extra
LSTM grid run on the three currently-untuned pairs and a per-split
base-rate table.

---

## 1. Internal numeric mismatches (must fix before submission)

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R1** | Abstract | `MCC = 0.326` | Table 4 H=5 reports 0.325. Pick one and propagate. |
| **R2** | Highlights | "Cross-pair mean MCC 0.326 beats…" | Same as R1. |
| **R5** | §6.3 | "P-MRQF-MS wins in 6 of 6 pairs against all four baselines (uncorrected p = 0.031 in every case)" | Table 1 shows Logistic 5/6 wins, p = 0.063. Statement must read "6/6 against LSTM-tuned, XGBoost-classical and Persistence; 5/6 against Logistic". |
| **R6** | §6.5 | "DM tests confirm significantly lower loss for P-MRQF-MS against all four baselines on 4 of the 6 pairs" | EUR/USD has DM = +0.96, p = 0.34 vs Persistence — fails "all four". Recount; the correct figure is 3 of 6 (GBP, CHF near-miss, AUD). |
| **R7** | §5.1 | "(2,236 observations after the 500-day warm-up)" | 500 + 2236 + 1254 + 3465 = **7455 ≠ 7115** stated total. Off by 340. Reconcile (likely the warm-up overlaps with training). |
| **R24** | §9 Conclusions | "0.326" | Same mismatch as R1. |

## 2. Unfair / mis-labelled comparisons

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R3, R4** | §1 / Abstract | "tuned LSTM ensemble" | Grid search was run only on EUR/USD, EUR/JPY, EUR/CAD; the other three pairs use the *untuned* LSTM. The cross-pair "tuned" mean is therefore a mixture. Either run the grid on all six pairs, or relabel as "partially-tuned LSTM" and add an explicit caveat in the table footnote. |
| **R23** | §5.4 / §8.2 | Tuning *reduced* mean MCC (0.305 → 0.284) | Counter-intuitive. Add an early-stopping / learning-rate audit, or remove the "tuned" claim altogether — readers will assume tuning should never hurt. |

## 3. Load-bearing citations to unpublished / mis-attributed work

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R8** | §3.4 | `re-fitted every 250 trading days [12, 18]` | [18] = Engle 1982 (ARCH), does not discuss rolling refit. Replace with Brownlees & Gallo (already in list, currently uncited) or Andersen et al. 2006. |
| **R9** | §2.2 | `pre-registered in the protocol of [29]` | [29] = Iovane 2026a is "Unpublished manuscript, available on request". A "pre-registration" claim must point to an archived artefact (OSF / Zenodo / arXiv). Either archive the manuscript, or weaken to "described in [29]". |
| **R10** | §4.1 | `operator taxonomy and cooperation primitive of [29]` | Same problem: foundational definition resting on an unpublished file. Archive or replace with [24] (peer-reviewed Iovane work) where possible. |
| **R11, R12** | §6.7 / §8.5 | "two-dimensional complex-time perspective introduced by Iovane and Iovane [31]" | [31] is on **LLM hallucination mitigation**, not FX regime predictability. Carrying its formalism to FX is speculative and out-of-scope. Demote to a "see also" footnote, or reframe as the authors' analogy (not a derivation). |
| **R13** | References | [29] Iovane 2026a — unpublished | Cited 7× in body. Either archive on Zenodo with a DOI, or remove the dependency. |
| **R14** | References | [30] Iovane 2026b — unpublished | Cited 4× in body. Same fix as R13. |

## 4. Sections to move to Appendix or shorten

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R15** | §6.6 (E,S) trajectories during COVID | Qualitative-only sanity check; one figure, no test. | Move to Appendix B; keep one sentence in the main text. |
| **R16** | §6.7 Temporal heterogeneity across horizons | Three-pair "regime" classification on n = 6 — sample-dependent. | Either expand to ≥ 10 pairs to give the classification statistical weight, or shrink to a single paragraph + table. |
| **R17** | §7.2 Stacking on EUR/USD | Single-pair negative result. | Compress to one sentence in the Discussion; full details already in Appendix A.2. |
| **R18** | §8.5 Complex-time interpretation | Speculative, links the paper to a non-FX, non-ML reference [31]. | Cut or move to a "Future work" footnote; mainstream ESWA reviewers will find this distracting. |

## 5. Defensive / hedging language to remove

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R19** | §4.1 "On the nature of the contribution" | Pre-emptive rebuttal to a reviewer that has not yet been seen. | Move to the cover letter; in the manuscript, condense to one sentence. |
| **R20** | §4.2 "On the economic meaning of the target" | Same pattern. | Same fix as R19. |
| **R21** | §7.1 "honest negative result" | Hedging adverbs ("honest", "frankly", "we acknowledge that") undercut the contribution. | Remove `honest`, `frankly`, `we humbly`, etc. throughout (≈6 occurrences). |

## 6. Missing citations

| Tag | Where | Problem | Fix |
|-----|-------|---------|-----|
| **R22** | §1 Introduction | "persistence-dominated baselines that inflate apparent predictive accuracy" | Strong empirical claim with no source. Cite Patton (2011) "Volatility forecast comparison using imperfect volatility proxies" or Hansen & Lunde (2005). |
| (—) | §3.1 first technical mention of D4 wavelets | No citation at point of use. | Cite [13] Daubechies (1992) — currently in list but uncited. |
| (—) | §3.4 EGARCH refit | See R8. | Add Brownlees & Gallo [6] (currently uncited). |
| (—) | §5 LSTM baseline | LSTM described without canonical reference. | Cite [21] Hochreiter & Schmidhuber (1997) — currently uncited. |

## 7. Reference-list hygiene

29 of the 53 entries (≈55 %) are **never cited in the body**:

```
Bachelier, Bessembinder–Seguin, Bollerslev, Bouchaud–Potters,
Brownlees–Gallo, Calvet–Fisher, Cont, Corsi, Daubechies, Deng et al.,
Embrechts–Maejima, Hochreiter–Schmidhuber, Hosseini Rad,
Iovane–Laserra–Tortoriello, Iovane et al. 2020, Iovane 2024a,
Iovane–Chinnici 2024, Kahneman–Tversky, Landis–Koch, Mallat,
Mandelbrot 1963, Mandelbrot 1997, Mantegna–Stanley, Moody–Saffell,
Morales et al., Peters, Poon–Granger, Schwert, Wooldridge.
```

Action: prune aggressively. Either cite each in the body where it
genuinely supports a claim, or remove from the reference list. ESWA
reviewers explicitly check this.

### Self-citation footprint

* 9 of 53 references are Iovane-authored (≈17 %).
* 5 of those 9 are cited in the body, accounting for ≈21 in-text mentions.
* 2 of the most-used Iovane refs ([29], [30]) are unpublished — see §3.
* 4 Iovane entries appear in the list but are never cited
  ([23], [25], [27], [28]) — remove or reuse them.

## 8. Methodological gaps a reviewer will request

1. **Class-imbalance handling.** Threshold optimisation on validation MCC
   is the only mechanism currently used. Add focal loss / class-weighted
   loss baseline, PR-AUC per pair, and reliability diagrams. Report the
   per-split transition rate to rule out concept drift.
2. **Multi-horizon justification.** §4.2 fixes H = 5 *a priori*; only §6.7
   varies H. Provide a sensitivity table of MCC on validation across
   H ∈ {1, 3, 5, 10, 22} to motivate the choice.
3. **EGARCH refit cadence ablation.** Refit period (250 days) is a free
   parameter; report MCC for refit ∈ {125, 250, 500} or a likelihood
   stability check.
4. **Multiple-comparison control.** The full family is 4 baselines × 6
   pairs × 3 tests (Wilcoxon, DM, MCC-bootstrap). Holm only corrects within
   each pair. Add a hierarchical-FDR statement covering the full family.
5. **Power analysis of the n = 6 cross-pair tests.** Wilcoxon at n = 6 has
   minimum p = 0.031 — which the paper saturates against three baselines.
   Either expand to ≥ 10 pairs, or move the cross-pair claims to the
   robustness section and lead with per-pair Diebold–Mariano.

## 9. Suggested additions (one-liners)

* PR-AUC and per-pair Brier score (currently only Brier @ H=10).
* Per-split transition base-rate table (rules out drift).
* Calibration / reliability diagrams (supports the "calibrated
  probabilities" claim).
* Random-seed protocol for the LSTM ensemble (5 seeds? same as XGBoost?).
* CPU/GPU runtime per pair (supports the "minimal integration effort"
  deployment claim).

## 10. ESWA-fit verdict

**Borderline.** The decision-support framing (Article 12 AI Act,
SHAP audit trail, dashboard / risk-officer scenarios) is on-target for
ESWA. But §2 (golden-mean exponent), §6.1 (empirical falsification of
φ = 0.618), §6.6 (COVID trajectories), §6.7 + §8.5 (complex-time) read
as *Quantitative Finance* / econophysics content. To make the ESWA fit
unambiguous:

1. Cut or appendicise the complex-time interpretation (R11, R12, R18).
2. Reduce §2 to a one-paragraph operational summary.
3. Add one concrete decision-support evaluation — e.g. cost-sensitive
   utility under a margin-call rule, or a desk-dashboard latency /
   throughput study.

## Bottom line

The paper is fixable in revision. Four edits unlock the largest gains:

1. Reconcile **0.326 / 0.325 / 6-of-6 / 4-of-6** numbers (R1, R2, R5, R6, R24).
2. Either run the LSTM grid on all six pairs or relabel "tuned" → "partially-tuned" (R3, R4, R23).
3. Replace the two unpublished Iovane refs ([29], [30]) with archived versions or peer-reviewed alternatives (R9, R10, R13, R14).
4. Move §6.6, §7.2, §8.5 to an appendix, and shrink §6.7 (R15, R16, R17, R18).

After those, the bibliographic prune (§7), the missing citations (§6),
and the small additions (§9) are mostly clerical.
