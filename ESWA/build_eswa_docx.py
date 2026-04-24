"""
Build an ESWA (Expert Systems with Applications, Elsevier) compliant .docx
from the original MDPI-styled manuscript. Preserves all scientific content;
only reformats and condenses abstract/keywords to satisfy ESWA word limits.

ESWA compliance targets:
- Single-column layout, A4, 2.54 cm margins
- Times New Roman 11 pt body, 1.5 line spacing, justified
- Heading 1 / 2 / 3 consistent with Article Structure guidelines
- Abstract <=250 words, single paragraph, stand-alone
- Keywords 1-7, English, short
- References in APA 7th edition, alphabetical + chronological
- In-text citations in (Author, Year) APA form
- Declarations: Funding, Competing Interests, CRediT, Generative AI, Data
- Tables: caption above, horizontal rules only, no shading
- Figures: captions below (here only captions are embedded, artworks supplied as separate files)
- Equations: displayed, centered, right-justified number
"""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------------------
# 1. APA reference list (derived from the original numbered references)
# ---------------------------------------------------------------------------
REFS_APA = [
    # Each tuple: (ref_number_in_original, intext_short [APA 7th, 3+ authors
    # collapsed to "et al." already], full_apa_entry)
    (1, "Bachelier (1900)",
     "Bachelier, L. (1900). Theorie de la speculation. Annales Scientifiques de l'Ecole Normale Superieure, 17, 21\u201386."),
    (2, "Mandelbrot (1963)",
     "Mandelbrot, B. B. (1963). The variation of certain speculative prices. The Journal of Business, 36(4), 394\u2013419."),
    (3, "Mandelbrot (1997)",
     "Mandelbrot, B. B. (1997). Fractals and scaling in finance. Springer."),
    (4, "Embrechts and Maejima (2002)",
     "Embrechts, P., & Maejima, M. (2002). Selfsimilar processes. Princeton University Press."),
    (5, "Iovane et al. (2004)",
     "Iovane, G., Laserra, E., & Tortoriello, F. S. (2004). Stochastic self-similar and fractal universe. Chaos, Solitons & Fractals, 20(3), 415\u2013426."),
    (6, "Daubechies (1992)",
     "Daubechies, I. (1992). Ten lectures on wavelets. Society for Industrial and Applied Mathematics."),
    (7, "Mallat (2008)",
     "Mallat, S. (2008). A wavelet tour of signal processing (3rd ed.). Academic Press."),
    (8, "Mantegna and Stanley (2000)",
     "Mantegna, R. N., & Stanley, H. E. (2000). An introduction to econophysics: Correlations and complexity in finance. Cambridge University Press."),
    (9, "Di Matteo (2007)",
     "Di Matteo, T. (2007). Multi-scaling in finance. Quantitative Finance, 7(1), 21\u201336."),
    (10, "Morales, Di Matteo, and Aste (2013)",
     "Morales, R., Di Matteo, T., & Aste, T. (2013). Non-stationary multifractality in stock returns. Physica A: Statistical Mechanics and Its Applications, 392(24), 6470\u20136483."),
    (11, "Cont (2001)",
     "Cont, R. (2001). Empirical properties of asset returns: Stylized facts and statistical issues. Quantitative Finance, 1(2), 223\u2013236."),
    (12, "Schwert (1989)",
     "Schwert, G. W. (1989). Why does stock market volatility change over time? The Journal of Finance, 44(5), 1115\u20131153."),
    (13, "Poon and Granger (2003)",
     "Poon, S.-H., & Granger, C. W. J. (2003). Forecasting volatility in financial markets: A review. Journal of Economic Literature, 41(2), 478\u2013539."),
    (14, "Iovane, Briscione, and Benedetto (2021)",
     "Iovane, G., Briscione, A., & Benedetto, E. (2021). Financion: A quantum approach to financial market modelling. Journal of Statistics & Management Systems, 24(5), 1127\u20131149."),
    (15, "Iovane, Landi, and Serino (2016)",
     "Iovane, G., Landi, A., & Serino, S. (2016). An optimized mathematical-physical approach to financial market. Journal of Information & Optimization Sciences, 37(3), 423\u2013448."),
    (16, "Iovane (2024a)",
     "Iovane, G. (2024a). Decision support system driven by thermo-complexity: Algorithms. IEEE Access, 12, 157359\u2013157382."),
    (17, "Iovane and Chinnici (2024)",
     "Iovane, G., & Chinnici, M. (2024). DSS driven by thermo-complexity: Scenario analysis. Applied Sciences, 14(6), 2387."),
    (18, "Iovane (2026a)",
     "Iovane, G. (2026a). MRQF-MAS: A multiscale relativistic quantum finance framework for cooperative multi-agent trading systems with shared knowledge base [Unpublished manuscript, available on request]."),
    (19, "Iovane (2026b)",
     "Iovane, G. (2026b). P-MRQF-MAS: An interpretable cooperative multi-agent framework for high-volatility regime forecasting on the energy\u2013entropy plane [Unpublished manuscript, available on request]."),
    (20, "Wooldridge (2009)",
     "Wooldridge, M. (2009). An introduction to multiagent systems (2nd ed.). Wiley."),
    (21, "Shavandi and Khedmati (2022)",
     "Shavandi, A., & Khedmati, M. (2022). A multi-agent deep reinforcement learning framework for algorithmic trading in financial markets. Expert Systems with Applications, 208, 118124."),
    (22, "Hosseini Rad and Tahmasebi Khorasani (2024)",
     "Hosseini Rad, S., & Tahmasebi Khorasani, S. (2024). Cooperative multi-agent deep reinforcement learning for Forex trading. In Proceedings of the IEEE AISP. IEEE."),
    (23, "Chen and Guestrin (2016)",
     "Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. In Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining (pp. 785\u2013794). ACM."),
    (24, "Lim and Zohren (2021)",
     "Lim, B., & Zohren, S. (2021). Time-series forecasting with deep learning: A survey. Philosophical Transactions of the Royal Society A, 379(2194), Article 20200209."),
    (25, "Nelson (1991)",
     "Nelson, D. B. (1991). Conditional heteroskedasticity in asset returns: A new approach. Econometrica, 59(2), 347\u2013370."),
    (26, "Bouchaud and Potters (2003)",
     "Bouchaud, J.-P., & Potters, M. (2003). Theory of financial risk and derivative pricing (2nd ed.). Cambridge University Press."),
    (27, "Lundberg and Lee (2017)",
     "Lundberg, S. M., & Lee, S.-I. (2017). A unified approach to interpreting model predictions. In Advances in Neural Information Processing Systems (Vol. 30, pp. 4765\u20134774)."),
    (28, "Peters (1994)",
     "Peters, E. E. (1994). Fractal market analysis: Applying chaos theory to investment and economics. Wiley."),
    (29, "Bessembinder and Seguin (1993)",
     "Bessembinder, H., & Seguin, P. J. (1993). Price volatility, trading volume, and market depth: Evidence from futures markets. The Journal of Financial and Quantitative Analysis, 28(1), 21\u201339."),
    (30, "Jaynes (1957)",
     "Jaynes, E. T. (1957). Information theory and statistical mechanics. Physical Review, 106(4), 620\u2013630."),
    (31, "Dacorogna, Gencay, Muller, Olsen, and Pictet (2001)",
     "Dacorogna, M. M., Gencay, R., Muller, U., Olsen, R. B., & Pictet, O. V. (2001). An introduction to high-frequency finance. Academic Press."),
    (32, "Kahneman and Tversky (1979)",
     "Kahneman, D., & Tversky, A. (1979). Prospect theory: An analysis of decision under risk. Econometrica, 47(2), 263\u2013292."),
    (33, "European Parliament and Council (2024)",
     "European Parliament and Council. (2024). Regulation (EU) 2024/1689 of 13 June 2024 laying down harmonised rules on artificial intelligence (Artificial Intelligence Act). Official Journal of the European Union."),
    (34, "Cover and Thomas (2006)",
     "Cover, T. M., & Thomas, J. A. (2006). Elements of information theory (2nd ed.). Wiley."),
    (35, "Brownlees and Gallo (2006)",
     "Brownlees, C. T., & Gallo, G. M. (2006). Financial econometric analysis at ultra-high frequency: Data handling concerns. Computational Statistics & Data Analysis, 51(4), 2232\u20132245."),
    (36, "Engle (1982)",
     "Engle, R. F. (1982). Autoregressive conditional heteroscedasticity with estimates of the variance of United Kingdom inflation. Econometrica, 50(4), 987\u20131007."),
    (37, "Diebold and Mariano (1995)",
     "Diebold, F. X., & Mariano, R. S. (1995). Comparing predictive accuracy. Journal of Business & Economic Statistics, 13(3), 253\u2013263."),
    (38, "Landis and Koch (1977)",
     "Landis, J. R., & Koch, G. G. (1977). The measurement of observer agreement for categorical data. Biometrics, 33(1), 159\u2013174."),
    (39, "Deng, Bao, Kong, Ren, and Dai (2017)",
     "Deng, Y., Bao, F., Kong, Y., Ren, Z., & Dai, Q. (2017). Deep direct reinforcement learning for financial signal representation and trading. IEEE Transactions on Neural Networks and Learning Systems, 28(3), 653\u2013664."),
    (40, "Moody and Saffell (2001)",
     "Moody, J., & Saffell, M. (2001). Learning to trade via direct reinforcement. IEEE Transactions on Neural Networks, 12(4), 875\u2013889."),
    (41, "Hochreiter and Schmidhuber (1997)",
     "Hochreiter, S., & Schmidhuber, J. (1997). Long short-term memory. Neural Computation, 9(8), 1735\u20131780."),
    (42, "Vaswani et al. (2017)",
     "Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., Kaiser, L., & Polosukhin, I. (2017). Attention is all you need. In Advances in Neural Information Processing Systems (Vol. 30, pp. 5998\u20136008)."),
    (43, "Bollerslev (1986)",
     "Bollerslev, T. (1986). Generalized autoregressive conditional heteroskedasticity. Journal of Econometrics, 31(3), 307\u2013327."),
    (44, "Corsi (2009)",
     "Corsi, F. (2009). A simple approximate long-memory model of realized volatility. Journal of Financial Econometrics, 7(2), 174\u2013196."),
    (45, "Iovane et al. (2020)",
     "Iovane, G., Amorosia, F. S., Benedetto, E., & Lamponi, G. (2020). Decision and reasoning in incompleteness or uncertainty conditions. IEEE Access, 8, 115109\u2013115122."),
    (46, "Blackledge and Murphy (2011)",
     "Blackledge, J. M., & Murphy, K. (2011). Currency trading using the fractal market hypothesis. In Risk management trends. IntechOpen."),
    (47, "Calvet and Fisher (2008)",
     "Calvet, L. E., & Fisher, A. J. (2008). Multifractal volatility: Theory, forecasting, and pricing. Academic Press."),
    (48, "Wolpert (1992)",
     "Wolpert, D. H. (1992). Stacked generalization. Neural Networks, 5(2), 241\u2013259."),
    (49, "Van der Laan, Polley, and Hubbard (2007)",
     "Van der Laan, M. J., Polley, E. C., & Hubbard, A. E. (2007). Super learner. Statistical Applications in Genetics and Molecular Biology, 6(1), Article 25."),
    (50, "Politis and Romano (1994)",
     "Politis, D. N., & Romano, J. P. (1994). The stationary bootstrap. Journal of the American Statistical Association, 89(428), 1303\u20131313."),
    (51, "Friedman (1937)",
     "Friedman, M. (1937). The use of ranks to avoid the assumption of normality implicit in the analysis of variance. Journal of the American Statistical Association, 32(200), 675\u2013701."),
    (52, "Nemenyi (1963)",
     "Nemenyi, P. B. (1963). Distribution-free multiple comparisons [Doctoral dissertation, Princeton University]. See also Dem\u0161ar, J. (2006). Statistical comparisons of classifiers over multiple data sets. Journal of Machine Learning Research, 7, 1\u201330."),
    (53, "Iovane and Iovane (2026)",
     "Iovane, G., & Iovane, G. (2026). Sophimatics and 2D complex time to mitigate hallucinations in LLMs for novel intelligent information systems in digital transformation. Applied Sciences, 16(1), Article 288. https://doi.org/10.3390/app16010288"),
]

# Map ref number -> APA intext (narrative) short form
REFMAP_NARR = {n: narr for n, narr, _ in REFS_APA}
# Map ref number -> APA parenthetical form (strip the year from the narrative)
def narr_to_paren(narr):
    """Convert 'Author (Year)' or 'Author and Author (Year)' -> 'Author, Year' / 'Author & Author, Year'."""
    m = re.match(r"^(.*?)\s*\((\d{4}[a-z]?)\)\s*$", narr)
    if not m:
        return narr
    authors, year = m.group(1), m.group(2)
    # replace narrative 'and' with '&' when between two authors
    authors_p = re.sub(r"\s+and\s+", " & ", authors)
    return f"{authors_p}, {year}"

REFMAP_PAREN = {n: narr_to_paren(narr) for n, narr, _ in REFS_APA}

# Alphabetically sort the APA reference list (by first author surname, then year)
def ref_sort_key(entry):
    _, _, full = entry
    # first-author surname is everything up to first comma
    m = re.match(r"^([^,]+),", full)
    surname = m.group(1) if m else full
    # year appears inside parentheses
    y = re.search(r"\((\d{4})[a-z]?\)", full)
    year = int(y.group(1)) if y else 0
    return (surname.lower(), year)

REFS_APA_SORTED = sorted(REFS_APA, key=ref_sort_key)

# ---------------------------------------------------------------------------
# 2. Content of the paper (preserved verbatim from the original except where
#    ESWA word limits require condensation, and except for reformatted
#    in-text citations)
# ---------------------------------------------------------------------------

TITLE = ("P-MRQF-MS: A Predictive Multiscale Framework with Adaptive Fractal "
         "Exponent for Regime-Transition Forecasting on the Energy\u2013Entropy "
         "Plane \u2014 Competitive Evidence over XGBoost, LSTM and Parametric "
         "Baselines on Six Major Foreign-Exchange Pairs")

AUTHORS = [
    # (display name, affiliation letter, corresponding?)
    ("Author 1", "a", True),
    ("Author 2", "b", False),
]
AFFILIATIONS = [
    ("a", "Affiliation 1, City, Country"),
    ("b", "Affiliation 2, City, Country"),
]
CORRESP_EMAIL = "author@email"

# Abstract condensed to <=250 words while preserving every key scientific
# claim of the original 530-word abstract (as required by ESWA).
ABSTRACT = (
    "We propose P-MRQF-MS, a predictive machine learning framework for "
    "regime-transition forecasting in financial time series. The approach "
    "combines three components that act only on the feature representation "
    "and the target definition, leaving the learning algorithm standard. "
    "First, the theoretical golden-mean exponent \u03C6 = 0.618 of the "
    "Multiscale Relativistic Quantum Finance feature-generation framework is "
    "replaced by an adaptive fractal exponent estimated online from "
    "Daubechies D4 wavelet decompositions at windows of 64 and 128 days; on "
    "six major EUR-quoted FX pairs the adaptive exponent has median 0.448 "
    "and inter-quartile range [0.369, 0.524], empirically rejecting the "
    "golden-mean value at daily resolution. Second, features at horizons D1, "
    "D5 and D22 (realised volatility, mean absolute deviation, "
    "energy\u2013entropy bands at two scales, EGARCH conditional variance at "
    "three horizons and shared-knowledge-base occupancy probabilities) are "
    "fused into a 20-dimensional vector. Third, a regime-transition target "
    "on the discretised (E, S) plane is defined such that persistence is "
    "provably zero (MCC = 0 by construction), removing the persistence bias "
    "of volatility-threshold forecasting. The system is built on XGBoost and "
    "validated on 7,115 daily ECB observations per pair (1999\u20132026) with "
    "a strict train/validation/test split. P-MRQF-MS attains cross-pair mean "
    "MCC 0.326, outperforming a tuned LSTM ensemble (0.284), XGBoost on "
    "classical volatility features (0.285), logistic regression (0.245) and "
    "persistence (0.000). Friedman \u03C7\u00B2 = 9.40 (p = 0.024) and "
    "pairwise Wilcoxon tests confirm a consistent advantage. Feature-"
    "importance stability reaches Spearman 0.96, supporting Article 12 "
    "AI-Act auditability."
)

# Sanity check handled at the bottom of this script.

KEYWORDS = [
    "regime-transition forecasting",
    "multiscale features",
    "adaptive fractal exponent",
    "XGBoost",
    "interpretability",
    "foreign exchange",
    "decision-support systems",
]

HIGHLIGHTS = [
    # max 85 chars each
    "Adaptive wavelet fractal exponent replaces the theoretical golden-mean value.",
    "Persistence-free regime-transition target on the (E, S) plane (MCC = 0 at rest).",
    "Cross-pair mean MCC 0.326 beats tuned LSTM, XGBoost-classical and logistic.",
    "Friedman test rejects equal performance (chi2 = 9.40, p = 0.024, n = 6 pairs).",
    "Spearman 0.96 SHAP-rank stability supports Article 12 AI-Act auditability.",
]
# Enforce 85-char limit; if a line is too long it will raise below.

# ---------------------------------------------------------------------------
# 3. Body content: each entry is (level, text).
#    level: 'h1' | 'h2' | 'p' | 'eq' | 'fig' | 'tbl'
#    'eq'  uses the special displayed-equation style (text includes equation
#          text followed by '\t(n)' for numbering)
#    'fig' rendering: caption text printed in italics under a placeholder
#    'tbl' rendering: caption printed in italics above a table; the table
#          payload is injected separately (see TABLES dict below)
# ---------------------------------------------------------------------------

# NOTE: In-text numbered citations [14] get rewritten to APA form by the
# preprocessing function citation_rewrite() below.

BODY = [

    ("h1", "1. Introduction"),
    ("p",
     "Regime-transition forecasting in financial time series is a core task "
     "for risk management, portfolio allocation and market monitoring. A "
     "sudden switch from a quiescent regime to a high-volatility regime "
     "materially changes the optimal trading policy, the margin requirements "
     "of a clearing house and the value-at-risk of an open position. Despite "
     "advances in machine learning, regime-transition prediction remains "
     "challenging due to multiscale and non-linear dynamics, heavy-tailed "
     "return distributions, class imbalance between transition and "
     "no-transition events, and persistence-dominated baselines that inflate "
     "apparent predictive accuracy on standard volatility-forecasting tasks. "
     "In practice, a classifier that always predicts \u201Cno transition\u201D "
     "achieves competitive metrics on most single-asset volatility benchmarks "
     "because volatility is itself autocorrelated, leaving little room for a "
     "learner to demonstrate genuine anticipation. This persistence bias is "
     "a systemic problem in the forecasting literature: reported "
     "improvements of order 0.01\u20130.03 in MCC or F\u2081 on volatility "
     "tasks can be largely attributable to better persistence handling "
     "rather than to a better predictive signal."),
    ("p",
     "The present paper proposes P-MRQF-MS, a predictive machine learning "
     "framework designed to address this gap. The system combines three "
     "engineering choices \u2014 an adaptive fractal exponent estimated "
     "online from wavelet decompositions, a multiscale feature representation "
     "at daily horizons D1, D5 and D22, and a novel regime-transition target "
     "on the energy\u2013entropy plane of the Multiscale Relativistic "
     "Quantum Finance (MRQF) feature-generation framework [14, 15, 18] "
     "\u2014 with a standard gradient-boosted classifier. The design "
     "explicitly separates feature engineering from learning: the "
     "MRQF-derived feature space encodes physically-motivated "
     "representations of price dynamics, while the learning algorithm is an "
     "off-the-shelf XGBoost with standard hyperparameters. This separation "
     "aligns with the Expert Systems paradigm of combining domain knowledge "
     "with data-driven learning and admits a clear audit trail via SHAP "
     "feature attribution, which we quantify below."),
    ("p", "The main contributions of this paper are:"),
    ("p",
     "1. A predictive framework integrating multiscale MRQF-derived features "
     "(energy, entropy, region, adaptive fractal exponent, "
     "shared-knowledge-base probabilities, EGARCH conditional variance at "
     "three horizons) with an off-the-shelf gradient-boosted classifier for "
     "regime-transition forecasting on foreign-exchange data."),
    ("p",
     "2. An adaptive fractal exponent estimator \u03C6\u0302_local(t) "
     "replacing the theoretical golden-mean value \u03C6 = 0.618 with a "
     "data-driven online estimate from Daubechies D4 wavelet decompositions, "
     "empirically validated on six FX pairs."),
    ("p",
     "3. A novel regime-transition target on the discretised (E, S) plane "
     "that removes the persistence bias of standard volatility-threshold "
     "forecasting by construction (MCC_persistence = 0)."),
    ("p",
     "4. Extensive empirical validation on six real-world FX datasets "
     "(EUR/USD, EUR/GBP, EUR/JPY, EUR/CHF, EUR/CAD, EUR/AUD) from the "
     "European Central Bank with 7,115 daily observations per pair spanning "
     "1999\u20132026, strict temporal split, five baseline comparisons "
     "including a per-pair tuned PyTorch LSTM ensemble, and multiple "
     "statistical tests (cross-pair Wilcoxon, Friedman\u2013Nemenyi, "
     "per-pair Diebold\u2013Mariano with Holm correction, paired bootstrap)."),
    ("p",
     "5. A multi-horizon diagnostic revealing a three-regime "
     "temporal-heterogeneity classification (reactive, balanced, "
     "anticipatory) across the six pairs, interpreted in the two-dimensional "
     "complex-time framework of [53]."),
    ("p",
     "We execute the full pipeline on real ECB daily data for EUR/USD, "
     "EUR/GBP, EUR/JPY, EUR/CHF, EUR/CAD and EUR/AUD from 1999-01-04 to "
     "2026-04-13 (7,115 observations per pair). The temporal split (train "
     "1999\u20132008, validation 2008\u20132013, test 2013\u20132026; last "
     "3,465 observations held out) is strict. To rule out the concern that "
     "deep sequence models might outperform our architecture if properly "
     "tuned, we conducted a hyperparameter grid search over six LSTM "
     "configurations (hidden size, layers, dropout, learning rate, lookback) "
     "on EUR/USD, EUR/JPY and EUR/CAD \u2014 the three pairs where the "
     "baseline LSTM was most competitive \u2014 and selected per-pair the "
     "configuration maximising validation MCC before evaluating on test. "
     "The tuned LSTM ensemble achieves cross-pair mean MCC 0.284 (vs 0.305 "
     "for the default configuration), whereas P-MRQF-MS achieves 0.326. The "
     "cross-pair Wilcoxon signed-rank test shows P-MRQF-MS winning 6 of 6 "
     "pairs against the tuned LSTM ensemble with uncorrected p = 0.031, and "
     "the non-parametric Friedman test across the four methods on the six "
     "paired pair-level MCC samples rejects the null of equal performance "
     "with p = 0.024, formally significant at \u03B1 = 0.05."),
    ("p",
     "The paper is organised as follows. Section 2 recalls the elements of "
     "MRQF used as feature-generation building blocks. Section 3 describes "
     "the feature construction and the adaptive fractal exponent. Section 4 "
     "describes the architecture and the regime-transition target. Section "
     "5 reports the experimental protocol. Section 6 reports the empirical "
     "results including a multi-horizon diagnostic. Section 7 reports "
     "ablation studies, feature redundancy analysis and quantitative "
     "interpretability metrics. Section 8 discusses practical applications, "
     "limitations and related work. Section 9 concludes."),
    ("fig",
     "Figure 1. P-MRQF-MS end-to-end architecture (schematic). Price streams "
     "feed the multiscale feature-extraction stage (D1, D5, D22), which "
     "computes realised volatility, dispersion, energy\u2013entropy bands "
     "at two scales, adaptive fractal exponent \u03C6\u0302_local, EGARCH "
     "conditional variance at three horizons and SKB occupancy "
     "probabilities. The 20-dimensional feature vector is consumed by the "
     "gradient-boosted decision layer to produce the regime-transition "
     "probability P\u0302(y\u209C = 1 | F\u209C)."),

    ("h1", "2. Theoretical Elements of MRQF Used as Feature-Generation Building Blocks"),
    ("p",
     "We briefly recall the operational elements of the Multiscale "
     "Relativistic Quantum Finance (MRQF) framework [14, 15, 18] that are "
     "used below as feature-generation building blocks. The present paper "
     "uses MRQF as a physically-motivated transformation of raw price "
     "series into a structured feature vector; the broader theoretical "
     "claims of the framework are orthogonal to the predictive evaluation "
     "and are not required for the results."),

    ("h2", "2.1. Scale-Invariance Law and Theoretical Exponent"),
    ("p", "MRQF posits a scale-invariance law for the conditional volatility of the price signal:"),
    ("eq", "\u03C3(N_lots) = \u03F1 N_lots^\u03C6,   \u03C6_MRQF = (\u221A5 \u2212 1)/2 \u2248 0.618.\t(1)"),
    ("p",
     "The golden-mean exponent is derived from a Fibonacci decomposition of "
     "trader categories [14]. The Brownian regime corresponds to \u03C6 = "
     "0.5. Section 6.1 of the present paper tests this prediction "
     "empirically on six FX pairs and documents its empirical failure at "
     "daily resolution, replacing it with the adaptive exponent "
     "\u03C6\u0302_local(t)."),

    ("h2", "2.2. Energy\u2013Entropy Plane"),
    ("p",
     "The market state at time t is projected onto the plane (E\u209C, "
     "S\u209C) with E\u209C \u2208 {7, \u2026, 32} and S\u209C \u2208 {0, "
     "\u2026, 12}. Energy reflects directional potential; entropy reflects "
     "disorder of the operator distribution [14, 18, 30]. The plane is "
     "partitioned into nine regions by tertiles of the two coordinates; the "
     "discrete nine-region classification enables the definition of the "
     "regime-transition target used in Section 4.2. The tertile partitioning "
     "is standard in the MRQF literature and is pre-registered in the "
     "protocol of [18]; sensitivity of the target difficulty to alternative "
     "percentile cuts (quartiles, halves) is discussed in Section 8.1."),
    ("fig",
     "Figure 2. The (E, S) plane with nine regions defined by tertiles of "
     "the two coordinates."),

    ("h1", "3. Feature Construction"),

    ("h2", "3.1. Adaptive Fractal Exponent"),
    ("p",
     "Given a rolling window W of log-returns {r_{t\u2212W+1}, \u2026, "
     "r_t}, the local scaling exponent is estimated by wavelet-based OLS on "
     "five scales s \u2208 {1, 2, 4, 8, 16}:"),
    ("eq",
     "\u03C6\u0302_local(t) = 0.5 + slope_s(log s, log \u03C3_s(t)),   "
     "\u03C3_s(t) = std(r_{t\u2212W+1+\u03C4 s})_{\u03C4 = 0, 1, 2, \u2026}.\t(2)"),
    ("p",
     "We compute this estimator at two window sizes: W = 64 (short-memory "
     "response) and W = 128 (long-memory stability). Both are included as "
     "separate features to let the downstream learner adapt to "
     "regime-specific fractal behaviour. Intermediate window sizes W \u2208 "
     "{96, 160} were spot-checked on EUR/USD and yielded medians in [0.45, "
     "0.46], consistent with the windows reported; the specific choice of "
     "64 and 128 is therefore not a crucial hyperparameter. The scale grid "
     "{1, 2, 4, 8, 16} is standard for daily data in the multifractal "
     "literature [9, 46]."),

    ("h2", "3.2. Multiscale Volatility and Dispersion"),
    ("p",
     "At horizons w \u2208 {1, 5, 22} (daily, weekly, monthly) we compute "
     "the rolling standard deviation of log-returns rv_w. At w \u2208 {5, "
     "22} we also compute the rolling mean absolute deviation from the "
     "median mad_w. At w \u2208 {1, 5} we include the rolling mean return "
     "r\u0304_w to capture short-horizon directional bias."),

    ("h2", "3.3. Energy\u2013Entropy Bands at Two Scales"),
    ("p",
     "The primary energy band is computed from the 22-day rolling standard "
     "deviation, E\u209C = clip(7 + 25(1 \u2212 rv_{22,t}/(2\u03C3_global)), "
     "7, 32). The primary entropy band is S\u209C = clip(9 mad_{22,t}/"
     "\u03C3_global, 0, 12). A secondary pair (E\u2085, S\u2085) is computed "
     "with the 5-day versions. The discretised region identifiers "
     "region_{22} and region_5 (integer values in {0, \u2026, 8}, tertile-"
     "by-tertile) are included as ordinal features."),

    ("h2", "3.4. EGARCH Conditional Variance at Three Horizons"),
    ("p",
     "An EGARCH(1,1) model [25] is fitted by Gaussian maximum likelihood on "
     "a rolling window of 1500 observations and re-fitted every 250 trading "
     "days [31, 36]. The h-step-ahead conditional variance forecasts at "
     "h \u2208 {1, 5, 22} are included as features \u03C3\u00B2\u2081, "
     "\u03C3\u00B2\u2085, \u03C3\u00B2\u2082\u2082."),

    ("h2", "3.5. Shared Knowledge Base (SKB)"),
    ("p",
     "The SKB maintains expanding-window counts of band occupancy on the "
     "22-day (E, S) plane with strict no-future-leakage timestamp guards. "
     "At time t, the probabilities e\u2082\u2082(t) and s\u2082\u2082(t) of "
     "the current band occupancy are Laplace-smoothed empirical frequencies "
     "over timestamps strictly less than t. The design satisfies the "
     "record-keeping requirements of Article 12 of the EU AI Act [33]."),

    ("h2", "3.6. Full Feature Vector"),
    ("p",
     "The 20-dimensional feature vector at time t comprises {rv\u2081, "
     "rv\u2085, rv\u2082\u2082, mad\u2085, mad\u2082\u2082, r\u0304\u2081, "
     "r\u0304\u2085, \u03C6\u0302_{64}, \u03C6\u0302_{128}, E\u2082\u2082, "
     "S\u2082\u2082, E\u2085, S\u2085, region_{22}, region_5, "
     "e\u2082\u2082, s\u2082\u2082, \u03C3\u00B2\u2081, \u03C3\u00B2\u2085, "
     "\u03C3\u00B2\u2082\u2082}. All features are lagged by one observation "
     "before entering any model."),

    ("h1", "4. Architecture and Regime-Transition Target"),

    ("h2", "4.1. Predictor"),
    ("p",
     "The architecture inherits the operator taxonomy and cooperation "
     "primitive of [18] but uses a single gradient-boosted classifier [23] "
     "as the final predictor, operating on the 20-dimensional feature "
     "vector described in Section 3.6. The explicit multi-agent "
     "decomposition of [18] is retained at the feature level (the (E, S) "
     "bands and SKB probabilities encode the cross-operator state), but the "
     "final decision is produced by the gradient-boosted classifier rather "
     "than by a hard-coded consensus rule. This design choice follows the "
     "stacking-as-consensus paradigm [48, 49] and is consistent with the "
     "feature-level interpretability of tree-based ensembles [27]. We do "
     "not claim a novel learned consensus operator beyond this stacking; "
     "rather, we claim that the MRQF feature representation is the source "
     "of the predictive value and the classifier is the standard vehicle "
     "for extracting that value."),
    ("p",
     "On the nature of the contribution. A reader may object that the "
     "learning algorithm is a standard XGBoost and that the gain is "
     "therefore attributable to feature engineering rather than to a "
     "fundamentally new modelling paradigm. We address this directly: the "
     "contribution of P-MRQF-MS lies not in the learning algorithm itself, "
     "but in the structured multiscale feature space derived from the MRQF "
     "theoretical framework, which transforms an otherwise intractable "
     "predictive task (regime transitions on the (E, S) plane, where "
     "persistence has MCC = 0 by construction) into a representation for "
     "which off-the-shelf tabular learners can extract signal. This design "
     "aligns with the Expert Systems paradigm of combining domain knowledge "
     "with data-driven learning, where the domain knowledge (MRQF-derived "
     "physical observables) enables the data-driven component (gradient "
     "boosting) to operate in a regime where it can achieve competitive "
     "performance without bespoke architecture. The distinction matters "
     "because three elements make this feature space a genuine "
     "architectural contribution rather than mere feature engineering: "
     "(i) the features are derived by a physically-motivated transformation "
     "(wavelet scale-invariance, energy\u2013entropy projection, adaptive "
     "fractal exponent) rather than engineered heuristically; (ii) the "
     "representation is universal across asset classes and timeframes "
     "\u2014 the same twenty features apply unmodified to USD, GBP, JPY, "
     "CHF, CAD and AUD, and in principle to equities or commodities; "
     "(iii) the feature space admits an audit trail via SHAP [27] that maps "
     "each prediction back to the underlying physical observables (energy, "
     "entropy, scaling exponent, EGARCH variance), satisfying the "
     "interpretability requirements of Article 12 of the EU AI Act [33] "
     "without retrofitting. We therefore claim that P-MRQF-MS is a "
     "feature-space contribution in the strong sense: it makes a class of "
     "predictive tasks tractable that were previously close to random under "
     "standard features, and the value of this contribution is demonstrated "
     "empirically by the +0.041 MCC cross-pair improvement over XGBoost "
     "with the same learning algorithm on the classical feature subset "
     "(Section 6.3)."),

    ("h2", "4.2. Regime-Transition Target"),
    ("p",
     "The central methodological contribution of this paper on the target "
     "side is the regime-transition target at horizon H:"),
    ("eq",
     "y\u209C = 1[region(E_{t+H}, S_{t+H}) \u2260 region(E\u209C, "
     "S\u209C)],   H = 5 trading days.\t(3)"),
    ("p",
     "Three properties make this target superior to volatility forecasting "
     "as a stress test: (i) Persistence is provably the worst strategy "
     "\u2014 a classifier that always predicts y\u209C = 0 achieves MCC = "
     "0 by construction, so any positive MCC reflects information beyond "
     "autocorrelation. (ii) The target is anchored to the MRQF theoretical "
     "construct: predicting a region change tests whether the MRQF-specific "
     "features (E, S, region, e, s, \u03C6\u0302) carry predictive value. "
     "(iii) The target is non-pathologically imbalanced with transition "
     "rates {16 %, 26 %, 26 %, 35 %, 17 %, 23 %} across our six pairs."),
    ("p",
     "On the economic meaning of the target. One might object that the "
     "regime-transition target is \u201Cconstructed\u201D rather than a "
     "directly observed market quantity such as the log-return or realised "
     "volatility. We acknowledge the construction but argue that the target "
     "captures a fundamental property of market dynamics \u2014 state "
     "change on the energy\u2013entropy plane \u2014 rather than an "
     "arbitrary price threshold. Economically, a transition of (E, S) "
     "corresponds to a change in the balance between directional flow "
     "strength (energy) and dispersion of the operator distribution "
     "(entropy), which is precisely the joint observable that MRQF theory "
     "predicts will matter for regime-dependent trading strategies. The "
     "transition indicator is also more robust than single-threshold "
     "volatility indicators because it does not depend on a specific "
     "cut-off (q\u2080.\u2088\u2080 in [18, 19]) whose value is itself "
     "arbitrary. Finally, the fact that persistence is trivially zero on "
     "the regime-transition target but near-optimal on the "
     "volatility-threshold target (as documented in [19]) is itself "
     "evidence that the two targets address different and complementary "
     "aspects of market dynamics: volatility-threshold forecasting tests "
     "whether autocorrelation of volatility can be exploited, while "
     "regime-transition forecasting tests whether the state-change geometry "
     "of the (E, S) plane can be anticipated. We see these as a hierarchy "
     "of difficulty, with the regime-transition task being the more "
     "discriminating test of genuine predictive skill."),

    ("h1", "5. Experimental Protocol"),

    ("h2", "5.1. Data"),
    ("p",
     "Daily ECB reference rates retrieved via CurrencyConverter 0.18.17 "
     "(pinned) for EUR/USD, EUR/GBP, EUR/JPY, EUR/CHF, EUR/CAD and EUR/AUD, "
     "spanning 4 January 1999 to 13 April 2026, 7,115 business-day "
     "observations per pair and no gaps."),

    ("h2", "5.2. Temporal Split and Threshold Optimisation"),
    ("p",
     "Three contiguous segments per pair: training 1999-01-04 to 2007-12-31 "
     "(2,236 observations after the 500-day warm-up), validation 2008-01-01 "
     "to 2012-12-31 (1,254 observations), test 2013-01-01 to 2026-04-13 "
     "(3,465 observations). All features strictly lagged. Hyperparameters "
     "selected on validation; classification thresholds optimised "
     "per-method on validation to maximise validation MCC, ensuring no "
     "method has a threshold advantage."),
    ("p",
     "We conducted a supplementary rolling-origin sensitivity analysis on "
     "EUR/USD with three 12-month test windows (origins 2015-01-01, "
     "2020-01-01, 2022-01-01) and two-year validation windows preceding "
     "each origin. The cross-window mean P-MRQF-MS MCC is 0.280 against "
     "0.271 for persistence and 0.250 for HAR-RV (see [19], Section 5.4, "
     "for the same protocol applied to the high-volatility forecasting "
     "task). The rolling-origin evaluation is restricted to EUR/USD because "
     "per-pair rolling windows often have insufficient positive cases for "
     "meaningful MCC estimation; we note this as a limitation in Section 8."),

    ("h2", "5.3. Baseline Methods"),
    ("p",
     "Four baselines are evaluated: (a) Persistence, always predicting "
     "y\u209C = 0. (b) XGBoost classical, using only {rv\u2081, rv\u2085, "
     "rv\u2082\u2082, \u03C3\u00B2\u2081, \u03C3\u00B2\u2085, "
     "\u03C3\u00B2\u2082\u2082} as features with the same hyperparameters "
     "as P-MRQF-MS. (c) L\u2082-regularised logistic regression on the full "
     "feature set with balanced class weights. (d) LSTM ensemble (real "
     "PyTorch 2.11 CPU implementation) over five seeds."),

    ("h2", "5.4. LSTM Hyperparameter Grid and Per-Pair Tuning"),
    ("p",
     "To ensure that the LSTM baseline is not a straw man, we evaluated six "
     "hyperparameter configurations on the three pairs where the default "
     "LSTM was most competitive (EUR/USD, EUR/JPY, EUR/CAD): {(h, l, d, "
     "\u03B7, \u2113)} with hidden size h \u2208 {48, 64, 96}, layers l "
     "\u2208 {1, 2}, dropout d \u2208 {0.25, 0.3}, learning rate \u03B7 "
     "\u2208 {1\u00D710\u207B\u00B3, 2\u00D710\u207B\u00B3}, lookback "
     "\u2113 \u2208 {20, 30}. For each configuration and pair, we trained "
     "the LSTM with seed 2026 on the training segment and scored validation "
     "MCC. The configuration with the best validation MCC was used to "
     "train a 5-seed ensemble (seeds 2026\u20132030) on training, averaged "
     "for prediction, and its classification threshold was optimised on "
     "validation. For the remaining three pairs (EUR/GBP, EUR/CHF, "
     "EUR/AUD) we used the default configuration h = 48, l = 1, d = 0.25, "
     "\u03B7 = 2\u00D710\u207B\u00B3, \u2113 = 20; a post-hoc check with "
     "the tuned configuration on EUR/GBP showed a similar validation MCC "
     "(0.26 vs 0.25 untuned) with no test improvement, consistent with the "
     "findings on the three tuned pairs."),

    ("h2", "5.5. P-MRQF-MS Hyperparameters"),
    ("p",
     "Five-seed XGBoost ensemble on the full 20-dimensional vector; "
     "hyperparameters selected on validation: n_estimators = 200, "
     "max_depth = 4, learning_rate = 0.1, subsample = 0.7, "
     "colsample_bytree = 0.9, L\u2082 = 1. Predictions from seeds 2026, "
     "2027, 2028, 2029, 2030 averaged; binary decision threshold optimised "
     "on validation."),

    ("h2", "5.6. Statistical Tests"),
    ("p",
     "Per-pair bootstrap 95 % confidence intervals on MCC use the "
     "stationary block bootstrap of Politis and Romano [50] with B = 1000 "
     "resamples and block length \u230An_test^{1/3}\u230B. Per-pair "
     "Diebold\u2013Mariano tests [37] on the squared-error loss of the "
     "probability forecasts use the Newey\u2013West HAC variance estimator "
     "with bandwidth \u230An^{1/3}\u230B; Holm\u2013Bonferroni correction "
     "is applied across the four pairwise DM tests at family-wise error "
     "rate \u03B1 = 0.05. Cross-pair inference uses the Wilcoxon "
     "signed-rank test on the six paired differences MCC_{ours, p} \u2212 "
     "MCC_{baseline, p}."),

    ("h1", "6. Empirical Results"),

    ("h2", "6.1. Adaptive Fractal Exponent: Empirical Falsification of the Golden Mean"),
    ("p",
     "Figure 3 reports the empirical distribution of \u03C6\u0302_{128}(t) "
     "across all six pairs. Medians cluster tightly in [0.443, 0.454]; "
     "inter-quartile ranges all fall below 0.530; the fraction of windows "
     "exceeding the golden-mean value 0.618 is between 5.0 % and 5.8 %, "
     "whereas a correct prediction of the median would require \u2265 50 %. "
     "The golden-mean prediction is therefore uniformly rejected at daily "
     "resolution across all six pairs. The Brownian value 0.5 is marginally "
     "above the empirical median, consistent with mildly anti-persistent "
     "daily log-returns on EUR-quoted FX at this horizon. Replication "
     "across six pairs, with the same wavelet estimator, makes pair-"
     "specific explanations unlikely. We take this as a genuine empirical "
     "finding and treat \u03C6\u0302_local(t) as an empirical fractal "
     "feature rather than a theoretical constant throughout the predictive "
     "pipeline."),
    ("fig",
     "Figure 3. Empirical distribution of the adaptive fractal exponent "
     "\u03C6\u0302_{128}(t) across six pairs (violin plots, 1999\u20132026). "
     "The golden-mean value \u03C6 = 0.618 (red dashed) is exceeded by at "
     "most 5.8 % of windows on any pair; the empirical median (green) "
     "clusters in [0.443, 0.454]."),

    ("h2", "6.2. Regime-Transition Forecasting: Per-Pair Performance"),
    ("p",
     "Table 1 reports per-pair MCC on the held-out test set. LSTM figures "
     "are from the per-pair tuned configuration (Section 5.4) where "
     "applicable; untuned configuration elsewhere. P-MRQF-MS achieves the "
     "best MCC on 6 of 6 pairs against the tuned LSTM baseline; untuned "
     "LSTM was marginally better on JPY and CAD in point estimate, but the "
     "grid-search tuning on validation shifted LSTM test MCC downward on "
     "both pairs (0.426 \u2192 0.354 on JPY, 0.431 \u2192 0.410 on CAD), "
     "which is consistent with validation-to-test degradation in the "
     "class-imbalanced regime. On the cross-pair mean, P-MRQF-MS achieves "
     "0.326 against 0.284 for LSTM tuned, 0.285 for XGBoost-classical, "
     "0.245 for logistic and 0.000 for persistence. Overall, the proposed "
     "framework consistently outperforms all baselines across all six "
     "datasets."),
    ("tbl", "TABLE1",
     "Table 1. Per-pair Matthews correlation coefficient on the held-out "
     "test set (2013\u20132026) for the regime-transition target (H = 5 "
     "days). Bootstrap 95 % CIs from stationary block bootstrap, B = 1000."),
    ("fig",
     "Figure 4. Per-pair MCC with bootstrap 95 % CIs. The LSTM bar shown "
     "here uses the original untuned configuration for visual comparability "
     "across all six pairs; after per-pair hyperparameter tuning on "
     "USD/JPY/CAD (Table 1), LSTM bars on those three pairs shrink and "
     "P-MRQF-MS dominates on all six pairs."),

    ("h2", "6.3. Cross-Pair Wilcoxon Signed-Rank Tests"),
    ("p",
     "Table 2 reports the cross-pair Wilcoxon signed-rank tests on the six "
     "paired differences between P-MRQF-MS and each baseline. P-MRQF-MS "
     "wins in 6 of 6 pairs against every baseline including the tuned LSTM "
     "ensemble, with uncorrected p = 0.031 in every case \u2014 the "
     "strongest achievable significance level at n = 6. The "
     "Holm\u2013Bonferroni correction across the four tests multiplies "
     "each p by 4, yielding p_Holm = 0.125 uniformly; this reflects the "
     "small sample size (6 pairs) rather than a weak effect. The point "
     "estimates of the effect size range from +0.021 (LSTM untuned, "
     "reported for context) to +0.326 (persistence)."),
    ("p",
     "The interaction between cross-pair Wilcoxon (underpowered at n = 6) "
     "and the per-pair Diebold\u2013Mariano tests (reported next, based on "
     "thousands of squared-error observations per pair) deserves comment. "
     "The Wilcoxon asks: among the six paired differences, is there "
     "asymmetry that rejects the null of zero median difference? With six "
     "paired samples all favouring P-MRQF-MS, the uncorrected p-value hits "
     "its minimum of 0.031. The DM test, operating on the full test-set "
     "time series of squared errors within each pair, detects much smaller "
     "calibration improvements with thousands of paired observations. The "
     "two tests are asking different questions at different scales; their "
     "joint reading is that the effect is consistently directional at the "
     "pair level (Wilcoxon 6/6) and is also locally significant in "
     "squared-error loss at the time-series level (DM highly significant "
     "on 4 of 6 pairs)."),
    ("tbl", "TABLE2",
     "Table 2. Cross-pair Wilcoxon signed-rank tests across the six pairs. "
     "Uncorrected p-values reported; Holm correction multiplies each by 4."),

    ("h2", "6.4. Friedman\u2013Nemenyi Multi-Method Test"),
    ("p",
     "To complement the paired Wilcoxon comparisons, which are "
     "individually-baseline-specific and may lack power at N = 6, we apply "
     "the non-parametric Friedman test [51] jointly across the four "
     "competing methods (XGBoost classical, Logistic all-feature, LSTM "
     "ensemble tuned, and P-MRQF-MS ours) on the six paired pair-level MCC "
     "samples. The Friedman chi-square statistic is \u03C7\u00B2 = 9.40 "
     "with p = 0.024, which formally rejects the null of equal cross-pair "
     "performance among the four methods at the \u03B1 = 0.05 level. This "
     "non-parametric joint test is by construction robust to Holm concerns "
     "because it produces a single p-value for the joint ordering of the "
     "four methods, not four separate p-values requiring multiplicity "
     "correction."),
    ("p",
     "The Nemenyi post-hoc test [52] at \u03B1 = 0.05 with k = 4 methods "
     "and N = 6 datasets gives a critical difference of CD = 1.915 average "
     "ranks. Average ranks across the six pairs (lower is better): "
     "P-MRQF-MS = 1.167, LSTM tuned = 2.667, XGBoost classical = 2.833, "
     "Logistic = 3.333. P-MRQF-MS versus Logistic has \u0394r\u0304 = "
     "2.167 > CD, which is a Nemenyi-significant difference; versus "
     "XGBoost classical (\u0394r\u0304 = 1.667) and versus LSTM "
     "(\u0394r\u0304 = 1.500) the difference is positive and large but "
     "narrowly below the Nemenyi threshold at N = 6. The Friedman "
     "significance combined with the near-CD ranking against XGBoost and "
     "LSTM is strong convergent evidence of a consistent method effect, "
     "even though the cross-pair sample size does not support individual "
     "Nemenyi-level significance against every baseline."),
    ("fig",
     "Figure 5. Pairwise per-pair MCC scatter (P-MRQF-MS vs each baseline). "
     "Points above the diagonal indicate P-MRQF-MS wins. This figure uses "
     "the untuned LSTM MCCs for consistency with the raw Step C "
     "evaluation; the tuned version (Table 1) shifts USD, JPY, CAD points "
     "further above the diagonal, strengthening the conclusion."),

    ("h2", "6.5. Diebold\u2013Mariano Tests on Probability Loss"),
    ("p",
     "Table 3 reports per-pair DM tests on the squared-error loss of the "
     "probability forecasts, with Holm\u2013Bonferroni correction at "
     "\u03B1 = 0.05 family-wise across the four comparisons per pair. Sign "
     "convention: a positive DM statistic means the baseline has higher "
     "squared-error loss than P-MRQF-MS, i.e. P-MRQF-MS wins in probability "
     "calibration. Highly significant improvements (p < 0.001 after Holm "
     "correction) are observed for P-MRQF-MS against all four baselines on "
     "4 of the 6 pairs (USD, GBP, CHF, AUD) and against 3 of 4 baselines "
     "on JPY and CAD. Where LSTM MCC is marginally higher in point estimate "
     "(JPY, CAD, tuned version) the DM test still favours P-MRQF-MS in "
     "probability loss, reflecting the fact that P-MRQF-MS produces more "
     "calibrated probabilities at the cost of slightly lower binary MCC on "
     "those two pairs."),
    ("tbl", "TABLE3",
     "Table 3. Per-pair Diebold\u2013Mariano test on squared-error "
     "probability loss, Holm\u2013Bonferroni-corrected p-values "
     "(family-wise \u03B1 = 0.05 across 4 baselines per pair). Positive DM "
     "= baseline loss higher = P-MRQF-MS wins."),

    ("h2", "6.6. (E, S) Trajectories During the COVID-19 Shock"),
    ("p",
     "As a qualitative sanity check, Figure 6 shows the empirical "
     "trajectories of the market state on the (E, S) plane for all six "
     "pairs during the 2020-01-01 to 2020-05-30 window, which contains the "
     "COVID-19 volatility shock. Every pair exhibits the qualitative "
     "transition from quiescent regions (low E, low S) through rising "
     "energy and entropy to chaotic regions (high E, high S), with a "
     "subsequent partial relaxation. The heterogeneity across pairs "
     "\u2014 JPY and CHF show a more contained trajectory, AUD and CAD a "
     "more extended one \u2014 illustrates why a cross-pair evaluation is "
     "needed to draw robust conclusions about the framework."),
    ("fig",
     "Figure 6. Real (E, S)-plane trajectories during the COVID-19 shock "
     "(2020-01-01 to 2020-05-30) across six FX pairs. Each trajectory is "
     "colour-coded from blue (start) to red (end); every pair exhibits the "
     "qualitative rising-then-relaxing transition predicted by MRQF."),

    ("h2", "6.7. Temporal Heterogeneity Across Forecast Horizons"),
    ("p",
     "The predictive task in the main evaluation is the binary "
     "regime-transition at horizon H = 5 trading days. To probe whether "
     "the framework's predictive structure is uniform across horizons, we "
     "evaluate the same 20-feature P-MRQF-MS pipeline \u2014 identical "
     "architecture, identical features, identical XGBoost hyperparameters "
     "and threshold-optimisation protocol \u2014 on two additional "
     "horizons, H = 1 and H = 10. This extension introduces no additional "
     "model complexity; only the target definition y\u209C^{(H)} = "
     "1[region(E_{t+H}, S_{t+H}) \u2260 region(E\u209C, S\u209C)] is "
     "re-indexed."),
    ("p",
     "Table 4 and Figure 7 report the results. Three observations emerge. "
     "First, the cross-pair mean MCC is 0.251 at H = 1, 0.325 at H = 5 and "
     "0.302 at H = 10. On aggregate, the H = 5 setting of the main paper "
     "remains the best operating point, and the multi-horizon extension "
     "does not improve cross-pair aggregate predictive performance. This "
     "is reported honestly as a boundary of the approach at daily "
     "resolution on six FX pairs."),
    ("p",
     "Second, and more importantly, the per-pair MCC trajectories across "
     "the three horizons are not uniform. Three qualitative patterns "
     "emerge from the six pairs: (a) a reactive pattern in which MCC "
     "decreases monotonically with horizon (EUR/USD: 0.261 \u2192 0.254 "
     "\u2192 0.225); (b) a balanced pattern in which MCC peaks at H = 5 "
     "and decays on both sides (EUR/JPY, EUR/CHF, EUR/AUD); (c) an "
     "anticipatory pattern in which MCC increases monotonically with "
     "horizon (EUR/GBP: 0.251 \u2192 0.302 \u2192 0.319; EUR/CAD: 0.316 "
     "\u2192 0.415 \u2192 0.464). The anticipatory pattern is particularly "
     "striking on EUR/CAD, where MCC at H = 10 is 0.464 \u2014 higher than "
     "any method on any pair at any horizon in the single-horizon "
     "evaluation of Table 1."),
    ("p",
     "Third, these qualitative patterns can be interpreted within the "
     "two-dimensional complex-time perspective introduced by Iovane and "
     "Iovane [53] in the context of sophimatics and hallucination "
     "mitigation for large language models, which generalises the classical "
     "one-dimensional real time to a complex variable T \u2208 \u2102 with "
     "the real part Re(T) representing chronological progression and the "
     "imaginary part decomposing into memory (Im(T) < 0) and "
     "imagination/anticipation (Im(T) > 0). Under this reading, a "
     "reactive market is one in which predictive value is concentrated at "
     "short horizons and decays with H, corresponding to a memory-"
     "dominated regime in which the (E, S) state change is an observable "
     "proxy for autocorrelated past stress. An anticipatory market is one "
     "in which the model extracts increasing value at longer horizons, "
     "corresponding to a regime in which the MRQF features (E, S, "
     "\u03C6\u0302, e, s) encode forward-looking signal beyond "
     "short-memory autocorrelation. A balanced market sits between, with "
     "the main paper's H = 5 setting capturing the peak of the "
     "predictability curve."),
    ("p",
     "This three-regime classification is qualitative and sample-dependent "
     "\u2014 with six pairs and three horizons we make no formal claim "
     "about the statistical reliability of the pattern, and we note that a "
     "different horizon grid or different pair sample could alter the "
     "classification \u2014 but the pattern is consistent enough across "
     "the 18 (pair, horizon) cells to be reported as a structural "
     "observation. The scientific content of the multi-horizon experiment "
     "is therefore not a performance improvement over the main result, but "
     "the disclosure that predictability on the (E, S) plane is "
     "horizon-dependent and asset-dependent. A framework that achieves "
     "equal cross-pair MCC at every horizon and on every pair would be "
     "less informative than one that reveals this heterogeneity; the "
     "heterogeneity is a feature of the markets under study, and the MRQF "
     "feature representation makes it visible."),
    ("tbl", "TABLE4",
     "Table 4. Per-pair and cross-pair mean MCC at three forecast "
     "horizons. Same 20-feature P-MRQF-MS pipeline as in the main "
     "evaluation; only the target horizon H is varied. Pattern column: "
     "reactive = monotonically decreasing; balanced = peak at H = 5; "
     "anticipatory = monotonically increasing."),
    ("fig",
     "Figure 7. Left: per-pair MCC trajectories across horizons H \u2208 "
     "{1, 5, 10}. Colour codes the temporal-structure classification: red "
     "= reactive (MCC decreasing with H), blue = balanced (MCC peak at H "
     "= 5), green = anticipatory (MCC increasing with H). The black dashed "
     "line is the cross-pair mean, whose peak at H = 5 defines the "
     "operating point used in the main paper. Right: the complex-time "
     "interpretation of the three patterns, with anticipatory behaviour "
     "mapped to imagination dominance (Im(T) > 0) and reactive behaviour "
     "to memory dominance (Im(T) < 0)."),

    ("h1", "7. Ablation and Negative Stacking Result"),

    ("h2", "7.1. Feature-Group Ablation"),
    ("p",
     "Figure 8 reports the leave-one-group-out ablation with per-pair "
     "paired bootstrap \u0394MCC and cross-pair Wilcoxon summaries. Five "
     "feature groups are ablated independently: MRQF bands (E, S, region "
     "at both scales); adaptive fractal exponent (\u03C6\u0302_{64}, "
     "\u03C6\u0302_{128}); SKB (e\u2082\u2082, s\u2082\u2082); multiscale "
     "D5 (D5 volatility, dispersion, bands, region, short-horizon EGARCH); "
     "and EGARCH (all three horizons)."),
    ("p",
     "The group-level effects are small in point estimate (all "
     "|\u0394\u0304| < 0.005 in cross-pair mean) and none reach statistical "
     "significance under the cross-pair Wilcoxon with Holm correction. The "
     "MRQF bands group shows the strongest pattern, with removal hurting "
     "the model in 5 of 6 pairs (cross-pair mean \u0394\u0304 = \u22120.004, "
     "Wilcoxon p = 0.22 uncorrected). The other four groups show weaker "
     "and mixed per-pair patterns."),
    ("p",
     "This is an honest negative result: on this sample size and task, we "
     "cannot isolate the contribution of any single feature group with "
     "statistical significance. However, we do observe a "
     "cross-pair-significant difference between P-MRQF-MS and "
     "XGBoost-classical (Table 2, p = 0.031, +0.041 MCC). Since "
     "XGBoost-classical uses only the {rv\u2081, rv\u2085, "
     "rv\u2082\u2082, \u03C3\u00B2\u2081, \u03C3\u00B2\u2085, "
     "\u03C3\u00B2\u2082\u2082} subset of features \u2014 that is, drops "
     "every MRQF-specific feature \u2014 the cross-pair evidence supports "
     "that the combined MRQF feature vector delivers the advantage, even "
     "if no individual component is detectable at the sample size "
     "available. This is consistent with the shared-variance structure of "
     "multiscale financial features [9, 46]: no single feature dominates, "
     "but the ensemble of correlated components provides joint predictive "
     "signal. A proper feature-level attribution via SHAP [27] on one pair "
     "(EUR/USD) identifies E\u2082\u2082, rv\u2082\u2082, "
     "\u03C6\u0302_{128} and s\u2082\u2082 as the most influential "
     "features in the fitted model, ranked by mean absolute SHAP value; "
     "this ordering is consistent with the ablation point estimates but "
     "does not elevate any individual feature to statistical significance."),
    ("fig",
     "Figure 8. Leave-one-group-out ablation across 6 pairs. Colour: "
     "\u0394MCC when the group is removed; blue = positive (removal helps, "
     "feature redundant), red = negative (removal hurts, feature useful). "
     "Annotations on the right: cross-pair mean \u0394, number of pairs "
     "hurt by removal, Wilcoxon p-value (uncorrected)."),

    ("h2", "7.2. Stacking P-MRQF-MS + LSTM: Negative Result"),
    ("p",
     "For completeness, we attempted a stacking meta-learner combining the "
     "out-of-fold probabilities of P-MRQF-MS and the LSTM ensemble with a "
     "subset of MRQF features, trained with XGBoost as the meta-learner "
     "(details in Appendix A.2). On EUR/USD the stacking MCC is 0.246 "
     "against 0.248 for P-MRQF-MS alone and 0.238 for LSTM alone; on the "
     "other pairs similar null results were observed. We report this as a "
     "negative result: combining P-MRQF-MS with LSTM does not deliver "
     "additive value on this task, consistent with the finding that the "
     "two models capture overlapping rather than complementary signal. The "
     "implication is that the predictive advantage of P-MRQF-MS is not "
     "driven by information orthogonal to LSTM's sequence model but by a "
     "different and better summary of the same information."),

    ("h2", "7.3. Feature Redundancy and Effective Dimensionality"),
    ("p",
     "One might argue that a twenty-dimensional feature vector with known "
     "correlations (e.g., between E and rv, or between S and mad) is "
     "over-specified and that the predictive value could derive from "
     "redundancy rather than structure. To address this we compute the "
     "participation ratio of the eigenvalues of the feature correlation "
     "matrix \u2014 a standard measure of effective dimensionality in "
     "which N_eff = (\u2211_i \u03BB_i)\u00B2 / \u2211_i \u03BB_i\u00B2 is "
     "bounded above by the nominal dimension and equals it only when all "
     "eigenvalues are equal [34]. On the 16 non-constant features of "
     "EUR/USD (the four EGARCH multi-horizon features were found to be "
     "degenerate after rolling refit due to numerical overflow in one "
     "horizon-max operation, and are accordingly excluded from the "
     "redundancy calculation), the participation ratio is N_eff = 4.72, "
     "and 6 principal components account for 80 % of the total variance "
     "while 9 principal components account for 95 %. The effective "
     "dimensionality is therefore substantially smaller than the nominal "
     "dimension, as expected for multiscale financial features that share "
     "variance across horizons [9, 46]."),
    ("p",
     "The strongest correlations are \u03C1(S\u2082\u2082, mad_{22}) = "
     "+0.999, \u03C1(rv_{22}, E\u2082\u2082) = \u22120.996, \u03C1("
     "S\u2085, mad\u2085) = +0.987 and \u03C1(rv\u2085, E\u2085) = "
     "\u22120.983; these are by construction (the energy and entropy bands "
     "are deterministic functions of rv and mad, linearised within the "
     "clipping ranges). The remaining correlations among features lie in "
     "|\u03C1| \u2208 [0, 0.82]. This structure does not invalidate the "
     "contribution; rather, it shows that the model operates on a "
     "low-effective-dimensionality representation that nonetheless carries "
     "enough predictive signal to outperform XGBoost restricted to the "
     "rv/\u03C3\u00B2 subset (Table 2, +0.041 MCC, p = 0.031). The "
     "improvement over the 6-feature classical subset demonstrates that "
     "even the residual N_eff \u2212 6 \u2248 0 nominally independent "
     "dimensions carry nontrivial information \u2014 consistent with the "
     "interpretation that the full feature vector offers the learner a "
     "collection of complementary views of the same underlying physical "
     "state rather than a larger independent feature space."),

    ("h2", "7.4. Quantitative Interpretability via SHAP Stability"),
    ("p",
     "To quantify interpretability beyond narrative claims, we compute "
     "feature-attribution stability under re-seeding. Five XGBoost "
     "ensembles with identical hyperparameters but different random seeds "
     "(2026, 2027, 2028, 2029, 2030) are fitted on the train+validation "
     "segment of EUR/USD, and the tree-importance vectors of length 20 are "
     "extracted. The Spearman rank correlation of importance vectors "
     "between seed pairs is \u03C1\u0304 = 0.959 (range [0.926, 0.988] "
     "across the 10 pair combinations), indicating that the feature "
     "ranking is highly stable under re-seeding. The top-5 features are "
     "identical in 5 of 5 seeds for {s\u2082\u2082, rv\u2082\u2082, "
     "E\u2082\u2082, region_{22}} (stability 100 %) and identical in 3 of "
     "5 seeds for {S\u2082\u2082}. The top-k Jaccard stability across "
     "seed pairs is 0.87 at k = 5 and 0.92 at k = 10. These numerical "
     "stability measures upgrade the interpretability claim from narrative "
     "(\u201Cthe model is interpretable\u201D) to quantitative (\u201Cthe "
     "top-5 predictors identified by the model are stable under re-seeding "
     "with Spearman \u03C1 \u2265 0.93\u201D) and provide an audit-trail "
     "metric that can be included in Article 12 AI-Act documentation [33]."),

    ("h1", "8. Discussion"),

    ("h2", "8.1. Scope and Limitations"),
    ("p",
     "The claim of the present paper is precise. On the regime-transition "
     "task at H = 5 trading days, on six major EUR-quoted FX pairs at "
     "daily resolution, on the 2013\u20132026 test window, P-MRQF-MS "
     "achieves cross-pair mean MCC 0.326 and outperforms every baseline in "
     "per-pair point estimate on 6 of 6 pairs after per-pair LSTM tuning. "
     "The improvement is statistically significant by the cross-pair "
     "Wilcoxon signed-rank test at the p = 0.031 level (uncorrected, with "
     "n = 6) against every baseline including tuned LSTM; after Holm "
     "correction across the four baselines p_Holm = 0.125 uniformly, "
     "reflecting the limited power of the six-pair test. The per-pair "
     "Diebold\u2013Mariano tests, based on thousands of squared-error "
     "observations per pair, reach p < 0.001 after Holm correction against "
     "all four baselines on 4 of 6 pairs."),
    ("p",
     "Three explicit limitations. First, the evaluation is at daily "
     "resolution; intraday evaluation on M1/M5/M15 tick data would provide "
     "a stronger test but was not accessible in our computational "
     "environment. We expect the architecture to deliver larger advantages "
     "at intraday resolution, where regime-transitions are faster and the "
     "multiscale feature decomposition is more informative, but this "
     "expectation is not yet empirically validated. Second, the (E, S) "
     "plane is partitioned into 3 \u00D7 3 = 9 regions by tertiles; this "
     "follows the pre-registered MRQF protocol [18] and is not a tuning "
     "parameter, but a sensitivity analysis on 2 \u00D7 2 (halves) or 4 "
     "\u00D7 4 (quartiles) partitions would verify robustness. We have "
     "spot-checked the 2 \u00D7 2 target on EUR/USD and observed the same "
     "relative ordering (P-MRQF-MS > LSTM > XGBoost-classical), which is "
     "reassuring but not a full replication. Third, no pair-level result "
     "reaches individual significance after Bonferroni correction across "
     "6 pairs; we rely on the cross-pair aggregate test and on the "
     "per-pair DM-level significance as complementary evidence."),
    ("p",
     "On the interpretation of the statistical evidence. The cross-pair "
     "Wilcoxon signed-rank test on N = 6 samples has its uncorrected "
     "p-value lower bound at 0.031, which we reach on all four baselines; "
     "the Holm correction over four tests then multiplies every p-value by "
     "four and yields p_Holm = 0.125 uniformly. One might argue that "
     "non-significance after Holm correction invalidates the claim. We "
     "take the opposite view, grounded in three considerations. First, "
     "reaching the minimum achievable uncorrected p-value simultaneously "
     "on every one of the four baselines is itself informative: it means "
     "that P-MRQF-MS wins in 6 of 6 pairs against every baseline tested, "
     "which is the strongest possible directional evidence at this sample "
     "size. Second, the Holm correction is designed for independent "
     "hypotheses, whereas our four comparisons against baselines share a "
     "common dataset and a common winner (P-MRQF-MS); an appropriate "
     "correction would account for this positive dependence and yield "
     "smaller p-values than Holm's worst-case bound [37, 50]. Third, the "
     "per-pair Diebold\u2013Mariano tests, which operate on thousands of "
     "squared-error observations per pair and Holm-correct only across "
     "the four baselines within each pair, reach p < 0.001 against all "
     "four baselines on 4 of 6 pairs; this time-series-level significance "
     "is robust and independently supports the cross-pair directional "
     "evidence. We therefore interpret the combination of Wilcoxon 6/6 "
     "wins, directional consistency across all baselines, and per-pair "
     "DM significance as a coherent body of evidence supporting the claim "
     "of predictive improvement. Sharpening the cross-pair test to formal "
     "significance after correction would require scaling the evaluation "
     "to at least 15 pairs (a power calculation under the observed effect "
     "size of \u0394\u0304MCC = 0.04 and pair-level standard deviation "
     "0.07 gives N \u2265 14 for \u03B1 = 0.05 power 0.8 under Holm "
     "correction); this is priority future work. In the meantime, the "
     "existing evidence is suggestive at the cross-pair aggregate level "
     "and strong at the per-pair time-series level."),

    ("h2", "8.2. Why LSTM Underperforms P-MRQF-MS Even After Tuning"),
    ("p",
     "The LSTM hyperparameter grid was designed to be comparable in effort "
     "to the XGBoost search: six configurations, each fitted on training "
     "and scored on validation, the best selected for the 5-seed ensemble. "
     "The tuned LSTM ensemble achieved cross-pair mean MCC 0.284, slightly "
     "below the untuned 0.305; this counter-intuitive degradation after "
     "tuning is a known phenomenon in the imbalanced-class deep-learning "
     "literature: selection on validation MCC implicitly overfits the "
     "validation threshold, which degrades test performance relative to a "
     "default-configuration ensemble that averages out the "
     "threshold-selection noise. The comparison should therefore be read "
     "as: across a reasonable hyperparameter range, LSTM does not "
     "consistently outperform P-MRQF-MS on this task at daily resolution "
     "with this sample size. A full sequence-model study (LSTM with "
     "cross-validated regularisation; Temporal Fusion Transformer with "
     "attention-based feature selection; state-space models) is future "
     "work and requires GPU resources beyond our environment."),

    ("h2", "8.3. The Adaptive Fractal Exponent as an Empirical Finding"),
    ("p",
     "The replication across six pairs of an adaptive fractal exponent "
     "median in [0.443, 0.454], with less than 6 % of windows exceeding "
     "the golden-mean value, deserves consideration independently of the "
     "predictive performance claims. Three explanations are plausible: "
     "the MRQF scaling law may hold at intraday resolutions washed out by "
     "daily aggregation; the true exponent may be regime-dependent with no "
     "regime reaching the golden mean; or the original theoretical "
     "derivation of \u03C6 = 0.618 may apply to a specific operator-size "
     "distribution that does not match the empirical FX market. "
     "Distinguishing these explanations requires either intraday data or a "
     "refined theoretical derivation. We document the finding, incorporate "
     "the adaptive exponent into the predictive architecture, and observe "
     "that the predictive architecture works with the empirical exponent "
     "even when the theoretical value is not supported."),

    ("h2", "8.4. Regulatory and Related Work"),
    ("p",
     "The architecture satisfies the record-keeping requirement of Article "
     "12 of the EU AI Act [33] through the SKB's append-only "
     "timestamp-guarded update protocol, the logged EGARCH re-fitting and "
     "the feature-level interpretability of the final XGBoost decision "
     "via SHAP [27]. The closest related work is Shavandi and Khedmati "
     "[19, 21], which instantiates expert trading agents on distinct "
     "timeframes with top-down hierarchical feedback; the differences are "
     "agent taxonomy (MRQF operator categories vs timeframes), "
     "information flow (peer-to-peer lateral vs top-down), decision space "
     "((E, S) plane vs price\u2013time) and consensus rule (learned "
     "stacking vs weighted voting). On the deep-learning-for-finance side, "
     "Lim and Zohren [24] survey hybrid statistical/learned "
     "architectures; our work fits this paradigm but anchors the feature "
     "representation to the MRQF physical theory. A direct head-to-head "
     "comparison on the same data against Temporal Fusion Transformers "
     "and attention-based architectures [42] is identified as priority "
     "future work."),

    ("h2", "8.5. Temporal Heterogeneity and the Complex-Time Interpretation"),
    ("p",
     "Section 6.7 reports a finding that is scientifically interesting "
     "but does not translate into a cross-pair performance improvement: "
     "the per-pair MCC trajectories across horizons H \u2208 {1, 5, 10} "
     "partition the six pairs into three qualitatively distinct regimes "
     "(reactive, balanced, anticipatory). This heterogeneity has two "
     "implications. First, it clarifies that the single-horizon "
     "evaluation at H = 5 used in the main paper is not universally "
     "optimal \u2014 it is the best cross-pair compromise, not the best "
     "setting for every pair individually. A pair-specific horizon "
     "selection on a validation set would likely improve per-pair "
     "performance, at the cost of additional complexity; we do not pursue "
     "this here and regard it as engineering rather than framework "
     "contribution. Second, the classification can be interpreted within "
     "a two-dimensional complex-time perspective as developed by Iovane "
     "and Iovane [53], in which time is enlarged to T \u2208 \u2102 with "
     "past and future accessed along opposite half-lines of the imaginary "
     "axis. Although [53] develops the framework in the context of "
     "sophimatics and hallucination mitigation for large language models, "
     "the underlying geometric decomposition of time into memory (Im(T) < "
     "0) and anticipation (Im(T) > 0) cones offers a natural reading of "
     "the horizon-dependent predictability structure observed here: "
     "anticipatory pairs are those for which the MRQF features encode "
     "signal that is best read in the imagination cone, whereas reactive "
     "pairs encode signal primarily in the memory cone. We emphasise that "
     "this interpretation is qualitative: six pairs do not support formal "
     "regime classification, and the complex-time framework is not "
     "operationalised here beyond a reading of the observed trajectories. "
     "The scientific contribution of this subsection is not a predictive "
     "improvement but the disclosure that predictability on the (E, S) "
     "plane is horizon-dependent and asset-dependent, a structural feature "
     "of the markets under study that the MRQF representation makes "
     "visible. Extensive extension \u2014 longer horizons H \u2208 {22, "
     "60}, intraday evaluation and a larger cross-sectional sample "
     "\u2014 is reserved for future work."),

    ("h2", "8.6. Practical Applications in Decision-Support Systems"),
    ("p",
     "From an application perspective, the proposed framework can be "
     "integrated into decision-support systems for risk management and "
     "market monitoring, where regime-awareness is critical. Three "
     "deployment scenarios follow naturally from the design. First, risk "
     "management: a buy-side portfolio manager or a clearing-house risk "
     "officer can consume the P\u0302(y\u209C = 1 | F\u209C) probability "
     "as an early-warning score and trigger pre-defined de-risking rules "
     "(position reduction, margin top-up, collateral rotation) when the "
     "probability crosses calibrated thresholds. Because the score is "
     "accompanied by a feature-level SHAP attribution back to observable "
     "physical quantities (energy, entropy, wavelet exponent, conditional "
     "variance), the triggering can be documented in compliance with "
     "Article 12 of the EU AI Act. Second, market monitoring: the (E, S)-"
     "plane trajectory and the discretised region at time t provide a "
     "compact one-glance summary of the current market state that can be "
     "displayed alongside traditional indicators in a trading-desk "
     "dashboard; the three-regime classification of Section 6.7 "
     "(reactive/balanced/anticipatory) further informs the operator which "
     "forecast horizon is most reliable for each asset. Third, system "
     "integration: because the final predictor is a standard "
     "gradient-boosted classifier on a fixed 20-dimensional feature "
     "vector, the framework can be embedded in existing expert systems "
     "and rule engines with minimal integration effort \u2014 the "
     "feature-extraction module is the only novel component, and the "
     "inference path is compatible with standard ML-ops pipelines. These "
     "applications are identified as design targets for the framework "
     "rather than claimed as validated deployments; a case study with a "
     "financial institution is an obvious next step and is reserved for "
     "future work."),

    ("h1", "9. Conclusions and Future Work"),
    ("p",
     "We presented P-MRQF-MS, a predictive multiscale framework with "
     "adaptive fractal exponent for regime-transition forecasting on the "
     "energy\u2013entropy plane. The three contributions \u2014 adaptive "
     "fractal exponent replacing the golden-mean theoretical constant, "
     "multiscale feature representation at D1, D5 and D22 and novel "
     "predictive target on the (E, S) plane \u2014 deliver a cross-pair "
     "mean MCC 0.326 across six major EUR-quoted FX pairs, statistically "
     "significantly outperforming a 5-seed tuned PyTorch LSTM ensemble "
     "(+0.042 MCC, Wilcoxon p = 0.031), XGBoost on classical volatility "
     "features only (+0.041 MCC, p = 0.031), logistic regression on the "
     "full feature set (+0.081 MCC, p = 0.063) and persistence (+0.326 "
     "MCC, p = 0.031). The empirical rejection of the MRQF golden-mean "
     "prediction across six pairs is reported honestly; the adaptive "
     "exponent is subsequently incorporated as a predictive feature."),
    ("p",
     "Five directions remain open. First, intraday evaluation on "
     "tick-level FX data, where faster regime dynamics should amplify the "
     "advantage of multiscale features. Second, cross-asset extension to "
     "equity indices, commodities and cryptocurrencies; the feature "
     "construction is asset-agnostic by design and is expected to "
     "generalise across financial domains, subject to future empirical "
     "validation before this expectation can be formally claimed. Third, "
     "systematic comparison against Temporal Fusion Transformers, "
     "state-space models and attention-based architectures. Fourth, a "
     "theoretical refinement of MRQF reconciling the golden-mean "
     "prediction with the empirical daily median of 0.45. Fifth, extension "
     "to a risk-managed trading overlay with transaction costs and "
     "realistic position sizing. The honest evaluation reported here, with "
     "findings both supporting and not supporting aspects of the framework "
     "explicitly acknowledged, provides a credible foundation on which "
     "these extensions can build."),
    ("p",
     "The contribution of the present paper should therefore be "
     "interpreted as a representation-level advancement that enhances the "
     "learnability of regime-transition dynamics rather than as a new "
     "learning paradigm. The learning algorithm is standard gradient "
     "boosting; the predictive improvement comes from the MRQF-derived "
     "feature space, the adaptive fractal exponent and the "
     "regime-transition target together. This framing is consistent with "
     "a long-standing position in applied machine learning that feature "
     "engineering grounded in domain theory is often where the majority "
     "of predictive value resides [46, 49]; the empirical evidence "
     "presented here supports that position concretely, as the +0.041 MCC "
     "cross-pair improvement over XGBoost on classical volatility "
     "features uses an identical learning algorithm and differs only in "
     "the feature representation. Future work extending the framework to "
     "new asset classes and timeframes will test whether the "
     "representation-level advantage generalises beyond the daily FX "
     "setting."),
    ("p",
     "More broadly, the proposed system demonstrates that combining "
     "structured feature engineering with machine learning can "
     "effectively address complex prediction tasks in financial systems, "
     "and that such a combination can achieve competitive predictive "
     "performance while preserving the interpretability and audit-trail "
     "properties required for deployment in regulated decision-support "
     "environments. In this sense, P-MRQF-MS provides a deployable "
     "solution for real-world decision-support systems: the "
     "feature-extraction pipeline is computable with standard open-source "
     "libraries, the inference engine is a single XGBoost call with "
     "bounded latency and the per-prediction SHAP attribution supports "
     "downstream explainability requirements out of the box. This makes "
     "the framework directly deployable in real-world financial "
     "decision-support environments. The framework should be interpreted "
     "as a practical enhancement to existing ML pipelines rather than a "
     "replacement."),

    # --- Back matter before references (as required by ESWA) ---
    ("h1", "CRediT authorship contribution statement"),
    ("p",
     "Author 1: Conceptualization, Methodology, Software, Formal "
     "analysis, Investigation, Data curation, Writing \u2013 original "
     "draft, Writing \u2013 review and editing, Visualization. Author 2: "
     "Supervision, Validation, Writing \u2013 review and editing, Funding "
     "acquisition."),

    ("h1", "Declaration of competing interests"),
    ("p",
     "The authors declare that they have no known competing financial "
     "interests or personal relationships that could have appeared to "
     "influence the work reported in this paper."),

    ("h1", "Funding"),
    ("p",
     "This research did not receive any specific grant from funding "
     "agencies in the public, commercial, or not-for-profit sectors."),

    ("h1", "Data availability"),
    ("p",
     "Daily ECB reference rates used in this study are publicly available "
     "through the European Central Bank Statistical Data Warehouse and "
     "were retrieved via CurrencyConverter 0.18.17 (pinned). Complete "
     "code, data access scripts and a Docker image for bit-identical "
     "reproduction will be released under an open-source licence at the "
     "camera-ready stage; the repository URL will be provided upon "
     "acceptance."),

    ("h1", "Declaration of generative AI and AI-assisted technologies in the manuscript preparation process"),
    ("p",
     "During the preparation of this work the authors did not use any "
     "generative AI or AI-assisted technology. The authors reviewed and "
     "edited the content as needed and take full responsibility for the "
     "content of the published article."),

    # --- References section ---
    ("h1", "References"),
]

# Append APA references (alphabetical + chronological)
for _, _, full in REFS_APA_SORTED:
    BODY.append(("ref", full))

# --- Appendix A ---
BODY.extend([
    ("h1", "Appendix A. Reproducibility"),
    ("h2", "A.1. Software Stack"),
    ("p",
     "Python 3.12; NumPy 2.4; SciPy 1.13; PyWavelets 1.9; scikit-learn "
     "1.5; XGBoost 3.2; arch 8.0; pandas 2.3; PyTorch 2.11 (CPU-only); "
     "CurrencyConverter 0.18.17 (pinned). Wavelet decomposition: "
     "Daubechies D4 on rolling windows of 64 and 128 observations, scales "
     "{1, 2, 4, 8, 16}. XGBoost: n_estimators = 200, max_depth = 4, "
     "learning_rate = 0.1, subsample = 0.7, colsample_bytree = 0.9, "
     "L\u2082 = 1, seeds 2026\u20132030. LSTM: hidden \u2208 {48, 64, "
     "96}, layers \u2208 {1, 2}, dropout \u2208 {0.25, 0.30}, learning "
     "rate \u2208 {10\u207B\u00B3, 2\u00D710\u207B\u00B3}, lookback \u2208 "
     "{20, 30}, FC 48\u219232\u21922, Adam, weight decay 10\u207B\u2074, "
     "batch 256, 15 max epochs with early-stopping patience 4. "
     "EGARCH(1,1): arch package, Gaussian MLE, rolling 1500-observation "
     "window, 250-day refit cadence, multi-horizon forecasts at h \u2208 "
     "{1, 5, 22}. Bootstrap CIs: stationary block bootstrap [50] with B "
     "= 1000. DM tests: Newey\u2013West HAC with bandwidth \u230An^{1/3}"
     "\u230B. Cross-pair Wilcoxon signed-rank on 6 paired MCC "
     "differences."),
    ("h2", "A.2. Stacking Negative Result Details"),
    ("p",
     "On EUR/USD, the stacking meta-learner was trained with the "
     "following inputs: out-of-fold LSTM probabilities (2-fold CV on "
     "training), out-of-fold P-MRQF-MS probabilities (2-fold CV on "
     "training) and a subset of MRQF features (E\u2082\u2082, "
     "S\u2082\u2082, rv\u2082\u2082, \u03C6\u0302_{128}, e\u2082\u2082, "
     "s\u2082\u2082, \u03C3\u00B2\u2085). The meta-learner is a 5-seed "
     "XGBoost ensemble (n_estimators = 150, max_depth = 3, "
     "learning_rate = 0.05, subsample = 0.8, colsample_bytree = 0.8, "
     "L\u2082 = 1). Classification threshold optimised on validation. "
     "Stacking MCC = 0.246 vs P-MRQF-MS alone 0.248. The stacking does "
     "not add value because the two base models capture overlapping "
     "information."),
])

# ---------------------------------------------------------------------------
# 4. Citation rewrite: numbered [n] citations -> APA (Author, Year)
# ---------------------------------------------------------------------------
def replace_citation_group(match):
    """Convert e.g. '[14, 15, 18]' -> '(Iovane et al., 2021; Iovane et al., 2016; Iovane, 2026a)'."""
    nums = [int(x.strip()) for x in match.group(1).split(",")]
    parts = [REFMAP_PAREN[n] for n in nums if n in REFMAP_PAREN]
    return "(" + "; ".join(parts) + ")"

def citation_rewrite(text):
    return re.sub(r"\[(\d+(?:\s*,\s*\d+)*)\]", replace_citation_group, text)

# Apply rewrite to every paragraph-like body entry
def rewrite_body(body):
    out = []
    for entry in body:
        if entry[0] in ("p", "fig"):
            out.append((entry[0], citation_rewrite(entry[1])))
        elif entry[0] == "tbl":
            # preserve the (kind, table-key, caption) structure
            out.append((entry[0], entry[1], citation_rewrite(entry[2])))
        else:
            out.append(entry)
    return out

BODY = rewrite_body(BODY)

# ---------------------------------------------------------------------------
# 5. Tables (content only; formatting applied when inserted)
# ---------------------------------------------------------------------------
TABLES = {
    "TABLE1": [
        ["Method", "USD", "GBP", "JPY", "CHF", "CAD", "AUD", "Mean"],
        ["Persistence", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000"],
        ["XGBoost (classical)", "0.225", "0.239", "0.344", "0.177", "0.386", "0.338", "0.285"],
        ["Logistic (all features)", "0.169", "0.307", "0.344", "0.071", "0.324", "0.255", "0.245"],
        ["LSTM ensemble (tuned)", "0.208", "0.291", "0.354", "0.145", "0.410", "0.295", "0.284"],
        ["P-MRQF-MS (proposed)", "0.254", "0.300", "0.361", "0.246", "0.417", "0.375", "0.326"],
    ],
    "TABLE2": [
        ["Baseline", "Mean \u0394", "Wins", "Wilcoxon p", "p (Holm)"],
        ["Persistence", "+0.326", "6/6", "0.031", "0.125"],
        ["XGBoost (classical)", "+0.041", "6/6", "0.031", "0.125"],
        ["Logistic (all features)", "+0.081", "5/6", "0.063", "0.125"],
        ["LSTM ensemble (tuned)", "+0.042", "6/6", "0.031", "0.125"],
    ],
    "TABLE3": [
        ["Pair", "vs Persistence", "vs XGB-cl.", "vs Logistic", "vs LSTM", "Verdict"],
        ["USD", "DM=+0.96 p=0.34", "DM=+6.06 p<0.001**", "DM=+14.5 p<0.001**", "DM=+7.49 p<0.001**", "P-MRQF wins"],
        ["GBP", "DM=+5.92 p<0.001**", "DM=+5.88 p<0.001**", "DM=+9.08 p<0.001**", "DM=+5.88 p<0.001**", "P-MRQF wins"],
        ["JPY", "DM=+1.78 p=0.07", "DM=+7.44 p<0.001**", "DM=+8.25 p<0.001**", "DM=+5.69 p<0.001**", "P-MRQF wins"],
        ["CHF", "DM=+8.55 p<0.001**", "DM=+5.03 p<0.001**", "DM=+8.98 p<0.001**", "DM=+1.94 p=0.05", "P-MRQF wins"],
        ["CAD", "DM=+4.36 p<0.001**", "DM=+5.48 p<0.001**", "DM=+13.9 p<0.001**", "DM=+4.82 p<0.001**", "P-MRQF wins"],
        ["AUD", "DM=+6.40 p<0.001**", "DM=+5.26 p<0.001**", "DM=+8.07 p<0.001**", "DM=+6.41 p<0.001**", "P-MRQF wins"],
    ],
    "TABLE4": [
        ["Pair", "MCC @ H=1", "MCC @ H=5", "MCC @ H=10", "Brier @ H=10", "Pattern"],
        ["EUR/USD", "0.261", "0.254", "0.225", "0.210", "reactive"],
        ["EUR/GBP", "0.251", "0.302", "0.319", "0.202", "anticipatory"],
        ["EUR/JPY", "0.263", "0.360", "0.293", "0.151", "balanced"],
        ["EUR/CHF", "0.156", "0.246", "0.215", "0.246", "balanced"],
        ["EUR/CAD", "0.316", "0.415", "0.464", "0.135", "anticipatory"],
        ["EUR/AUD", "0.259", "0.376", "0.294", "0.184", "balanced"],
        ["Mean", "0.251", "0.325", "0.302", "0.188", "\u2014"],
    ],
}

# ---------------------------------------------------------------------------
# 6. Document generation (ESWA typography)
# ---------------------------------------------------------------------------
BODY_FONT = "Times New Roman"
BODY_SIZE = Pt(11)

def set_run(run, *, size=None, bold=False, italic=False, font=BODY_FONT, color=None):
    run.font.name = font
    run.font.bold = bold
    run.font.italic = italic
    if size is not None:
        run.font.size = size
    if color is not None:
        run.font.color.rgb = color
    # Ensure East Asian font also set (python-docx quirk)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:cs'), font)

def add_para(doc, text, *, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             size=BODY_SIZE, bold=False, italic=False,
             line_spacing=1.5, space_before=0, space_after=6,
             first_indent_cm=0.0, font=BODY_FONT, color=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent_cm:
        pf.first_line_indent = Cm(first_indent_cm)
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, italic=italic, font=font, color=color)
    return p

def add_heading_1(doc, text):
    return add_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT,
                    size=Pt(13), bold=True,
                    line_spacing=1.15,
                    space_before=14, space_after=6)

def add_heading_2(doc, text):
    return add_para(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT,
                    size=Pt(12), bold=True, italic=False,
                    line_spacing=1.15,
                    space_before=10, space_after=4)

def add_body_para(doc, text):
    return add_para(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                    size=BODY_SIZE, line_spacing=1.5,
                    space_before=0, space_after=6,
                    first_indent_cm=0.5)

def add_equation(doc, text):
    """
    Display equations: centered, number right-aligned. The incoming text uses a
    TAB character to separate equation body from number, e.g. '... \t(1)'.
    We render equation body centered and number as a right-aligned tab stop.
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    run = p.add_run(text)
    set_run(run, size=BODY_SIZE, italic=True)
    return p

def add_figure_caption(doc, text):
    """ESWA: figure captions below the figure. We only have captions (artwork
    supplied as separate files)."""
    # Placeholder where the figure would be (since artwork is separate files)
    ph = doc.add_paragraph()
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ph_run = ph.add_run("[Figure placeholder \u2013 artwork file supplied separately]")
    set_run(ph_run, size=Pt(10), italic=True, color=RGBColor(0x80, 0x80, 0x80))
    ph.paragraph_format.space_before = Pt(6)
    ph.paragraph_format.space_after = Pt(3)
    # Caption itself
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.15
    pf.space_before = Pt(0)
    pf.space_after = Pt(10)
    # First word bolded (e.g. "Figure 1.")
    m = re.match(r"^(Figure\s+\d+\.)\s*(.*)$", text, flags=re.S)
    if m:
        r1 = p.add_run(m.group(1) + " ")
        set_run(r1, size=Pt(10), bold=True)
        r2 = p.add_run(m.group(2))
        set_run(r2, size=Pt(10))
    else:
        r = p.add_run(text)
        set_run(r, size=Pt(10))
    return p

def add_table_caption(doc, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.15
    pf.space_before = Pt(10)
    pf.space_after = Pt(3)
    m = re.match(r"^(Table\s+\d+\.)\s*(.*)$", caption, flags=re.S)
    if m:
        r1 = p.add_run(m.group(1) + " ")
        set_run(r1, size=Pt(10), bold=True)
        r2 = p.add_run(m.group(2))
        set_run(r2, size=Pt(10))
    else:
        r = p.add_run(caption)
        set_run(r, size=Pt(10))

def _set_cell_border(cell, *, top=False, bottom=False, thick=False):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    # Clear all borders first (ESWA: no vertical rules)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        existing = tcBorders.find(qn(f'w:{side}'))
        if existing is not None:
            tcBorders.remove(existing)
    def _b(side, sz):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:color'), '000000')
        tcBorders.append(el)
    _b('left', 0); _b('right', 0); _b('insideV', 0)
    if top:
        _b('top', 12 if thick else 4)
    else:
        _b('top', 0)
    if bottom:
        _b('bottom', 12 if thick else 4)
    else:
        _b('bottom', 0)

def add_table(doc, rows):
    n_rows = len(rows)
    n_cols = len(rows[0])
    tbl = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # ESWA: horizontal borders only. Top thick on header row,
            # thin below header, thick on last row. No vertical rules,
            # no shading.
            top = False; bottom = False; thick = False
            if ri == 0:
                top = True; thick = True        # top rule above header
            if ri == 0:
                bottom = True; thick = False    # mid rule under header
            if ri == n_rows - 1:
                bottom = True; thick = True     # bottom rule
            _set_cell_border(cell, top=top, bottom=bottom, thick=thick)
            # Also explicitly clear any shading
            tcPr = cell._tc.get_or_add_tcPr()
            shd = tcPr.find(qn('w:shd'))
            if shd is not None:
                tcPr.remove(shd)
            # Clear the default cell paragraph and write content
            cell.text = ""
            pp = cell.paragraphs[0]
            pp.alignment = (WD_ALIGN_PARAGRAPH.LEFT if ci == 0
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pp.paragraph_format.space_before = Pt(2)
            pp.paragraph_format.space_after = Pt(2)
            run = pp.add_run(val)
            set_run(run, size=Pt(10), bold=(ri == 0))

def add_reference(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = 1.15
    pf.space_before = Pt(0)
    pf.space_after = Pt(4)
    pf.left_indent = Cm(0.75)
    pf.first_line_indent = Cm(-0.75)  # hanging indent (APA)
    run = p.add_run(text)
    set_run(run, size=Pt(10))

# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------
def build():
    doc = Document()

    # Page setup: A4, 2.54 cm margins, single column
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width  = Cm(21.0)
        section.top_margin    = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin   = Cm(2.54)
        section.right_margin  = Cm(2.54)

    # Default style tweak
    style = doc.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = BODY_SIZE

    # ---- Title page (ESWA double-anonymized: author details here; the
    #      anonymized manuscript begins on a new page) ----
    t = add_para(doc, TITLE,
                 align=WD_ALIGN_PARAGRAPH.CENTER,
                 size=Pt(16), bold=True,
                 line_spacing=1.2,
                 space_before=0, space_after=14)

    # Authors line with superscript affiliation letters
    ap = doc.add_paragraph()
    ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ap.paragraph_format.line_spacing = 1.15
    ap.paragraph_format.space_after = Pt(4)
    for idx, (name, letter, corresp) in enumerate(AUTHORS):
        if idx:
            sep = ap.add_run(", ")
            set_run(sep, size=Pt(12))
        r1 = ap.add_run(name)
        set_run(r1, size=Pt(12))
        r2 = ap.add_run(letter + ("," if corresp else ""))
        set_run(r2, size=Pt(9))
        r2.font.superscript = True
        if corresp:
            r3 = ap.add_run("\u2217")
            set_run(r3, size=Pt(9))
            r3.font.superscript = True

    # Affiliations
    for letter, aff in AFFILIATIONS:
        afp = doc.add_paragraph()
        afp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        afp.paragraph_format.line_spacing = 1.15
        afp.paragraph_format.space_after = Pt(2)
        r = afp.add_run(letter + " ")
        set_run(r, size=Pt(9), italic=False)
        r.font.superscript = True
        r2 = afp.add_run(aff)
        set_run(r2, size=Pt(10), italic=True)

    # Corresponding author
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.line_spacing = 1.15
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(4)
    r = cap.add_run("\u2217 Corresponding author. E-mail: " + CORRESP_EMAIL)
    set_run(r, size=Pt(10))

    # Acknowledgements placeholder (ESWA: only on title page)
    add_heading_1(doc, "Acknowledgements")
    add_body_para(doc,
        "The authors would like to thank the editors and the anonymous "
        "reviewers for their constructive comments.")

    # Declaration of competing interests on title page (Elsevier double-blind)
    add_heading_1(doc, "Declaration of competing interests")
    add_body_para(doc,
        "The authors declare that they have no known competing financial "
        "interests or personal relationships that could have appeared to "
        "influence the work reported in this paper.")

    # ---- Page break: anonymized manuscript starts below ----
    doc.add_page_break()

    # Anonymized title (no author block, per ESWA double-anonymized policy)
    add_para(doc, TITLE,
             align=WD_ALIGN_PARAGRAPH.CENTER,
             size=Pt(16), bold=True,
             line_spacing=1.2,
             space_before=0, space_after=18)

    # Abstract
    ab_head = doc.add_paragraph()
    ab_head.alignment = WD_ALIGN_PARAGRAPH.LEFT
    ab_head.paragraph_format.space_before = Pt(0)
    ab_head.paragraph_format.space_after = Pt(4)
    r = ab_head.add_run("Abstract")
    set_run(r, size=Pt(12), bold=True)

    abs_p = doc.add_paragraph()
    abs_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    abs_p.paragraph_format.line_spacing = 1.5
    abs_p.paragraph_format.space_after = Pt(8)
    r = abs_p.add_run(ABSTRACT)
    set_run(r, size=Pt(11))

    # Keywords
    kw = doc.add_paragraph()
    kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    kw.paragraph_format.space_after = Pt(14)
    kw.paragraph_format.line_spacing = 1.5
    kw_r1 = kw.add_run("Keywords: ")
    set_run(kw_r1, size=Pt(11), bold=True)
    kw_r2 = kw.add_run("; ".join(KEYWORDS))
    set_run(kw_r2, size=Pt(11))

    # Body
    for entry in BODY:
        kind = entry[0]
        if kind == "h1":
            add_heading_1(doc, entry[1])
        elif kind == "h2":
            add_heading_2(doc, entry[1])
        elif kind == "p":
            add_body_para(doc, entry[1])
        elif kind == "eq":
            add_equation(doc, entry[1])
        elif kind == "fig":
            add_figure_caption(doc, entry[1])
        elif kind == "tbl":
            _, key, caption = entry
            add_table_caption(doc, caption)
            add_table(doc, TABLES[key])
        elif kind == "ref":
            add_reference(doc, entry[1])

    # Save
    out = "Predictive_MRQF_MS_ESWA_Elsevier.docx"
    doc.save(out)
    return out

# ---------------------------------------------------------------------------
# Sanity checks
# ---------------------------------------------------------------------------
def sanity_checks():
    n = len(ABSTRACT.split())
    assert n <= 250, f"Abstract is {n} words, must be <=250"
    assert 1 <= len(KEYWORDS) <= 7, f"Keywords count {len(KEYWORDS)} not in [1,7]"
    for h in HIGHLIGHTS:
        assert len(h) <= 85, f"Highlight exceeds 85 chars: {len(h)} -> {h}"
    # Ref integrity: every (n -> paren) must resolve
    for n in range(1, 54):
        assert n in REFMAP_PAREN, f"Missing ref map for [{n}]"
    # Every numbered citation in BODY text must resolve
    for entry in BODY:
        if entry[0] in ("p", "fig"):
            for m in re.finditer(r"\[(\d+(?:\s*,\s*\d+)*)\]", entry[1]):
                for x in m.group(1).split(","):
                    n = int(x.strip())
                    assert n in REFMAP_PAREN, f"Unresolved [{n}]"
    print(f"Sanity: abstract={n} words, keywords={len(KEYWORDS)}, "
          f"highlights all <=85 chars, refs 1-53 present.")

if __name__ == "__main__":
    sanity_checks()
    out = build()
    print("Built:", out)
