from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Styles ────────────────────────────────────────────────────────────────────
styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Arial"
normal.font.size = Pt(10.5)

for h_name, size, bold in [("Heading 1", 16, True), ("Heading 2", 13, True), ("Heading 3", 11, True)]:
    s = styles[h_name]
    s.font.name  = "Arial"
    s.font.size  = Pt(size)
    s.font.bold  = bold
    s.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

# ── Helper functions ──────────────────────────────────────────────────────────
def add_heading(text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    return p

def add_para(text, bold_prefix=None, space_after=6):
    p = doc.add_paragraph()
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10.5)
        p.add_run(" " + text).font.name = "Arial"
    else:
        run = p.add_run(text)
        run.font.name = "Arial"
        run.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(space_after)
    return p

def add_bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(10.5)
        p.add_run(" " + text).font.name = "Arial"
    else:
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(3)
    return p

def shade_cell(cell, fill_hex):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_hex)
    tcPr.append(shd)

def make_table(headers, rows, col_widths, header_fill="1F497D", header_text_color="FFFFFF"):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    # Header row
    hdr = t.rows[0]
    for i, (h, w) in enumerate(zip(headers, col_widths)):
        cell = hdr.cells[i]
        cell.width = Inches(w)
        shade_cell(cell, header_fill)
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(*bytes.fromhex(header_text_color))
        run.font.name = "Arial"
        run.font.size = Pt(10)
        p.paragraph_format.space_after = Pt(0)
    # Data rows
    for ri, row_data in enumerate(rows):
        row = t.rows[ri + 1]
        fill = "EBF3FB" if ri % 2 == 0 else "FFFFFF"
        for ci, (val, w) in enumerate(zip(row_data, col_widths)):
            cell = row.cells[ci]
            cell.width = Inches(w)
            shade_cell(cell, fill)
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.name = "Arial"
            run.font.size = Pt(9.5)
            p.paragraph_format.space_after = Pt(0)
    doc.add_paragraph()  # spacing after table

# ══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════════════════════
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run("CLV Prediction — Technical Development Log")
run.bold = True
run.font.name  = "Arial"
run.font.size  = Pt(22)
run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
title_p.paragraph_format.space_after = Pt(6)

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run("Advanced Analytics Assignment 1  |  Group: Lowie")
run.font.name  = "Arial"
run.font.size  = Pt(13)
run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
sub_p.paragraph_format.space_after = Pt(20)

doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# 1. PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
add_heading("1. Project Overview", 1)
add_para("Goal: Predict 2018–2019 revenue (Customer Lifetime Value) for customers of a "
         "European fashion shoe shop, using only their 2016–2017 transaction history.")

add_para("Competition metrics:", space_after=3)
add_bullet("Primary: Mean Absolute Error (MAE) — lower is better")
add_bullet("Secondary: Spearman rank correlation — higher is better")
add_bullet("Professor benchmark: MAE 61.146 / Spearman 0.41")

add_para("Data:", space_after=3)
add_bullet("Training set: ~116,000 customers with known 2018–2019 revenue labels")
add_bullet("Test set: ~29,000 customers (no labels)")
add_bullet("Transaction data: ~1.2 M rows covering 2016–2017")
add_bullet("Key challenge: ~80% of customers churned (revenue = 0 in 2018–2019), "
           "making this a highly zero-inflated regression problem")

# ══════════════════════════════════════════════════════════════════════════════
# 2. APPROACH EVOLUTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading("2. Approach Evolution", 1)

# 2.1
add_heading("2.1  Initial Two-Stage Pipeline", 2)
add_para("Architecture: Binary churn classifier → revenue regressor.")
add_para("Stage 1 — Churn classifier:", bold_prefix="Stage 1 — Churn classifier:",
         space_after=3)
add_para("Three gradient boosting models (LightGBM, XGBoost, CatBoost) trained to predict "
         "whether a customer would return (revenue > 0). Ensemble averaged probabilities.")
add_para("Stage 2 — Revenue regressor:", bold_prefix="Stage 2 — Revenue regressor:",
         space_after=3)
add_para("Trained on returning customers only. Predicted log1p(revenue) with MAE objective.")
add_para("Prediction: If P(return) > threshold → predict revenue × P(return), else predict 0. "
         "Threshold grid-searched over [0.10, 0.90].")
add_para("Results: MAE 62.65 / Spearman 0.3964", bold_prefix="Results:")

add_para("Problems identified:", space_after=3)
add_bullet("Churn classifier AUC 0.72 — churner/returner distributions overlap heavily in RFM feature space.")
add_bullet("Hard threshold creates 'arms' in the scatter plot: horizontal arm (true returners predicted as churned) "
           "and vertical arm (true churners predicted as returning). Both inflate MAE directly.")
add_bullet("No feature set can fully resolve this ambiguity — churn in fashion retail is inherently probabilistic.")

# 2.2
add_heading("2.2  Probability Calibration", 2)
add_para("Change: Replaced deprecated CalibratedClassifierCV(cv='prefit') with IsotonicRegression "
         "fitted on a 20% calibration split (three-way split: 60% train / 20% calibration / 20% validation).")
add_para("Why it matters: Raw classifier probabilities are overconfident. Isotonic regression maps "
         "them to calibrated posteriors without changing AUC (calibration is a monotonic transformation).")
add_para("Results: Log loss improved 0.6013 → 0.5758. AUC unchanged at 0.72. "
         "MAE 62.50 / Spearman 0.4002. Small improvement; arm problem remained.", bold_prefix="Results:")

# 2.3
add_heading("2.3  Tweedie Regression  (Discarded)", 2)
add_para("Idea: Replace the two-stage architecture with a single Tweedie distribution model, "
         "which naturally handles zero-inflated positive data.")
add_para("Why it failed: Tweedie regression never predicts exactly 0. The ~80% of churned customers "
         "all receive small positive predictions, systematically inflating MAE.")
add_para("Results: MAE 69.88 — significantly worse. Discarded entirely.", bold_prefix="Results:")

# 2.4
add_heading("2.4  Feature Engineering", 2)
add_para("Original: Basic RFM aggregates (frequency, recency, monetary value, return rates). "
         "The following were added iteratively:")

add_bullet("Quarterly breakdown (18 features): 6 quarters (2016Q1–2017Q2) × 3 metrics "
           "(revenue, orders, returns). Lets the model weight recent quarters more heavily — "
           "a customer active in 2017Q2 is a stronger retention signal than one active only in 2016Q1.",
           bold_prefix="Quarterly breakdown (18 features):")
add_bullet("Recency normalisation: recency_normalized = recency_days / (avg_days_between_orders + 1), "
           "clipped at 99th percentile. A 60-day gap means different things for a weekly vs annual shopper.",
           bold_prefix="Recency normalisation:")
add_bullet("Target encoding: brand_target_enc and category_target_enc via 5-fold CV with Laplace "
           "smoothing (MIN_SAMPLES=5) to prevent label leakage.",
           bold_prefix="Target encoding:")
add_bullet("Return rates: brand_return_rate and category_return_rate from transaction data.",
           bold_prefix="Return rates:")

add_para("Total features after engineering: 74. Impact was marginal on final MAE — the bottleneck "
         "was churn identification, not revenue prediction for identified returners.")

# 2.5
add_heading("2.5  BG/NBD Model", 2)
add_para("Model: Beta-Geometric / Negative Binomial Distribution (Fader, Hardie & Lee, 2005) — "
         "the academic standard for CLV in non-contractual settings.")
add_para("Two components:", space_after=3)
add_bullet("BG/NBD: Models purchase frequency (Poisson) and permanent dropout (geometric). "
           "Outputs p_alive — the posterior probability a customer has NOT permanently dropped out.",
           bold_prefix="BG/NBD:")
add_bullet("Gamma-Gamma: Models expected revenue per transaction conditional on activity. "
           "Fitted only on customers with ≥ 1 repeat purchase.",
           bold_prefix="Gamma-Gamma:")

add_para("Implementation: lifetimes library (BetaGeoFitter + GammaGammaFitter). "
         "Time unit: weeks (days cause numerical instability). penalizer_coef=0.001.")
add_para("Technical fix: lifetimes stores a local lambda (generate_new_data) that pickle cannot "
         "serialise. Fixed by popping it before joblib.dump and restoring it after.")
add_para("8 features extracted: p_alive, exp_purchases_24m, exp_avg_revenue, bgnbd_clv, "
         "bgnbd_frequency, bgnbd_recency_weeks, bgnbd_T_weeks, bgnbd_monetary_value. "
         "Merged with 74 engineered features → 82 features total (customer_features_v3.csv).")

# 2.6
add_heading("2.6  Single-Stage ML with BG/NBD Features", 2)
add_para("Architecture change: Single model trained on ALL customers (including churners with revenue = 0).")
add_para("Target: log1p(revenue). Loss: MAE objective (regression_l1 / reg:absoluteerror / MAE). "
         "With MAE loss, each tree leaf predicts the median of its samples — leaves dominated "
         "by churners naturally predict 0 without needing a separate classifier or threshold.")
add_para("Methodology:", space_after=3)
add_bullet("40-iteration RandomizedSearchCV + 5-fold CV per model for hyperparameter search")
add_bullet("5-fold out-of-fold (OOF) predictions for robust MAE estimation without leakage")
add_bullet("Nelder-Mead optimisation with softmax parameterisation for ensemble weights (LGB/XGB/CAT)")
add_bullet("Scalar alpha blend: final = alpha × bgnbd_clv + (1-alpha) × ML_ensemble, alpha optimised on OOF")
add_para("Results: MAE ~62.50 / Spearman ~0.40 — similar to two-stage pipeline.", bold_prefix="Results:")
add_para("Why similar: The bottleneck is feature-space overlap, not architecture. The arms persisted "
         "in the scatter plot, confirming the model faces the same fundamental ambiguity.")

# 2.7
add_heading("2.7  P_alive Soft Scaling  (Discarded)", 2)
add_para("Idea: prediction = p_alive × E[revenue | customer is active]. Train revenue model on "
         "returning customers only, predict for all, multiply by p_alive — mirrors BG/NBD + "
         "Gamma-Gamma but with a stronger ML revenue estimator.")
add_para("Results: Worse than single-stage ensemble.", bold_prefix="Results:")
add_para("Why it failed: p_alive is already one of the 82 input features. The trees learned the "
         "optimal relationship between p_alive and the target implicitly. Multiplying by it again "
         "adds distortion. The conditional model (trained only on returners) also cannot learn "
         "when to predict near-zero.")

# 2.8
add_heading("2.8  Optuna Hyperparameter Optimisation  (Current)", 2)
add_para("Motivation: RandomizedSearchCV samples blindly. Optuna's TPE (Tree-structured Parzen "
         "Estimator) sampler learns which hyperparameter regions produce low loss and focuses "
         "exploration there — far more efficient with 8+ interacting parameters.")
add_para("Setup:", space_after=3)
add_bullet("60 trials for LightGBM and XGBoost, 40 for CatBoost (slower per trial)")
add_bullet("3-fold CV inside Optuna (fast proxy); full 5-fold OOF with winning params")
add_bullet("Wider search ranges: learning_rate 0.005–0.15 (log), n_estimators 300–3000, "
           "num_leaves 20–300, reg_alpha/lambda 1e-8 to 10 (log), XGBoost adds gamma parameter")
add_para("Expected impact: 0.5–1.5 MAE points. The gap to the professor benchmark is ~1.35 MAE, "
         "consistent with a hyperparameter tuning improvement.", bold_prefix="Expected impact:")

# ══════════════════════════════════════════════════════════════════════════════
# 3. RESULTS SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading("3. Results Summary", 1)
make_table(
    headers=["Approach", "Val MAE", "Spearman", "Notes"],
    col_widths=[2.8, 0.9, 0.9, 2.4],
    rows=[
        ["Two-stage: ML churn + revenue",       "62.65",  "0.3964", "Baseline"],
        ["+ Probability calibration",           "62.50",  "0.4002", "IsotonicRegression"],
        ["Tweedie regression",                  "69.88",  "—",      "Discarded: never predicts 0"],
        ["Single-stage + 82 features",          "~62.50", "~0.40",  "BG/NBD features included"],
        ["+ BG/NBD blend (alpha tuned)",        "~62.50", "~0.40",  "Minimal alpha improvement"],
        ["P_alive soft scaling",                "worse",  "—",      "Discarded: p_alive already a feature"],
        ["Optuna optimised (current)",          "TBD",    "TBD",    "Expected best"],
        ["Professor benchmark",                 "61.146", "0.41",   "Target to beat"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
# 4. TECHNICAL CHALLENGES
# ══════════════════════════════════════════════════════════════════════════════
add_heading("4. Technical Challenges & Fixes", 1)
make_table(
    headers=["Problem", "Cause", "Fix"],
    col_widths=[2.2, 2.2, 2.6],
    rows=[
        ["CalibratedClassifierCV(cv='prefit') error",
         "Removed in newer scikit-learn",
         "IsotonicRegression fitted directly on calibration set"],
        ["lifetimes PicklingError on joblib.dump",
         "Local lambda (generate_new_data) in fitted model",
         "Pop attribute before dump, restore after"],
        ["SHAP shape mismatch with CatBoost",
         "CatBoost adds a bias column to SHAP output",
         "shap_values[:, :-1] to strip bias column"],
        ["X_val_ret undefined in notebook 06",
         "Variable renamed during refactor",
         "Replaced with X_shap throughout cells 7–9"],
        ["FileNotFoundError for feature_cols_ml.pkl",
         "Running old main-branch notebook instead of worktree version",
         "Switch to correct worktree notebook path"],
        ["plot_probability_alive_matrix TypeError",
         "lifetimes creates its own axes; ax= kwarg forwarded to imshow",
         "Remove ax=ax; let lifetimes manage its own figure"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
# 5. ARCHITECTURE & DESIGN DECISIONS
# ══════════════════════════════════════════════════════════════════════════════
add_heading("5. Architecture & Design Decisions", 1)

add_para("Why log1p(revenue) as target?", bold_prefix="Why log1p(revenue) as target?")
add_para("Revenue is zero-inflated and right-skewed. log1p compresses the scale (€1000→6.9, "
         "€100→4.6, €0→0), reducing outlier influence and helping trees split evenly. "
         "The 0→0 mapping preserves the natural zero boundary.")

add_para("Why MAE loss?", bold_prefix="Why MAE loss?")
add_para("The competition evaluates on MAE. MAE loss (predicting the leaf median) directly aligns "
         "training and evaluation. RMSE would optimise the mean, overfitting to high-revenue outliers.")

add_para("Why weekly time units for BG/NBD?", bold_prefix="Why weekly time units for BG/NBD?")
add_para("BG/NBD assumes a Poisson purchase process. With daily units and customers buying "
         "only monthly/quarterly, parameter estimation becomes numerically unstable. "
         "Weeks strike the right balance.")

add_para("Why 5-fold OOF for ensemble weights?", bold_prefix="Why 5-fold OOF for ensemble weights?")
add_para("Using a held-out validation set to find weights risks overfitting that specific split. "
         "OOF covers all training customers and gives a less biased MAE estimate.")

add_para("Why Nelder-Mead with softmax?", bold_prefix="Why Nelder-Mead with softmax?")
add_para("Constrained optimisation (weights sum to 1, all positive) becomes unconstrained "
         "via softmax over logits. Nelder-Mead is derivative-free and robust to the non-smooth MAE objective.")

# ══════════════════════════════════════════════════════════════════════════════
# 6. NOTEBOOK STRUCTURE
# ══════════════════════════════════════════════════════════════════════════════
add_heading("6. Notebook Structure", 1)
make_table(
    headers=["Notebook", "Purpose"],
    col_widths=[2.8, 4.2],
    rows=[
        ["02_feature_engineering.ipynb",
         "RFM + quarterly + recency_normalized + target encoding → customer_features_v2.csv"],
        ["03_bgnbd_model.ipynb",
         "BG/NBD fit, p_alive extraction, merge → customer_features_v3.csv + bgf/ggf models"],
        ["03_churn_model.ipynb",
         "Old ML churn classifier (kept for reference, not in active pipeline)"],
        ["04_revenue_model.ipynb",
         "Optuna HPO, 5-fold OOF, Nelder-Mead ensemble, soft scaling comparison, retrain, save"],
        ["05_test_prediction.ipynb",
         "Load best models, predict test set → submission_v7_*.csv"],
        ["06_interpretability.ipynb",
         "SHAP beeswarm, p_alive dependence plot, quarterly importance, error analysis"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
# 7. KEY LESSONS LEARNED
# ══════════════════════════════════════════════════════════════════════════════
add_heading("7. Key Lessons Learned", 1)

lessons = [
    ("Architecture bottleneck vs feature bottleneck",
     "The arms problem is not caused by two-stage vs single-stage — it is caused by "
     "inherent overlap between churner and returner feature distributions. No architecture "
     "can cleanly separate what the data cannot."),
    ("Calibration ≠ discrimination",
     "Isotonic calibration improved log loss but not AUC or MAE, because it only "
     "rescales probabilities monotonically."),
    ("Soft scaling can hurt when the signal is already embedded",
     "p_alive was already a feature. Multiplying by it again added distortion rather than signal."),
    ("BG/NBD is best as a feature, not a replacement",
     "The probabilistic model provides valuable p_alive and CLV estimates that tree models "
     "can use contextually. Using BG/NBD alone loses the 74 engineered features."),
    ("MAE in log-space ≠ MAE in original space",
     "Training on log1p(revenue) with MAE loss minimises median log1p error, which approximately "
     "minimises original-space MAE. Far better than raw revenue (where leaf medians are mostly 0)."),
    ("RandomizedSearchCV is insufficient for 8+ interacting hyperparameters",
     "40 random trials cover a tiny fraction of a high-dimensional space. "
     "Optuna TPE makes each trial informative for the next."),
]

for i, (title, body) in enumerate(lessons, 1):
    p = doc.add_paragraph()
    r = p.add_run(f"{i}.  {title}: ")
    r.bold = True
    r.font.name = "Arial"
    r.font.size = Pt(10.5)
    r2 = p.add_run(body)
    r2.font.name = "Arial"
    r2.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(6)

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
out = ("/Users/lowievanlitsenborg/Library/Mobile Documents/"
       "com~apple~CloudDocs/lessen/1e Ma/Advanced Analytics/Assignments/1/"
       "AdvancedAnalyticsTabular/.claude/worktrees/laughing-swartz/"
       "Code_Lowie/Development_Report.docx")
doc.save(out)
print(f"Saved: {out}")
