"""
ECN Cycle Time Predictor
========================
Trains Random Forest and XGBoost models on FY27 ECN data
to predict processing cycle time (ProcCT) for new ECNs.

Outputs:
  - ecn_ml_results/model_comparison.png   : Model performance comparison chart
  - ecn_ml_results/feature_importance.png : Top feature drivers chart
  - ecn_ml_results/predictions.csv        : All predictions vs actuals
  - ecn_ml_results/new_ecn_estimate.py    : Standalone estimator script
  - ecn_ml_results/report.txt             : Summary report for stakeholders
"""

import os
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec

from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer

import pickle

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
DATA_FILE   = "Reference_Filtered_Export.xlsx"
OUTPUT_DIR  = "ecn_ml_results"
TARGET      = "ProcCT(days)"          # what we predict
RANDOM_SEED = 42

os.makedirs(OUTPUT_DIR, exist_ok=True)

COLORS = {
    "blue":   "#2563eb",
    "green":  "#059669",
    "orange": "#f59e0b",
    "red":    "#ef4444",
    "purple": "#7c3aed",
    "gray":   "#6b7280",
    "light":  "#f0f7ff",
}

print("=" * 60)
print("  ECN CYCLE TIME ML PREDICTOR")
print("=" * 60)


# ─────────────────────────────────────────────
# 1. LOAD DATA
# ─────────────────────────────────────────────
print("\n[1/7] Loading data...")
df = pd.read_excel(DATA_FILE, sheet_name="Document_TB11")
print(f"      Loaded {len(df):,} rows × {df.shape[1]} columns")


# ─────────────────────────────────────────────
# 2. FEATURE ENGINEERING
# ─────────────────────────────────────────────
print("[2/7] Engineering features...")

# Keep only Closed ECNs with valid cycle times for training
df_model = df[
    (df["State"] == "Closed") &
    (df[TARGET].notna()) &
    (df[TARGET] >= 0)
].copy()
print(f"      Using {len(df_model):,} closed ECNs for training")

# ── Date features ──
df_model["SubmitDate"] = pd.to_datetime(df_model["SubmitDate"])
df_model["SubmitMonth"]     = df_model["SubmitDate"].dt.month
df_model["SubmitQuarter"]   = df_model["SubmitDate"].dt.quarter
df_model["SubmitDayOfWeek"] = df_model["SubmitDate"].dt.dayofweek   # 0=Mon
df_model["SubmitHour"]      = df_model["SubmitDate"].dt.hour
df_model["SubmitWeek"]      = df_model["SubmitDate"].dt.isocalendar().week.astype(int)

# ── Is it a Q4/end-of-quarter crunch? ──
df_model["IsQ4"]            = (df_model["SubmitQuarter"] == 4).astype(int)
df_model["IsMonthEnd"]      = (df_model["SubmitDate"].dt.day >= 25).astype(int)

# ── Rush flag ──
df_model["IsRush"]          = (df_model["RushRequest"] == "Y").astype(int)

# ── Was it ever held? ──
df_model["WasHeld"]         = df_model["HoldDate"].notna().astype(int)

# ── Has pending days? ──
df_model["HasPending"]      = df_model["PendingDays"].notna().astype(int)
df_model["PendingDaysFill"] = df_model["PendingDays"].fillna(0)

# ── Topic features — split EcnTopic into main topic + subtopic ──
def parse_topic(val):
    if pd.isna(val):
        return "UNKNOWN", "UNKNOWN"
    parts = str(val).split("~")
    main = parts[0].strip() if len(parts) > 0 else "UNKNOWN"
    sub  = parts[1].strip() if len(parts) > 1 else "UNKNOWN"
    # Shorten: extract number + label e.g. "(3) MATERIAL DISPOSITION..."
    import re
    m = re.match(r'\((\w+)\)\s*(.*)', main)
    if m:
        main = f"T{m.group(1)}"
    m2 = re.match(r'\((\w+)\)\s*(.*)', sub)
    if m2:
        sub = f"S{m2.group(1)}"
    return main[:40], sub[:40]

df_model[["MainTopic", "SubTopic"]] = pd.DataFrame(
    df_model["EcnTopic"].apply(parse_topic).tolist(),
    index=df_model.index
)

# ── Coordinator workload — how many ECNs does this coordinator typically handle? ──
coord_workload = df_model.groupby("ECNCoordinator")["RequestNum"].count().rename("CoordWorkload")
df_model = df_model.join(coord_workload, on="ECNCoordinator")

# ── Coordinator average CT (target encoding with 5-fold to avoid leakage) ──
coord_avg_ct = df_model.groupby("ECNCoordinator")[TARGET].mean().rename("CoordAvgCT")
df_model = df_model.join(coord_avg_ct, on="ECNCoordinator")

# ── MFG site fill ──
df_model["MFGSiteCode"] = df_model["MFGSiteCode"].fillna("UNKNOWN")

# ── Label encode categoricals ──
cat_cols = ["ECNCoordinatorSite", "ECNCoordinator", "MainTopic", "SubTopic", "MFGSiteCode"]
encoders = {}
for col in cat_cols:
    le = LabelEncoder()
    df_model[f"{col}_enc"] = le.fit_transform(df_model[col].astype(str))
    encoders[col] = le

# ── Final feature set ──
FEATURES = [
    # Time
    "SubmitMonth", "SubmitQuarter", "SubmitDayOfWeek", "SubmitHour",
    "SubmitWeek", "IsQ4", "IsMonthEnd",
    # Request type
    "IsRush", "WasHeld", "HasPending", "PendingDaysFill",
    # Coordinator
    "ECNCoordinatorSite_enc", "ECNCoordinator_enc",
    "CoordWorkload", "CoordAvgCT",
    # Topic
    "MainTopic_enc", "SubTopic_enc",
    # Site
    "MFGSiteCode_enc",
]

X = df_model[FEATURES].copy()
y = df_model[TARGET].copy()

# Cap extreme outliers at 99th percentile for training stability
p99 = y.quantile(0.99)
mask = y <= p99
X, y = X[mask], y[mask]
print(f"      Features: {len(FEATURES)} | Training rows (after outlier cap): {len(X):,}")
print(f"      Target range: 0 – {y.max():.1f} days  |  Median: {y.median():.1f} days")


# ─────────────────────────────────────────────
# 3. TRAIN / TEST SPLIT
# ─────────────────────────────────────────────
print("[3/7] Splitting train/test (80/20)...")
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=RANDOM_SEED
)
print(f"      Train: {len(X_train):,}  |  Test: {len(X_test):,}")


# ─────────────────────────────────────────────
# 4. TRAIN MODELS
# ─────────────────────────────────────────────
print("[4/7] Training models (this may take 1–2 minutes)...")

models = {
    "Linear Regression": LinearRegression(),
    "Random Forest":     RandomForestRegressor(
                            n_estimators=200,
                            max_depth=12,
                            min_samples_leaf=5,
                            n_jobs=-1,
                            random_state=RANDOM_SEED
                         ),
    "Gradient Boosting": GradientBoostingRegressor(
                            n_estimators=200,
                            learning_rate=0.05,
                            max_depth=5,
                            min_samples_leaf=10,
                            random_state=RANDOM_SEED
                         ),
}

results = {}
for name, model in models.items():
    print(f"      -> Training {name}...", end=" ", flush=True)
    # Fill any remaining NaNs
    X_tr = X_train.fillna(X_train.median())
    X_te = X_test.fillna(X_train.median())

    model.fit(X_tr, y_train)
    preds = model.predict(X_te)
    preds = np.clip(preds, 0, None)   # no negative predictions

    mae  = mean_absolute_error(y_test, preds)
    rmse = np.sqrt(mean_squared_error(y_test, preds))
    r2   = r2_score(y_test, preds)

    results[name] = {
        "model":  model,
        "preds":  preds,
        "mae":    mae,
        "rmse":   rmse,
        "r2":     r2,
        "X_median": X_train.median(),
    }
    print(f"MAE={mae:.2f}d  RMSE={rmse:.2f}d  R²={r2:.3f}")

# Pick best model by MAE
best_name = min(results, key=lambda k: results[k]["mae"])
best      = results[best_name]
print(f"\n      Best model: {best_name}  (MAE={best['mae']:.2f} days)")


# ─────────────────────────────────────────────
# 5. SAVE PREDICTIONS CSV
# ─────────────────────────────────────────────
print("[5/7] Saving predictions...")
pred_df = pd.DataFrame({
    "RequestNum":   df_model.loc[X_test.index, "RequestNum"].values,
    "SubmitDate":   df_model.loc[X_test.index, "SubmitDate"].values,
    "ECNTopic":     df_model.loc[X_test.index, "EcnTopic"].values,
    "Coordinator":  df_model.loc[X_test.index, "ECNCoordinator"].values,
    "Site":         df_model.loc[X_test.index, "ECNCoordinatorSite"].values,
    "Actual_ProcCT":  y_test.values,
    "Predicted_RF":   results["Random Forest"]["preds"],
    "Predicted_GB":   results["Gradient Boosting"]["preds"],
    "Error_RF":       results["Random Forest"]["preds"] - y_test.values,
    "Error_GB":       results["Gradient Boosting"]["preds"] - y_test.values,
})
pred_df.to_csv(f"{OUTPUT_DIR}/predictions.csv", index=False)
print(f"      Saved {len(pred_df):,} predictions -> {OUTPUT_DIR}/predictions.csv")


# ─────────────────────────────────────────────
# 6. CHARTS
# ─────────────────────────────────────────────
print("[6/7] Generating charts...")

plt.rcParams.update({
    "font.family": "Segoe UI",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "figure.dpi": 150,
})

# ── Chart 1: Model Comparison ──────────────────
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle("ECN Cycle Time — ML Model Performance Comparison", fontsize=14, fontweight="bold", y=1.02)

model_names  = list(results.keys())
mae_vals     = [results[m]["mae"]  for m in model_names]
rmse_vals    = [results[m]["rmse"] for m in model_names]
r2_vals      = [results[m]["r2"]   for m in model_names]
bar_colors   = [COLORS["blue"], COLORS["green"], COLORS["purple"]]

for ax, vals, title, ylabel, fmt in zip(
    axes,
    [mae_vals, rmse_vals, r2_vals],
    ["Mean Absolute Error (days)\nLower is better",
     "RMSE (days)\nLower is better",
     "R² Score\nHigher is better (max 1.0)"],
    ["Days", "Days", "R²"],
    [".2f", ".2f", ".3f"]
):
    bars = ax.bar(model_names, vals, color=bar_colors, width=0.5, zorder=3)
    ax.set_title(title, fontsize=10, pad=10)
    ax.set_ylabel(ylabel, fontsize=9)
    ax.tick_params(axis="x", labelsize=8, rotation=10)
    ax.grid(axis="y", alpha=0.3, zorder=0)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(vals)*0.02,
                f"{v:{fmt}}", ha="center", va="bottom", fontsize=9, fontweight="bold")

plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/model_comparison.png", bbox_inches="tight")
plt.close()
print(f"      -> {OUTPUT_DIR}/model_comparison.png")

# ── Chart 2: Feature Importance ────────────────
rf_model  = results["Random Forest"]["model"]
fi        = pd.Series(rf_model.feature_importances_, index=FEATURES).sort_values(ascending=True)
top_fi    = fi.tail(15)

# Friendly labels
label_map = {
    "CoordAvgCT":           "Coordinator Avg Cycle Time",
    "MainTopic_enc":        "ECN Main Topic",
    "SubTopic_enc":         "ECN Sub-Topic",
    "ECNCoordinator_enc":   "Coordinator",
    "SubmitMonth":          "Submit Month",
    "SubmitQuarter":        "Submit Quarter",
    "CoordWorkload":        "Coordinator Workload",
    "IsRush":               "Rush Request Flag",
    "WasHeld":              "Was Ever Held",
    "PendingDaysFill":      "Pending Days",
    "SubmitDayOfWeek":      "Submit Day of Week",
    "SubmitHour":           "Submit Hour",
    "IsQ4":                 "Q4 Submission",
    "IsMonthEnd":           "Month-End Submission",
    "MFGSiteCode_enc":      "Manufacturing Site",
    "ECNCoordinatorSite_enc": "Region (US/GT)",
    "SubmitWeek":           "Submit Week Number",
    "HasPending":           "Has Pending Phase",
}
top_fi.index = [label_map.get(i, i) for i in top_fi.index]

fig, ax = plt.subplots(figsize=(10, 7))
bar_colors_fi = [COLORS["blue"] if v >= top_fi.quantile(0.7) else COLORS["gray"] for v in top_fi.values]
bars = ax.barh(top_fi.index, top_fi.values, color=bar_colors_fi, height=0.65, zorder=3)
ax.set_title("Top Feature Importance — Random Forest\n(What drives ECN cycle time?)",
             fontsize=13, fontweight="bold", pad=15)
ax.set_xlabel("Importance Score", fontsize=10)
ax.grid(axis="x", alpha=0.3, zorder=0)
for bar, v in zip(bars, top_fi.values):
    ax.text(v + top_fi.max()*0.01, bar.get_y() + bar.get_height()/2,
            f"{v:.3f}", va="center", fontsize=8)
blue_patch  = mpatches.Patch(color=COLORS["blue"],  label="High importance")
gray_patch  = mpatches.Patch(color=COLORS["gray"],  label="Lower importance")
ax.legend(handles=[blue_patch, gray_patch], fontsize=9, loc="lower right")
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/feature_importance.png", bbox_inches="tight")
plt.close()
print(f"      -> {OUTPUT_DIR}/feature_importance.png")

# ── Chart 3: Predicted vs Actual scatter ───────
fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Predicted vs Actual Cycle Time", fontsize=13, fontweight="bold")

for ax, mname, color in zip(
    axes,
    ["Random Forest", "Gradient Boosting"],
    [COLORS["blue"], COLORS["purple"]]
):
    preds  = results[mname]["preds"]
    actual = y_test.values
    lim    = max(actual.max(), preds.max()) * 1.05

    ax.scatter(actual, preds, alpha=0.15, s=8, color=color, zorder=3)
    ax.plot([0, lim], [0, lim], "r--", linewidth=1.5, label="Perfect prediction", zorder=4)
    ax.set_xlim(0, lim); ax.set_ylim(0, lim)
    ax.set_xlabel("Actual Cycle Time (days)", fontsize=10)
    ax.set_ylabel("Predicted Cycle Time (days)", fontsize=10)
    ax.set_title(f"{mname}\nMAE={results[mname]['mae']:.2f}d  R²={results[mname]['r2']:.3f}",
                 fontsize=10)
    ax.legend(fontsize=8)
    ax.grid(alpha=0.2, zorder=0)

plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/predicted_vs_actual.png", bbox_inches="tight")
plt.close()
print(f"      -> {OUTPUT_DIR}/predicted_vs_actual.png")

# ── Chart 4: Error distribution ────────────────
fig, ax = plt.subplots(figsize=(10, 5))
errors_rf = results["Random Forest"]["preds"] - y_test.values
errors_gb = results["Gradient Boosting"]["preds"] - y_test.values
bins = np.linspace(-50, 50, 60)
ax.hist(errors_rf, bins=bins, alpha=0.6, color=COLORS["blue"],   label="Random Forest",     zorder=3)
ax.hist(errors_gb, bins=bins, alpha=0.6, color=COLORS["purple"], label="Gradient Boosting", zorder=3)
ax.axvline(0, color="red", linewidth=2, linestyle="--", label="Zero error")
ax.set_title("Prediction Error Distribution\n(Predicted − Actual days)", fontsize=12, fontweight="bold")
ax.set_xlabel("Error (days)", fontsize=10)
ax.set_ylabel("Count", fontsize=10)
ax.legend(fontsize=9)
ax.grid(alpha=0.2, zorder=0)
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/error_distribution.png", bbox_inches="tight")
plt.close()
print(f"      -> {OUTPUT_DIR}/error_distribution.png")

# ── Chart 5: Avg actual vs predicted by topic ──
topic_summary = pred_df.groupby(
    pred_df["ECNTopic"].str.extract(r'\((\w+)\)')[0]
).agg(
    Actual=("Actual_ProcCT", "mean"),
    Predicted_RF=("Predicted_RF", "mean"),
    Count=("Actual_ProcCT", "count")
).dropna().sort_values("Actual", ascending=True)
topic_summary = topic_summary[topic_summary["Count"] >= 50]

fig, ax = plt.subplots(figsize=(10, 5))
x = np.arange(len(topic_summary))
w = 0.35
ax.bar(x - w/2, topic_summary["Actual"],       width=w, color=COLORS["blue"],   label="Actual Avg CT",    zorder=3)
ax.bar(x + w/2, topic_summary["Predicted_RF"], width=w, color=COLORS["green"],  label="Predicted Avg CT", zorder=3)
ax.set_xticks(x)
ax.set_xticklabels([f"Topic {i}" for i in topic_summary.index], fontsize=9)
ax.set_ylabel("Average Cycle Time (days)", fontsize=10)
ax.set_title("Actual vs Predicted Cycle Time by ECN Topic\n(Random Forest)", fontsize=12, fontweight="bold")
ax.legend(fontsize=9)
ax.grid(axis="y", alpha=0.3, zorder=0)
plt.tight_layout()
plt.savefig(f"{OUTPUT_DIR}/topic_comparison.png", bbox_inches="tight")
plt.close()
print(f"      -> {OUTPUT_DIR}/topic_comparison.png")


# ─────────────────────────────────────────────
# 7. SAVE MODEL + STANDALONE ESTIMATOR + REPORT
# ─────────────────────────────────────────────
print("[7/7] Saving model and report...")

# Save the best model
with open(f"{OUTPUT_DIR}/best_model.pkl", "wb") as f:
    pickle.dump({
        "model":    best["model"],
        "features": FEATURES,
        "encoders": encoders,
        "X_median": best["X_median"],
        "name":     best_name,
    }, f)
print(f"      -> {OUTPUT_DIR}/best_model.pkl")

# ── Standalone estimator script ──────────────
estimator_code = f'''"""
ECN Cycle Time Estimator
========================
Run this script to predict cycle time for a new ECN.
Generated by ecn_cycle_time_predictor.py
"""
import pickle, numpy as np

MODEL_PATH = "ecn_ml_results/best_model.pkl"

def predict_cycle_time(
    ecn_coordinator_site: str,   # "US" or "GT"
    ecn_coordinator: str,        # coordinator login e.g. "DSENG"
    submit_month: int,            # 1-12
    submit_quarter: int,          # 1-4
    submit_day_of_week: int,      # 0=Mon, 6=Sun
    submit_hour: int,             # 0-23
    submit_week: int,             # ISO week 1-53
    is_rush: int,                 # 1 or 0
    was_held: int,                # 1 or 0
    has_pending: int,             # 1 or 0
    pending_days: float,          # 0 if none
    main_topic: str,              # e.g. "T3"
    sub_topic: str,               # e.g. "S3E"
    mfg_site_code: str,           # e.g. "AEK-ALL" or "UNKNOWN"
    coord_workload: float = 100,  # approx ECNs handled by this coordinator
    coord_avg_ct: float = 5.0,    # coordinator historical avg CT
) -> dict:
    with open(MODEL_PATH, "rb") as f:
        bundle = pickle.load(f)

    model    = bundle["model"]
    encoders = bundle["encoders"]
    features = bundle["features"]
    X_median = bundle["X_median"]

    def encode(col, val):
        le = encoders.get(col)
        if le is None:
            return 0
        if val in le.classes_:
            return int(le.transform([val])[0])
        return int(X_median.get(f"{{col}}_enc", 0))

    row = {{
        "SubmitMonth":          submit_month,
        "SubmitQuarter":        submit_quarter,
        "SubmitDayOfWeek":      submit_day_of_week,
        "SubmitHour":           submit_hour,
        "SubmitWeek":           submit_week,
        "IsQ4":                 1 if submit_quarter == 4 else 0,
        "IsMonthEnd":           0,
        "IsRush":               is_rush,
        "WasHeld":              was_held,
        "HasPending":           has_pending,
        "PendingDaysFill":      pending_days,
        "ECNCoordinatorSite_enc": encode("ECNCoordinatorSite", ecn_coordinator_site),
        "ECNCoordinator_enc":   encode("ECNCoordinator", ecn_coordinator),
        "CoordWorkload":        coord_workload,
        "CoordAvgCT":           coord_avg_ct,
        "MainTopic_enc":        encode("MainTopic", main_topic),
        "SubTopic_enc":         encode("SubTopic", sub_topic),
        "MFGSiteCode_enc":      encode("MFGSiteCode", mfg_site_code),
    }}

    X = np.array([[row[f] for f in features]], dtype=float)
    pred = float(model.predict(X)[0])
    pred = max(0, pred)

    return {{
        "predicted_proc_ct_days": round(pred, 1),
        "estimated_range_days":   f"{{max(0, pred-3):.1f}} – {{pred+3:.1f}}",
        "model_used":             bundle["name"],
    }}


if __name__ == "__main__":
    # Example: US coordinator, Topic 4, submitted in July, not a rush
    result = predict_cycle_time(
        ecn_coordinator_site = "US",
        ecn_coordinator      = "DSENG",
        submit_month         = 7,
        submit_quarter       = 1,
        submit_day_of_week   = 1,   # Tuesday
        submit_hour          = 9,
        submit_week          = 28,
        is_rush              = 0,
        was_held             = 0,
        has_pending          = 0,
        pending_days         = 0.0,
        main_topic           = "T4",
        sub_topic            = "S4A",
        mfg_site_code        = "AEK-ALL",
        coord_workload       = 500,
        coord_avg_ct         = 3.5,
    )
    print("\\nECN Cycle Time Prediction")
    print("=" * 35)
    for k, v in result.items():
        print(f"  {{k:<30}} {{v}}")
'''

with open(f"{OUTPUT_DIR}/new_ecn_estimate.py", "w") as f:
    f.write(estimator_code)
print(f"      -> {OUTPUT_DIR}/new_ecn_estimate.py")

# ── Text report ──────────────────────────────
top3_features = list(
    pd.Series(rf_model.feature_importances_, index=FEATURES)
    .sort_values(ascending=False)
    .head(3)
    .index
)
top3_friendly = [label_map.get(f, f) for f in top3_features]

report = f"""
ECN CYCLE TIME ML MODEL — SUMMARY REPORT
=========================================
Generated: 2026-07-17
Dataset:   Reference_Filtered_Export.xlsx
Records:   {len(df_model):,} closed ECNs (FY27: Jul 2025 – Jun 2026)
Target:    Processing Cycle Time (ProcCT days)

MODEL PERFORMANCE
-----------------
{'Model':<25} {'MAE (days)':<15} {'RMSE (days)':<15} {'R² Score':<10}
{'-'*65}
"""
for mname in model_names:
    r = results[mname]
    report += f"{mname:<25} {r['mae']:<15.2f} {r['rmse']:<15.2f} {r['r2']:<10.3f}\n"

report += f"""
Best Model: {best_name}
  -> On average, predictions are within ±{best['mae']:.1f} days of actual cycle time
  -> R² = {best['r2']:.3f} means the model explains {best['r2']*100:.1f}% of cycle time variance

TOP CYCLE TIME DRIVERS (Feature Importance)
--------------------------------------------
  1. {top3_friendly[0]}
  2. {top3_friendly[1]}
  3. {top3_friendly[2]}

These are the strongest predictors of how long an ECN will take.
Coordinators, ECN topic type, and submission timing are key levers.

OUTPUT FILES
------------
  predictions.csv         — All test predictions vs actuals
  best_model.pkl          — Saved model (load with pickle)
  new_ecn_estimate.py     — Standalone estimator for new ECNs
  model_comparison.png    — Model performance bar charts
  feature_importance.png  — What drives cycle time
  predicted_vs_actual.png — Scatter plot of predictions
  error_distribution.png  — Prediction error histogram
  topic_comparison.png    — Avg CT by ECN topic

NEXT STEPS
----------
  1. Review feature_importance.png with DBA/leadership
  2. Run new_ecn_estimate.py to predict any new ECN
  3. Integrate predictions.csv into the dashboard for live tracking
  4. Retrain monthly as new ECN data is added
"""

with open(f"{OUTPUT_DIR}/report.txt", "w") as f:
    f.write(report)
print(f"      -> {OUTPUT_DIR}/report.txt")

print("\n" + "=" * 60)
print("  DONE! All outputs saved to:", OUTPUT_DIR)
print("=" * 60)
print(report)
