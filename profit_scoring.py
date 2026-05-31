import re
import numpy as np
import pandas as pd
import xlwings as xw

from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.ensemble import HistGradientBoostingRegressor


# ======================
# SHEETS
# ======================
SHEET_HISTORY_CLEAN = "Job_History_Clean"
SHEET_UPCOMING_CLEAN = "Upcoming_Jobs_Clean"

CONFIG_SHEET_PRIMARY = "Job_History"          # AI4/AI5 + confusion matrix lives here
CONFIG_SHEET_FALLBACK = SHEET_HISTORY_CLEAN

DEFAULT_K = 10
DEFAULT_MODE = "profit"  # profit | margin | both

# >>> CLIPPED MODEL ADDITIONS
CLIP_LO_PCT = 1    # 1st percentile
CLIP_HI_PCT = 99   # 99th percentile
OUT_PROFIT_CLIPPED = "Predicted Profit (Clipped)"
OUT_MARGIN_CLIPPED = "Predicted Profit Margin (Clipped)"


# ======================
# DIVISION MAPPING
# ======================
DIV_MAP = {
    "0": "Int Painting",
    "int painting": "Int Painting",
    "wallpaper": "Int Painting",
    "drywall": "Carpentry",
    "remodel": "Carpentry",
    "carpentry": "Carpentry",
    "wash": "Ext Painting",
    "ext painting": "Ext Painting",
}

def clean_division(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    s = str(x).strip()
    if not s:
        return np.nan
    return DIV_MAP.get(s.lower(), s)


# ======================
# HELPERS
# ======================
def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())

def pick_col(df: pd.DataFrame, candidates, required=True):
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    norm_map = {_norm(c): c for c in cols}
    for c in candidates:
        key = _norm(c)
        if key in norm_map:
            return norm_map[key]
    if required:
        raise KeyError(f"Missing required column. Tried: {candidates}\nAvailable: {cols}")
    return None

def money_to_float(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(r"[^0-9\.\-]", "", regex=True).replace("", np.nan)
    return pd.to_numeric(s, errors="coerce")

def fill_blank(a: pd.Series, b: pd.Series) -> pd.Series:
    a2 = a.astype(str)
    blank = a.isna() | (a2.str.strip() == "") | (a2.str.lower().str.strip() == "nan")
    return a.where(~blank, b)

def nonblank_text(s: pd.Series) -> pd.Series:
    s2 = s.astype(str).str.strip()
    return s.notna() & (s2 != "") & (s2.str.lower() != "nan")

def valid_zip(z: pd.Series) -> pd.Series:
    zz = pd.to_numeric(z, errors="coerce")
    return zz.notna() & (zz >= 10000) & (zz <= 99999)

def get_config(wb: xw.Book):
    sh = None
    for name in [CONFIG_SHEET_PRIMARY, CONFIG_SHEET_FALLBACK]:
        try:
            sh = wb.sheets[name]
            break
        except Exception:
            continue

    k_val = DEFAULT_K
    mode = DEFAULT_MODE

    if sh is not None:
        raw_k = sh.range("AI4").value
        raw_mode = sh.range("AI5").value

        try:
            if raw_k is not None and str(raw_k).strip() != "":
                k_val = int(float(raw_k))
        except Exception:
            k_val = DEFAULT_K

        if raw_mode is not None:
            m = str(raw_mode).strip().lower()
            if m in {"profit", "margin", "both"}:
                mode = m

    if k_val <= 0:
        k_val = DEFAULT_K

    return k_val, mode


# ---- Excel reading/writing (formula-safe) ----
def excel_last_row(sheet: xw.Sheet) -> int:
    api = sheet.api
    xlByRows = 1
    xlPrevious = 2
    xlFormulas = -4123
    found = api.Cells.Find(What="*", LookIn=xlFormulas, SearchOrder=xlByRows, SearchDirection=xlPrevious)
    return int(found.Row) if found is not None else 1

def last_header_col(sheet: xw.Sheet, header_row: int = 1, max_cols: int = 250) -> int:
    vals = sheet.range((header_row, 1), (header_row, max_cols)).value
    if not isinstance(vals, list):
        vals = [vals]
    last = 0
    for i, v in enumerate(vals, start=1):
        if v is None:
            continue
        if str(v).strip() != "":
            last = i
    return last if last > 0 else 1

def read_table(sheet: xw.Sheet, header_row: int = 1) -> pd.DataFrame:
    last_col = last_header_col(sheet, header_row=header_row)
    headers = sheet.range((header_row, 1), (header_row, last_col)).value
    if not isinstance(headers, list):
        headers = [headers]
    headers = [("" if h is None else str(h).strip()) for h in headers]
    if all(h == "" for h in headers):
        return pd.DataFrame()

    last_row = excel_last_row(sheet)
    if last_row <= header_row:
        return pd.DataFrame(columns=[h for h in headers if h != ""])

    rng = sheet.range((header_row, 1), (last_row, last_col))
    values = rng.value
    if not isinstance(values, list) or (values and not isinstance(values[0], list)):
        return pd.DataFrame(columns=[h for h in headers if h != ""])

    df = pd.DataFrame(values[1:], columns=[str(c).strip() if c is not None else "" for c in values[0]])
    df = df.loc[:, df.columns != ""].copy()
    df = df.dropna(how="all").copy()
    return df

def ensure_output_columns(sheet: xw.Sheet, out_cols, header_row: int = 1) -> dict:
    last_col = last_header_col(sheet, header_row=header_row)
    header_vals = sheet.range((header_row, 1), (header_row, last_col)).value
    if not isinstance(header_vals, list):
        header_vals = [header_vals]
    header_vals = [("" if v is None else str(v).strip()) for v in header_vals]

    col_map = {name: (i + 1) for i, name in enumerate(header_vals) if name != ""}
    cur_last = last_col
    for c in out_cols:
        if c not in col_map:
            cur_last += 1
            sheet.range((header_row, cur_last)).value = c
            col_map[c] = cur_last
    return col_map

def clear_output_columns(sheet: xw.Sheet, col_indices, from_row: int, to_row: int):
    for ci in col_indices:
        sheet.range((from_row, ci), (to_row, ci)).clear_contents()

def write_column(sheet: xw.Sheet, col_index: int, start_row: int, values):
    sheet.range((start_row, col_index)).options(transpose=True).value = values

def build_profit_pipeline():
    cat_cols = ["Estimator", "Crew Leader", "Division_clean"]
    num_cols = ["Cust Zip Code", "Contracted Revenue", "Contract Estimated Hours"]

    preprocess = ColumnTransformer(
        transformers=[
            ("num", Pipeline([("imp", SimpleImputer(strategy="median"))]), num_cols),
            ("cat", Pipeline([
                ("imp", SimpleImputer(strategy="most_frequent")),
                ("oh", OneHotEncoder(handle_unknown="ignore", sparse_output=False))
            ]), cat_cols),
        ],
        remainder="drop",
    )

    model = HistGradientBoostingRegressor(
        random_state=42,
        max_depth=6,
        learning_rate=0.05,
        max_leaf_nodes=31,
        min_samples_leaf=20
    )

    return Pipeline([("prep", preprocess), ("model", model)])


def write_confusion_matrix(wb: xw.Book, train_df: pd.DataFrame, X_cols_df: pd.DataFrame, y: pd.Series):
    """
    Confusion matrix for LOSS detection:
      Actual Loss = PROFIT <= 0
      Pred  Loss = predicted PROFIT <= 0
    Time-based holdout = last 20% (uses Date if available and parseable, else row order).

    Written to Job_History!AJ8:AK9 with labels around it.
    Layout (rows=actual, cols=pred):
        AJ8 TP  | AK8 FN
        AJ9 FP  | AK9 TN
    """
    # pick sheet to write to
    try:
        sh = wb.sheets[CONFIG_SHEET_PRIMARY]
    except Exception:
        sh = wb.sheets[CONFIG_SHEET_FALLBACK]

    raw = sh.range("AK4").value
    try:
        test_frac = float(raw)
    except Exception:
        test_frac = 0.20

    # guardrails
    if test_frac < 0.05 or test_frac > 0.50:
        test_frac = 0.20

    # Need enough rows to make this meaningful
    n = len(train_df)
    if n < 20:
        # clear cells if not enough data
        sh.range("AJ8:AK9").clear_contents()
        sh.range("AI7:AK7").clear_contents()
        sh.range("AI8:AI9").clear_contents()
        sh.range("AI7").value = f"Confusion matrix: not enough rows (n={n})"
        return

    # Sort by Date if we can
    df_eval = train_df.copy()
    date_col = None
    for c in ["Date", "date"]:
        if c in df_eval.columns:
            date_col = c
            break

    if date_col is not None:
        dt = pd.to_datetime(df_eval[date_col], errors="coerce")
        if dt.notna().any():
            df_eval = df_eval.assign(_dt=dt).sort_values("_dt").drop(columns=["_dt"])
        # else keep row order

    split_idx = int(n * (1.0 - test_frac))
    df_tr = df_eval.iloc[:split_idx].copy()
    df_te = df_eval.iloc[split_idx:].copy()

    X_tr = X_cols_df.loc[df_tr.index]
    y_tr = y.loc[df_tr.index]
    X_te = X_cols_df.loc[df_te.index]
    y_te = y.loc[df_te.index]

    pipe = build_profit_pipeline()
    pipe.fit(X_tr, y_tr)
    yhat = pipe.predict(X_te)

    actual_loss = (y_te.values <= 0)
    pred_loss = (yhat <= 0)

    TP = int(np.sum(pred_loss & actual_loss))
    FN = int(np.sum((~pred_loss) & actual_loss))
    FP = int(np.sum(pred_loss & (~actual_loss)))
    TN = int(np.sum((~pred_loss) & (~actual_loss)))

    # labels + matrix
    sh.range("AI7").value = "Confusion (Loss vs Profit) | rows=Actual, cols=Pred"
    sh.range("AJ7").value = "Pred Loss"
    sh.range("AK7").value = "Pred Profit"
    sh.range("AI8").value = "Actual Loss"
    sh.range("AI9").value = "Actual Profit"

    sh.range("AJ8").value = TP
    sh.range("AK8").value = FN
    sh.range("AJ9").value = FP
    sh.range("AK9").value = TN


# ======================
# MAIN
# ======================
def main():
    wb = xw.Book.caller()
    k, mode = get_config(wb)

    sh_hist = wb.sheets[SHEET_HISTORY_CLEAN]
    sh_upc  = wb.sheets[SHEET_UPCOMING_CLEAN]

    hist = read_table(sh_hist, header_row=1)
    if hist.empty:
        raise ValueError(f"{SHEET_HISTORY_CLEAN} unreadable/empty.")

    # ----- HISTORY (training) -----
    col_date_h = pick_col(hist, ["Date"], required=False)

    col_est_h  = pick_col(hist, ["Estimator"])
    col_crew_h = pick_col(hist, ["Crew Leader", "Crew Lead", "CrewLead"])
    col_zip_h  = pick_col(hist, ["Cust Zip Code", "Zip", "Zip Code"])
    col_div_h  = pick_col(hist, ["Division/ Type", "Division / Type", "Division/Type", "Division Type", "Division"])
    col_rev_h  = pick_col(hist, ["Contracted Revenue", "Contract Revenue", "Revenue"])
    col_hrs_h  = pick_col(hist, ["Contract Estimated Hours", "Contract Est Hours", "Estimated Hours", "Contract Hours"])

    col_totalrev_h = pick_col(hist, ["Total Revenue", "TotalRevenue"], required=False)
    col_exp_h      = pick_col(hist, ["Total Job Expense", "Total Expense", "Job Expense", "TotalJobExpense"], required=False)
    col_profit_h   = pick_col(hist, ["PROFIT", "Profit"], required=False)

    hist_w = hist.copy()
    hist_w[col_crew_h] = fill_blank(hist_w[col_crew_h], hist_w[col_est_h])
    hist_w["Division_clean"] = hist_w[col_div_h].apply(clean_division)

    hist_w["_ContractedRevenue"] = money_to_float(hist_w[col_rev_h])
    hist_w["_EstHours"] = pd.to_numeric(hist_w[col_hrs_h], errors="coerce")
    hist_w["_Zip"] = pd.to_numeric(hist_w[col_zip_h], errors="coerce")

    # Treat formula-blank zeros as missing for training
    hist_w.loc[hist_w["_ContractedRevenue"] <= 0, "_ContractedRevenue"] = np.nan
    hist_w.loc[hist_w["_EstHours"] <= 0, "_EstHours"] = np.nan

    if col_profit_h and hist_w[col_profit_h].notna().any():
        hist_w["_PROFIT"] = money_to_float(hist_w[col_profit_h])
    else:
        if not col_totalrev_h or not col_exp_h:
            raise ValueError("History needs PROFIT OR Total Revenue + Total Job Expense.")
        hist_w["_TotalRevenue"] = money_to_float(hist_w[col_totalrev_h])
        hist_w["_Expense"] = money_to_float(hist_w[col_exp_h])
        hist_w["_PROFIT"] = hist_w["_TotalRevenue"] - hist_w["_Expense"]

        zero_zero = (hist_w["_TotalRevenue"].fillna(0) == 0) & (hist_w["_Expense"].fillna(0) == 0)
        hist_w.loc[zero_zero, "_PROFIT"] = np.nan

    # STRICT training filter
    train_mask = (
        nonblank_text(hist_w[col_est_h]) &
        hist_w["Division_clean"].notna() &
        valid_zip(hist_w["_Zip"]) &
        hist_w["_ContractedRevenue"].notna() &
        hist_w["_EstHours"].notna() &
        hist_w["_PROFIT"].notna()
    )
    train = hist_w.loc[train_mask].copy()

    # Build X/y
    X_all = pd.DataFrame({
        "Estimator": train[col_est_h].astype(str),
        "Crew Leader": train[col_crew_h].astype(str),
        "Cust Zip Code": train["_Zip"],
        "Division_clean": train["Division_clean"].astype(str),
        "Contracted Revenue": train["_ContractedRevenue"],
        "Contract Estimated Hours": train["_EstHours"],
    }, index=train.index)
    y_all = train["_PROFIT"].astype(float)

    # Train final model used for scoring (original)
    pipe = build_profit_pipeline()
    pipe.fit(X_all, y_all)

    # >>> CLIPPED MODEL ADDITIONS
    # Winsorize/clamp the PROFIT target so rare disasters don't yank the model around as much.
    lo = float(np.nanpercentile(y_all.values, CLIP_LO_PCT))
    hi = float(np.nanpercentile(y_all.values, CLIP_HI_PCT))
    y_clipped = y_all.clip(lower=lo, upper=hi)

    pipe_clipped = build_profit_pipeline()
    pipe_clipped.fit(X_all, y_clipped)
    # <<< CLIPPED MODEL ADDITIONS

    # Confusion matrix to Job_History!AJ8:AK9 (time-based holdout)
    write_confusion_matrix(wb, train.copy(), X_all, y_all)

    # ----- UPCOMING (scoring) -----
    upc = read_table(sh_upc, header_row=1)

    out_map = ensure_output_columns(
        sh_upc,
        [
            "Predicted Profit", "Predicted Profit Margin",
            OUT_PROFIT_CLIPPED, OUT_MARGIN_CLIPPED,  # >>> CLIPPED MODEL ADDITIONS
            "BottomK_Profit_Flag", "BottomK_Margin_Flag", "Bottom_K_Flag"
        ],
        header_row=1
    )

    # Clear old outputs down the sheet so stale scores don’t linger
    last_row_upc = excel_last_row(sh_upc)
    clear_output_columns(
        sh_upc,
        [out_map[c] for c in [
            "Predicted Profit", "Predicted Profit Margin",
            OUT_PROFIT_CLIPPED, OUT_MARGIN_CLIPPED,  # >>> CLIPPED MODEL ADDITIONS
            "BottomK_Profit_Flag", "BottomK_Margin_Flag", "Bottom_K_Flag"
        ]],
        from_row=2,
        to_row=max(last_row_upc, 2)
    )

    if upc.empty:
        wb.app.status_bar = f"Trained on {len(train)} real jobs; no upcoming rows detected."
        return

    col_est_u  = pick_col(upc, ["Estimator"])
    col_zip_u  = pick_col(upc, ["Cust Zip Code", "Zip", "Zip Code"])
    col_div_u  = pick_col(upc, ["Division/ Type", "Division / Type", "Division/Type", "Division Type", "Division"])
    col_rev_u  = pick_col(upc, ["Contracted Revenue", "Contract Revenue", "Revenue"])
    col_hrs_u  = pick_col(upc, ["Contract Estimated Hours", "Contract Est Hours", "Estimated Hours", "Contract Hours"])
    col_crew_u = pick_col(upc, ["Crew Leader", "Crew Lead", "CrewLead"], required=False)

    upc_w = upc.copy()
    if col_crew_u is None:
        upc_w["Crew Leader"] = np.nan
        col_crew_u = "Crew Leader"

    upc_w[col_crew_u] = fill_blank(upc_w[col_crew_u], upc_w[col_est_u])
    upc_w["Division_clean"] = upc_w[col_div_u].apply(clean_division)
    upc_w["_ContractedRevenue"] = money_to_float(upc_w[col_rev_u])
    upc_w["_EstHours"] = pd.to_numeric(upc_w[col_hrs_u], errors="coerce")
    upc_w["_Zip"] = pd.to_numeric(upc_w[col_zip_u], errors="coerce")

    upc_w.loc[upc_w["_ContractedRevenue"] <= 0, "_ContractedRevenue"] = np.nan
    upc_w.loc[upc_w["_EstHours"] <= 0, "_EstHours"] = np.nan

    score_mask = (
        nonblank_text(upc_w[col_est_u]) &
        upc_w["Division_clean"].notna() &
        valid_zip(upc_w["_Zip"]) &
        upc_w["_ContractedRevenue"].notna() &
        upc_w["_EstHours"].notna()
    )

    pred_profit = np.array([np.nan] * len(upc_w), dtype=float)
    pred_margin = np.array([np.nan] * len(upc_w), dtype=float)

    # >>> CLIPPED MODEL ADDITIONS
    pred_profit_clip = np.array([np.nan] * len(upc_w), dtype=float)
    pred_margin_clip = np.array([np.nan] * len(upc_w), dtype=float)
    # <<< CLIPPED MODEL ADDITIONS

    if score_mask.sum() > 0:
        X_new = pd.DataFrame({
            "Estimator": upc_w.loc[score_mask, col_est_u].astype(str),
            "Crew Leader": upc_w.loc[score_mask, col_crew_u].astype(str),
            "Cust Zip Code": upc_w.loc[score_mask, "_Zip"],
            "Division_clean": upc_w.loc[score_mask, "Division_clean"].astype(str),
            "Contracted Revenue": upc_w.loc[score_mask, "_ContractedRevenue"],
            "Contract Estimated Hours": upc_w.loc[score_mask, "_EstHours"],
        })
        p = pipe.predict(X_new)
        pred_profit[score_mask.values] = p
        denom = upc_w.loc[score_mask, "_ContractedRevenue"].values
        pred_margin[score_mask.values] = np.where(denom > 0, p / denom, np.nan)

        # >>> CLIPPED MODEL ADDITIONS
        pc = pipe_clipped.predict(X_new)
        pred_profit_clip[score_mask.values] = pc
        pred_margin_clip[score_mask.values] = np.where(denom > 0, pc / denom, np.nan)
        # <<< CLIPPED MODEL ADDITIONS

    # Bottom-K among scored rows (based on ORIGINAL predictions)
    bottom_profit_flag = np.zeros(len(upc_w), dtype=bool)
    bottom_margin_flag = np.zeros(len(upc_w), dtype=bool)

    scored_idx = np.where(~np.isnan(pred_profit))[0]
    if scored_idx.size > 0:
        k_eff = min(k, scored_idx.size)
        order_profit = scored_idx[np.argsort(pred_profit[scored_idx])]
        bottom_profit_flag[order_profit[:k_eff]] = True

        scored_idx_m = np.where(~np.isnan(pred_margin))[0]
        if scored_idx_m.size > 0:
            k_eff2 = min(k, scored_idx_m.size)
            order_margin = scored_idx_m[np.argsort(pred_margin[scored_idx_m])]
            bottom_margin_flag[order_margin[:k_eff2]] = True

    if mode == "profit":
        bottom_k_flag = bottom_profit_flag
    elif mode == "margin":
        bottom_k_flag = bottom_margin_flag
    else:
        bottom_k_flag = bottom_profit_flag | bottom_margin_flag

    def nan_to_none(arr):
        return [None if (v is None or (isinstance(v, float) and np.isnan(v))) else float(v) for v in arr]

    write_column(sh_upc, out_map["Predicted Profit"],        start_row=2, values=nan_to_none(pred_profit))
    write_column(sh_upc, out_map["Predicted Profit Margin"], start_row=2, values=nan_to_none(pred_margin))

    # >>> CLIPPED MODEL ADDITIONS
    write_column(sh_upc, out_map[OUT_PROFIT_CLIPPED],        start_row=2, values=nan_to_none(pred_profit_clip))
    write_column(sh_upc, out_map[OUT_MARGIN_CLIPPED],        start_row=2, values=nan_to_none(pred_margin_clip))
    # <<< CLIPPED MODEL ADDITIONS

    write_column(sh_upc, out_map["BottomK_Profit_Flag"],     start_row=2, values=list(bottom_profit_flag))
    write_column(sh_upc, out_map["BottomK_Margin_Flag"],     start_row=2, values=list(bottom_margin_flag))
    write_column(sh_upc, out_map["Bottom_K_Flag"],           start_row=2, values=list(bottom_k_flag))

    wb.app.status_bar = (
        f"Trained on {len(train)} real jobs; scored {int(score_mask.sum())} upcoming rows. "
        f"K={k} mode={mode}. "
        f"Clipped target p{CLIP_LO_PCT}={lo:,.0f}, p{CLIP_HI_PCT}={hi:,.0f}."
    )