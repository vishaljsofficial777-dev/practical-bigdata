"""
Credit EDA Case Study — end-to-end analysis script.

Reads application_data.csv and previous_application.csv, performs data
cleaning, outlier detection, univariate/bivariate analysis by TARGET,
correlation analysis by segment, and merges previous_application history
to test whether prior loan refusals predict current payment difficulty.

Usage:
    python eda_analysis.py --data-dir /path/to/csvs --out-dir ./output_csvs

All numbers referenced in credit_eda_case_study.xlsx and the README were
produced by this script.
"""

import argparse
import os
import numpy as np
import pandas as pd


def load_application_data(path):
    df = pd.read_csv(path, encoding="latin-1")
    return df


def clean_application_data(df):
    """Drop very sparse columns, fix known anomalies, add derived fields."""
    high_miss = df.isnull().mean()
    drop_cols = high_miss[high_miss > 0.40].index.tolist()
    df = df.drop(columns=drop_cols)

    # DAYS_EMPLOYED anomaly: 365243 is a placeholder (pensioners/unemployed)
    df["DAYS_EMPLOYED_ANOM"] = df["DAYS_EMPLOYED"] == 365243
    df["DAYS_EMPLOYED"] = df["DAYS_EMPLOYED"].replace(365243, np.nan)

    # Convert DAYS_* (negative, days-before-application) to readable years
    for col in ["DAYS_BIRTH", "DAYS_EMPLOYED", "DAYS_REGISTRATION", "DAYS_ID_PUBLISH"]:
        if col in df.columns:
            df[col.replace("DAYS_", "YEARS_")] = (-df[col] / 365).round(1)

    # Derived affordability ratios
    df["CREDIT_INCOME_RATIO"] = df["AMT_CREDIT"] / df["AMT_INCOME_TOTAL"]
    df["ANNUITY_INCOME_RATIO"] = df["AMT_ANNUITY"] / df["AMT_INCOME_TOTAL"]

    return df, drop_cols


def outlier_scan(df, cols):
    rows = []
    for col in cols:
        q1, q3 = df[col].quantile([0.25, 0.75])
        iqr = q3 - q1
        upper = q3 + 1.5 * iqr
        above = df[col] > upper
        rows.append({
            "Variable": col,
            "Median": df[col].median(),
            "Upper_Whisker": upper,
            "Count_Above_Whisker": above.sum(),
            "Pct_Above_Whisker": round(above.mean() * 100, 2),
            "Max_Value": df[col].max(),
        })
    return pd.DataFrame(rows).set_index("Variable")


def segment_medians(df, num_vars):
    return df.groupby("TARGET")[num_vars].median().T


def categorical_default_rates(df, cat_vars, min_count=200):
    out = {}
    for col in cat_vars:
        grp = df.groupby(col)["TARGET"].agg(["mean", "count"])
        grp = grp[grp["count"] >= min_count].sort_values("mean", ascending=False)
        grp["mean"] = (grp["mean"] * 100).round(2)
        out[col] = grp.rename(columns={"mean": "default_rate_pct"})
    return out


def bucketed_default_rate(df, col, bins, labels):
    bucket_col = f"{col}_BUCKET"
    df[bucket_col] = pd.cut(df[col], bins=bins, labels=labels)
    return df.groupby(bucket_col, observed=True)["TARGET"].agg(["mean", "count"])


def correlation_by_target(df, num_cols):
    corr0 = df[df["TARGET"] == 0][num_cols].corr()
    corr1 = df[df["TARGET"] == 1][num_cols].corr()
    return corr0, corr1


def correlation_shift(corr0, corr1):
    diff = (corr1 - corr0).abs().copy()
    vals = diff.values.copy()
    np.fill_diagonal(vals, 0)
    diff = pd.DataFrame(vals, index=diff.index, columns=diff.columns)
    pairs = []
    for i in range(len(diff.columns)):
        for j in range(i + 1, len(diff.columns)):
            pairs.append((diff.columns[i], diff.columns[j], diff.iloc[i, j],
                          corr0.iloc[i, j], corr1.iloc[i, j]))
    return pd.DataFrame(
        pairs, columns=["var1", "var2", "abs_diff", "corr_target0", "corr_target1"]
    ).sort_values("abs_diff", ascending=False)


def merge_previous_application(app_df, prev_path):
    """Test whether prior refusal/cancellation history predicts current default."""
    prev = pd.read_csv(
        prev_path, encoding="latin-1",
        usecols=["SK_ID_PREV", "SK_ID_CURR", "NAME_CONTRACT_STATUS"],
    )
    status_counts = (
        prev.groupby(["SK_ID_CURR", "NAME_CONTRACT_STATUS"]).size().unstack(fill_value=0)
    )
    status_counts.columns = [
        f"PREV_{c.upper().replace(' ', '_')}_CNT" for c in status_counts.columns
    ]
    status_counts["PREV_TOTAL_APPS"] = status_counts.sum(axis=1)

    merged = app_df[["SK_ID_CURR", "TARGET"]].merge(
        status_counts, on="SK_ID_CURR", how="left"
    ).fillna(0)

    merged["HAD_PRIOR_REFUSAL"] = merged["PREV_REFUSED_CNT"] > 0
    merged["refusal_bucket"] = merged["PREV_REFUSED_CNT"].clip(upper=5)

    by_flag = merged.groupby("HAD_PRIOR_REFUSAL")["TARGET"].agg(["mean", "count"])
    by_dose = merged.groupby("refusal_bucket")["TARGET"].agg(["mean", "count"])
    status_dist = prev["NAME_CONTRACT_STATUS"].value_counts()

    return by_flag, by_dose, status_dist


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-dir", default=".", help="Folder containing the CSV files")
    parser.add_argument("--out-dir", default="./output_csvs", help="Where to write result CSVs")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    app_path = os.path.join(args.data_dir, "application_data.csv")
    prev_path = os.path.join(args.data_dir, "previous_application.csv")

    print("Loading application_data.csv ...")
    df = load_application_data(app_path)
    print(f"Shape: {df.shape}")

    target_dist = df["TARGET"].value_counts(normalize=True) * 100
    print("\nTARGET distribution (%):")
    print(target_dist)

    df, dropped = clean_application_data(df)
    print(f"\nDropped {len(dropped)} columns with >40% missing values.")

    outlier_cols = ["AMT_INCOME_TOTAL", "AMT_CREDIT", "AMT_ANNUITY",
                     "AMT_GOODS_PRICE", "CNT_CHILDREN", "YEARS_EMPLOYED"]
    outliers = outlier_scan(df, outlier_cols)
    outliers.to_csv(os.path.join(args.out_dir, "outlier_scan.csv"))
    print("\nOutlier scan saved.")

    num_vars = ["AMT_INCOME_TOTAL", "AMT_CREDIT", "AMT_ANNUITY", "AMT_GOODS_PRICE",
                "CREDIT_INCOME_RATIO", "ANNUITY_INCOME_RATIO", "CNT_CHILDREN",
                "YEARS_BIRTH", "YEARS_EMPLOYED", "REGION_RATING_CLIENT",
                "EXT_SOURCE_2", "EXT_SOURCE_3"]
    seg = segment_medians(df, num_vars)
    seg.to_csv(os.path.join(args.out_dir, "segment_medians.csv"))
    print("\nSegment medians by TARGET saved.")

    cat_vars = ["NAME_CONTRACT_TYPE", "CODE_GENDER", "NAME_INCOME_TYPE",
                "NAME_EDUCATION_TYPE", "NAME_FAMILY_STATUS", "NAME_HOUSING_TYPE",
                "OCCUPATION_TYPE", "ORGANIZATION_TYPE", "FLAG_OWN_CAR",
                "FLAG_OWN_REALTY", "NAME_TYPE_SUITE"]
    cat_results = categorical_default_rates(df, cat_vars)
    for name, tbl in cat_results.items():
        tbl.to_csv(os.path.join(args.out_dir, f"default_rate_by_{name}.csv"))
    print("\nCategorical default-rate tables saved.")

    income_buckets = bucketed_default_rate(
        df, "AMT_INCOME_TOTAL", [0, 100000, 150000, 200000, 300000, np.inf],
        ["<100K", "100-150K", "150-200K", "200-300K", "300K+"],
    )
    income_buckets.to_csv(os.path.join(args.out_dir, "default_rate_by_income_bucket.csv"))

    age_buckets = bucketed_default_rate(
        df, "YEARS_BIRTH", [20, 30, 40, 50, 60, 70],
        ["20-30", "30-40", "40-50", "50-60", "60-70"],
    )
    age_buckets.to_csv(os.path.join(args.out_dir, "default_rate_by_age_bucket.csv"))
    print("\nIncome/age bucket tables saved.")

    corr_num_cols = ["AMT_INCOME_TOTAL", "AMT_CREDIT", "AMT_ANNUITY", "AMT_GOODS_PRICE",
                      "CREDIT_INCOME_RATIO", "ANNUITY_INCOME_RATIO", "CNT_CHILDREN",
                      "CNT_FAM_MEMBERS", "YEARS_BIRTH", "YEARS_EMPLOYED",
                      "REGION_RATING_CLIENT", "REGION_POPULATION_RELATIVE",
                      "EXT_SOURCE_2", "EXT_SOURCE_3", "OBS_30_CNT_SOCIAL_CIRCLE",
                      "DEF_30_CNT_SOCIAL_CIRCLE"]
    corr0, corr1 = correlation_by_target(df, corr_num_cols)
    corr0.to_csv(os.path.join(args.out_dir, "corr_target0.csv"))
    corr1.to_csv(os.path.join(args.out_dir, "corr_target1.csv"))
    corr_diff = correlation_shift(corr0, corr1)
    corr_diff.to_csv(os.path.join(args.out_dir, "corr_diff.csv"), index=False)
    print("\nCorrelation matrices + shift table saved.")

    if os.path.exists(prev_path):
        print("\nLoading previous_application.csv (this is a large file, may take a while) ...")
        by_flag, by_dose, status_dist = merge_previous_application(df, prev_path)
        by_flag.to_csv(os.path.join(args.out_dir, "default_rate_by_prior_refusal_flag.csv"))
        by_dose.to_csv(os.path.join(args.out_dir, "default_rate_by_num_prior_refusals.csv"))
        status_dist.to_csv(os.path.join(args.out_dir, "prior_application_status_distribution.csv"))
        print("\nPrevious-application merge analysis saved.")
    else:
        print(f"\n{prev_path} not found — skipping previous_application merge step.")

    print(f"\nAll result tables written to: {args.out_dir}")


if __name__ == "__main__":
    main()
