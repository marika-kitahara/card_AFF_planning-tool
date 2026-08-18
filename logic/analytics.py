import numpy as np
import pandas as pd


def _clean_month_label(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip()


def prepare_monthly_performance(history_df: pd.DataFrame) -> pd.DataFrame:
    """月度別の発生件数・発行数・コスト・発行CPAを返す。"""
    required = {"date", "月度", "cv", "cost", "approved_cv"}
    missing = required - set(history_df.columns)
    if missing:
        raise ValueError(
            "月度別実績の集計に必要な列がありません: "
            + ", ".join(sorted(missing))
        )

    work = history_df[["date", "月度", "cv", "cost", "approved_cv"]].copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce")
    work["月度"] = _clean_month_label(work["月度"])
    work["cv"] = pd.to_numeric(work["cv"], errors="coerce").fillna(0.0)
    work["cost"] = pd.to_numeric(work["cost"], errors="coerce").fillna(0.0)
    work["approved_cv"] = pd.to_numeric(
        work["approved_cv"], errors="coerce"
    ).fillna(0.0)

    work = work[
        work["date"].notna()
        & work["月度"].notna()
        & work["月度"].ne("")
        & work["月度"].ne("未設定")
    ].copy()

    if work.empty:
        return pd.DataFrame(
            columns=["月度", "月度開始日", "発生件数", "発行数", "コスト", "発行CPA"]
        )

    monthly = (
        work.groupby("月度", as_index=False)
        .agg(
            月度開始日=("date", "min"),
            発生件数=("cv", "sum"),
            発行数=("approved_cv", "sum"),
            コスト=("cost", "sum"),
        )
        .sort_values(["月度開始日", "月度"], kind="stable")
        .reset_index(drop=True)
    )

    monthly["発行CPA"] = np.where(
        monthly["発行数"] > 0,
        monthly["コスト"] / monthly["発行数"],
        np.nan,
    )
    return monthly


def prepare_unit_price_band_matrix(
    history_df: pd.DataFrame,
    band_step: int = 4000,
    value_metric: str = "発行数",
) -> pd.DataFrame:
    """
    月度×単価帯の積み上げグラフ用データを返す。

    単価 = cost / cv（発生1件あたりのグロス単価）
    value_metric: 発行数 / 発生件数
    """
    if band_step <= 0:
        raise ValueError("単価帯の刻み幅は1円以上にしてください。")

    required = {"date", "月度", "cv", "cost", "approved_cv"}
    missing = required - set(history_df.columns)
    if missing:
        raise ValueError(
            "単価帯集計に必要な列がありません: "
            + ", ".join(sorted(missing))
        )

    metric_col = {
        "発行数": "approved_cv",
        "発生件数": "cv",
    }.get(value_metric)
    if metric_col is None:
        raise ValueError("単価帯の集計指標は『発行数』または『発生件数』を指定してください。")

    work = history_df[["date", "月度", "cv", "cost", "approved_cv"]].copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce")
    work["月度"] = _clean_month_label(work["月度"])
    for col in ["cv", "cost", "approved_cv"]:
        work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0.0)

    work = work[
        work["date"].notna()
        & work["月度"].notna()
        & work["月度"].ne("")
        & work["月度"].ne("未設定")
        & (work["cv"] > 0)
    ].copy()

    if work.empty:
        return pd.DataFrame()

    work["unit_price"] = work["cost"] / work["cv"]
    work = work[np.isfinite(work["unit_price"]) & (work["unit_price"] >= 0)].copy()
    if work.empty:
        return pd.DataFrame()

    work["band_lower"] = (
        np.floor(work["unit_price"] / band_step).astype(int) * band_step
    )
    work["band_upper"] = work["band_lower"] + band_step - 1
    work["単価帯"] = work.apply(
        lambda r: f"¥{int(r['band_lower']):,}–¥{int(r['band_upper']):,}", axis=1
    )

    month_order = (
        work.groupby("月度", as_index=False)
        .agg(月度開始日=("date", "min"))
        .sort_values(["月度開始日", "月度"], kind="stable")
    )
    month_labels = month_order["月度"].tolist()

    band_order = (
        work[["単価帯", "band_lower"]]
        .drop_duplicates()
        .sort_values("band_lower", kind="stable")["単価帯"]
        .tolist()
    )

    grouped = (
        work.groupby(["月度", "単価帯"], as_index=False)[metric_col]
        .sum()
        .rename(columns={metric_col: value_metric})
    )

    matrix = grouped.pivot(index="月度", columns="単価帯", values=value_metric).fillna(0.0)
    matrix = matrix.reindex(index=month_labels, columns=band_order, fill_value=0.0)
    matrix.index.name = "月度"
    return matrix
