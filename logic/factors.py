import numpy as np
import pandas as pd

from config.constants import (
    MAGITOKU_AFTER_FACTOR,
    PREMIUM_MEDIA_KEYWORDS,
    TRANSITION_NORMAL_DAYS,
)


def _safe_ratio(numerator: pd.Series, denominator: pd.Series | float) -> pd.Series:
    ratio = numerator / denominator
    return ratio.replace([np.inf, -np.inf], np.nan).fillna(1.0)


def _daily_media(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["date", "media", "cv"])
    return df.groupby(["date", "media"], as_index=False).agg(cv=("cv", "sum"))


def _ordinary_weekday_mask(df: pd.DataFrame) -> pd.Series:
    weekday_no = pd.to_datetime(df["date"]).dt.weekday
    return (
        weekday_no.lt(5)
        & df["is_month_start"].eq(0)
        & df["is_month_end"].eq(0)
    )


def calculate_transition_cpn_factor(
    history_df: pd.DataFrame,
    selected_cpn: str,
    reference_date: pd.Timestamp,
) -> pd.DataFrame:
    """前年同一CPNの直前定常→CPN切替率を媒体別に算出する。"""
    cpn_dates = (
        history_df.loc[history_df["CPN名"] == selected_cpn, "date"]
        .drop_duplicates()
        .sort_values()
    )
    columns = [
        "media", "normal_daily_cv", "cpn_daily_cv", "cpn_factor",
        "reference_cpn_start", "reference_cpn_end",
    ]
    if cpn_dates.empty:
        return pd.DataFrame(columns=columns)

    gaps = cpn_dates.diff().dt.days.fillna(1).gt(1)
    block_id = gaps.cumsum()
    periods = cpn_dates.groupby(block_id).agg(["min", "max"]).reset_index(drop=True)
    target = pd.Timestamp(reference_date) - pd.DateOffset(years=1)
    periods["distance"] = (periods["min"] - target).abs()
    chosen = periods.sort_values("distance").iloc[0]
    cpn_start = pd.Timestamp(chosen["min"])
    cpn_end = pd.Timestamp(chosen["max"])

    cpn_hist = history_df[
        history_df["date"].between(cpn_start, cpn_end)
        & history_df["CPN名"].eq(selected_cpn)
    ]
    normal_start = cpn_start - pd.Timedelta(days=TRANSITION_NORMAL_DAYS)
    normal_hist = history_df[
        history_df["date"].between(normal_start, cpn_start, inclusive="left")
        & history_df["CPN名"].eq("通常")
    ]

    def daily_average(df: pd.DataFrame) -> pd.Series:
        daily = _daily_media(df)
        return daily.groupby("media")["cv"].mean() if not daily.empty else pd.Series(dtype=float)

    normal_avg = daily_average(normal_hist)
    cpn_avg = daily_average(cpn_hist)
    media = sorted(set(normal_avg.index) | set(cpn_avg.index))
    result = pd.DataFrame({"media": media})
    result["normal_daily_cv"] = result["media"].map(normal_avg)
    result["cpn_daily_cv"] = result["media"].map(cpn_avg)
    result["cpn_factor"] = _safe_ratio(result["cpn_daily_cv"], result["normal_daily_cv"])
    result["reference_cpn_start"] = cpn_start
    result["reference_cpn_end"] = cpn_end
    return result[columns]


def calculate_dynamic_factor_tables(history_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    通常CPN実績から各係数を媒体別に算出する。
    基準は月初・月末を除く通常平日の媒体別1日平均CV=1.0。
    """
    normal = history_df[history_df["CPN名"].eq("通常")].copy()
    daily = _daily_media(normal)
    if daily.empty:
        empty = pd.DataFrame()
        return {"weekday": empty, "month_edge": empty, "season": empty, "line_oa": empty}

    day_attrs = normal.groupby(["date", "media"], as_index=False).agg(
        weekday=("weekday", "first"),
        is_month_start=("is_month_start", "max"),
        is_month_end=("is_month_end", "max"),
        line_oa_flag=("line_oa_flag", "max") if "line_oa_flag" in normal.columns else ("cv", lambda _: 0),
    )
    daily = daily.merge(day_attrs, on=["date", "media"], how="left")
    daily["month"] = daily["date"].dt.month

    ordinary = daily[_ordinary_weekday_mask(daily)]
    base_by_media = ordinary.groupby("media")["cv"].mean()
    global_base = ordinary["cv"].mean()
    if pd.isna(global_base) or global_base == 0:
        global_base = daily["cv"].mean()
    all_media = pd.Index(sorted(daily["media"].unique()), name="media")
    base_by_media = base_by_media.reindex(all_media).fillna(global_base).replace(0, np.nan)

    # 曜日係数
    weekday_avg = daily.groupby(["media", "weekday"])["cv"].mean().rename("actual_daily_cv").reset_index()
    weekday_avg["base_weekday_cv"] = weekday_avg["media"].map(base_by_media)
    weekday_avg["factor"] = _safe_ratio(weekday_avg["actual_daily_cv"], weekday_avg["base_weekday_cv"])

    # 月初・月末4営業日係数
    edge_rows = []
    for label, flag in [("月初4営業日", "is_month_start"), ("月末4営業日", "is_month_end")]:
        avg = daily[daily[flag].eq(1)].groupby("media")["cv"].mean()
        tmp = pd.DataFrame({"media": all_media})
        tmp["区分"] = label
        tmp["actual_daily_cv"] = tmp["media"].map(avg)
        tmp["base_weekday_cv"] = tmp["media"].map(base_by_media)
        tmp["factor"] = _safe_ratio(tmp["actual_daily_cv"], tmp["base_weekday_cv"])
        edge_rows.append(tmp)
    month_edge = pd.concat(edge_rows, ignore_index=True)

    # 需要期係数（月別）
    season_avg = daily.groupby(["media", "month"])["cv"].mean().rename("actual_daily_cv").reset_index()
    season_avg["base_weekday_cv"] = season_avg["media"].map(base_by_media)
    season_avg["factor"] = _safe_ratio(season_avg["actual_daily_cv"], season_avg["base_weekday_cv"])

    # LINE OA係数。履歴に配信日がない場合は1.0。
    line_hist = daily[daily["media"].str.contains("LINE", case=False, na=False)]
    line_rows = []
    for media, group in line_hist.groupby("media"):
        oa = group[group["line_oa_flag"].eq(1)]["cv"].mean()
        non_oa = group[group["line_oa_flag"].eq(0)]["cv"].mean()
        factor = 1.0 if pd.isna(oa) or pd.isna(non_oa) or non_oa == 0 else oa / non_oa
        line_rows.append({"media": media, "oa_daily_cv": oa, "non_oa_daily_cv": non_oa, "factor": factor})
    line_oa = pd.DataFrame(line_rows)

    return {
        "weekday": weekday_avg,
        "month_edge": month_edge,
        "season": season_avg,
        "line_oa": line_oa,
    }


def apply_dynamic_factors(
    df: pd.DataFrame,
    factor_tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    out = df.copy()

    weekday_map = factor_tables["weekday"].set_index(["media", "weekday"])["factor"] if not factor_tables["weekday"].empty else pd.Series(dtype=float)
    season_map = factor_tables["season"].set_index(["media", "month"])["factor"] if not factor_tables["season"].empty else pd.Series(dtype=float)
    edge_map = factor_tables["month_edge"].set_index(["media", "区分"])["factor"] if not factor_tables["month_edge"].empty else pd.Series(dtype=float)
    line_map = factor_tables["line_oa"].set_index("media")["factor"] if not factor_tables["line_oa"].empty else pd.Series(dtype=float)

    out["weekday_factor"] = [weekday_map.get((m, w), 1.0) for m, w in zip(out["media"], out["weekday"])]
    out["season_factor"] = [season_map.get((m, month), 1.0) for m, month in zip(out["media"], out["date"].dt.month)]
    out["month_edge_factor"] = 1.0
    start_mask = out["is_month_start"].eq(1)
    end_mask = out["is_month_end"].eq(1)
    out.loc[start_mask, "month_edge_factor"] = [edge_map.get((m, "月初4営業日"), 1.0) for m in out.loc[start_mask, "media"]]
    out.loc[end_mask, "month_edge_factor"] *= [edge_map.get((m, "月末4営業日"), 1.0) for m in out.loc[end_mask, "media"]]

    out["line_factor"] = 1.0
    line_mask = out["media"].str.contains("LINE", case=False, na=False) & out["line_oa_flag"].eq(1)
    out.loc[line_mask, "line_factor"] = out.loc[line_mask, "media"].map(line_map).fillna(1.0)

    out["after_factor"] = 1.0
    out.loc[out["magitoku_after_flag"].eq(1), "after_factor"] = MAGITOKU_AFTER_FACTOR
    return out


def enforce_premium_media_cost(df: pd.DataFrame) -> pd.DataFrame:
    """
    ハピタス・モッピー・LINEが中小媒体より高単価になるよう調整する。
    固定倍率は使わず、中小媒体の実績上限+1円を最低ラインにする。
    """
    out = df.copy()
    is_premium = out["media"].apply(
        lambda x: any(k.lower() in str(x).lower() for k in PREMIUM_MEDIA_KEYWORDS)
    )
    small_media_cost = out.loc[~is_premium].groupby("media")["cost"].mean()
    if small_media_cost.empty:
        return out
    floor = float(small_media_cost.max()) + 1.0
    out.loc[is_premium, "cost"] = out.loc[is_premium, "cost"].clip(lower=floor)
    return out
