import numpy as np
import pandas as pd

from config.constants import (
    MAGITOKU_AFTER_FACTOR,
    PREMIUM_MEDIA_KEYWORDS,
    RECENCY_MONTH_DECAY,
    OUTLIER_MIN_DAYS,
    OUTLIER_MAD_Z_THRESHOLD,
    FACTOR_PRIOR_DAYS,
)


def _safe_ratio(numerator: pd.Series, denominator: pd.Series | float) -> pd.Series:
    ratio = numerator / denominator
    return ratio.replace([np.inf, -np.inf], np.nan).fillna(1.0)




def _shrink_factor(raw_factor: float, sample_days: int, prior_days: int = FACTOR_PRIOR_DAYS) -> float:
    """少数実績の係数を1.0側へ縮め、偶然の振れを過信しない。"""
    if pd.isna(raw_factor) or not np.isfinite(raw_factor) or raw_factor <= 0:
        return 1.0
    n = max(int(sample_days), 0)
    confidence = n / float(n + max(int(prior_days), 1))
    return 1.0 + confidence * (float(raw_factor) - 1.0)

def _weighted_mean(values: pd.Series, weights: pd.Series) -> float:
    values = pd.to_numeric(values, errors="coerce")
    weights = pd.to_numeric(weights, errors="coerce")
    valid = values.notna() & weights.notna() & weights.gt(0)
    if not valid.any():
        return np.nan
    return float(np.average(values[valid], weights=weights[valid]))


def _daily_media(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["date", "media", "cv", "cost"])
    return df.groupby(["date", "media"], as_index=False).agg(
        cv=("cv", "sum"),
        cost=("cost", "sum"),
    )


def _ordinary_weekday_mask(df: pd.DataFrame) -> pd.Series:
    weekday_no = pd.to_datetime(df["date"]).dt.weekday
    return (
        weekday_no.lt(5)
        & df["is_month_start"].eq(0)
        & df["is_month_end"].eq(0)
        & df.get("line_oa_flag", pd.Series(0, index=df.index)).eq(0)
    )


def _normalize_selected_months(selected_months: list[str] | None) -> list[str]:
    if not selected_months:
        return []
    return [str(x).strip() for x in selected_months if str(x).strip()]


def _add_recency_weights(df: pd.DataFrame) -> pd.DataFrame:
    """月度ではなく実日付順で、直近月ほど大きい学習ウェイトを付与する。

    最新の実績月=1.0、1か月古いごとにRECENCY_MONTH_DECAY倍。
    同じ月度内の日は同じ重みとし、説明可能性を優先する。
    """
    out = df.copy()
    if out.empty:
        out["recency_weight"] = pd.Series(dtype=float)
        return out

    dates = pd.to_datetime(out["date"], errors="coerce")
    month_period = dates.dt.to_period("M")
    latest = month_period.max()
    if pd.isna(latest):
        out["recency_weight"] = 1.0
        return out

    month_age = (latest.year - month_period.dt.year) * 12 + (latest.month - month_period.dt.month)
    out["recency_weight"] = np.power(float(RECENCY_MONTH_DECAY), month_age.astype(float))
    return out


def _detect_media_daily_outliers(daily: pd.DataFrame) -> pd.DataFrame:
    """媒体別の日次CV異常値を頑健に検出する。

    CV分布の歪みを抑えるためlog1p(CV)に対してMADベースのrobust z-scoreを使う。
    サンプル不足の媒体やMAD=0の媒体は自動除外しない。
    """
    if daily.empty:
        return pd.DataFrame(columns=["date", "media", "cv", "outlier_score", "reason"])

    rows: list[pd.DataFrame] = []
    for media, group in daily.groupby("media", sort=False):
        g = group.copy()
        if len(g) < int(OUTLIER_MIN_DAYS):
            continue

        x = np.log1p(pd.to_numeric(g["cv"], errors="coerce").fillna(0.0).clip(lower=0.0))
        median = float(x.median())
        mad = float((x - median).abs().median())
        if np.isfinite(mad) and mad > 1e-12:
            # 0.6745 * deviation / MAD は正規分布のz-score相当にスケールされる。
            score = 0.6745 * (x - median) / mad
            mask = score.abs().gt(float(OUTLIER_MAD_Z_THRESHOLD))
        else:
            # 同値が多い媒体ではMAD=0になり得るため、IQRへフォールバック。
            q1 = float(x.quantile(0.25))
            q3 = float(x.quantile(0.75))
            iqr = q3 - q1
            if not np.isfinite(iqr) or iqr <= 1e-12:
                continue
            lower = q1 - 1.5 * iqr
            upper = q3 + 1.5 * iqr
            mask = x.lt(lower) | x.gt(upper)
            # 監査表示用。IQR幅を1単位とした符号付き距離。
            score = pd.Series(0.0, index=x.index)
            score.loc[x.gt(upper)] = (x.loc[x.gt(upper)] - upper) / iqr
            score.loc[x.lt(lower)] = (x.loc[x.lt(lower)] - lower) / iqr

        if not mask.any():
            continue

        flagged = g.loc[mask, ["date", "media", "cv"]].copy()
        flagged["outlier_score"] = score.loc[mask].to_numpy()
        flagged["reason"] = np.where(flagged["outlier_score"].gt(0), "日次CVの異常高値", "日次CVの異常低値")
        rows.append(flagged)

    if not rows:
        return pd.DataFrame(columns=["date", "media", "cv", "outlier_score", "reason"])
    return pd.concat(rows, ignore_index=True).sort_values(["date", "media"], kind="stable")


def _prepare_normal_learning_data(
    history_df: pd.DataFrame,
    selected_months: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """定常学習用データを作成し、マジ得直後・異常値を除外する。

    Returns:
        clean_rows: 元粒度（媒体×商品ID等）の学習対象行
        clean_daily: 媒体×日の日次集計（係数算出用）
        excluded: 除外監査用の日次一覧
    """
    work = history_df.copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce").dt.normalize()
    work["CPN名"] = work["CPN名"].astype("string").str.strip()
    if "月度" in work.columns:
        work["月度"] = work["月度"].astype("string").str.strip()

    normal = work[work["CPN名"].isin(["通常", "定常"])].copy()
    months = _normalize_selected_months(selected_months)
    if months and "月度" in normal.columns:
        normal = normal[normal["月度"].isin(months)].copy()

    if normal.empty:
        empty_daily = pd.DataFrame(columns=["date", "media", "cv", "cost"])
        empty_excluded = pd.DataFrame(columns=["date", "media", "cv", "reason", "outlier_score"])
        return normal, empty_daily, empty_excluded

    excluded_parts: list[pd.DataFrame] = []

    # マジ得直後は需要先食い等の特殊期間として定常平均から完全除外。
    if "magitoku_after_flag" in normal.columns:
        after_mask = pd.to_numeric(normal["magitoku_after_flag"], errors="coerce").fillna(0).eq(1)
        after_daily = _daily_media(normal.loc[after_mask])
        if not after_daily.empty:
            after_daily["reason"] = "マジ得直後"
            after_daily["outlier_score"] = np.nan
            excluded_parts.append(after_daily[["date", "media", "cv", "reason", "outlier_score"]])
        normal = normal.loc[~after_mask].copy()

    if normal.empty:
        excluded = pd.concat(excluded_parts, ignore_index=True) if excluded_parts else pd.DataFrame()
        return normal, _daily_media(normal), excluded

    daily = _daily_media(normal)
    outliers = _detect_media_daily_outliers(daily)
    if not outliers.empty:
        excluded_parts.append(outliers[["date", "media", "cv", "reason", "outlier_score"]])
        bad_keys = pd.MultiIndex.from_frame(outliers[["date", "media"]])
        row_keys = pd.MultiIndex.from_frame(normal[["date", "media"]])
        normal = normal.loc[~row_keys.isin(bad_keys)].copy()

    normal = _add_recency_weights(normal)
    clean_daily = _daily_media(normal)

    # 日次属性は元粒度から復元する。
    if not clean_daily.empty:
        attrs_spec = {
            "weekday": ("weekday", "first"),
            "is_month_start": ("is_month_start", "max"),
            "is_month_end": ("is_month_end", "max"),
            "recency_weight": ("recency_weight", "first"),
        }
        if "line_oa_flag" in normal.columns:
            attrs_spec["line_oa_flag"] = ("line_oa_flag", "max")
        day_attrs = normal.groupby(["date", "media"], as_index=False).agg(**attrs_spec)
        clean_daily = clean_daily.merge(day_attrs, on=["date", "media"], how="left")
        if "line_oa_flag" not in clean_daily.columns:
            clean_daily["line_oa_flag"] = 0
        clean_daily["month"] = clean_daily["date"].dt.month

    excluded = (
        pd.concat(excluded_parts, ignore_index=True)
        .drop_duplicates(subset=["date", "media", "reason"])
        .sort_values(["date", "media", "reason"], kind="stable")
        if excluded_parts
        else pd.DataFrame(columns=["date", "media", "cv", "reason", "outlier_score"])
    )
    return normal, clean_daily, excluded


def get_cpn_reference_periods(
    history_df: pd.DataFrame,
    selected_cpn: str,
) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """実績内に存在する指定CPNの日付を、連続したCPN期間ごとに返す。"""
    work = history_df.copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce").dt.normalize()
    dates = (
        work.loc[work["CPN名"].eq(selected_cpn), "date"]
        .dropna()
        .drop_duplicates()
        .sort_values()
    )
    if dates.empty:
        return []

    block_id = dates.diff().dt.days.fillna(1).gt(1).cumsum()
    periods = dates.groupby(block_id).agg(["min", "max"])
    return [(pd.Timestamp(row["min"]), pd.Timestamp(row["max"])) for _, row in periods.iterrows()]


def calculate_selected_cpn_base(
    history_df: pd.DataFrame,
    selected_cpn: str,
    selected_periods: list[tuple[pd.Timestamp, pd.Timestamp]],
) -> pd.DataFrame:
    """選択した複数CPN期間の合計実績を総日数で割り、媒体×商品ID別の日平均を返す。"""
    columns = ["media", "商品ID", "base_cv", "cost"]
    if not selected_periods:
        return pd.DataFrame(columns=columns)

    work = history_df.copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce").dt.normalize()

    mask = pd.Series(False, index=work.index)
    total_days = 0
    for start, end in selected_periods:
        start = pd.Timestamp(start).normalize()
        end = pd.Timestamp(end).normalize()
        if end < start:
            continue
        mask |= work["date"].between(start, end)
        total_days += (end - start).days + 1

    selected = work[mask & work["CPN名"].eq(selected_cpn)].copy()
    if selected.empty or total_days <= 0:
        return pd.DataFrame(columns=columns)

    result = (
        selected.groupby(["media", "商品ID"], as_index=False)
        .agg(total_cv=("cv", "sum"), total_cost=("cost", "sum"))
    )
    result["base_cv"] = result["total_cv"] / total_days
    result["cost"] = result["total_cost"] / total_days
    return result[columns]


def calculate_dynamic_factor_tables(
    history_df: pd.DataFrame,
    selected_months: list[str] | None = None,
) -> dict[str, pd.DataFrame]:
    """定常実績から媒体別係数を算出する。

    - 学習月度を基礎値と揃える
    - マジ得直後・媒体別異常日を除外
    - 直近月ほど重く評価
    - 曜日係数は媒体別かつ月水準を正規化して算出し、季節との二重取りを抑える
    """
    normal, daily, excluded = _prepare_normal_learning_data(history_df, selected_months)
    if daily.empty:
        empty = pd.DataFrame()
        return {"weekday": empty, "month_edge": empty, "season": empty, "line_oa": empty, "excluded": excluded}

    # 媒体ごとの全体基準CV。基礎値と同じく全定常日を直近加重で評価する。
    base_by_media = daily.groupby("media")[["cv", "recency_weight"]].apply(
        lambda g: _weighted_mean(g["cv"], g["recency_weight"])
    )
    global_base = _weighted_mean(daily["cv"], daily["recency_weight"])
    all_media = pd.Index(sorted(daily["media"].unique()), name="media")
    base_by_media = base_by_media.reindex(all_media).fillna(global_base).replace(0, np.nan)

    # 月ごとの媒体水準を先に除去し、曜日と需要期の同じ変動を二重計上しにくくする。
    media_month_base = daily.groupby(["media", "month"])[["cv", "recency_weight"]].apply(
        lambda g: _weighted_mean(g["cv"], g["recency_weight"])
    )
    daily["media_month_base"] = [media_month_base.get((m, mon), np.nan) for m, mon in zip(daily["media"], daily["month"])]
    daily["weekday_residual"] = _safe_ratio(daily["cv"], daily["media_month_base"])

    weekday_rows = []
    for (media, weekday), group in daily.groupby(["media", "weekday"]):
        raw_factor = _weighted_mean(group["weekday_residual"], group["recency_weight"])
        sample_days = int(group["date"].nunique())
        factor = _shrink_factor(raw_factor, sample_days)
        weekday_rows.append({
            "media": media,
            "weekday": weekday,
            "actual_daily_cv": _weighted_mean(group["cv"], group["recency_weight"]),
            "base_weekday_cv": base_by_media.get(media, np.nan),
            "raw_factor": raw_factor,
            "factor": factor,
            "sample_days": sample_days,
        })
    weekday_avg = pd.DataFrame(weekday_rows)

    # 月初・月末は曜日影響を除去した残差から算出。
    weekday_map = weekday_avg.set_index(["media", "weekday"])["factor"] if not weekday_avg.empty else pd.Series(dtype=float)
    daily["weekday_factor_hist"] = [weekday_map.get((m, w), 1.0) for m, w in zip(daily["media"], daily["weekday"])]
    daily["edge_residual"] = _safe_ratio(daily["cv"], daily["media_month_base"] * daily["weekday_factor_hist"])

    edge_rows = []
    for label, flag in [("月初4営業日", "is_month_start"), ("月末4営業日", "is_month_end")]:
        for media in all_media:
            group = daily[(daily["media"].eq(media)) & (daily[flag].eq(1))]
            raw_factor = _weighted_mean(group["edge_residual"], group["recency_weight"]) if not group.empty else np.nan
            sample_days = int(group["date"].nunique()) if not group.empty else 0
            factor = _shrink_factor(raw_factor, sample_days)
            edge_rows.append({
                "media": media,
                "区分": label,
                "actual_daily_cv": _weighted_mean(group["cv"], group["recency_weight"]) if not group.empty else np.nan,
                "base_weekday_cv": base_by_media.get(media, np.nan),
                "raw_factor": raw_factor,
                "factor": factor,
                "sample_days": sample_days,
            })
    month_edge = pd.DataFrame(edge_rows)

    # 需要期係数は、曜日・月初月末を除去したCVから算出する。
    # これにより「9月の土日が弱い」を曜日と9月の両方へ重複計上しにくくする。
    edge_map_hist = month_edge.set_index(["media", "区分"])["factor"] if not month_edge.empty else pd.Series(dtype=float)
    daily["edge_factor_hist"] = 1.0
    start_mask_hist = daily["is_month_start"].eq(1)
    end_mask_hist = daily["is_month_end"].eq(1)
    daily.loc[start_mask_hist, "edge_factor_hist"] = [
        edge_map_hist.get((m, "月初4営業日"), 1.0) for m in daily.loc[start_mask_hist, "media"]
    ]
    daily.loc[end_mask_hist, "edge_factor_hist"] *= [
        edge_map_hist.get((m, "月末4営業日"), 1.0) for m in daily.loc[end_mask_hist, "media"]
    ]
    daily["calendar_adjusted_cv"] = _safe_ratio(
        daily["cv"],
        daily["weekday_factor_hist"] * daily["edge_factor_hist"],
    )
    adjusted_base_by_media = daily.groupby("media")[["calendar_adjusted_cv", "recency_weight"]].apply(
        lambda g: _weighted_mean(g["calendar_adjusted_cv"], g["recency_weight"])
    )

    season_rows = []
    for (media, month), group in daily.groupby(["media", "month"]):
        actual = _weighted_mean(group["calendar_adjusted_cv"], group["recency_weight"])
        base = adjusted_base_by_media.get(media, np.nan)
        raw_factor = 1.0 if pd.isna(base) or base == 0 or pd.isna(actual) else actual / base
        sample_days = int(group["date"].nunique())
        factor = _shrink_factor(raw_factor, sample_days)
        season_rows.append({
            "media": media,
            "month": month,
            "actual_daily_cv": actual,
            "base_weekday_cv": base,
            "raw_factor": raw_factor,
            "factor": factor,
            "sample_days": sample_days,
        })
    season_avg = pd.DataFrame(season_rows)

    # LINE OAは曜日・月初月末・需要期を除去した残差で算出する。
    season_map_hist = season_avg.set_index(["media", "month"])["factor"] if not season_avg.empty else pd.Series(dtype=float)
    daily["season_factor_hist"] = [
        season_map_hist.get((m, mon), 1.0) for m, mon in zip(daily["media"], daily["month"])
    ]
    line_rows = []
    line_hist = daily[daily["media"].str.contains("LINE", case=False, na=False)].copy()
    line_hist["normalized_cv"] = _safe_ratio(
        line_hist["cv"],
        line_hist["weekday_factor_hist"]
        * line_hist["edge_factor_hist"]
        * line_hist["season_factor_hist"],
    )
    for media, group in line_hist.groupby("media"):
        oa_group = group[group["line_oa_flag"].eq(1)]
        non_group = group[group["line_oa_flag"].eq(0)]
        oa = _weighted_mean(oa_group["normalized_cv"], oa_group["recency_weight"]) if not oa_group.empty else np.nan
        non_oa = _weighted_mean(non_group["normalized_cv"], non_group["recency_weight"]) if not non_group.empty else np.nan
        raw_factor = 1.0 if pd.isna(oa) or pd.isna(non_oa) or non_oa == 0 else oa / non_oa
        oa_sample_days = int(oa_group["date"].nunique())
        factor = _shrink_factor(raw_factor, oa_sample_days)
        line_rows.append({
            "media": media,
            "oa_daily_cv": _weighted_mean(oa_group["cv"], oa_group["recency_weight"]) if not oa_group.empty else np.nan,
            "non_oa_daily_cv": _weighted_mean(non_group["cv"], non_group["recency_weight"]) if not non_group.empty else np.nan,
            "raw_factor": raw_factor,
            "factor": factor,
            "oa_sample_days": oa_sample_days,
        })
    line_oa = pd.DataFrame(line_rows)

    return {
        "weekday": weekday_avg,
        "month_edge": month_edge,
        "season": season_avg,
        "line_oa": line_oa,
        "excluded": excluded,
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
    """ハピタス・モッピー・LINEが中小媒体より高単価になるよう調整する。"""
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


def calculate_normal_month_base(
    history_df: pd.DataFrame,
    selected_months: list[str],
) -> pd.DataFrame:
    """選択月度の定常実績から媒体×商品IDの加重日平均を返す。

    改修後仕様:
    - 直近月ほど重く評価
    - マジ得直後を除外
    - 媒体別の日次異常値を除外
    - 除外は媒体単位なので、他媒体の正常日は巻き込まない
    """
    columns = ["media", "商品ID", "base_cv", "cost"]
    if not selected_months:
        return pd.DataFrame(columns=columns)

    required = {"date", "media", "商品ID", "月度", "CPN名", "cv", "cost"}
    missing = required - set(history_df.columns)
    if missing:
        raise ValueError("定常月度学習に必要な列がありません: " + ", ".join(sorted(missing)))

    clean, _, _ = _prepare_normal_learning_data(history_df, selected_months)
    if clean.empty:
        return pd.DataFrame(columns=columns)

    # 媒体ごとに有効な学習日と重みを分母にする。
    # 商品IDの行がない日は0件として既存仕様を維持する。
    media_day_weights = (
        clean[["media", "date", "recency_weight"]]
        .drop_duplicates(["media", "date"])
        .groupby("media")["recency_weight"]
        .sum()
    )

    weighted = clean.copy()
    weighted["weighted_cv"] = pd.to_numeric(weighted["cv"], errors="coerce").fillna(0.0) * weighted["recency_weight"]
    weighted["weighted_cost"] = pd.to_numeric(weighted["cost"], errors="coerce").fillna(0.0) * weighted["recency_weight"]
    result = (
        weighted.groupby(["media", "商品ID"], as_index=False)
        .agg(weighted_cv=("weighted_cv", "sum"), weighted_cost=("weighted_cost", "sum"))
    )
    result["weight_days"] = result["media"].map(media_day_weights)
    result["base_cv"] = result["weighted_cv"] / result["weight_days"]
    result["cost"] = result["weighted_cost"] / result["weight_days"]
    return result[columns]
