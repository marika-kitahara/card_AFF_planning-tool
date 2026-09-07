import numpy as np
import pandas as pd

from config.constants import (
    MAGITOKU_AFTER_FACTOR,
    PREMIUM_MEDIA_KEYWORDS,
    RECENCY_MONTH_DECAY,
    OUTLIER_MIN_DAYS,
    OUTLIER_MAD_Z_THRESHOLD,
    FACTOR_PRIOR_DAYS,
    INACTIVE_MEDIA_LOOKBACK_DAYS,
    UNIT_PRICE_MIN_MONTHS,
    UNIT_PRICE_PRIOR_MONTHS,
    UNIT_PRICE_ELASTICITY_MIN,
    UNIT_PRICE_ELASTICITY_MAX,
    UNIT_PRICE_FACTOR_MIN,
    UNIT_PRICE_FACTOR_MAX,
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

    # 重要: ローデータに媒体行が存在しない日も、その媒体の0CV日として日次学習へ含める。
    # 従来は「CVが出た/行が存在した日」だけが分母になり、断続掲載媒体の基礎CVを
    # 過大評価していた。ここでは学習対象の定常日 × 媒体を展開し、欠損を0CVで補完する。
    observed_daily = _daily_media(normal)
    if not observed_daily.empty:
        media_list = pd.DataFrame({"media": sorted(normal["media"].dropna().astype(str).unique())})
        calendar = pd.DataFrame({"date": sorted(normal["date"].dropna().unique())})
        calendar = _add_recency_weights(calendar)
        calendar["weekday"] = calendar["date"].dt.day_name()

        # 月初/月末・LINE OAは日付共通属性として、元データから復元する。
        date_attrs_spec = {}
        if "is_month_start" in normal.columns:
            date_attrs_spec["is_month_start"] = ("is_month_start", "max")
        if "is_month_end" in normal.columns:
            date_attrs_spec["is_month_end"] = ("is_month_end", "max")
        if "line_oa_flag" in normal.columns:
            date_attrs_spec["line_oa_flag"] = ("line_oa_flag", "max")
        if date_attrs_spec:
            date_attrs = normal.groupby("date", as_index=False).agg(**date_attrs_spec)
            calendar = calendar.merge(date_attrs, on="date", how="left")
        for col in ["is_month_start", "is_month_end", "line_oa_flag"]:
            if col not in calendar.columns:
                calendar[col] = 0
            calendar[col] = pd.to_numeric(calendar[col], errors="coerce").fillna(0).astype(int)

        clean_daily = media_list.merge(calendar, how="cross").merge(
            observed_daily, on=["date", "media"], how="left"
        )
        clean_daily["cv"] = pd.to_numeric(clean_daily["cv"], errors="coerce").fillna(0.0)
        clean_daily["cost"] = pd.to_numeric(clean_daily["cost"], errors="coerce").fillna(0.0)

        # 異常値として除外した媒体×日は、0CVに置換せず学習母集団そのものから外す。
        if not outliers.empty:
            bad_keys = pd.MultiIndex.from_frame(outliers[["date", "media"]])
            daily_keys = pd.MultiIndex.from_frame(clean_daily[["date", "media"]])
            clean_daily = clean_daily.loc[~daily_keys.isin(bad_keys)].copy()
        clean_daily["month"] = clean_daily["date"].dt.month
    else:
        clean_daily = observed_daily

    excluded = (
        pd.concat(excluded_parts, ignore_index=True)
        .drop_duplicates(subset=["date", "media", "reason"])
        .sort_values(["date", "media", "reason"], kind="stable")
        if excluded_parts
        else pd.DataFrame(columns=["date", "media", "cv", "reason", "outlier_score"])
    )
    return normal, clean_daily, excluded



def identify_inactive_media(
    history_df: pd.DataFrame,
    selected_months: list[str] | None = None,
    lookback_days: int = INACTIVE_MEDIA_LOOKBACK_DAYS,
) -> pd.DataFrame:
    """定常学習期間内で、直近N日CV=0の媒体を休眠扱いとして返す。

    過去にCV>0の実績がある媒体だけを対象にするため、単なる未稼働媒体は含めない。
    判定基準日は、渡されたデータ内の最新定常日。バックテストでは未来月を
    切り落としたtrainingが渡されるため、その時点の状態だけで判定できる。
    """
    columns = ["media", "last_positive_date", "recent_cv", "lookback_days", "reason"]
    if history_df.empty:
        return pd.DataFrame(columns=columns)

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
        return pd.DataFrame(columns=columns)

    normal["cv"] = pd.to_numeric(normal["cv"], errors="coerce").fillna(0.0)
    ref_date = normal["date"].max()
    if pd.isna(ref_date):
        return pd.DataFrame(columns=columns)

    lookback_days = max(int(lookback_days), 1)
    recent_start = pd.Timestamp(ref_date).normalize() - pd.Timedelta(days=lookback_days - 1)
    daily = _daily_media(normal)
    historical_positive = daily[daily["cv"].gt(0)].groupby("media")["date"].max()
    recent_cv = (
        daily[daily["date"].ge(recent_start)]
        .groupby("media")["cv"]
        .sum()
    )

    rows = []
    for media, last_date in historical_positive.items():
        cv30 = float(recent_cv.get(media, 0.0))
        if cv30 <= 0:
            rows.append({
                "media": media,
                "last_positive_date": pd.Timestamp(last_date).normalize(),
                "recent_cv": cv30,
                "lookback_days": lookback_days,
                "reason": f"直近{lookback_days}日CVなし",
            })
    return pd.DataFrame(rows, columns=columns).sort_values(
        ["last_positive_date", "media"], kind="stable"
    ) if rows else pd.DataFrame(columns=columns)


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



def calculate_unit_price_response_table(
    history_df: pd.DataFrame,
    selected_months: list[str] | None = None,
) -> pd.DataFrame:
    """媒体別に、1件あたり単価と日平均CVの関係を月度単位で推定する。

    COST合計はCV増加に伴って増えるため、そのまま説明変数にはしない。
    月度ごとの unit_price = total_cost / total_cv を使い、
    log(日平均CV) ~ elasticity * log(unit_price) の傾きを安全に縮小して返す。
    データ不足・単価変動不足では elasticity=0（補正なし）。
    """
    columns = [
        "media", "reference_unit_price", "latest_unit_price", "elasticity_raw",
        "elasticity", "sample_months", "price_variation", "fit_r2",
    ]
    clean, daily, _ = _prepare_normal_learning_data(history_df, selected_months)
    if clean.empty or daily.empty:
        return pd.DataFrame(columns=columns)

    # 媒体×月度のCV/COSTを元粒度から集計。日平均CVの分母は媒体行がない日も含む全定常日。
    work = clean.copy()
    work["cv"] = pd.to_numeric(work["cv"], errors="coerce").fillna(0.0)
    work["cost"] = pd.to_numeric(work["cost"], errors="coerce").fillna(0.0)
    work["month_key"] = work["月度"].astype("string").str.strip()
    monthly = work.groupby(["media", "month_key"], as_index=False).agg(
        total_cv=("cv", "sum"), total_cost=("cost", "sum"), last_date=("date", "max")
    )
    eligible = daily.copy()
    if "月度" not in eligible.columns:
        # dailyには月度がない旧データ互換。dateからcleanの月度を戻す。
        date_month = clean[["date", "月度"]].drop_duplicates("date")
        eligible = eligible.merge(date_month, on="date", how="left")
    eligible["month_key"] = eligible["月度"].astype("string").str.strip()
    day_counts = eligible.groupby(["media", "month_key"])["date"].nunique().rename("eligible_days")
    monthly = monthly.merge(day_counts.reset_index(), on=["media", "month_key"], how="left")
    monthly["eligible_days"] = pd.to_numeric(monthly["eligible_days"], errors="coerce").fillna(0).clip(lower=1)
    monthly["daily_cv"] = monthly["total_cv"] / monthly["eligible_days"]
    monthly["unit_price"] = monthly["total_cost"] / monthly["total_cv"].replace(0, np.nan)
    monthly = monthly.replace([np.inf, -np.inf], np.nan)
    monthly = monthly.loc[
        monthly["daily_cv"].gt(0) & monthly["unit_price"].gt(0)
    ].copy()
    if monthly.empty:
        return pd.DataFrame(columns=columns)

    # 選択月度の新しい月ほど重くするため、last_dateから月順位を作る。
    max_period = monthly["last_date"].dt.to_period("M").max()
    month_period = monthly["last_date"].dt.to_period("M")
    monthly["month_age"] = month_period.map(lambda x: int(max_period.ordinal - x.ordinal) if pd.notna(x) else 0)
    monthly["w"] = RECENCY_MONTH_DECAY ** monthly["month_age"].clip(lower=0)

    rows = []
    for media, g in monthly.groupby("media", sort=False):
        g = g.sort_values("last_date")
        n = int(len(g))
        ref_price = _weighted_mean(g["unit_price"], g["w"])
        latest_price = float(g.iloc[-1]["unit_price"]) if n else np.nan
        raw = 0.0
        shrunk = 0.0
        fit_r2 = 0.0
        variation = 0.0
        if n >= UNIT_PRICE_MIN_MONTHS:
            x = np.log(pd.to_numeric(g["unit_price"], errors="coerce").to_numpy(dtype=float))
            y = np.log(pd.to_numeric(g["daily_cv"], errors="coerce").to_numpy(dtype=float))
            w = pd.to_numeric(g["w"], errors="coerce").fillna(0).to_numpy(dtype=float)
            valid = np.isfinite(x) & np.isfinite(y) & np.isfinite(w) & (w > 0)
            x, y, w = x[valid], y[valid], w[valid]
            if len(x) >= UNIT_PRICE_MIN_MONTHS and np.ptp(x) >= 0.05:
                variation = float(np.exp(np.max(x) - np.min(x)) - 1.0)
                wx = np.average(x, weights=w)
                wy = np.average(y, weights=w)
                varx = np.average((x - wx) ** 2, weights=w)
                if varx > 1e-8:
                    raw = float(np.average((x - wx) * (y - wy), weights=w) / varx)
                    raw = float(np.clip(raw, UNIT_PRICE_ELASTICITY_MIN, UNIT_PRICE_ELASTICITY_MAX))
                    yhat = wy + raw * (x - wx)
                    sst = float(np.sum(w * (y - wy) ** 2))
                    sse = float(np.sum(w * (y - yhat) ** 2))
                    fit_r2 = max(0.0, min(1.0, 1.0 - sse / sst)) if sst > 1e-8 else 0.0
                    # 月数が少ない・説明力が弱いほど0（補正なし）へ強く縮小。
                    confidence = n / float(n + UNIT_PRICE_PRIOR_MONTHS)
                    shrunk = raw * confidence * fit_r2
        rows.append({
            "media": media,
            "reference_unit_price": float(ref_price) if pd.notna(ref_price) else np.nan,
            "latest_unit_price": latest_price,
            "elasticity_raw": raw,
            "elasticity": float(shrunk),
            "sample_months": n,
            "price_variation": variation,
            "fit_r2": fit_r2,
        })
    return pd.DataFrame(rows, columns=columns)

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
        return {"weekday": empty, "month_edge": empty, "season": empty, "line_oa": empty, "unit_price": empty, "excluded": excluded}

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

    inactive_media = identify_inactive_media(history_df, selected_months)
    unit_price = calculate_unit_price_response_table(history_df, selected_months)

    return {
        "weekday": weekday_avg,
        "month_edge": month_edge,
        "season": season_avg,
        "line_oa": line_oa,
        "excluded": excluded,
        "inactive_media": inactive_media,
        "unit_price": unit_price,
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
    unit_tbl = factor_tables.get("unit_price", pd.DataFrame())
    unit_ref_map = unit_tbl.set_index("media")["reference_unit_price"] if not unit_tbl.empty else pd.Series(dtype=float)
    unit_latest_map = unit_tbl.set_index("media")["latest_unit_price"] if not unit_tbl.empty else pd.Series(dtype=float)
    unit_elasticity_map = unit_tbl.set_index("media")["elasticity"] if not unit_tbl.empty else pd.Series(dtype=float)

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

    # 単価補正: 指定単価があればそれを、なければ学習期間の直近単価を使用。
    # 基準単価は学習期間の直近加重平均。データ不足ではelasticity=0なので必ず1.0。
    out["unit_price_reference"] = out["media"].map(unit_ref_map)
    out["unit_price_latest"] = out["media"].map(unit_latest_map)
    out["unit_price_elasticity"] = out["media"].map(unit_elasticity_map).fillna(0.0)
    if "planned_unit_price" in out.columns:
        planned = pd.to_numeric(out["planned_unit_price"], errors="coerce")
    else:
        planned = pd.Series(np.nan, index=out.index, dtype=float)
    out["unit_price_planned"] = planned.fillna(out["unit_price_latest"]).fillna(out["unit_price_reference"])
    ratio = out["unit_price_planned"] / out["unit_price_reference"].replace(0, np.nan)
    ratio = ratio.replace([np.inf, -np.inf], np.nan).fillna(1.0).clip(lower=0.25, upper=4.0)
    out["unit_price_factor"] = np.power(ratio, out["unit_price_elasticity"])
    out["unit_price_factor"] = pd.to_numeric(out["unit_price_factor"], errors="coerce").fillna(1.0).clip(
        lower=UNIT_PRICE_FACTOR_MIN, upper=UNIT_PRICE_FACTOR_MAX
    )
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

    # 過去実績だけ残っている休眠媒体を予測対象から外す。
    inactive = identify_inactive_media(history_df, selected_months)
    if not inactive.empty:
        clean = clean.loc[~clean["media"].isin(inactive["media"])].copy()
    if clean.empty:
        return pd.DataFrame(columns=columns)

    # 媒体ごとの「全定常日」を分母にする。行が無い日は0CV=非稼働日。
    # base_cv = 稼働率 × 稼働時CV と同値だが、診断値も保持して説明可能にする。
    _, full_daily, _ = _prepare_normal_learning_data(history_df, selected_months)
    if not inactive.empty and not full_daily.empty:
        full_daily = full_daily.loc[~full_daily["media"].isin(inactive["media"])].copy()
    media_day_weights = full_daily.groupby("media")["recency_weight"].sum()

    weighted = clean.copy()
    weighted["cv"] = pd.to_numeric(weighted["cv"], errors="coerce").fillna(0.0)
    weighted["cost"] = pd.to_numeric(weighted["cost"], errors="coerce").fillna(0.0)
    weighted["weighted_cv"] = weighted["cv"] * weighted["recency_weight"]
    weighted["weighted_cost"] = weighted["cost"] * weighted["recency_weight"]
    result = weighted.groupby(["media", "商品ID"], as_index=False).agg(
        weighted_cv=("weighted_cv", "sum"),
        weighted_cost=("weighted_cost", "sum"),
    )
    result["weight_days"] = result["media"].map(media_day_weights)
    result["base_cv"] = result["weighted_cv"] / result["weight_days"]

    # costは掲載日の水準を維持する（0CV日を混ぜて単価まで薄めない）。
    active_cost = weighted.loc[weighted["cv"].gt(0)].copy()
    if active_cost.empty:
        result["cost"] = 0.0
    else:
        active_cost["cost_w"] = active_cost["cost"] * active_cost["recency_weight"]
        cost_sum = active_cost.groupby(["media", "商品ID"])["cost_w"].sum()
        cost_weight = active_cost.groupby(["media", "商品ID"])["recency_weight"].sum()
        cost_mean = (cost_sum / cost_weight).replace([np.inf, -np.inf], np.nan)
        idx = pd.MultiIndex.from_frame(result[["media", "商品ID"]])
        result["cost"] = cost_mean.reindex(idx).to_numpy()
        result["cost"] = pd.to_numeric(result["cost"], errors="coerce").fillna(0.0)
    return result[columns]


def calculate_normal_media_diagnostics(
    history_df: pd.DataFrame,
    selected_months: list[str],
) -> pd.DataFrame:
    """定常予測の媒体別基礎値を、稼働率×稼働時CVに分解して返す。"""
    columns = ["media", "activity_rate", "active_daily_cv", "expected_daily_cv", "eligible_days", "active_days"]
    clean, daily, _ = _prepare_normal_learning_data(history_df, selected_months)
    if daily.empty:
        return pd.DataFrame(columns=columns)
    inactive = identify_inactive_media(history_df, selected_months)
    if not inactive.empty:
        daily = daily.loc[~daily["media"].isin(inactive["media"])].copy()
    rows = []
    for media, g in daily.groupby("media", sort=False):
        w = pd.to_numeric(g["recency_weight"], errors="coerce").fillna(0.0)
        cv = pd.to_numeric(g["cv"], errors="coerce").fillna(0.0)
        valid = w.gt(0)
        if not valid.any():
            continue
        active = cv.gt(0)
        activity_rate = float(w[active].sum() / w.sum()) if w.sum() else 0.0
        active_daily_cv = _weighted_mean(cv[active], w[active]) if active.any() else 0.0
        rows.append({
            "media": media,
            "activity_rate": activity_rate,
            "active_daily_cv": float(active_daily_cv) if pd.notna(active_daily_cv) else 0.0,
            "expected_daily_cv": activity_rate * (float(active_daily_cv) if pd.notna(active_daily_cv) else 0.0),
            "eligible_days": int(g["date"].nunique()),
            "active_days": int(g.loc[active, "date"].nunique()),
        })
    return pd.DataFrame(rows, columns=columns)
