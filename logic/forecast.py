import pandas as pd

from logic.factors import (
    calc_month_edge_factor,
    calc_line_oa_factor,
    calc_magitoku_factor,
    calc_cpn_factor
)


def calc_weekday_factor_from_history(future_df, history_df):
    """
    過去データから曜日係数を作って未来に適用
    """

    weekday_avg = history_df.groupby(["media", "weekday"])["cv"].mean()
    overall_avg = history_df.groupby("media")["cv"].mean()

    ratio = (weekday_avg / overall_avg).to_dict()

    future_df["weekday_factor"] = future_df.set_index(
        ["media", "weekday"]
    ).index.map(ratio)

    return future_df["weekday_factor"].fillna(1.0)


def calc_cpn_factor_from_history(future_df, history_df):
    """
    商品IDベースの係数（過去→未来適用）
    """

    media_avg = history_df.groupby("media")["cv"].mean()
    item_avg = history_df.groupby(["media", "商品ID"])["cv"].mean()

    ratio = (item_avg / media_avg).to_dict()

    future_df["cpn_factor"] = future_df.set_index(
        ["media", "商品ID"]
    ).index.map(ratio)

    return future_df["cpn_factor"].fillna(1.0)


def forecast_cv(future_df, history_df):
    """
    未来の件数予測
    """

    df = future_df.copy()

    # -----------------------
    # ✅ ベースCV（過去から）
    # -----------------------
    base_cv = history_df.groupby("media")["cv"].mean()
    df["base_cv"] = df["media"].map(base_cv)

    # -----------------------
    # ✅ 係数（過去から生成）
    # -----------------------
    df["weekday_factor"] = calc_weekday_factor_from_history(df, history_df)
    df["cpn_factor"] = calc_cpn_factor_from_history(df, history_df)

    # -----------------------
    # ✅ フラグ系（未来データにそのまま適用）
    # -----------------------
    df["month_edge_factor"] = calc_month_edge_factor(df)
    df["line_factor"] = calc_line_oa_factor(df)
    df["magitoku_factor"] = calc_magitoku_factor(df)

    # -----------------------
    # ✅ 予測
    # -----------------------
    df["forecast_cv"] = (
        df["base_cv"]
        * df["weekday_factor"]
        * df["cpn_factor"]
        * df["month_edge_factor"]
        * df["line_factor"]
        * df["magitoku_factor"]
    )

    # -----------------------
    # ✅ ベースコスト（過去から）
    # -----------------------
    base_cost = history_df.groupby("media")["cost"].mean()
    df["cost"] = df["media"].map(base_cost)

    return df