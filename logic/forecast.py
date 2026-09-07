import pandas as pd

from logic.factors import apply_dynamic_factors


def forecast_cv(
    future_df: pd.DataFrame,
    factor_tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """ベース値へ、実績から都度算出した各係数を一度だけ掛ける。"""
    df = apply_dynamic_factors(future_df, factor_tables)
    # 診断用に、補正を1段ずつ適用した途中値を保持する。
    # 最終 forecast_cv の式自体は従来と同じ。
    df["stage_base_cv"] = df["base_cv"] * df["cpn_factor"]
    df["stage_unit_price_cv"] = df["stage_base_cv"] * df.get("unit_price_factor", 1.0)
    df["stage_weekday_cv"] = df["stage_unit_price_cv"] * df["weekday_factor"]
    df["stage_season_cv"] = df["stage_weekday_cv"] * df["season_factor"]
    df["stage_month_edge_cv"] = df["stage_season_cv"] * df["month_edge_factor"]
    df["stage_after_cv"] = df["stage_month_edge_cv"] * df["after_factor"]
    df["stage_line_cv"] = df["stage_after_cv"] * df["line_factor"]
    df["forecast_cv"] = df["stage_line_cv"]
    return df
