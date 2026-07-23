import pandas as pd

from logic.factors import apply_dynamic_factors


def forecast_cv(
    future_df: pd.DataFrame,
    factor_tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """ベース値へ、実績から都度算出した各係数を一度だけ掛ける。"""
    df = apply_dynamic_factors(future_df, factor_tables)
    df["forecast_cv"] = (
        df["base_cv"]
        * df["cpn_factor"]
        * df["weekday_factor"]
        * df["season_factor"]
        * df["month_edge_factor"]
        * df["after_factor"]
        * df["line_factor"]
    )
    return df
