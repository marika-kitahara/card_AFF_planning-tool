def calc_weekday_factor(df):
    """
    媒体別 × 曜日係数
    """

    overall_avg = df.groupby("media")["cv"].transform("mean")
    weekday_avg = df.groupby(["media", "weekday"])["cv"].transform("mean")

    factor = weekday_avg / overall_avg

    # 念のためNaN対策
    factor = factor.fillna(1.0)

    return factor


def calc_month_edge_factor(df):
    """
    月初・月末ブースト
    """

    df["month_edge_factor"] = 1.0

    df.loc[df["is_month_start"] == 1, "month_edge_factor"] *= 1.1
    df.loc[df["is_month_end"] == 1, "month_edge_factor"] *= 1.2

    return df["month_edge_factor"]


def calc_line_oa_factor(df):
    """
    LINE配信日ブースト
    """

    df["line_factor"] = 1.0

    df.loc[
        (df["media"] == "LINE") & (df["line_oa_flag"] == 1),
        "line_factor"
    ] *= 1.3

    return df["line_factor"]


def calc_magitoku_factor(df):
    """
    マジ得後減衰
    """

    df["magitoku_factor"] = 1.0

    df.loc[df["magitoku_flag"] == 1, "magitoku_factor"] *= 0.9

    return df["magitoku_factor"]

def calc_cpn_factor(df):
    """
    商品IDベースで係数計算（cpn使わない）
    """

    # 媒体別平均
    media_avg = df.groupby("media")["cv"].transform("mean")

    # 媒体 × 商品ID平均
    item_avg = df.groupby(["media", "商品ID"])["cv"].transform("mean")

    factor = item_avg / media_avg

    return factor.fillna(1.0)