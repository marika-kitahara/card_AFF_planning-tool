import pandas as pd
try:
    import jpholiday
except ImportError:  # ローカル確認時のフォールバック。クラウドではrequirementsから導入。
    jpholiday = None

REQUIRED_HISTORY_COLUMNS = {
    "成果発生日時",
    "パートナーサイト名",
    "件数",
    "報酬額",
    "商品ID",
}


def is_business_day(ts: pd.Timestamp) -> bool:
    date_value = pd.Timestamp(ts).date()
    is_holiday = jpholiday.is_holiday(date_value) if jpholiday is not None else False
    return date_value.weekday() < 5 and not is_holiday


def add_business_edge_flags(df: pd.DataFrame) -> pd.DataFrame:
    """日本の土日祝を除き、月初・月末4営業日を判定する。"""
    df = df.copy()
    dates = pd.to_datetime(df["date"])
    unique_months = dates.dt.to_period("M").dropna().unique()

    start_dates: set[pd.Timestamp] = set()
    end_dates: set[pd.Timestamp] = set()

    for period in unique_months:
        month_dates = pd.date_range(period.start_time, period.end_time, freq="D")
        business_days = [d.normalize() for d in month_dates if is_business_day(d)]
        start_dates.update(business_days[:4])
        end_dates.update(business_days[-4:])

    normalized = dates.dt.normalize()
    df["is_month_start"] = normalized.isin(start_dates).astype(int)
    df["is_month_end"] = normalized.isin(end_dates).astype(int)
    return df


def add_flags(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["date"] = pd.to_datetime(df["成果発生日時"], errors="coerce")
    df["month"] = df["date"].dt.month
    df["weekday"] = df["date"].dt.day_name()
    df = add_business_edge_flags(df)
    return df


def _read_csv_with_fallback(file) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSVの文字コードを判定できませんでした。UTF-8またはShift-JISで保存してください。")


def load_data(file) -> pd.DataFrame:
    df = _read_csv_with_fallback(file)

    missing = REQUIRED_HISTORY_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"実績CSVに必要な列がありません: {', '.join(sorted(missing))}")

    df = df.copy()
    df["media"] = df["パートナーサイト名"].astype(str).str.strip()
    df["cv"] = pd.to_numeric(df["件数"], errors="coerce")
    df["cost"] = pd.to_numeric(df["報酬額"], errors="coerce")
    df["商品ID"] = df["商品ID"].astype(str).str.strip()
    df = add_flags(df)

    df = df.dropna(subset=["date", "cv", "cost"])
    if df.empty:
        raise ValueError("有効な実績データがありません。日付・件数・報酬額を確認してください。")

    return (
        df.groupby(["date", "media", "商品ID"], dropna=False)
        .agg(
            cv=("cv", "sum"),
            cost=("cost", "sum"),
            weekday=("weekday", "first"),
            is_month_start=("is_month_start", "max"),
            is_month_end=("is_month_end", "max"),
        )
        .reset_index()
    )
