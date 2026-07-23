import pandas as pd

REQUIRED_HISTORY_COLUMNS = {
    "成果発生日時",
    "パートナーサイト名",
    "件数",
    "報酬額",
    "商品ID",
}


def add_flags(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["date"] = pd.to_datetime(df["成果発生日時"], errors="coerce")
    df["day"] = df["date"].dt.day
    df["month"] = df["date"].dt.month
    df["weekday"] = df["date"].dt.day_name()

    # 現状は暦日ベース。厳密な「営業日」にする場合は休日マスタが必要。
    df["is_month_start"] = df["day"].isin([1, 2, 3, 4]).astype(int)
    df["is_month_end"] = df["day"].isin([28, 29, 30, 31]).astype(int)
    df["magitoku_flag"] = 0
    df["line_oa_flag"] = 0
    return df


def _read_csv_with_fallback(file) -> pd.DataFrame:
    errors = []
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=encoding)
        except UnicodeDecodeError as exc:
            errors.append(f"{encoding}: {exc}")
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
            magitoku_flag=("magitoku_flag", "max"),
            line_oa_flag=("line_oa_flag", "max"),
        )
        .reset_index()
    )
