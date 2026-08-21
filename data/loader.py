import pandas as pd

try:
    import jpholiday
except ImportError:
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

    start_dates = set()
    end_dates = set()

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
    raise ValueError(
        "CSVの文字コードを判定できませんでした。"
        "UTF-8またはShift-JISで保存してください。"
    )


def _normalize_id_value(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def load_media_master(file, sheet_name: str = "媒体名マスタ") -> pd.DataFrame:
    """CPNマスタExcel内の媒体名マスタを読み込む。A:SID / B:媒体名 / C:カテゴリ。"""
    if hasattr(file, "seek"):
        file.seek(0)
    media_master = pd.read_excel(file, sheet_name=sheet_name, engine="openpyxl")

    required_columns = {"SID", "媒体名"}
    missing = required_columns - set(media_master.columns)
    if missing:
        raise ValueError("媒体名マスタに必要な列がありません: " + ", ".join(sorted(missing)))

    media_master = media_master.copy()
    media_master["SID"] = media_master["SID"].map(_normalize_id_value)
    media_master["媒体名"] = media_master["媒体名"].astype("string").str.strip()

    if "カテゴリ" not in media_master.columns:
        media_master["カテゴリ"] = "未分類"

    media_master["カテゴリ"] = (
        media_master["カテゴリ"].astype("string").str.strip().replace("", pd.NA).fillna("未分類")
    )
    media_master = media_master[media_master["SID"].ne("")].copy()
    media_master = media_master.drop_duplicates(subset=["SID"], keep="last")
    return media_master[["SID", "媒体名", "カテゴリ"]]


def load_data(file, exclude_compensation: bool = True) -> pd.DataFrame:
    """
    実績CSVをプランニング用の日次データへ整形する。

    位置指定ルール（ユーザー仕様）:
      A列: 成果フラグ Y / D / N。Nは完全除外。
      C列: 商品ID（列名『商品ID』も必須）。
      F列: SID。
      P列: 補填情報。文字列が入っている行は補填対象。
      W列: グロス。0円は成果対象外として完全除外。

    集計ルール:
      発生件数(cv): Y + D（ただしN・グロス0・補填除外行は対象外）
      発行数(approved_cv): Yのみ
      コスト(cost): Yのみ
      承認率分母(approval_base_cv): Y + D の発生件数
    """
    df = _read_csv_with_fallback(file)

    missing = REQUIRED_HISTORY_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(f"実績CSVに必要な列がありません: {', '.join(sorted(missing))}")

    if df.shape[1] < 23:
        raise ValueError(
            "実績CSVは少なくともW列まで必要です。"
            "A列=成果フラグ、F列=SID、P列=補填、W列=グロスを使用します。"
        )

    df = df.copy()

    # 位置指定列を先に保持。列名変更の影響を受けないようにする。
    result_flag = df.iloc[:, 0].astype("string").fillna("").str.strip().str.upper()
    compensation_raw = df.iloc[:, 15]
    gross_raw = pd.to_numeric(df.iloc[:, 22], errors="coerce").fillna(0.0)

    df["SID"] = df.iloc[:, 5].map(_normalize_id_value)
    df["media"] = df["パートナーサイト名"].astype(str).str.strip()
    df["cv"] = pd.to_numeric(df["件数"], errors="coerce")
    raw_cost = pd.to_numeric(df["報酬額"], errors="coerce").fillna(0.0)
    df["商品ID"] = df["商品ID"].map(_normalize_id_value)

    # 補填判定: P列に「空白以外の文字列」が入っている行。
    comp_text = compensation_raw.astype("string").fillna("").str.strip()
    df["is_compensation"] = comp_text.ne("")

    # Nは常に対象外。W列グロス0も成果対象外。
    valid_mask = result_flag.isin(["Y", "D"]) & gross_raw.ne(0)
    if exclude_compensation:
        valid_mask &= ~df["is_compensation"]
    df = df.loc[valid_mask].copy()
    result_flag = result_flag.loc[df.index]
    raw_cost = raw_cost.loc[df.index]

    # Yのみ発行数・コスト対象。Dは発生件数のみに残す。
    y_mask = result_flag.eq("Y")
    df["approved_cv"] = 0.0
    df.loc[y_mask, "approved_cv"] = df.loc[y_mask, "cv"]
    df["approval_base_cv"] = df["cv"]
    df["cost"] = 0.0
    df.loc[y_mask, "cost"] = raw_cost.loc[y_mask]
    df["approval_source_column"] = "A列成果フラグ"

    df = add_flags(df)
    df = df.dropna(subset=["date", "cv", "cost"])

    if df.empty:
        raise ValueError(
            "有効な実績データがありません。"
            "N除外・グロス0除外・補填除外後の日付/件数/報酬額を確認してください。"
        )

    grouped = (
        df.groupby(["date", "media", "SID", "商品ID"], dropna=False)
        .agg(
            cv=("cv", "sum"),
            cost=("cost", "sum"),
            approved_cv=("approved_cv", "sum"),
            approval_base_cv=("approval_base_cv", "sum"),
            approval_source_column=("approval_source_column", "first"),
            weekday=("weekday", "first"),
            is_month_start=("is_month_start", "max"),
            is_month_end=("is_month_end", "max"),
        )
        .reset_index()
    )
    return grouped
