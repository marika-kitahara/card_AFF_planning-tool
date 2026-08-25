from io import BytesIO

import pandas as pd
import streamlit as st

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

# load_data() で位置指定して使う列。
# A=0 / C=2 / F=5 / P=15 / W=22
REQUIRED_HISTORY_POSITIONS = {0, 2, 5, 15, 22}


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


def _file_to_bytes(file) -> bytes:
    """UploadedFile / BytesIO / 通常のfile-likeをbytesへ統一する。"""
    if isinstance(file, bytes):
        return file
    if isinstance(file, bytearray):
        return bytes(file)
    if hasattr(file, "getvalue"):
        return file.getvalue()
    if hasattr(file, "seek"):
        file.seek(0)
    data = file.read()
    if hasattr(file, "seek"):
        file.seek(0)
    return data


def _read_csv_selected_columns(raw_bytes: bytes, encoding: str) -> pd.DataFrame:
    """必要列だけ読むことで巨大CSVのピークメモリを抑える。"""
    header = pd.read_csv(BytesIO(raw_bytes), encoding=encoding, nrows=0)

    if len(header.columns) < 23:
        raise ValueError(
            "実績CSVは少なくともW列まで必要です。"
            "A列=成果フラグ、F列=SID、P列=補填、W列=グロスを使用します。"
        )

    missing = REQUIRED_HISTORY_COLUMNS - set(header.columns)
    if missing:
        raise ValueError(f"実績CSVに必要な列がありません: {', '.join(sorted(missing))}")

    named_positions = {header.columns.get_loc(name) for name in REQUIRED_HISTORY_COLUMNS}
    usecols = sorted(REQUIRED_HISTORY_POSITIONS | named_positions)

    return pd.read_csv(
        BytesIO(raw_bytes),
        encoding=encoding,
        usecols=usecols,
        low_memory=False,
    )


def _read_csv_with_fallback(raw_bytes: bytes) -> pd.DataFrame:
    last_unicode_error = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return _read_csv_selected_columns(raw_bytes, encoding)
        except UnicodeDecodeError as exc:
            last_unicode_error = exc
            continue

    raise ValueError(
        "CSVの文字コードを判定できませんでした。"
        "UTF-8またはShift-JISで保存してください。"
    ) from last_unicode_error


def _normalize_id_value(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _normalize_media_master(media_master: pd.DataFrame) -> pd.DataFrame:
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


@st.cache_data(show_spinner=False, max_entries=2)
def _load_master_workbook_cached(raw_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    """CPNマスタExcelを1回だけ開き、2シートをまとめてキャッシュする。"""
    sheets = pd.read_excel(
        BytesIO(raw_bytes),
        sheet_name=["CPNマスタ", "媒体名マスタ"],
        engine="openpyxl",
    )
    return sheets["CPNマスタ"], _normalize_media_master(sheets["媒体名マスタ"])


def load_master_workbook(file) -> tuple[pd.DataFrame, pd.DataFrame]:
    """CPNマスタと媒体名マスタを同一Excelからまとめて読み込む。"""
    raw_bytes = _file_to_bytes(file)
    cpn_master, media_master = _load_master_workbook_cached(raw_bytes)
    # 呼び出し側で加工するため、キャッシュ本体を直接変更しないようcopyして返す。
    return cpn_master.copy(), media_master.copy()


def load_media_master(file, sheet_name: str = "媒体名マスタ") -> pd.DataFrame:
    """後方互換用。通常は load_master_workbook() を使用する。"""
    if sheet_name == "媒体名マスタ":
        _, media_master = load_master_workbook(file)
        return media_master

    raw_bytes = _file_to_bytes(file)
    media_master = pd.read_excel(BytesIO(raw_bytes), sheet_name=sheet_name, engine="openpyxl")
    return _normalize_media_master(media_master)


@st.cache_data(show_spinner=False, max_entries=2)
def _load_data_cached(raw_bytes: bytes, exclude_compensation: bool = True) -> pd.DataFrame:
    """巨大CSVの読込～日次集計をキャッシュし、rerun時の再読込を防ぐ。"""
    df = _read_csv_with_fallback(raw_bytes)

    # usecolsで列数が減っているため、位置指定は元CSVの列名から取得する。
    # ヘッダーを再度軽量に読み、A/F/P/Wの元列名を確定する。
    header = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            header = pd.read_csv(BytesIO(raw_bytes), encoding=encoding, nrows=0)
            break
        except UnicodeDecodeError:
            continue
    if header is None:
        raise ValueError("CSVの文字コードを判定できませんでした。")

    result_flag_col = header.columns[0]
    sid_col = header.columns[5]
    compensation_col = header.columns[15]
    gross_col = header.columns[22]

    result_flag = df[result_flag_col].astype("string").fillna("").str.strip().str.upper()
    compensation_raw = df[compensation_col]
    gross_raw = pd.to_numeric(df[gross_col], errors="coerce").fillna(0.0)

    df["SID"] = df[sid_col].map(_normalize_id_value)
    df["media"] = df["パートナーサイト名"].astype(str).str.strip()
    df["cv"] = pd.to_numeric(df["件数"], errors="coerce")
    raw_cost = pd.to_numeric(df["報酬額"], errors="coerce").fillna(0.0)
    df["商品ID"] = df["商品ID"].map(_normalize_id_value)

    comp_text = compensation_raw.astype("string").fillna("").str.strip()
    df["is_compensation"] = comp_text.ne("")

    valid_mask = result_flag.isin(["Y", "D"]) & gross_raw.ne(0)
    if exclude_compensation:
        valid_mask &= ~df["is_compensation"]

    df = df.loc[valid_mask].copy()
    result_flag = result_flag.loc[df.index]
    raw_cost = raw_cost.loc[df.index]

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

    メモリ対策:
      - 必要列だけCSVから読み込む。
      - 同じファイル・同じ補填条件なら日次集計結果を再利用する。
    """
    raw_bytes = _file_to_bytes(file)
    return _load_data_cached(raw_bytes, exclude_compensation=exclude_compensation).copy()


def _read_tabular_with_fallback(raw_bytes: bytes, filename: str = "") -> pd.DataFrame:
    """AF実績用。CSV/Excelのどちらでも全列を読み込む。"""
    lower_name = str(filename or "").lower()
    if lower_name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(BytesIO(raw_bytes), engine="openpyxl")

    last_unicode_error = None
    for encoding in ("utf-8-sig", "cp932", "utf-8"):
        try:
            return pd.read_csv(BytesIO(raw_bytes), encoding=encoding, low_memory=False)
        except UnicodeDecodeError as exc:
            last_unicode_error = exc
            continue
    raise ValueError(
        "AF実績ファイルの文字コードを判定できませんでした。UTF-8またはShift-JISで保存してください。"
    ) from last_unicode_error


def _normalize_af_code(value) -> str:
    return _normalize_id_value(value)


@st.cache_data(show_spinner=False, max_entries=4)
def _load_af_data_cached(raw_bytes: bytes, filename: str, metric_name: str, valid_codes: tuple[str, ...]) -> pd.DataFrame:
    """AF横持ち実績を対象AFコード列だけ合計し、日次件数へ集計する。"""
    df = _read_tabular_with_fallback(raw_bytes, filename=filename)
    if df.empty:
        raise ValueError(f"{metric_name}実績データが空です。")
    if len(df.columns) < 2:
        raise ValueError(
            f"{metric_name}実績データは、A列=日付、B列以降=AFコードの横持ち形式である必要があります。"
        )

    date_col = df.columns[0]
    valid_code_set = set(valid_codes)

    # B列以降のヘッダーがAFコード。AFコードマスタA列と一致する列だけ集計対象にする。
    # pandasが重複ヘッダーへ付ける '.1' 等は、元コード判定時だけ除去する。
    matched_columns = []
    for col in df.columns[1:]:
        normalized_col = _normalize_af_code(col)
        base_col = normalized_col.rsplit(".", 1)[0] if normalized_col.rsplit(".", 1)[-1].isdigit() else normalized_col
        if normalized_col in valid_code_set or base_col in valid_code_set:
            matched_columns.append(col)

    if not matched_columns:
        # 対象AFコード列がゼロの場合もエラーにはせず、日付だけ保持して0件として返す。
        # AFコードの追加・削除で列構成が変動してもアプリ全体を止めないため。
        raw_date = df[date_col].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
        parsed_yyyymmdd = pd.to_datetime(raw_date, format="%Y%m%d", errors="coerce")
        parsed_fallback = pd.to_datetime(df[date_col], errors="coerce")
        dates = parsed_yyyymmdd.fillna(parsed_fallback).dt.normalize()
        result = pd.DataFrame({"date": dates}).dropna(subset=["date"])
        result[metric_name] = 0
        return result.groupby("date", as_index=False)[metric_name].sum()

    work = df[[date_col] + matched_columns].copy()

    # A列の日付は 20260101 のような8桁数値/文字列を想定。
    raw_date = work[date_col].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
    parsed_yyyymmdd = pd.to_datetime(raw_date, format="%Y%m%d", errors="coerce")
    parsed_fallback = pd.to_datetime(work[date_col], errors="coerce")
    work["date"] = parsed_yyyymmdd.fillna(parsed_fallback).dt.normalize()

    # 各AFコード列のセル値が件数。文字列や空欄は0として扱い、行方向に合計する。
    numeric_counts = work[matched_columns].apply(pd.to_numeric, errors="coerce").fillna(0)
    work[metric_name] = numeric_counts.sum(axis=1)
    work = work.dropna(subset=["date"])

    if work.empty:
        return pd.DataFrame(columns=["date", metric_name])

    result = work.groupby("date", as_index=False)[metric_name].sum()
    # 件数は整数想定だが、元データに小数が混ざっても情報を落とさない。
    if (result[metric_name] % 1 == 0).all():
        result[metric_name] = result[metric_name].astype(int)
    return result


def load_af_data(file, af_code_master: pd.DataFrame, metric_name: str) -> pd.DataFrame:
    """
    AF計測データを読み込む。

    - A列: 日付。20260101形式を日付へ変換
    - B列以降の1行目: AFコード
    - AFコードマスタA列と一致するAFコード列だけを集計対象にする
    - 各セルの数値が件数。対象列を行方向に合計して日次件数を作る
    """
    if af_code_master is None or af_code_master.empty:
        raise ValueError("CPNマスタの『AFコードマスタ』シートにAFコードがありません。")
    code_col = af_code_master.columns[0]
    valid_codes = tuple(
        sorted({
            _normalize_af_code(v)
            for v in af_code_master[code_col].tolist()
            if _normalize_af_code(v) != ""
        })
    )
    if not valid_codes:
        raise ValueError("CPNマスタの『AFコードマスタ』A列に有効なAFコードがありません。")
    raw_bytes = _file_to_bytes(file)
    filename = getattr(file, "name", "")
    return _load_af_data_cached(raw_bytes, filename, metric_name, valid_codes).copy()


@st.cache_data(show_spinner=False, max_entries=2)
def _load_master_workbook_with_af_cached(raw_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """CPN/媒体/AFコードの3シートをまとめて読み込む。AFコードマスタは任意。"""
    book = pd.ExcelFile(BytesIO(raw_bytes), engine="openpyxl")
    required = {"CPNマスタ", "媒体名マスタ"}
    missing = required - set(book.sheet_names)
    if missing:
        raise ValueError("CPNマスタExcelに必要なシートがありません: " + ", ".join(sorted(missing)))

    cpn_master = pd.read_excel(book, sheet_name="CPNマスタ")
    media_master = _normalize_media_master(pd.read_excel(book, sheet_name="媒体名マスタ"))
    af_code_master = (
        pd.read_excel(book, sheet_name="AFコードマスタ")
        if "AFコードマスタ" in book.sheet_names
        else pd.DataFrame(columns=["AFコード"])
    )
    return cpn_master, media_master, af_code_master


def load_master_workbook_with_af(file) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw_bytes = _file_to_bytes(file)
    cpn_master, media_master, af_code_master = _load_master_workbook_with_af_cached(raw_bytes)
    return cpn_master.copy(), media_master.copy(), af_code_master.copy()
