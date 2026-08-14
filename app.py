# -*- coding: utf-8 -*-
# SUBMISSION_TEMPLATE_ALL_SHEETS_V5 = 2026-08-14
# MergedCell-safe version

import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pathlib import Path
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from data.loader import load_data
from logic.forecast import forecast_cv
from logic.simulation import simulate_plan
from logic.optimize import optimize_budget


# -----------------------
# ✅ 日付フォーマット
# -----------------------
def format_date(df, col="date"):
    return (
        pd.to_datetime(df[col])
        .dt.strftime("%Y/%m/%d")
        .str.replace("/0", "/", regex=False)
    )


# -----------------------
# ✅ 帳票形式
# -----------------------
def create_report_table(df):

    pivot = df.pivot_table(
        index=["media", "plan"],
        columns="date",
        values=["cv", "cost", "cpa"],
        aggfunc="sum"
    )

    pivot = pivot.sort_index(axis=1)

    rows = []

    for (media, plan) in pivot.index:

        sub = pivot.loc[(media, plan)]

        cv = sub["cv"]
        cost = sub["cost"]
        cpa = sub["cpa"]

        cv_df = pd.DataFrame([cv])
        cost_df = pd.DataFrame([cost])
        cpa_df = pd.DataFrame([cpa])

        cv_df["media"] = media
        cv_df["plan"] = plan
        cv_df["metric"] = "CV"

        cost_df["media"] = media
        cost_df["plan"] = plan
        cost_df["metric"] = "COST"

        cpa_df["media"] = media
        cpa_df["plan"] = plan
        cpa_df["metric"] = "CPA"

        rows.extend([cv_df, cost_df, cpa_df])

    result = pd.concat(rows)

    cols = ["media", "plan", "metric"] + [
        c for c in result.columns if c not in ["media", "plan", "metric"]
    ]

    result = result[cols]

    result["media"] = result["media"].mask(result["media"].duplicated())
    result["plan"] = result["plan"].mask(
        (result["plan"].shift() == result["plan"]) &
        (result["media"].shift() == result["media"])
    )

    return result.reset_index(drop=True)


# -----------------------
# ✅ Excel
# -----------------------
def to_excel_multi(sim_df, opt_df):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        create_report_table(sim_df).to_excel(writer, sheet_name="松竹梅", index=False)
        create_report_table(opt_df).to_excel(writer, sheet_name="最適", index=False)

    return output.getvalue()


# -----------------------
# ✅ 提出用Excel（最適プランのみ / 添付テンプレ全シート再現）
# -----------------------
def _template_path() -> Path:
    base = Path(__file__).resolve().parent
    candidates = [
        base / "assets" / "submission_template.xlsx",
        base / "submission_template.xlsx",
    ]
    for path in candidates:
        if path.exists():
            return path
    raise FileNotFoundError(
        "提出用テンプレートが見つかりません。assets/submission_template.xlsx を配置してください。"
    )


def _safe_number(value, default=0.0):
    value = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    return default if pd.isna(value) else float(value)


def _copy_style(src, dst):
    """テンプレの見た目をそのまま複製するための最小スタイルコピー。"""
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)
    if src.protection:
        dst.protection = copy(src.protection)


def _set_date_slots(ws, row, first_col, slot_count, dates, total_col=None):
    """テンプレの日付セル書式を維持しつつ、予測期間の日付へ差し替える。"""
    for i in range(slot_count):
        cell = ws.cell(row, first_col + i)

        # 結合セルの左上以外には書き込まない
        if isinstance(cell, MergedCell):
            continue

        if i < len(dates):
            cell.value = dates[i].to_pydatetime()
            cell.number_format = "m/d"
        else:
            cell.value = None

    if total_col:
        cell = ws.cell(row, total_col)
        if not isinstance(cell, MergedCell):
            cell.value = "Total"


def _clear_values(ws, min_row, max_row, min_col, max_col):
    """
    書式・罫線・セル結合・行高・列幅を維持したまま値だけ消す。
    openpyxl の MergedCell は value を変更できないためスキップする。
    """
    for row in ws.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.value = None


def _set_value(ws, row, col, value):
    """
    結合セル対策付きの値セット。
    左上セル以外の MergedCell に当たった場合は何もしない。
    """
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def create_submission_excel(opt_summary, history_df, start_date, end_date, selected_cpn, opt_mode):
    """
    添付された提出用Excelそのものをテンプレートとして使い、
    最適プランの結果だけを全7シートへ反映する。

    - SID: 実績データF列を loader.py で保持した history_df["SID"]
    - 件数: 最適プランCV
    - コスト: 最適プランcost
    - 発行: 現行アプリが承認状況を加味しないため、最適プランCVをそのまま表示
    - Actual: 提出時点では未入力

    ※テンプレ内の既存媒体名や数値は、出力時に対象範囲をクリアしてから
      最適プランの値へ置き換える。
    ※元テンプレートそのものは変更しない。
    """
    plan = opt_summary.copy()
    plan["date"] = pd.to_datetime(plan["date"], errors="coerce").dt.normalize()
    plan = plan.dropna(subset=["date", "media"])
    plan["media"] = plan["media"].astype(str)
    plan["cv"] = pd.to_numeric(plan["cv"], errors="coerce").fillna(0)
    plan["cost"] = pd.to_numeric(plan["cost"], errors="coerce").fillna(0)

    dates = list(pd.date_range(pd.Timestamp(start_date), pd.Timestamp(end_date), freq="D"))
    if len(dates) > 33:
        raise ValueError("提出用テンプレートは最大33日分です。予測期間を33日以内にしてください。")

    daily = (
        plan.groupby(["date", "media"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    daily["cpa"] = (
        (daily["cost"] / daily["cv"])
        .replace([float("inf"), float("-inf")], 0)
        .fillna(0)
    )
    media_list = list(dict.fromkeys(daily["media"].tolist()))

    cv_map = {(r.date, r.media): float(r.cv) for r in daily.itertuples()}
    cost_map = {(r.date, r.media): float(r.cost) for r in daily.itertuples()}
    cpa_map = {(r.date, r.media): float(r.cpa) for r in daily.itertuples()}
    total_cv_by_date = daily.groupby("date")["cv"].sum().to_dict()
    total_cost_by_date = daily.groupby("date")["cost"].sum().to_dict()

    # SIDは実績CSVのF列を正とする。
    sid_map = {}
    if "SID" in history_df.columns:
        sid_source = history_df[["media", "SID"]].copy()
        sid_source["media"] = sid_source["media"].astype(str)
        sid_source["SID"] = sid_source["SID"].fillna("").astype(str).str.strip()
        sid_source = sid_source[sid_source["SID"] != ""]
        sid_map = (
            sid_source.groupby("media")["SID"]
            .agg(lambda x: " / ".join(dict.fromkeys(x.tolist())))
            .to_dict()
        )

    template = _template_path()
    wb_out = load_workbook(template)

    # Excelを開いたときに数式があれば再計算させる。
    try:
        wb_out.calculation.fullCalcOnLoad = True
        wb_out.calculation.forceFullCalc = True
        wb_out.calculation.calcMode = "auto"
    except Exception:
        pass

    main_old_name = "10月（既存移管合算）"
    main_ws = (
        wb_out[main_old_name]
        if main_old_name in wb_out.sheetnames
        else wb_out.worksheets[2]
    )

    # テンプレにあるSID→媒体区分の対応を先に取得してから、
    # 明細の既存値をクリアする。
    template_type_by_sid = {}
    template_type_by_media = {}
    for r in range(1, main_ws.max_row + 1):
        sid = main_ws.cell(r, 3).value
        media = main_ws.cell(r, 4).value
        media_type = main_ws.cell(r, 19).value
        if sid not in (None, "") and media_type not in (None, ""):
            template_type_by_sid[str(sid).strip()] = str(media_type).strip()
        if media not in (None, "") and media_type not in (None, ""):
            template_type_by_media[str(media).strip()] = str(media_type).strip()

    media_type_map = {}
    for media in media_list:
        sid = sid_map.get(media, "").split(" / ")[0].strip()
        media_type_map[media] = (
            template_type_by_sid.get(sid)
            or template_type_by_media.get(media)
            or "ポイントサイト"
        )

    # =========================================================
    # 1) メインシート：○月（既存移管合算）
    # =========================================================
    new_main_name = f"{pd.Timestamp(start_date).month}月（既存移管合算）"
    main_ws.title = new_main_name

    # 最小テンプレートは他シートから旧シート名を参照する数式を持たないため、
    # 全ワークシート・全セルの走査は行わない。
    # これにより提出用Excel生成時の処理時間を大幅に削減する。

    first_date_col = 25  # Y
    total_col = 58       # BF
    date_slots = 33

    # 上部サマリの日付・Target/Actual/GAP
    _set_date_slots(main_ws, 2, first_date_col, date_slots, dates, total_col)
    _set_date_slots(main_ws, 7, first_date_col, date_slots, dates, total_col)

    for i in range(date_slots):
        col = first_date_col + i
        if i < len(dates):
            dt = dates[i]
            target = round(total_cv_by_date.get(dt.normalize(), 0))
            _set_value(main_ws, 3, col, target)
            _set_value(main_ws, 4, col, None)
            _set_value(main_ws, 5, col, -target)
            _set_value(main_ws, 8, col, "月火水木金土日"[dt.weekday()])
        else:
            for rr in (3, 4, 5, 8):
                _set_value(main_ws, rr, col, None)

    _set_value(main_ws, 3, total_col, round(sum(total_cv_by_date.values())))
    _set_value(main_ws, 4, total_col, 0)
    _set_value(main_ws, 5, total_col, -round(sum(total_cv_by_date.values())))

    # 媒体区分ごとの上部4ブロックを最適プランから再集計。
    group_rows = {}
    for r in range(8, 21):
        label = main_ws.cell(r, 4).value
        metric = main_ws.cell(r, 24).value
        if (
            isinstance(label, str)
            and "合計" in label
            and metric == "Daily Target (Initiative)"
        ):
            key = label.replace("【", "").replace("】合計", "").strip()
            group_rows[key] = r

    for group_name, start_row in group_rows.items():
        members = [m for m in media_list if media_type_map.get(m) == group_name]

        for i in range(date_slots):
            col = first_date_col + i
            if i < len(dates):
                dt = dates[i].normalize()
                val = round(sum(cv_map.get((dt, m), 0) for m in members))
                _set_value(main_ws, start_row, col, val)
                _set_value(main_ws, start_row + 1, col, None)
                _set_value(main_ws, start_row + 2, col, -val)
            else:
                for rr in range(
                    start_row,
                    min(start_row + 3, main_ws.max_row + 1),
                ):
                    _set_value(main_ws, rr, col, None)

        total_val = round(
            sum(
                cv_map.get((d.normalize(), m), 0)
                for d in dates
                for m in members
            )
        )
        _set_value(main_ws, start_row, total_col, total_val)
        _set_value(main_ws, start_row + 1, total_col, 0)
        _set_value(main_ws, start_row + 2, total_col, -total_val)

    # 明細エリアはテンプレの先頭4行ブロックの見た目を全媒体にコピーして再構築。
    detail_start = 21
    detail_end = main_ws.max_row
    style_source_rows = [21, 22, 23, 24]

    style_snapshots = []
    for src_r in style_source_rows:
        row_styles = []
        for c in range(1, total_col + 1):
            src = main_ws.cell(src_r, c)
            row_styles.append(
                (
                    copy(src._style),
                    copy(src.alignment),
                    copy(src.font),
                    copy(src.fill),
                    copy(src.border),
                    src.number_format,
                )
            )
        style_snapshots.append(row_styles)

    # 既存値だけ消し、列幅・罫線等のテンプレ設定は保持。
    _clear_values(main_ws, detail_start, detail_end, 1, total_col)

    # 必要な明細行を最初にまとめて確保する。
    # insert_rows() を媒体ごとに繰り返すと、openpyxl が既存セルを毎回移動するため非常に重くなる。
    required_main_end = detail_start + max(len(media_list) - 1, 0) * 4 + 3
    if media_list and required_main_end > main_ws.max_row:
        main_ws.insert_rows(
            main_ws.max_row + 1,
            amount=required_main_end - main_ws.max_row,
        )

    for idx, media in enumerate(media_list, start=0):
        r0 = detail_start + idx * 4

        # 4行すべてにテンプレの同じ行パターンを適用
        for off in range(4):
            target_r = r0 + off
            for c in range(1, total_col + 1):
                cell = main_ws.cell(target_r, c)
                if isinstance(cell, MergedCell):
                    continue

                stl, algn, font, fill, border, numfmt = style_snapshots[off][c - 1]
                cell._style = copy(stl)
                cell.alignment = copy(algn)
                cell.font = copy(font)
                cell.fill = copy(fill)
                cell.border = copy(border)
                cell.number_format = numfmt

        sid = sid_map.get(media, "")
        media_type = media_type_map.get(media, "ポイントサイト")
        total_cv = round(
            sum(cv_map.get((d.normalize(), media), 0) for d in dates)
        )
        total_cost = round(
            sum(cost_map.get((d.normalize(), media), 0) for d in dates)
        )
        overall_cpa = round(total_cost / total_cv) if total_cv else 0

        # 左側情報
        _set_value(main_ws, r0, 2, idx + 1)
        _set_value(main_ws, r0, 3, sid)
        _set_value(main_ws, r0, 4, media)
        _set_value(main_ws, r0, 7, total_cv)
        _set_value(main_ws, r0, 8, 0)
        _set_value(main_ws, r0, 10, 0)
        _set_value(main_ws, r0, 11, 1)
        _set_value(main_ws, r0, 12, 1)
        _set_value(main_ws, r0, 17, total_cost)
        _set_value(main_ws, r0, 19, media_type)

        metrics = [
            "Daily Target (Initiative)",
            "Actural",
            "Promotion Detail",
            "GAP",
        ]
        for off, metric in enumerate(metrics):
            _set_value(main_ws, r0 + off, 24, metric)

        for i in range(date_slots):
            col = first_date_col + i
            if i < len(dates):
                dt = dates[i].normalize()
                cv = round(cv_map.get((dt, media), 0))
                cpa = round(cpa_map.get((dt, media), 0)) if cv else 0

                _set_value(main_ws, r0, col, cv)
                _set_value(main_ws, r0 + 1, col, None)
                _set_value(main_ws, r0 + 2, col, cpa)
                _set_value(main_ws, r0 + 3, col, -cv)
            else:
                for off in range(4):
                    _set_value(main_ws, r0 + off, col, None)

        _set_value(main_ws, r0, total_col, total_cv)
        _set_value(main_ws, r0 + 1, total_col, 0)
        _set_value(main_ws, r0 + 2, total_col, overall_cpa)
        _set_value(main_ws, r0 + 3, total_col, -total_cv)

    # =========================================================
    # 2) 件数(合算）
    # =========================================================
    count_ws = wb_out["件数(合算）"]
    count_first_date_col = 16  # P
    count_total_col = 49       # AW

    _set_date_slots(
        count_ws,
        7,
        count_first_date_col,
        date_slots,
        dates,
        count_total_col,
    )
    _clear_values(count_ws, 8, count_ws.max_row, 1, count_total_col)

    required_count_end = 8 + max(len(media_list) - 1, 0)
    if media_list and required_count_end > count_ws.max_row:
        count_ws.insert_rows(
            count_ws.max_row + 1,
            amount=required_count_end - count_ws.max_row,
        )

    for idx, media in enumerate(media_list):
        r = 8 + idx

        sid = sid_map.get(media, "")
        _set_value(count_ws, r, 1, sid)
        _set_value(count_ws, r, 2, media)
        _set_value(count_ws, r, 11, 1)
        _set_value(count_ws, r, 13, 1)

        total_cv = 0
        for i in range(date_slots):
            c = count_first_date_col + i
            if i < len(dates):
                val = round(cv_map.get((dates[i].normalize(), media), 0))
                _set_value(count_ws, r, c, val)
                total_cv += val
            else:
                _set_value(count_ws, r, c, None)

        _set_value(count_ws, r, count_total_col, total_cv)

    # =========================================================
    # 3) 発行(合算） : 現行アプリは承認加味なしのためCVをそのまま表示
    # =========================================================
    issue_ws = wb_out["発行(合算）"]
    issue_first_date_col = 15  # O
    issue_total_col = 48       # AV

    _set_date_slots(
        issue_ws,
        5,
        issue_first_date_col,
        date_slots,
        dates,
        issue_total_col,
    )
    _clear_values(issue_ws, 6, issue_ws.max_row, 1, issue_total_col)

    required_issue_end = 6 + max(len(media_list) - 1, 0)
    if media_list and required_issue_end > issue_ws.max_row:
        issue_ws.insert_rows(
            issue_ws.max_row + 1,
            amount=required_issue_end - issue_ws.max_row,
        )

    for idx, media in enumerate(media_list):
        r = 6 + idx

        _set_value(issue_ws, r, 1, sid_map.get(media, ""))
        _set_value(issue_ws, r, 2, media)
        _set_value(issue_ws, r, 3, 1)
        _set_value(issue_ws, r, 4, 1)

        total_cv = 0
        for i in range(date_slots):
            c = issue_first_date_col + i
            if i < len(dates):
                val = round(cv_map.get((dates[i].normalize(), media), 0))
                _set_value(issue_ws, r, c, val)
                total_cv += val
            else:
                _set_value(issue_ws, r, c, None)

        _set_value(issue_ws, r, issue_total_col, total_cv)

    # =========================================================
    # 4) コスト計算用(合算） : 最適プランcostを直接反映
    # =========================================================
    cost_ws = wb_out["コスト計算用(合算）"]
    cost_first_date_col = 15  # O
    cost_total_col = 48       # AV

    _set_date_slots(
        cost_ws,
        6,
        cost_first_date_col,
        date_slots,
        dates,
        cost_total_col,
    )
    _clear_values(cost_ws, 7, cost_ws.max_row, 1, cost_total_col)

    required_cost_end = 7 + max(len(media_list) - 1, 0)
    if media_list and required_cost_end > cost_ws.max_row:
        cost_ws.insert_rows(
            cost_ws.max_row + 1,
            amount=required_cost_end - cost_ws.max_row,
        )

    for idx, media in enumerate(media_list):
        r = 7 + idx

        _set_value(cost_ws, r, 1, sid_map.get(media, ""))
        _set_value(cost_ws, r, 2, media)
        _set_value(cost_ws, r, 11, 1)
        _set_value(cost_ws, r, 12, 1)
        _set_value(cost_ws, r, 13, 1)
        _set_value(cost_ws, r, 14, 1)

        total_cost = 0
        for i in range(date_slots):
            c = cost_first_date_col + i
            if i < len(dates):
                val = round(cost_map.get((dates[i].normalize(), media), 0))
                _set_value(cost_ws, r, c, val)
                total_cost += val
            else:
                _set_value(cost_ws, r, c, None)

        _set_value(cost_ws, r, cost_total_col, total_cost)

    # =========================================================
    # 5) 全体サマリ / 6) 短縮承認除外
    # =========================================================
    for summary_name in ["全体サマリ（定常期間サマリ）", "短縮承認除外"]:
        sws = wb_out[summary_name]
        _set_value(
            sws,
            1,
            2,
            f"{pd.Timestamp(start_date).month}月度サマリ 件数コスト",
        )

        # 日次33行をテンプレの3行目から使用
        for i in range(33):
            r = 3 + i

            if i < len(dates):
                dt = dates[i].normalize()
                cv = round(total_cv_by_date.get(dt, 0))
                cost = round(total_cost_by_date.get(dt, 0))
                cpa = round(cost / cv) if cv else 0

                _set_value(sws, r, 1, dates[i].to_pydatetime())
                if not isinstance(sws.cell(r, 1), MergedCell):
                    sws.cell(r, 1).number_format = "m/d"

                _set_value(sws, r, 2, cv)
                _set_value(sws, r, 3, cv)
                _set_value(sws, r, 5, cv)
                _set_value(sws, r, 7, 1 if cv else 0)
                _set_value(sws, r, 8, cost)
                _set_value(sws, r, 9, cpa)
            else:
                for c in range(1, 10):
                    _set_value(sws, r, c, None)

    # =========================================================
    # 7) 短縮承認日程
    # =========================================================
    sched_ws = wb_out["短縮承認日程"]
    start_ts = pd.Timestamp(start_date)
    sched_start = (
        (start_ts - pd.DateOffset(months=1))
        .replace(day=1)
        .normalize()
    )
    sched_dates = pd.date_range(sched_start, periods=64, freq="D")

    for i, dt in enumerate(sched_dates, start=2):
        _set_value(sched_ws, i, 1, dt.to_pydatetime())
        if not isinstance(sched_ws.cell(i, 1), MergedCell):
            sched_ws.cell(i, 1).number_format = "m/d"

        _set_value(sched_ws, i, 2, "月火水木金土日"[dt.weekday()])

        _set_value(sched_ws, i, 7, dt.to_pydatetime())
        if not isinstance(sched_ws.cell(i, 7), MergedCell):
            sched_ws.cell(i, 7).number_format = "m/d"

        _set_value(sched_ws, i, 8, "月火水木金土日"[dt.weekday()])

        for c in range(3, 7):
            _set_value(sched_ws, i, c, None)

        for c in range(9, 13):
            _set_value(sched_ws, i, c, None)

    output = BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output.getvalue()


# -----------------------
# ✅ 補助関数
# -----------------------
def _truthy(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin(
        {"1", "true", "yes", "y", "○", "〇", "あり", "有", "実施"}
    )


def _daily_pair_average(df: pd.DataFrame) -> pd.DataFrame:
    daily = (
        df.groupby(["date", "media", "商品ID"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    return (
        daily.groupby(["media", "商品ID"], as_index=False)
        .agg(base_cv=("cv", "mean"), cost=("cost", "mean"))
    )


# -----------------------
# ✅ UI
# -----------------------
st.set_page_config(page_title="AFプランニングツール", layout="wide")
st.title("📊 件数予測＆プランニングツール")
st.caption(
    "実績CSVと最新のCPNマスタをアップロードしてください。"
    "ファイルはGitHubには保存されません。"
)

col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader("① 実績CSV", type=["csv"])

with col2:
    uploaded_master = st.file_uploader(
        "② CPNマスタ",
        type=["xlsx", "xlsm"],
    )

if uploaded_file and uploaded_master:
    try:
        from data.loader import add_business_edge_flags
        from logic.factors import (
            calculate_dynamic_factor_tables,
            calculate_selected_cpn_base,
            get_cpn_reference_periods,
            enforce_premium_media_cost,
        )
        from config.constants import RECENT_NORMAL_DAYS

        history_df = load_data(uploaded_file)
        history_df["date"] = pd.to_datetime(
            history_df["date"],
            errors="coerce",
        ).dt.normalize()

        cpn_master = pd.read_excel(
            uploaded_master,
            engine="openpyxl",
        )

        required_master_columns = {"日付", "CPN名"}
        missing_master = required_master_columns - set(cpn_master.columns)

        if missing_master:
            raise ValueError(
                f"CPNマスタに必要な列がありません: "
                f"{', '.join(sorted(missing_master))}"
            )

        cpn_master = cpn_master.copy()
        cpn_master["日付"] = pd.to_datetime(
            cpn_master["日付"],
            errors="coerce",
        ).dt.normalize()
        cpn_master["CPN名"] = (
            cpn_master["CPN名"]
            .astype("string")
            .str.strip()
        )
        cpn_master = cpn_master.dropna(
            subset=["日付", "CPN名"]
        )
        cpn_master = cpn_master.drop_duplicates(
            subset=["日付"],
            keep="last",
        )

        if cpn_master.empty:
            raise ValueError(
                "CPNマスタに有効な日付・CPN名がありません。"
            )

        # 任意列。未登録なら補正なし。
        cpn_master["line_oa_flag"] = (
            _truthy(cpn_master["LINE OA配信"])
            if "LINE OA配信" in cpn_master
            else 0
        )
        cpn_master["magitoku_after_flag"] = (
            _truthy(cpn_master["マジ得後"])
            if "マジ得後" in cpn_master
            else 0
        )

    except Exception as exc:
        st.error(f"ファイルの読み込みに失敗しました: {exc}")
        st.stop()

    history_df = history_df.merge(
        cpn_master[
            [
                "日付",
                "CPN名",
                "line_oa_flag",
                "magitoku_after_flag",
            ]
        ],
        left_on="date",
        right_on="日付",
        how="left",
    )

    history_df["CPN名"] = history_df["CPN名"].fillna("通常")
    history_df["line_oa_flag"] = (
        history_df["line_oa_flag"]
        .fillna(0)
        .astype(int)
    )
    history_df["magitoku_after_flag"] = (
        history_df["magitoku_after_flag"]
        .fillna(0)
        .astype(int)
    )

    st.sidebar.header("媒体選択")
    all_media = sorted(history_df["media"].unique())
    default_media = [
        m for m in all_media
        if "計測" not in m
    ]

    selected_media = st.sidebar.multiselect(
        "媒体",
        all_media,
        default=default_media,
    )

    if not selected_media:
        st.stop()

    history_df = history_df[
        history_df["media"].isin(selected_media)
    ].copy()

    st.sidebar.header("📊 CPN選択")

    cpn_list = sorted(
        cpn_master["CPN名"]
        .dropna()
        .astype(str)
        .unique()
    )

    default_cpn_index = (
        cpn_list.index("マジ得")
        if "マジ得" in cpn_list
        else 0
    )

    selected_cpn = st.sidebar.selectbox(
        "CPN",
        cpn_list,
        index=default_cpn_index,
    )

    today = datetime.date.today()

    start_date = st.sidebar.date_input(
        "予測開始",
        today,
    )

    end_date = st.sidebar.date_input(
        "予測終了",
        today + datetime.timedelta(days=7),
    )

    if start_date > end_date:
        st.error(
            "予測期間の開始日は終了日以前にしてください。"
        )
        st.stop()

    normal_labels = {"通常", "定常"}
    is_normal_selected = selected_cpn in normal_labels

    if is_normal_selected:
        # 定常は予測期間を365日前へそのままずらした同期間を参照する。
        normal_reference_start = (
            pd.Timestamp(start_date).normalize()
            - pd.Timedelta(days=365)
        )
        normal_reference_end = (
            pd.Timestamp(end_date).normalize()
            - pd.Timedelta(days=365)
        )

        normal_reference = history_df[
            history_df["date"].between(
                normal_reference_start,
                normal_reference_end,
            )
            & history_df["CPN名"].isin(normal_labels)
        ].copy()

        base_pair = _daily_pair_average(normal_reference)
        base_pair = base_pair[
            base_pair["media"].isin(selected_media)
        ]

        if base_pair.empty:
            st.error(
                "365日前の同期間に定常実績がありません。"
                "実績CSVとCPNマスタの『定常／通常』登録を確認してください。"
            )
            st.stop()

        st.subheader("📈 前年同期間の定常実績")
        st.caption(
            f"参照期間: "
            f"{normal_reference_start.date()} ～ "
            f"{normal_reference_end.date()}"
            "（予測期間の365日前）"
        )

        normal_display = (
            normal_reference.groupby(
                ["date", "media"],
                as_index=False,
            )["cv"]
            .sum()
            .groupby(
                "media",
                as_index=False,
            )["cv"]
            .mean()
            .rename(
                columns={
                    "media": "媒体",
                    "cv": "前年同期間の定常CV/日",
                }
            )
        )

        normal_display["前年同期間の定常CV/日"] = (
            normal_display["前年同期間の定常CV/日"]
            .round(2)
        )

        st.dataframe(
            normal_display,
            use_container_width=True,
            hide_index=True,
        )

    else:
        # 実績内に存在する同一CPNの連続期間を候補化し、複数選択できるようにする。
        available_periods = get_cpn_reference_periods(
            history_df,
            selected_cpn,
        )

        if not available_periods:
            st.error(
                f"実績内に『{selected_cpn}』のキャンペーン期間がありません。"
            )
            st.stop()

        period_options = {
            f"{start.strftime('%Y/%m/%d')} ～ "
            f"{end.strftime('%Y/%m/%d')}": (start, end)
            for start, end in available_periods
        }

        selected_period_labels = st.sidebar.multiselect(
            "CPN参照期間（複数選択可）",
            options=list(period_options.keys()),
            default=list(period_options.keys()),
        )

        if not selected_period_labels:
            st.warning(
                "CPN参照期間を1つ以上選択してください。"
            )
            st.stop()

        selected_periods = [
            period_options[label]
            for label in selected_period_labels
        ]

        base_pair = calculate_selected_cpn_base(
            history_df,
            selected_cpn,
            selected_periods,
        )

        base_pair = base_pair[
            base_pair["media"].isin(selected_media)
        ]

        if base_pair.empty:
            st.error(
                "選択したCPN参照期間に対象媒体の実績がありません。"
            )
            st.stop()

        total_reference_days = sum(
            (end - start).days + 1
            for start, end in selected_periods
        )

        st.subheader("📈 選択CPN期間の実績")
        st.caption(
            f"選択期間: {len(selected_periods)}期間 / "
            f"合計 {total_reference_days}日。"
            " 選択期間のCV・COST合計を合計日数で割った"
            "日平均を予測ベースに使用します。"
        )

        for label in selected_period_labels:
            st.write(f"・{label}")

        cpn_display = (
            base_pair.groupby(
                "media",
                as_index=False,
            )
            .agg(
                **{
                    "選択期間CV/日": ("base_cv", "sum"),
                    "選択期間COST/日": ("cost", "sum"),
                }
            )
            .rename(
                columns={"media": "媒体"}
            )
        )

        cpn_display["選択期間CV/日"] = (
            cpn_display["選択期間CV/日"]
            .round(2)
        )

        cpn_display["選択期間COST/日"] = (
            cpn_display["選択期間COST/日"]
            .round(0)
        )

        st.dataframe(
            cpn_display,
            use_container_width=True,
            hide_index=True,
        )

    factor_tables = calculate_dynamic_factor_tables(
        history_df
    )

    st.subheader("📐 実績から算出した変動係数")

    tab1, tab2, tab3, tab4 = st.tabs(
        ["曜日", "月初・月末", "需要期", "LINE OA"]
    )

    with tab1:
        st.dataframe(
            factor_tables["weekday"].round(3),
            use_container_width=True,
            hide_index=True,
        )

    with tab2:
        st.dataframe(
            factor_tables["month_edge"].round(3),
            use_container_width=True,
            hide_index=True,
        )

    with tab3:
        st.dataframe(
            factor_tables["season"].round(3),
            use_container_width=True,
            hide_index=True,
        )

    with tab4:
        if factor_tables["line_oa"].empty:
            st.info(
                "過去のLINE OA配信実績がないため、"
                "LINE OA係数は1.0です。"
            )
        else:
            st.dataframe(
                factor_tables["line_oa"].round(3),
                use_container_width=True,
                hide_index=True,
            )

    future_dates = pd.date_range(
        start=start_date,
        end=end_date,
    )

    future_df = pd.DataFrame(
        {"date": future_dates}
    ).merge(
        base_pair,
        how="cross",
    )

    future_df["weekday"] = (
        future_df["date"].dt.day_name()
    )

    future_df = add_business_edge_flags(
        future_df
    )

    future_df = future_df.merge(
        cpn_master[
            [
                "日付",
                "CPN名",
                "line_oa_flag",
                "magitoku_after_flag",
            ]
        ],
        left_on="date",
        right_on="日付",
        how="left",
    )

    future_df["CPN名"] = (
        future_df["CPN名"]
        .fillna(selected_cpn)
    )

    future_df["line_oa_flag"] = (
        future_df["line_oa_flag"]
        .fillna(0)
        .astype(int)
    )

    future_df["magitoku_after_flag"] = (
        future_df["magitoku_after_flag"]
        .fillna(0)
        .astype(int)
    )

    # キャンペーン平均をbase_cvとして直接使用するため、CPN倍率は掛けない。
    future_df["cpn_factor"] = 1.0

    # 曜日・月初月末・月別需要期・LINE OAは、アップロード実績から毎回算出。
    forecast_df = forecast_cv(
        future_df,
        factor_tables,
    )

    forecast_df = enforce_premium_media_cost(
        forecast_df
    )

    # 係数確認用の明細
    with st.expander("予測係数の確認"):
        factor_cols = [
            "date",
            "media",
            "商品ID",
            "base_cv",
            "cpn_factor",
            "weekday_factor",
            "season_factor",
            "month_edge_factor",
            "after_factor",
            "line_factor",
            "forecast_cv",
            "cost",
        ]

        st.dataframe(
            forecast_df[factor_cols],
            use_container_width=True,
            hide_index=True,
        )

    forecast_df["date"] = format_date(
        forecast_df
    )

    sim_df = simulate_plan(
        forecast_df
    )

    sim_summary = (
        sim_df.groupby(
            ["date", "media", "plan"],
            as_index=False,
        )
        .agg(
            cv=("cv", "sum"),
            cost=("cost", "sum"),
        )
    )

    sim_summary["cpa"] = (
        (sim_summary["cost"] / sim_summary["cv"])
        .replace(
            [float("inf"), float("-inf")],
            0,
        )
        .fillna(0)
    )

    sim_summary["date"] = format_date(
        sim_summary
    )

    st.subheader("📊 松竹梅")
    st.dataframe(
        create_report_table(sim_summary),
        use_container_width=True,
    )

    st.sidebar.header("🎯 最適化ロジック")

    opt_mode = st.sidebar.radio(
        "最適基準",
        ["CPA最小", "CV最大"],
        index=0,
    )

    opt_df = optimize_budget(
        sim_df,
        opt_mode,
    )

    opt_summary = (
        opt_df.groupby(
            ["date", "media", "plan"],
            as_index=False,
        )
        .agg(
            cv=("cv", "sum"),
            cost=("cost", "sum"),
        )
    )

    opt_summary["cpa"] = (
        (opt_summary["cost"] / opt_summary["cv"])
        .replace(
            [float("inf"), float("-inf")],
            0,
        )
        .fillna(0)
    )

    opt_summary["date"] = format_date(
        opt_summary
    )

    st.subheader("🚀 最適プラン")
    st.dataframe(
        create_report_table(opt_summary),
        use_container_width=True,
    )

    target_cv = st.sidebar.number_input(
        "目標",
        min_value=0,
        value=1000,
    )

    gap = target_cv - forecast_df["forecast_cv"].sum()
    st.write(f"差分: {gap:.0f}")

    submission_filename = (
        f"【提出用】楽天カード"
        f"{pd.Timestamp(start_date).year}年"
        f"{pd.Timestamp(start_date).month}月"
        f"プランニング.xlsx"
    )

    # ---------------------------------------------------------
    # 提出用Excelは画面表示のたびに自動生成しない。
    # Streamlitはウィジェット操作のたびに上から再実行されるため、
    # create_submission_excel() を常時実行するとDLボタン表示まで重くなる。
    # 「生成」ボタンを押した時だけ作成し、session_stateに保持する。
    # ---------------------------------------------------------
    submission_key = (
        f"{start_date}_{end_date}_{selected_cpn}_{opt_mode}_"
        f"{','.join(map(str, selected_media))}"
    )

    # 条件が変わったら古い成果物を破棄
    if st.session_state.get("submission_key") != submission_key:
        st.session_state.pop("submission_excel", None)
        st.session_state["submission_key"] = submission_key

    if st.button(
        "📄 提出用Excelを生成",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("提出用Excelを作成しています..."):
            st.session_state["submission_excel"] = create_submission_excel(
                opt_summary=opt_summary,
                history_df=history_df,
                start_date=start_date,
                end_date=end_date,
                selected_cpn=selected_cpn,
                opt_mode=opt_mode,
            )

    if "submission_excel" in st.session_state:
        st.success("提出用Excelの作成が完了しました。")
        st.download_button(
            "📥 最適プランを提出用ExcelでDL",
            data=st.session_state["submission_excel"],
            file_name=submission_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )
