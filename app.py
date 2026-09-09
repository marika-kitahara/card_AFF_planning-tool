# -*- coding: utf-8 -*-
# SUBMISSION_TEMPLATE_ALL_SHEETS_V5 = 2026-08-14
# MergedCell-safe version

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import datetime
import base64
import glob
from io import BytesIO

import matplotlib.pyplot as plt
from matplotlib import font_manager

# Streamlit Cloudでもaptに依存せず日本語を描画できるようにする。
try:
    import japanize_matplotlib  # noqa: F401
except Exception:
    japanize_matplotlib = None
from pathlib import Path
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from data.loader import load_data, load_master_workbook_with_af, load_af_data, load_af_code_data
from logic.forecast import forecast_cv
from logic.simulation import simulate_plan
from logic.optimize import optimize_budget
from logic.analytics import prepare_monthly_performance, prepare_unit_price_band_matrix


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
    """
    松竹梅・最適プラン表示用の高速版。
    媒体×プランごとのループとDataFrame大量生成をやめ、
    melt + pivot_table で一括変換する。
    """
    if df.empty:
        return pd.DataFrame()

    work = df[["date", "media", "plan", "cv", "cost", "cpa"]].copy()

    long_df = work.melt(
        id_vars=["media", "plan", "date"],
        value_vars=["cv", "cost", "cpa"],
        var_name="metric",
        value_name="value",
    )

    metric_map = {
        "cv": "CV",
        "cost": "COST",
        "cpa": "単価",
    }
    metric_order = {
        "CV": 0,
        "COST": 1,
        "単価": 2,
    }

    long_df["metric"] = long_df["metric"].map(metric_map)
    long_df["_metric_order"] = long_df["metric"].map(metric_order)

    result = (
        long_df.pivot_table(
            index=["media", "plan", "metric", "_metric_order"],
            columns="date",
            values="value",
            aggfunc="sum",
        )
        .reset_index()
        .sort_values(
            ["media", "plan", "_metric_order"],
            kind="stable",
        )
        .drop(columns="_metric_order")
        .reset_index(drop=True)
    )

    date_cols = [
        c for c in result.columns
        if c not in ["media", "plan", "metric"]
    ]
    result = result[["media", "plan", "metric"] + date_cols]

    result["media"] = result["media"].mask(
        result["media"].duplicated()
    )

    # planは3行（CV/COST/単価）の先頭だけ表示
    result["plan"] = result["plan"].mask(
        result["plan"].eq(result["plan"].shift())
        & result["media"].isna()
    )

    return result


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
        base / "assets" / "submission_template_v2.xlsx",
        base / "submission_template_v2.xlsx",
        base / "assets" / "submission_template_fast.xlsx",
        base / "submission_template_fast.xlsx",
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


def _set_value(ws, row, col, value):
    """
    結合セル対策付きの値セット。
    左上セル以外の MergedCell に当たった場合は何もしない。
    """
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value


def _set_percent(ws, row, col, value):
    """0.583 をExcel上で58.3%表示にする。"""
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        return
    cell.value = value
    cell.number_format = "0.0%"




def _calculate_media_approval_rates(history_df: pd.DataFrame) -> dict:
    """
    過去実績の『成果承認フラグ = Y』を発行数として、媒体別承認率を算出する。

    承認率 = 成果承認フラグYの件数合計 / 全発生件数合計

    媒体に承認実績がない場合は全媒体の加重承認率をフォールバックに使う。
    全体でも承認実績がない場合はエラーにする。
    """
    required = {"media", "approved_cv", "approval_base_cv"}
    missing = required - set(history_df.columns)

    if missing:
        raise ValueError(
            "承認率算出用データがありません。"
            "data/loader.py を承認率対応版へ更新してください。"
        )

    work = history_df[
        ["media", "approved_cv", "approval_base_cv"]
    ].copy()

    work["approved_cv"] = pd.to_numeric(
        work["approved_cv"],
        errors="coerce",
    )
    work["approval_base_cv"] = pd.to_numeric(
        work["approval_base_cv"],
        errors="coerce",
    )

    valid = work[
        work["approval_base_cv"].notna()
        & (work["approval_base_cv"] > 0)
    ].copy()

    if valid.empty:
        source_col = ""
        if "approval_source_column" in history_df.columns:
            sources = (
                history_df["approval_source_column"]
                .dropna()
                .astype(str)
                .str.strip()
            )
            sources = sources[sources != ""]
            if not sources.empty:
                source_col = sources.iloc[0]

        if source_col:
            raise ValueError(
                f"承認列『{source_col}』は見つかりましたが、"
                "Y/N・1/0・承認/否認として判定できる実績がありません。"
            )

        raise ValueError(
            "実績CSVから承認判定列を見つけられませんでした。"
            "『承認フラグ』『承認状況』『承認ステータス』『承認』"
            "などの列を確認してください。"
        )

    total_base = valid["approval_base_cv"].sum()
    total_approved = valid["approved_cv"].fillna(0).sum()

    overall_rate = (
        float(total_approved / total_base)
        if total_base > 0
        else 0.0
    )
    overall_rate = min(max(overall_rate, 0.0), 1.0)

    media_rates = (
        valid.groupby("media", as_index=False)
        .agg(
            approved_cv=("approved_cv", "sum"),
            approval_base_cv=("approval_base_cv", "sum"),
        )
    )

    media_rates["approval_rate"] = (
        media_rates["approved_cv"]
        / media_rates["approval_base_cv"]
    ).clip(0, 1)

    rate_map = dict(
        zip(
            media_rates["media"].astype(str),
            media_rates["approval_rate"].astype(float),
        )
    )

    # どの媒体でも必ず率が取れるよう全体率を保持
    rate_map["__overall__"] = overall_rate
    return rate_map



def _calculate_period_media_metrics(history_df: pd.DataFrame):
    """
    媒体別に定常・マジ得の過去実績指標を作る。

    承認率 = Y件数 / (Y + D)件数
    グロス単価 = ローデータW列「グロス」の直近実績

    COST/発生CVから単価を逆算しない。
    営業プランで使う「単価」はW列グロスを正とする。
    """
    required = {
        "media", "CPN名", "cv", "cost",
        "approved_cv", "approval_base_cv",
    }
    missing = required - set(history_df.columns)
    if missing:
        raise ValueError(
            "定常・マジ得指標の算出に必要な列がありません: "
            + ", ".join(sorted(missing))
        )

    use_cols = [
        "media", "CPN名", "date", "cv", "cost",
        "approved_cv", "approval_base_cv",
    ]
    if "gross_unit" in history_df.columns:
        use_cols.append("gross_unit")

    work = history_df[use_cols].copy()
    work["media"] = work["media"].astype(str)
    work["CPN名"] = work["CPN名"].astype(str).str.strip()
    work["date"] = pd.to_datetime(work["date"], errors="coerce")
    for col in ["cv", "cost", "approved_cv", "approval_base_cv"]:
        work[col] = pd.to_numeric(work[col], errors="coerce").fillna(0)
    if "gross_unit" not in work.columns:
        work["gross_unit"] = 0.0
    work["gross_unit"] = pd.to_numeric(work["gross_unit"], errors="coerce").fillna(0.0)

    def build(mask):
        sub = work[mask].copy()
        if sub.empty:
            return {}, 0.0, {}, 0.0

        agg = (
            sub.groupby("media", as_index=False)
            .agg(
                cv=("cv", "sum"),
                cost=("cost", "sum"),
                approved_cv=("approved_cv", "sum"),
                approval_base_cv=("approval_base_cv", "sum"),
            )
        )
        agg["approval_rate"] = (
            agg["approved_cv"]
            / agg["approval_base_cv"].replace(0, pd.NA)
        ).fillna(0).clip(0, 1)

        total_base = sub["approval_base_cv"].sum()
        overall_rate = (
            float(sub["approved_cv"].sum() / total_base)
            if total_base > 0 else 0.0
        )

        # 媒体ごとの「最新の正のグロス単価」を採用。
        # 同一最新日に複数行ある場合はCV加重平均。
        positive = sub[sub["gross_unit"].gt(0) & sub["date"].notna()].copy()
        unit_map = {}
        if not positive.empty:
            positive["gross_weight"] = positive["gross_unit"] * positive["cv"].clip(lower=0)
            daily_unit = (
                positive.groupby(["media", "date"], as_index=False)
                .agg(gross_weight=("gross_weight", "sum"), unit_cv=("cv", "sum"))
            )
            daily_unit["unit"] = (
                daily_unit["gross_weight"]
                / daily_unit["unit_cv"].replace(0, pd.NA)
            )
            daily_unit = daily_unit[daily_unit["unit"].gt(0)].sort_values("date")
            latest = daily_unit.groupby("media", as_index=False).tail(1)
            unit_map = dict(zip(latest["media"], latest["unit"].astype(float)))

        # フォールバックは全該当実績のCV加重グロス単価。
        total_weight = float((positive["gross_unit"] * positive["cv"].clip(lower=0)).sum()) if not positive.empty else 0.0
        total_unit_cv = float(positive["cv"].clip(lower=0).sum()) if not positive.empty else 0.0
        overall_unit = total_weight / total_unit_cv if total_unit_cv > 0 else 0.0

        rate_map = dict(zip(agg["media"], agg["approval_rate"].astype(float)))
        return rate_map, overall_rate, unit_map, overall_unit

    normal_mask = work["CPN名"].isin(["通常", "定常"])
    magi_mask = work["CPN名"].eq("マジ得")

    normal_rate_map, normal_rate_all, normal_unit_map, normal_unit_all = build(normal_mask)
    magi_rate_map, magi_rate_all, magi_unit_map, magi_unit_all = build(magi_mask)

    return {
        "normal_rate": normal_rate_map,
        "normal_rate_all": normal_rate_all,
        "normal_unit": normal_unit_map,
        "normal_unit_all": normal_unit_all,
        "magi_rate": magi_rate_map,
        "magi_rate_all": magi_rate_all,
        "magi_unit": magi_unit_map,
        "magi_unit_all": magi_unit_all,
    }


def _future_period_name(date_value, future_cpn_map, selected_cpn):
    """
    未来日を定常/マジ得に分類する。
    CPNマスタに当日登録があればそれを優先。
    未登録の場合はUI選択CPNを使用。
    """
    cpn_name = future_cpn_map.get(
        pd.Timestamp(date_value).normalize(),
        selected_cpn,
    )
    cpn_name = str(cpn_name).strip()
    return "マジ得" if cpn_name == "マジ得" else "定常"



def _build_manual_settings_defaults(
    opt_summary: pd.DataFrame,
    history_df: pd.DataFrame,
    selected_cpn: str,
) -> pd.DataFrame:
    """
    最適プラン / 提案用Excelの計算値を、手動設定テーブルの初期値へ変換する。

    編集可能:
      今回プラン採用グロス単価
      今回プラン採用承認率
      今回採用件数

    自動計算:
      承認件数 = 今回採用件数 × 今回プラン採用承認率
      費用     = 承認件数 × 今回プラン採用グロス単価
      発行CPA = 費用 ÷ 承認件数
    """
    plan = opt_summary.copy()
    plan["media"] = plan["media"].astype(str)
    plan["cv"] = pd.to_numeric(plan["cv"], errors="coerce").fillna(0)
    plan["cost"] = pd.to_numeric(plan["cost"], errors="coerce").fillna(0)

    plan["cpa"] = pd.to_numeric(plan.get("cpa", 0), errors="coerce").fillna(0)
    plan["_unit_weight"] = plan["cpa"] * plan["cv"]
    totals = (
        plan.groupby("media", as_index=False)
        .agg(
            plan_cv=("cv", "sum"),
            plan_cost=("cost", "sum"),
            unit_weight=("_unit_weight", "sum"),
        )
    )
    totals["opt_unit"] = (
        totals["unit_weight"]
        / totals["plan_cv"].replace(0, pd.NA)
    ).fillna(0)

    period = _calculate_period_media_metrics(history_df)

    # SID
    sid_map = {}
    if "SID" in history_df.columns:
        sid_source = history_df[["media", "SID"]].copy()
        sid_source["media"] = sid_source["media"].astype(str)
        sid_source["SID"] = (
            sid_source["SID"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        sid_source = sid_source[sid_source["SID"] != ""]
        sid_map = (
            sid_source.groupby("media")["SID"]
            .agg(lambda x: " / ".join(dict.fromkeys(x.tolist())))
            .to_dict()
        )

    rows = []
    for r in totals.itertuples():
        media = str(r.media)
        opt_unit = float(r.opt_unit or 0)

        # 複合施策時は、施策別予測CVを重みにして承認率・グロス単価をブレンドする。
        media_plan = plan[plan["media"].eq(media)].copy()
        if "CPN名" in media_plan.columns and not media_plan.empty:
            mix = media_plan.groupby("CPN名", as_index=False)["cv"].sum()
            weighted_rate = 0.0
            weighted_gross = 0.0
            weight_total = float(mix["cv"].sum())
            for mix_row in mix.itertuples():
                cpn_name = str(mix_row.CPN名).strip()
                weight = float(mix_row.cv or 0)
                if cpn_name == "マジ得":
                    cpn_rate = period["magi_rate"].get(media, period["magi_rate_all"])
                    cpn_gross = period["magi_unit"].get(media, period["magi_unit_all"])
                    if cpn_gross <= 0:
                        cpn_gross = opt_unit
                else:
                    cpn_rate = period["normal_rate"].get(media, period["normal_rate_all"])
                    cpn_gross = period["normal_unit"].get(media, period["normal_unit_all"])
                    if cpn_gross <= 0:
                        cpn_gross = opt_unit
                weighted_rate += weight * float(cpn_rate or 0)
                weighted_gross += weight * float(cpn_gross or 0)
            rate = weighted_rate / weight_total if weight_total else 0.0
            gross_unit = weighted_gross / weight_total if weight_total else opt_unit
        elif selected_cpn == "マジ得":
            rate = period["magi_rate"].get(media, period["magi_rate_all"])
            gross_unit = period["magi_unit"].get(media, period["magi_unit_all"])
            if gross_unit <= 0:
                gross_unit = opt_unit
        else:
            rate = period["normal_rate"].get(media, period["normal_rate_all"])
            gross_unit = period["normal_unit"].get(media, period["normal_unit_all"])
            if gross_unit <= 0:
                gross_unit = opt_unit

        adopted_count = float(r.plan_cv or 0)
        approved_count = adopted_count * float(rate or 0)
        cost = approved_count * float(gross_unit or 0)
        issue_cpa = cost / approved_count if approved_count else 0

        rows.append(
            {
                "SID": sid_map.get(media, ""),
                "媒体名": media,
                "今回プラン採用グロス単価": round(gross_unit),
                "今回プラン採用承認率": float(rate or 0),
                "今回採用件数": round(adopted_count),
                "費用": round(cost),
                "承認件数": round(approved_count, 1),
                "発行CPA": round(issue_cpa),
            }
        )

    result = pd.DataFrame(rows)

    if not result.empty:
        result = result.sort_values(
            ["今回採用件数", "費用", "媒体名"],
            ascending=[False, False, True],
        ).reset_index(drop=True)

    return result


def _normalize_manual_settings(df: pd.DataFrame) -> pd.DataFrame:
    """手動設定の入力値を安全に数値化し、派生値を再計算する。"""
    out = df.copy()

    required_cols = [
        "SID",
        "媒体名",
        "今回プラン採用グロス単価",
        "今回プラン採用承認率",
        "今回採用件数",
        "費用",
    ]
    for col in required_cols:
        if col not in out.columns:
            out[col] = "" if col in {"SID", "媒体名"} else 0

    for col in [
        "今回プラン採用グロス単価",
        "今回プラン採用承認率",
        "今回採用件数",
        "費用",
    ]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    out["今回プラン採用グロス単価"] = (
        out["今回プラン採用グロス単価"].clip(lower=0)
    )
    out["今回プラン採用承認率"] = (
        out["今回プラン採用承認率"].clip(lower=0, upper=1)
    )
    out["今回採用件数"] = out["今回採用件数"].clip(lower=0)
    out["費用"] = out["費用"].clip(lower=0)

    out["承認件数"] = (
        out["今回採用件数"]
        * out["今回プラン採用承認率"]
    )
    # 営業プランの定義に合わせ、費用は編集値ではなく自動計算する。
    # 発行コスト = 発行見込み(承認件数) × 今回プラン採用グロス単価
    out["費用"] = (
        out["承認件数"]
        * out["今回プラン採用グロス単価"]
    )

    out["発行CPA"] = (
        out["費用"]
        / out["承認件数"].replace(0, pd.NA)
    ).fillna(0)

    out["今回採用件数"] = out["今回採用件数"].round(0)
    out["費用"] = out["費用"].round(0)
    out["承認件数"] = out["承認件数"].round(1)
    out["発行CPA"] = out["発行CPA"].round(0)

    return out[
        [
            "SID",
            "媒体名",
            "今回プラン採用グロス単価",
            "今回プラン採用承認率",
            "今回採用件数",
            "費用",
            "承認件数",
            "発行CPA",
        ]
    ]


def _manual_settings_signature(df: pd.DataFrame):
    """session_state更新判定用。"""
    cols = [
        "媒体名",
        "今回プラン採用グロス単価",
        "今回プラン採用承認率",
        "今回採用件数",
        "費用",
    ]
    if df is None or df.empty:
        return ()
    work = df[cols].copy()
    return tuple(
        tuple(row)
        for row in work.astype(object).itertuples(index=False, name=None)
    )


def render_manual_settings(
    opt_summary,
    history_df,
    selected_cpn,
    calc_key,
):
    """
    手動設定エディタの安全版。

    st.fragment / st.rerun は使用しない。
    data_editor の編集時はStreamlit標準の再実行に任せる。
    予測・最適化結果は既存のsession_stateキャッシュを再利用するため、
    重い再計算は発生しない。
    """
    if st.session_state.get("_manual_calc_key") != calc_key:
        st.session_state["_manual_calc_key"] = calc_key
        st.session_state["_manual_settings"] = _build_manual_settings_defaults(
            opt_summary=opt_summary,
            history_df=history_df,
            selected_cpn=selected_cpn,
        )

        # 計算条件が変わった時だけEditorのwidget stateもリセット
        old_widget_key = st.session_state.get("_manual_widget_key")
        if old_widget_key:
            st.session_state.pop(old_widget_key, None)

        st.session_state["_manual_widget_key"] = (
            f"manual_settings_editor_{abs(hash(str(calc_key)))}"
        )

    current = _normalize_manual_settings(
        st.session_state.get(
            "_manual_settings",
            _build_manual_settings_defaults(
                opt_summary=opt_summary,
                history_df=history_df,
                selected_cpn=selected_cpn,
            ),
        )
    )

    widget_key = st.session_state.get(
        "_manual_widget_key",
        f"manual_settings_editor_{abs(hash(str(calc_key)))}",
    )

    edited = st.data_editor(
        current,
        key=widget_key,
        width="stretch",
        hide_index=True,
        num_rows="fixed",
        disabled=[
            "SID",
            "媒体名",
            "費用",
            "承認件数",
            "発行CPA",
        ],
        column_config={
            "SID": st.column_config.TextColumn(
                "SID",
                width="small",
            ),
            "媒体名": st.column_config.TextColumn(
                "媒体名",
                width="large",
            ),
            "今回プラン採用グロス単価": st.column_config.NumberColumn(
                "今回プラン採用グロス単価",
                min_value=0,
                step=100,
                format="¥%d",
            ),
            "今回プラン採用承認率": st.column_config.NumberColumn(
                "今回プラン採用承認率",
                min_value=0.0,
                max_value=1.0,
                step=0.001,
                format="%.3f",
                help="0.50 = 50% として入力",
            ),
            "今回採用件数": st.column_config.NumberColumn(
                "今回採用件数",
                min_value=0,
                step=1,
                format="%d",
            ),
            "費用": st.column_config.NumberColumn(
                "費用",
                min_value=0,
                step=1000,
                format="¥%d",
            ),
            "承認件数": st.column_config.NumberColumn(
                "承認件数",
                format="%.1f",
            ),
            "発行CPA": st.column_config.NumberColumn(
                "発行CPA",
                format="¥%d",
            ),
        },
    )

    normalized = _normalize_manual_settings(edited)
    st.session_state["_manual_settings"] = normalized

    # 派生値はEditorの下に必ず最新値を表示。
    total_count = normalized["今回採用件数"].sum()
    total_issue = normalized["承認件数"].sum()
    total_cost = normalized["費用"].sum()
    overall_rate = total_issue / total_count if total_count else 0
    overall_cpa = total_cost / total_issue if total_issue else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("今回採用件数 合計", f"{total_count:,.0f}")
    c2.metric("承認件数 合計", f"{total_issue:,.1f}")
    c3.metric("全体承認率", f"{overall_rate:.1%}")
    c4.metric("全体発行CPA", f"¥{overall_cpa:,.0f}")

    # 自動計算結果の確認用
    derived = normalized[
        [
            "媒体名",
            "費用",
            "承認件数",
            "発行CPA",
        ]
    ].copy()

    with st.expander("自動計算結果を確認"):
        st.dataframe(
            derived,
            width="stretch",
            hide_index=True,
            column_config={
                "媒体名": st.column_config.TextColumn("媒体名"),
                "承認件数": st.column_config.NumberColumn(
                    "承認件数",
                    format="%.1f",
                ),
                "発行CPA": st.column_config.NumberColumn(
                    "発行CPA",
                    format="¥%d",
                ),
            },
        )



def _resolve_submission_cpn_month(cpn_master, start_date, end_date=None):
    """提出用Excelの月度をCPNマスタ基準で解決する。"""
    fallback_ts = pd.Timestamp(start_date)
    fallback_label = f"{fallback_ts.year}年{fallback_ts.month}月度"

    if cpn_master is None or len(cpn_master) == 0:
        return fallback_label, fallback_ts.year, fallback_ts.month
    if "日付" not in cpn_master.columns or "月度" not in cpn_master.columns:
        return fallback_label, fallback_ts.year, fallback_ts.month

    master = cpn_master[["日付", "月度"]].copy()
    master["日付"] = pd.to_datetime(master["日付"], errors="coerce").dt.normalize()
    master["月度"] = master["月度"].astype("string").str.strip()
    master = master.dropna(subset=["日付", "月度"])
    master = master.loc[master["月度"].ne("") & master["月度"].ne("未設定")]
    if master.empty:
        return fallback_label, fallback_ts.year, fallback_ts.month

    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize() if end_date is not None else start

    # まず予測開始日そのもののCPN月度を正とする。
    exact = master.loc[master["日付"].eq(start), "月度"]
    if not exact.empty:
        label = str(exact.iloc[0])
    else:
        # 開始日に行がない場合だけ、予測対象期間内で最頻の月度を採用する。
        in_range = master.loc[master["日付"].between(start, end), "月度"]
        if in_range.empty:
            return fallback_label, fallback_ts.year, fallback_ts.month
        label = str(in_range.mode().iloc[0])

    import re
    m = re.search(r"(\d{4})年\s*(\d{1,2})月度?", label)
    if m:
        return label, int(m.group(1)), int(m.group(2))
    return label, fallback_ts.year, fallback_ts.month


def _copy_cell_style_and_format(src, dst):
    """セルの値以外の表示設定を複製する。"""
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.protection:
        dst.protection = copy(src.protection)


def _extend_sheet_date_columns(ws, first_date_col, base_slots, required_slots, total_col):
    """33日テンプレートを必要日数まで右方向へ拡張し、Total列を後ろへ送る。"""
    required_slots = max(int(required_slots), int(base_slots))
    extra = required_slots - int(base_slots)
    if extra <= 0:
        return total_col

    source_col = total_col - 1  # 既存の最後の日付列
    ws.insert_cols(total_col, amount=extra)

    src_letter = get_column_letter(source_col)
    src_width = ws.column_dimensions[src_letter].width
    for offset in range(extra):
        new_col = total_col + offset
        new_letter = get_column_letter(new_col)
        if src_width is not None:
            ws.column_dimensions[new_letter].width = src_width
        for row in range(1, ws.max_row + 1):
            _copy_cell_style_and_format(ws.cell(row, source_col), ws.cell(row, new_col))
            ws.cell(row, new_col).value = None

    return total_col + extra


def _extend_summary_date_rows(ws, base_slots, required_slots, first_data_row=3, total_row=36):
    """全体サマリの33日分行を必要日数まで下方向へ拡張する。"""
    required_slots = max(int(required_slots), int(base_slots))
    extra = required_slots - int(base_slots)
    if extra <= 0:
        return total_row

    source_row = total_row - 1
    ws.insert_rows(total_row, amount=extra)
    src_height = ws.row_dimensions[source_row].height
    for offset in range(extra):
        new_row = total_row + offset
        if src_height is not None:
            ws.row_dimensions[new_row].height = src_height
        for col in range(1, ws.max_column + 1):
            _copy_cell_style_and_format(ws.cell(source_row, col), ws.cell(new_row, col))
            ws.cell(new_row, col).value = None

    return total_row + extra


def create_submission_excel(opt_summary, history_df, cpn_master, manual_settings, start_date, end_date, selected_cpn, opt_mode):
    """
    添付された提出用Excelそのものをテンプレートとして使い、
    最適プランを初期値とした手動設定の結果を提出用Excelへ反映する。

    - SID: 実績データF列を loader.py で保持した history_df["SID"]
    - 件数: 最適プランCV
    - コスト: 最適プランcost
    - 承認率: 過去実績から媒体別に算出
    - 発行: 最適プランForecast × 媒体別承認率
    - Actual: 提出時点では未入力

    ※テンプレ内の既存媒体名や数値は、出力時に対象範囲をクリアしてから
      最適プランの値へ置き換える。
    ※元テンプレートそのものは変更しない。
    """
    submission_month_label, submission_year, submission_month = _resolve_submission_cpn_month(
        cpn_master, start_date, end_date
    )

    plan = opt_summary.copy()
    plan["date"] = pd.to_datetime(plan["date"], errors="coerce").dt.normalize()
    plan = plan.dropna(subset=["date", "media"])
    plan["media"] = plan["media"].astype(str)
    plan["cv"] = pd.to_numeric(plan["cv"], errors="coerce").fillna(0)
    plan["cost"] = pd.to_numeric(plan["cost"], errors="coerce").fillna(0)

    # 複合施策では「開始～終了の連続日」ではなく、実際に予測した対象日だけを出力する。
    dates = sorted(pd.Timestamp(x).normalize() for x in plan["date"].dropna().unique())

    daily = (
        plan.groupby(["date", "media"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    daily["cpa"] = (
        (daily["cost"] / daily["cv"])
        .replace([float("inf"), float("-inf")], 0)
        .fillna(0)
    )

    base_media_totals = (
        daily.groupby("media", as_index=False)
        .agg(
            total_cv=("cv", "sum"),
            total_cost=("cost", "sum"),
        )
    )
    base_media_totals["media"] = base_media_totals["media"].astype(str)

    # ---------------------------------------------------------
    # 手動設定を最終提出値として採用。
    # 日次件数は、元の最適プラン日次構成比を維持して再配分する。
    # ---------------------------------------------------------
    manual = _normalize_manual_settings(manual_settings)
    manual["媒体名"] = manual["媒体名"].astype(str)

    base_media_set = set(base_media_totals["media"].tolist())
    manual = manual[manual["媒体名"].isin(base_media_set)].copy()

    manual = manual.sort_values(
        ["今回採用件数", "費用", "媒体名"],
        ascending=[False, False, True],
    ).reset_index(drop=True)

    media_list = manual["媒体名"].tolist()

    MAX_TEMPLATE_MEDIA = 150
    if len(media_list) > MAX_TEMPLATE_MEDIA:
        raise ValueError(
            f"提出用Excelは最大{MAX_TEMPLATE_MEDIA}媒体まで対応しています。"
            f"現在は{len(media_list)}媒体です。"
        )

    manual_count_map = dict(
        zip(
            manual["媒体名"],
            manual["今回採用件数"].astype(float),
        )
    )
    manual_rate_map = dict(
        zip(
            manual["媒体名"],
            manual["今回プラン採用承認率"].astype(float),
        )
    )
    manual_cost_total_map = dict(
        zip(
            manual["媒体名"],
            manual["費用"].astype(float),
        )
    )
    manual_gross_unit_map = dict(
        zip(
            manual["媒体名"],
            manual["今回プラン採用グロス単価"].astype(float),
        )
    )

    original_total_cv_map = dict(
        zip(
            base_media_totals["media"],
            base_media_totals["total_cv"].astype(float),
        )
    )

    # 元日次Forecastを今回採用件数へ比例配分。
    cv_map = {}
    for row in daily.itertuples():
        dt = pd.Timestamp(row.date).normalize()
        media = str(row.media)
        original_total = original_total_cv_map.get(media, 0.0)
        adopted_total = manual_count_map.get(media, 0.0)

        scale = (
            adopted_total / original_total
            if original_total > 0
            else 0.0
        )
        cv_map[(dt, media)] = float(row.cv) * scale

    # 端数差が出ても媒体Totalが手動入力値と一致するよう、最後の日へ差分を寄せる。
    dates_by_media = {}
    for dt, media in cv_map.keys():
        dates_by_media.setdefault(media, []).append(dt)

    for media in media_list:
        media_dates = sorted(dates_by_media.get(media, []))
        if not media_dates:
            continue

        current_total = sum(
            cv_map.get((dt, media), 0.0)
            for dt in media_dates
        )
        diff = manual_count_map.get(media, 0.0) - current_total
        cv_map[(media_dates[-1], media)] = (
            cv_map.get((media_dates[-1], media), 0.0)
            + diff
        )

    # 過去実績の定常・マジ得指標は、提出用Excelの参考列用に残す。
    period_metrics = _calculate_period_media_metrics(history_df)
    normal_rate_map = period_metrics["normal_rate"]
    magi_rate_map = period_metrics["magi_rate"]
    normal_rate_all = period_metrics["normal_rate_all"]
    magi_rate_all = period_metrics["magi_rate_all"]
    magi_unit_map = period_metrics["magi_unit"]
    magi_unit_all = period_metrics["magi_unit_all"]

    # 未来日ごとのCPN区分。複合施策時はUI選択を保持したplan側を最優先。
    future_cpn_map = {}
    if "CPN名" in plan.columns:
        fm = plan[["date", "CPN名"]].dropna(subset=["date"]).drop_duplicates(subset=["date"], keep="last")
        future_cpn_map = dict(zip(fm["date"], fm["CPN名"]))
    elif cpn_master is not None and not cpn_master.empty:
        fm = cpn_master[["日付", "CPN名"]].copy()
        fm["日付"] = pd.to_datetime(fm["日付"], errors="coerce").dt.normalize()
        fm = fm.dropna(subset=["日付"])
        future_cpn_map = dict(zip(fm["日付"], fm["CPN名"]))

    # 発行数 = 手動設定後Forecast × 手動承認率
    issue_map = {}
    for (dt, media), cv in cv_map.items():
        rate = manual_rate_map.get(media, 0.0)
        issue_map[(dt, media)] = cv * rate

    # 費用はユーザー入力総額を正として、日次承認件数の構成比で配分。
    expected_cost_map = {}
    for media in media_list:
        media_dates = sorted(dates_by_media.get(media, []))
        total_cost = manual_cost_total_map.get(media, 0.0)
        total_issue = sum(
            issue_map.get((dt, media), 0.0)
            for dt in media_dates
        )

        if total_issue > 0:
            for dt in media_dates:
                share = issue_map.get((dt, media), 0.0) / total_issue
                expected_cost_map[(dt, media)] = total_cost * share
        else:
            total_cv = sum(
                cv_map.get((dt, media), 0.0)
                for dt in media_dates
            )
            for dt in media_dates:
                share = (
                    cv_map.get((dt, media), 0.0) / total_cv
                    if total_cv > 0
                    else 0.0
                )
                expected_cost_map[(dt, media)] = total_cost * share

        # 費用Totalの丸め差は最後の日へ寄せる。
        if media_dates:
            assigned = sum(
                expected_cost_map.get((dt, media), 0.0)
                for dt in media_dates
            )
            diff = total_cost - assigned
            last_dt = media_dates[-1]
            expected_cost_map[(last_dt, media)] = (
                expected_cost_map.get((last_dt, media), 0.0)
                + diff
            )

    # Promotion Detail等は今回採用グロス単価を表示。
    period_unit_map = {
        (dt, media): manual_gross_unit_map.get(media, 0.0)
        for (dt, media) in cv_map.keys()
    }

    total_cv_by_date = {}
    total_issue_by_date = {}
    total_cost_by_date = {}

    for (dt, media), cv in cv_map.items():
        total_cv_by_date[dt] = total_cv_by_date.get(dt, 0.0) + cv
        total_issue_by_date[dt] = (
            total_issue_by_date.get(dt, 0.0)
            + issue_map.get((dt, media), 0.0)
        )
        total_cost_by_date[dt] = (
            total_cost_by_date.get(dt, 0.0)
            + expected_cost_map.get((dt, media), 0.0)
        )

    # 参考列用：今回採用件数を未来CPN区分に応じて分割。
    normal_forecast_by_media = {m: 0.0 for m in media_list}
    magi_forecast_by_media = {m: 0.0 for m in media_list}

    for (dt, media), cv in cv_map.items():
        period = _future_period_name(
            dt,
            future_cpn_map,
            selected_cpn,
        )
        if period == "マジ得":
            magi_forecast_by_media[media] += cv
        else:
            normal_forecast_by_media[media] += cv

    normal_issue_by_media = {
        m: (
            normal_forecast_by_media[m]
            * manual_rate_map.get(m, 0.0)
        )
        for m in media_list
    }
    magi_issue_by_media = {
        m: (
            magi_forecast_by_media[m]
            * manual_rate_map.get(m, 0.0)
        )
        for m in media_list
    }

    # 後段互換用。
    opt_unit_map = manual_gross_unit_map
    cpa_map = {
        key: (
            expected_cost_map.get(key, 0.0)
            / issue_map.get(key, 0.0)
            if issue_map.get(key, 0.0) > 0
            else 0.0
        )
        for key in cv_map.keys()
    }

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

    # 今回は短縮承認系2シートを成果物から除外。
    for remove_name in ["短縮承認除外", "短縮承認日程"]:
        if remove_name in wb_out.sheetnames:
            del wb_out[remove_name]

    # 高速テンプレートは数式依存を最小化しているため、
    # 保存時の強制フル再計算指定は行わない。

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
    base_date_slots = 33
    date_slots = max(base_date_slots, len(dates))
    total_col = 58       # BF（33日テンプレート時）
    total_col = _extend_sheet_date_columns(
        main_ws,
        first_date_col=first_date_col,
        base_slots=base_date_slots,
        required_slots=date_slots,
        total_col=total_col,
    )

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

    # 明細エリア
    # 高速テンプレート側に150媒体×4行の空枠・書式を事前作成済み。
    # ここでは行追加・全セルクリア・スタイルコピーを一切行わず、値だけ書く。
    detail_start = 21

    for idx, media in enumerate(media_list, start=0):
        r0 = detail_start + idx * 4

        sid = sid_map.get(media, "")
        media_type = media_type_map.get(media, "ポイントサイト")
        total_cv = round(
            sum(cv_map.get((d.normalize(), media), 0) for d in dates)
        )
        total_cost = round(
            sum(expected_cost_map.get((d.normalize(), media), 0) for d in dates)
        )
        overall_cpa = round(total_cost / total_cv) if total_cv else 0

        # 左側情報
        _set_value(main_ws, r0, 2, idx + 1)
        _set_value(main_ws, r0, 3, sid)
        _set_value(main_ws, r0, 4, media)
        _set_value(main_ws, r0, 7, total_cv)
        _set_value(main_ws, r0, 8, 0)
        _set_value(main_ws, r0, 10, 0)
        normal_rate = normal_rate_map.get(media, normal_rate_all)
        selected_rate = manual_rate_map.get(media, 0.0)
        _set_percent(main_ws, r0, 11, normal_rate)
        _set_percent(main_ws, r0, 12, selected_rate)
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
                unit_price = round(period_unit_map.get((dt, media), 0))

                _set_value(main_ws, r0, col, cv)
                _set_value(main_ws, r0 + 1, col, None)
                _set_value(main_ws, r0 + 2, col, unit_price)
                _set_value(main_ws, r0 + 3, col, -cv)
            else:
                for off in range(4):
                    _set_value(main_ws, r0 + off, col, None)

        _set_value(main_ws, r0, total_col, total_cv)
        _set_value(main_ws, r0 + 1, total_col, 0)
        _set_value(main_ws, r0 + 2, total_col, overall_cpa)
        _set_value(main_ws, r0 + 3, total_col, -total_cv)

    # =========================================================
    # 2) 件数 / 3) 発行 / 4) コスト計算
    # 3シートとも同じ列構造:
    # A SID / B 媒体名 / C 単価①最適 / D 単価②マジ得 /
    # E 定常承認率 / F 定常発行数 /
    # G マジ得承認率 / H マジ得発行数 /
    # I:AO 日次 / AP Total
    # =========================================================
    metric_specs = [
        ("件数(合算）", "count"),
        ("発行(合算）", "issue"),
        ("コスト計算用(合算）", "cost"),
    ]

    metric_first_date_col = 9   # I
    metric_base_date_slots = 33
    metric_header_row = 3
    metric_data_start = 4

    for sheet_name, metric_kind in metric_specs:
        ws = wb_out[sheet_name]
        metric_total_col = 42  # AP（33日テンプレート時）
        metric_total_col = _extend_sheet_date_columns(
            ws,
            first_date_col=metric_first_date_col,
            base_slots=metric_base_date_slots,
            required_slots=date_slots,
            total_col=metric_total_col,
        )
        # 1行目のタイトル結合範囲も拡張後のTotal列まで広げる。
        try:
            for rng in list(ws.merged_cells.ranges):
                if rng.min_row == 1 and rng.max_row == 1 and rng.min_col == 1:
                    ws.unmerge_cells(str(rng))
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=metric_total_col)
        except Exception:
            pass

        _set_date_slots(
            ws,
            metric_header_row,
            metric_first_date_col,
            date_slots,
            dates,
            metric_total_col,
        )

        # 固定列ヘッダーを毎回明示。
        headers = [
            "SID",
            "媒体名",
            "単価①\n最適プラン",
            "単価②\nマジ得",
            "定常承認率",
            "定常発行数",
            "マジ得承認率",
            "マジ得発行数",
        ]
        for c, header in enumerate(headers, start=1):
            _set_value(ws, metric_header_row, c, header)

        for idx, media in enumerate(media_list):
            r = metric_data_start + idx

            normal_rate = normal_rate_map.get(media, normal_rate_all)
            manual_rate = manual_rate_map.get(media, 0.0)

            opt_unit = manual_gross_unit_map.get(media, 0.0)
            magi_unit = magi_unit_map.get(media, magi_unit_all)
            if magi_unit <= 0:
                magi_unit = opt_unit

            _set_value(ws, r, 1, sid_map.get(media, ""))
            _set_value(ws, r, 2, media)

            # Cは今回の手動採用グロス単価。
            _set_value(ws, r, 3, round(opt_unit))
            _set_value(ws, r, 4, round(magi_unit))

            # Eは過去定常の参考値、Gは今回採用承認率を反映。
            _set_percent(ws, r, 5, normal_rate)
            _set_value(ws, r, 6, round(normal_issue_by_media.get(media, 0)))
            _set_percent(ws, r, 7, manual_rate)
            _set_value(ws, r, 8, round(magi_issue_by_media.get(media, 0)))

            row_total = 0.0

            for i in range(date_slots):
                c = metric_first_date_col + i

                if i >= len(dates):
                    _set_value(ws, r, c, None)
                    continue

                dt = dates[i].normalize()

                if metric_kind == "count":
                    value = cv_map.get((dt, media), 0)
                elif metric_kind == "issue":
                    value = issue_map.get((dt, media), 0)
                else:
                    value = expected_cost_map.get((dt, media), 0)

                value = round(value)
                _set_value(ws, r, c, value)
                row_total += value

            _set_value(ws, r, metric_total_col, round(row_total))

    # =========================================================
    # 5) 全体サマリ
    # =========================================================
    sws = wb_out["全体サマリ（定常期間サマリ）"]
    summary_total_row = _extend_summary_date_rows(
        sws,
        base_slots=33,
        required_slots=date_slots,
        first_data_row=3,
        total_row=36,
    )
    _set_value(
        sws,
        1,
        1,
        f"{pd.Timestamp(start_date).month}月度サマリ 件数・発行・コスト",
    )

    sum_forecast = 0
    sum_issue = 0
    sum_cost = 0

    for i in range(date_slots):
        r = 3 + i

        if i < len(dates):
            dt = dates[i].normalize()
            forecast = round(total_cv_by_date.get(dt, 0))
            issue = round(total_issue_by_date.get(dt, 0))
            cost = round(total_cost_by_date.get(dt, 0))

            approval_rate = issue / forecast if forecast else 0
            issue_cpa = round(cost / issue) if issue else 0

            _set_value(sws, r, 1, dates[i].to_pydatetime())
            if not isinstance(sws.cell(r, 1), MergedCell):
                sws.cell(r, 1).number_format = "m/d"

            # A日付 B目標 CForecast DGAP E発行 F発行GAP G承認率 H発行コスト I発行CPA
            _set_value(sws, r, 2, forecast)
            _set_value(sws, r, 3, forecast)
            _set_value(sws, r, 4, 0)
            _set_value(sws, r, 5, issue)
            _set_value(sws, r, 6, 0)
            _set_percent(sws, r, 7, approval_rate)
            _set_value(sws, r, 8, cost)
            _set_value(sws, r, 9, issue_cpa)

            sum_forecast += forecast
            sum_issue += issue
            sum_cost += cost
        else:
            for c in range(1, 10):
                _set_value(sws, r, c, None)

    total_rate = sum_issue / sum_forecast if sum_forecast else 0
    total_cpa = round(sum_cost / sum_issue) if sum_issue else 0

    _set_value(sws, summary_total_row, 1, "Total")
    _set_value(sws, summary_total_row, 2, sum_forecast)
    _set_value(sws, summary_total_row, 3, sum_forecast)
    _set_value(sws, summary_total_row, 4, 0)
    _set_value(sws, summary_total_row, 5, sum_issue)
    _set_value(sws, summary_total_row, 6, 0)
    _set_percent(sws, summary_total_row, 7, total_rate)
    _set_value(sws, summary_total_row, 8, sum_cost)
    _set_value(sws, summary_total_row, 9, total_cpa)

    # =========================================================
    # 既存移管合算シート：不要列 E:W を最終出力時に削除
    # =========================================================
    # すべての値を書き込み終えた後に削除するため、
    # 既存の列番号ベースの出力処理には影響しない。
    main_ws.delete_cols(5, 19)  # E:W（19列）

    output = BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output.getvalue()


# -----------------------
# ✅ 提出用Excelダウンロード専用Fragment
# -----------------------
def _submission_download_body(
    opt_summary,
    history_df,
    cpn_master,
    manual_settings,
    start_date,
    end_date,
    selected_cpn,
    opt_mode,
    submission_filename,
):
    """
    提出用Excelはクリックされた時だけ生成する。
    data に callable を渡すため、画面描画時にはExcelを作らない。
    """

    def build_submission_excel():
        return create_submission_excel(
            opt_summary=opt_summary,
            history_df=history_df,
            cpn_master=cpn_master,
            manual_settings=manual_settings,
            start_date=start_date,
            end_date=end_date,
            selected_cpn=selected_cpn,
            opt_mode=opt_mode,
        )

    st.download_button(
        "📥 提出用Excelを生成してDL",
        data=build_submission_excel,
        file_name=submission_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        on_click="ignore",
        type="primary",
        width="stretch",
    )


# st.fragment が使えるStreamlitでは、
# このボタン操作だけを独立再実行する。
if hasattr(st, "fragment"):
    render_submission_download = st.fragment(
        _submission_download_body
    )
else:
    # 古いStreamlitでも起動自体は可能。
    render_submission_download = _submission_download_body


# -----------------------
# ✅ 補助関数
# -----------------------
def _truthy(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin(
        {"1", "true", "yes", "y", "○", "〇", "あり", "有", "実施"}
    )


def _normalize_sid(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _figure_png_bytes(fig, dpi: int = 180) -> bytes:
    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=dpi, bbox_inches="tight")
    buffer.seek(0)
    return buffer.getvalue()


def _get_japanese_font_properties():
    """日本語フォントを安全に取得する。

    Streamlit CloudではOSパッケージに依存せず、japanize-matplotlibに
    同梱されているIPAexGothicを直接登録して利用する。
    """
    if japanize_matplotlib is not None:
        try:
            pkg_dir = Path(japanize_matplotlib.__file__).resolve().parent
            bundled = [
                pkg_dir / "fonts" / "ipaexg.ttf",
                pkg_dir / "fonts" / "ipaexg.ttf",
            ]
            for path in bundled:
                if path.is_file():
                    font_manager.fontManager.addfont(str(path))
                    prop = font_manager.FontProperties(fname=str(path))
                    family = prop.get_name()
                    plt.rcParams["font.family"] = family
                    plt.rcParams["axes.unicode_minus"] = False
                    return prop
            # 同梱フォントのパス構成が変わっても探索できるようにする。
            for path in pkg_dir.rglob("*.ttf"):
                try:
                    font_manager.fontManager.addfont(str(path))
                    prop = font_manager.FontProperties(fname=str(path))
                    family = prop.get_name()
                    if family:
                        plt.rcParams["font.family"] = family
                        plt.rcParams["axes.unicode_minus"] = False
                        return prop
                except Exception:
                    continue
        except Exception:
            pass

    candidates = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    ]
    candidates.extend(glob.glob("/usr/share/fonts/**/*NotoSansCJK*", recursive=True))
    candidates.extend(glob.glob("/usr/share/fonts/**/*NotoSansJP*", recursive=True))
    candidates.extend(glob.glob("/usr/share/fonts/**/*IPA*Gothic*", recursive=True))

    seen = set()
    for path in candidates:
        if path in seen:
            continue
        seen.add(path)
        try:
            if Path(path).is_file():
                font_manager.fontManager.addfont(str(path))
                prop = font_manager.FontProperties(fname=path)
                family = prop.get_name()
                plt.rcParams["font.family"] = family
                plt.rcParams["axes.unicode_minus"] = False
                return prop
        except Exception:
            continue
    return None


def _render_png_actions(png_bytes: bytes, file_name: str, download_key: str, copy_key: str):
    """PNG保存とクリップボードへの画像コピーを横並びで表示する。"""
    left, right = st.columns([1, 1])
    with left:
        st.download_button(
            "画像をPNG保存",
            data=png_bytes,
            file_name=file_name,
            mime="image/png",
            key=download_key,
            width="stretch",
        )
    with right:
        b64 = base64.b64encode(png_bytes).decode("ascii")
        button_id = f"copy_{copy_key}".replace("-", "_")
        html = f"""
        <div style="width:100%;">
          <button id="{button_id}" style="
            width:100%; height:38px; border:1px solid rgba(49,51,63,.2);
            border-radius:8px; background:white; cursor:pointer; font-size:14px;
          ">画像をコピー</button>
          <div id="{button_id}_msg" style="font-size:12px; margin-top:3px; min-height:16px;"></div>
        </div>
        <script>
        const btn = document.getElementById('{button_id}');
        const msg = document.getElementById('{button_id}_msg');
        btn.addEventListener('click', async () => {{
          try {{
            const response = await fetch('data:image/png;base64,{b64}');
            const blob = await response.blob();
            if (!navigator.clipboard || typeof ClipboardItem === 'undefined') {{
              throw new Error('clipboard_api_unavailable');
            }}
            await navigator.clipboard.write([new ClipboardItem({{'image/png': blob}})]);
            msg.textContent = 'コピーしました';
          }} catch (e) {{
            msg.textContent = 'ブラウザの権限でコピーできませんでした';
          }}
        }});
        </script>
        """
        components.html(html, height=62)


def _chart_month_label(value) -> str:
    """日本語フォントが無い場合にも文字化けしない月度ラベルへ変換する。"""
    text = str(value).strip()
    import re
    match = re.search(r"(\d{4})年\s*(\d{1,2})月度?", text)
    if match:
        return f"{match.group(1)}/{int(match.group(2)):02d}"
    return text


def render_history_analytics(
    history_df: pd.DataFrame,
    current_plan_df: pd.DataFrame | None = None,
    current_plan_magitoku_df: pd.DataFrame | None = None,
):
    st.subheader("📈 月度別 過去実績分析")
    st.caption(
        "月度はCPNマスタの『月度』を正として集計。"
        "発行数 = 成果承認フラグYの件数、発行CPA = コスト ÷ 発行数です。"
    )

    analysis_scope = st.radio(
        "集計対象",
        ["全期間", "マジ得期間のみ"],
        horizontal=True,
        key="history_analytics_scope",
    )

    analytics_df = history_df
    current_chart_plan = current_plan_df
    scope_suffix = "全期間"
    if analysis_scope == "マジ得期間のみ":
        if "CPN名" not in history_df.columns:
            st.warning("CPN名が付与されていないため、マジ得期間のみの集計ができません。")
            return
        analytics_df = history_df.loc[
            history_df["CPN名"].astype(str).str.strip().eq("マジ得")
        ].copy()
        current_chart_plan = current_plan_magitoku_df
        scope_suffix = "マジ得期間のみ"
        if analytics_df.empty:
            st.info("マジ得期間に該当する過去実績がありません。")
            return

    monthly = prepare_monthly_performance(analytics_df)
    if monthly.empty:
        st.info("月度が付与された過去実績がないため、月度別グラフを表示できません。")
        return

    preview = monthly[["月度", "発生件数", "発行数", "コスト", "発行CPA"]].copy()
    preview["発行CPA"] = preview["発行CPA"].map(
        lambda x: "" if pd.isna(x) else f"¥{x:,.0f}"
    )
    preview["コスト"] = preview["コスト"].map(lambda x: f"¥{x:,.0f}")

    tab_issue, tab_band, tab_table = st.tabs([
        "発行数・発行CPA",
        "単価帯",
        "集計表",
    ])

    with tab_issue:
        jp_font = _get_japanese_font_properties()
        x = range(len(monthly))
        labels_raw = monthly["月度"].astype(str).tolist()
        labels = labels_raw if jp_font else [_chart_month_label(v) for v in labels_raw]
        scope_suffix_display = scope_suffix if jp_font else ("Magi-toku period" if analysis_scope == "マジ得" else "Full period")

        fig_count, ax_count = plt.subplots(figsize=(12, 5))
        ax_count.bar(x, monthly["発行数"])
        ax_count.set_title(
            f"月度別 発行数（{scope_suffix}）" if jp_font else f"Issued count by month ({scope_suffix_display})",
            fontproperties=jp_font,
        )
        ax_count.set_ylabel("発行数" if jp_font else "Issued count", fontproperties=jp_font)
        ax_count.set_xticks(list(x))
        ax_count.set_xticklabels(labels, rotation=45, ha="right", fontproperties=jp_font)
        ax_count.grid(axis="y", alpha=0.25)
        fig_count.tight_layout()
        st.pyplot(fig_count, width="stretch")
        count_png = _figure_png_bytes(fig_count)
        _render_png_actions(
            count_png,
            f"月度別_発行数_{scope_suffix}.png",
            f"download_monthly_issue_count_{analysis_scope}",
            f"copy_monthly_issue_count_{analysis_scope}",
        )
        plt.close(fig_count)

        fig_cpa, ax_cpa = plt.subplots(figsize=(12, 5))
        ax_cpa.plot(x, monthly["発行CPA"], marker="o")
        ax_cpa.set_title(
            f"月度別 発行CPA（{scope_suffix}）" if jp_font else f"Issued CPA by month ({scope_suffix_display})",
            fontproperties=jp_font,
        )
        ax_cpa.set_ylabel("発行CPA（円）" if jp_font else "Issued CPA (JPY)", fontproperties=jp_font)
        ax_cpa.set_xticks(list(x))
        ax_cpa.set_xticklabels(labels, rotation=45, ha="right", fontproperties=jp_font)
        ax_cpa.grid(axis="y", alpha=0.25)
        fig_cpa.tight_layout()
        st.pyplot(fig_cpa, width="stretch")
        cpa_png = _figure_png_bytes(fig_cpa)
        _render_png_actions(
            cpa_png,
            f"月度別_発行CPA_{scope_suffix}.png",
            f"download_monthly_issue_cpa_{analysis_scope}",
            f"copy_monthly_issue_cpa_{analysis_scope}",
        )
        plt.close(fig_cpa)

    with tab_band:
        control1, control2 = st.columns([1, 1])
        with control1:
            band_step = st.number_input(
                "単価帯の刻み幅（円）",
                min_value=500,
                max_value=50000,
                value=4000,
                step=500,
            )
        with control2:
            band_metric = st.selectbox(
                "積み上げる指標",
                ["発行数", "発生件数"],
                index=0,
            )

        # 1,000円刻み等で単価帯が大量発生してもクラッシュしないよう、
        # 表示系列数は最大40帯に制御。超過分は高単価側を「○円以上」に集約する。
        band_matrix = prepare_unit_price_band_matrix(
            analytics_df,
            band_step=int(band_step),
            value_metric=band_metric,
            max_bands=40,
        )

        # 今回プラン（手動設定後）を同じ単価帯グラフへ追加。
        if current_chart_plan is not None and not current_chart_plan.empty:
            cp = _normalize_manual_settings(current_chart_plan)
            cp["unit_price"] = pd.to_numeric(
                cp["今回プラン採用グロス単価"], errors="coerce"
            ).fillna(0.0)
            metric_col = "承認件数" if band_metric == "発行数" else "今回採用件数"
            cp[metric_col] = pd.to_numeric(cp[metric_col], errors="coerce").fillna(0.0)
            cp = cp[(cp["unit_price"] >= 0) & (cp[metric_col] > 0)].copy()
            if not cp.empty:
                cp["band_lower"] = (
                    (cp["unit_price"] // int(band_step)).astype(int) * int(band_step)
                )
                current = (
                    cp.groupby("band_lower", as_index=False)[metric_col]
                    .sum()
                )
                current_labels = {}
                for r in current.itertuples():
                    lower = int(r.band_lower)
                    upper = lower + int(band_step) - 1
                    current_labels[lower] = f"¥{lower:,}–¥{upper:,}"

                # 既存履歴側の帯と今回帯を和集合にする。
                if band_matrix.empty:
                    band_matrix = pd.DataFrame(index=["今回"])
                    band_matrix.index.name = "月度"
                all_cols = list(band_matrix.columns)
                for label in current_labels.values():
                    if label not in all_cols:
                        all_cols.append(label)
                band_matrix = band_matrix.reindex(columns=all_cols, fill_value=0.0)
                band_matrix.loc["今回"] = 0.0
                for _, r in current.iterrows():
                    band_matrix.loc["今回", current_labels[int(r["band_lower"])]] = float(r[metric_col])

        # 今回分を足した結果も最大40帯に制御し、細かい刻みでのメモリ急増を防ぐ。
        if not band_matrix.empty and len(band_matrix.columns) > 40:
            import re
            def _band_lower_from_label(label):
                nums = re.findall(r"[\d,]+", str(label))
                return int(nums[0].replace(",", "")) if nums else 0
            ordered_cols = sorted(band_matrix.columns, key=_band_lower_from_label)
            keep_cols = ordered_cols[:39]
            overflow_cols = ordered_cols[39:]
            overflow_lower = _band_lower_from_label(overflow_cols[0])
            overflow_label = f"¥{overflow_lower:,}以上"
            overflow_values = band_matrix[overflow_cols].sum(axis=1)
            band_matrix = band_matrix[keep_cols].copy()
            band_matrix[overflow_label] = overflow_values

        if band_matrix.empty:
            st.info("単価帯グラフを作成できる実績がありません。")
        else:
            jp_font = _get_japanese_font_properties()
            plot_matrix = band_matrix.copy()
            if not jp_font:
                plot_matrix.index = [_chart_month_label(v) for v in plot_matrix.index]
            fig_band, ax_band = plt.subplots(figsize=(12, 6))
            plot_matrix.plot(kind="bar", stacked=True, ax=ax_band, width=0.8)
            ax_band.set_title(
                f"月度別 単価帯構成（{scope_suffix} / {int(band_step):,}円刻み / {band_metric}）"
                if jp_font
                else f"Unit price bands by month ({int(band_step):,} JPY step)",
                fontproperties=jp_font,
            )
            ax_band.set_xlabel("月度" if jp_font else "Month", fontproperties=jp_font)
            ax_band.set_ylabel(band_metric if jp_font else ("Issued count" if band_metric == "発行数" else "Conversions"), fontproperties=jp_font)
            ax_band.tick_params(axis="x", rotation=45)
            if jp_font:
                for tick in ax_band.get_xticklabels():
                    tick.set_fontproperties(jp_font)
            # 凡例の縦長化でPNGが巨大化しないよう、系列数に応じて複数列化。
            legend_cols = max(1, min(4, (len(plot_matrix.columns) + 14) // 15))
            legend_prop = jp_font.copy() if jp_font is not None else None
            if legend_prop is not None:
                legend_prop.set_size(8)
            legend = ax_band.legend(
                title="単価帯" if jp_font else "Unit price band",
                bbox_to_anchor=(1.02, 1),
                loc="upper left",
                ncol=legend_cols,
                fontsize=8,
                prop=legend_prop,
            )
            if jp_font and legend is not None:
                legend.get_title().set_fontproperties(jp_font)
            ax_band.grid(axis="y", alpha=0.25)
            fig_band.tight_layout()
            st.pyplot(fig_band, width="stretch")
            # 積み上げグラフは系列数が多いため、PNG生成時のメモリ使用量も抑える。
            band_png = _figure_png_bytes(fig_band, dpi=140)
            _render_png_actions(
                band_png,
                f"月度別_単価帯_{int(band_step)}円刻み_{scope_suffix}.png",
                f"download_unit_price_band_{analysis_scope}",
                f"copy_unit_price_band_{analysis_scope}",
            )
            plt.close(fig_band)

    with tab_table:
        st.dataframe(preview, width="stretch", hide_index=True)


def _daily_pair_average(df: pd.DataFrame) -> pd.DataFrame:
    daily = (
        df.groupby(["date", "media", "商品ID"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    return (
        daily.groupby(["media", "商品ID"], as_index=False)
        .agg(base_cv=("cv", "mean"), cost=("cost", "mean"))
    )


def _calculate_normal_month_base(
    history_df: pd.DataFrame,
    selected_months: list[str],
    calendar_dates=None,
) -> pd.DataFrame:
    """logic.factorsの定常学習ロジックを呼び出す互換ラッパー。"""
    from logic.factors import calculate_normal_month_base

    return calculate_normal_month_base(history_df, selected_months, calendar_dates=calendar_dates)





def _cpn_month_sort_key(label: str):
    """CPNマスタの月度ラベルを日付範囲ではなくラベル自体で時系列化する。"""
    import re
    text = str(label).strip()
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月度", text)
    if m:
        return (int(m.group(1)), int(m.group(2)), text)
    m = re.search(r"(\d{4})[-/](\d{1,2})", text)
    if m:
        return (int(m.group(1)), int(m.group(2)), text)
    return (9999, 99, text)


def _get_cpn_normal_months(cpn_master: pd.DataFrame) -> list[str]:
    """CPNマスタ上の定常月度を、CPN月度ラベル順で返す。"""
    normal_labels = {"通常", "定常"}
    master = cpn_master.copy()
    master["月度"] = master["月度"].astype("string").str.strip()
    master["CPN名"] = master["CPN名"].astype("string").str.strip()
    labels = (
        master.loc[
            master["CPN名"].isin(normal_labels)
            & master["月度"].notna()
            & master["月度"].astype(str).ne("未設定"),
            "月度",
        ]
        .astype(str)
        .drop_duplicates()
        .tolist()
    )
    return sorted(labels, key=_cpn_month_sort_key)


def _candidate_learning_counts(available_prior: int) -> list[int]:
    """ユーザー操作ではなく、内部バックテストで比較する候補。"""
    base = [2, 3, 4, 6, 9, 12]
    candidates = [n for n in base if n <= available_prior]
    if not candidates and available_prior >= 1:
        candidates = [available_prior]
    return candidates


def _select_learning_month_count(
    history_df: pd.DataFrame,
    cpn_master: pd.DataFrame,
    target_month: str,
) -> tuple[int, pd.DataFrame]:
    """対象月より前の月度だけを使い、過去バックテストから学習期間を自動選択する。"""
    month_labels = _get_cpn_normal_months(cpn_master)
    if str(target_month) not in month_labels:
        raise ValueError(f"CPNマスタに {target_month} の定常月度がありません。")
    target_idx = month_labels.index(str(target_month))
    candidates = _candidate_learning_counts(target_idx)
    if not candidates:
        raise ValueError("バックテスト対象月より前の学習月度がありません。")

    # 対象月の直前から最大3月度を検証。各検証月度の実績はその予測には使わない。
    validation_months = month_labels[max(0, target_idx - 3):target_idx]
    rows = []
    for n in candidates:
        scores = []
        wapes = []
        tested = []
        for vm in validation_months:
            vm_idx = month_labels.index(vm)
            if vm_idx < n:
                continue
            try:
                r = _run_normal_backtest(
                    history_df,
                    cpn_master,
                    vm,
                    learning_month_count=n,
                    auto_select_learning=False,
                )
            except Exception:
                continue
            if pd.notna(r.get("total_error_rate")):
                scores.append(abs(float(r["total_error_rate"])))
                tested.append(vm)
            if pd.notna(r.get("wape")):
                wapes.append(float(r["wape"]))
        if scores:
            rows.append({
                "学習月度数": n,
                "平均総量誤差率": float(np.mean(scores)),
                "平均日次WAPE": float(np.mean(wapes)) if wapes else np.nan,
                "検証月度数": len(scores),
                "検証月度": "、".join(tested),
            })

    score_df = pd.DataFrame(rows)
    if score_df.empty:
        # 十分な過去検証ができない初期月度は、最大6月度までを安全な既定値にする。
        fallback = min(6, max(candidates))
        return fallback, score_df

    score_df = score_df.sort_values(
        ["平均総量誤差率", "平均日次WAPE", "学習月度数"],
        ascending=[True, True, True],
        na_position="last",
        kind="stable",
    ).reset_index(drop=True)
    return int(score_df.iloc[0]["学習月度数"]), score_df


def _run_normal_backtest(
    history_df: pd.DataFrame,
    cpn_master: pd.DataFrame,
    target_month: str,
    learning_month_count: int | None = None,
    auto_select_learning: bool = True,
) -> dict:
    """指定月を未来扱いし、それ以前の実績だけで定常予測を再現する。"""
    from logic.factors import calculate_dynamic_factor_tables, calculate_normal_month_base, calculate_normal_media_diagnostics, calculate_unit_price_response_table, calculate_media_continuation_table

    normal_labels = {"通常", "定常"}
    work = history_df.copy()
    work["date"] = pd.to_datetime(work["date"], errors="coerce").dt.normalize()
    work["月度"] = work["月度"].astype("string").str.strip()
    work["CPN名"] = work["CPN名"].astype("string").str.strip()

    # バックテストの「月度」は暦月ではなくCPNマスタの月度を唯一の基準にする。
    master = cpn_master.copy()
    master["日付"] = pd.to_datetime(master["日付"], errors="coerce").dt.normalize()
    master["月度"] = master["月度"].astype("string").str.strip()
    master["CPN名"] = master["CPN名"].astype("string").str.strip()
    # 月度の順番は対象日のmin/maxではなく、CPNマスタの月度ラベルを正とする。
    month_labels = _get_cpn_normal_months(master)
    if str(target_month) not in month_labels:
        raise ValueError(f"CPNマスタに {target_month} の定常月度がありません。")
    target_idx = month_labels.index(str(target_month))
    prior_months = month_labels[:target_idx]
    if not prior_months:
        raise ValueError("バックテスト対象月より前の定常実績がありません。")

    target_calendar = master.loc[
        master["月度"].astype(str).eq(str(target_month)) & master["CPN名"].isin(normal_labels),
        ["日付", "月度", "line_oa_flag", "magitoku_after_flag"],
    ].drop_duplicates("日付")
    if target_calendar.empty:
        raise ValueError(f"CPNマスタに {target_month} の定常日がありません。")
    target_dates = pd.to_datetime(target_calendar["日付"], errors="coerce").dropna().dt.normalize()
    target_start = pd.Timestamp(target_dates.min()).normalize()
    target_end = pd.Timestamp(target_dates.max()).normalize()

    # 学習はCPN月度数ではなく、予測基準日以前の直近60暦日に固定する。
    # CPN月度の日数・その中の定常日数が不均一でも、同じ時間幅で現在の媒体力を評価する。
    learning_score_table = pd.DataFrame()
    cutoff_date = target_start - pd.Timedelta(days=1)
    requested_learning_start = cutoff_date - pd.Timedelta(days=59)
    # 最大60日を使う。ただしローデータが60日未満しかない場合は、
    # 実際に存在する最古日から学習し、存在しない過去日を0CV扱いしない。
    available_before_cutoff = work.loc[work["date"].le(cutoff_date), "date"].dropna()
    if available_before_cutoff.empty:
        raise ValueError("予測基準日以前の実績データがありません。")
    earliest_available_date = pd.Timestamp(available_before_cutoff.min()).normalize()
    learning_start = max(requested_learning_start, earliest_available_date)
    training = work.loc[work["date"].between(learning_start, cutoff_date)].copy()

    # 実際にデータが存在する学習期間内のCPN定常日だけを0CV日の分母に使う。
    learning_calendar_dates = (
        master.loc[
            master["日付"].between(learning_start, cutoff_date)
            & master["CPN名"].isin(normal_labels),
            "日付",
        ]
        .dropna().drop_duplicates().sort_values().tolist()
    )
    if not learning_calendar_dates:
        raise ValueError("予測基準日以前の利用可能データ内にCPNマスタの定常日がありません。")

    learning_months = (
        training.loc[training["CPN名"].isin(normal_labels), "月度"]
        .dropna().astype(str).loc[lambda x: x.ne("未設定")].drop_duplicates().tolist()
    )
    learning_month_count = len(learning_months)

    base_pair = calculate_normal_month_base(
        training, learning_months, calendar_dates=learning_calendar_dates
    )
    if base_pair.empty:
        raise ValueError("予測基準日以前の利用可能期間に定常実績がありません。")
    factor_tables = calculate_dynamic_factor_tables(
        training, learning_months, calendar_dates=learning_calendar_dates
    )

    # 学習データ監査: 予測ロジックは変えず、「実際に何を学習対象にしたか」だけ可視化する。
    # CPNマスタ上の対象日数と、ローデータ上の実績量を月度別に並べる。
    learning_audit_rows = []
    for lm in learning_months:
        lm_master_dates = (
            master.loc[
                master["月度"].astype(str).eq(str(lm)) & master["CPN名"].isin(normal_labels),
                "日付",
            ]
            .dropna().drop_duplicates().sort_values()
        )
        lm_raw = training.loc[
            training["月度"].astype(str).eq(str(lm))
            & training["CPN名"].isin(normal_labels)
        ].copy()
        lm_raw["cv"] = pd.to_numeric(lm_raw.get("cv", 0), errors="coerce").fillna(0.0)
        learning_audit_rows.append({
            "学習月度": str(lm),
            "CPN対象日数": int(lm_master_dates.nunique()),
            "CPN最初の日": lm_master_dates.min() if not lm_master_dates.empty else pd.NaT,
            "CPN最後の日": lm_master_dates.max() if not lm_master_dates.empty else pd.NaT,
            "実績存在日数": int(lm_raw["date"].nunique()) if not lm_raw.empty else 0,
            "CV合計": float(lm_raw["cv"].sum()) if not lm_raw.empty else 0.0,
            "稼働媒体数": int(lm_raw.loc[lm_raw["cv"].gt(0), "media"].nunique()) if not lm_raw.empty else 0,
            "元データ行数": int(len(lm_raw)),
        })
    learning_audit = pd.DataFrame(learning_audit_rows)

    selected_training = training.loc[
        training["月度"].astype(str).isin([str(x) for x in learning_months])
        & training["CPN名"].isin(normal_labels)
    ].copy()
    selected_training["cv"] = pd.to_numeric(selected_training.get("cv", 0), errors="coerce").fillna(0.0)
    excluded_audit = factor_tables.get("excluded", pd.DataFrame()).copy()
    inactive_audit = factor_tables.get("inactive_media", pd.DataFrame()).copy()
    learning_summary = {
        "最初の日": selected_training["date"].min() if not selected_training.empty else pd.NaT,
        "最後の日": selected_training["date"].max() if not selected_training.empty else pd.NaT,
        "実績存在日数": int(selected_training["date"].nunique()) if not selected_training.empty else 0,
        "CV合計": float(selected_training["cv"].sum()) if not selected_training.empty else 0.0,
        "媒体数": int(selected_training["media"].nunique()) if not selected_training.empty else 0,
        "元データ行数": int(len(selected_training)),
        "除外媒体日数": int(len(excluded_audit)),
        "休眠除外媒体数": int(len(inactive_audit)),
    }

    # 対象月のカレンダー条件は上でCPNマスタから確定済み。
    future = pd.DataFrame({"date": sorted(target_calendar["日付"].dropna().unique())}).merge(base_pair, how="cross")
    future["date"] = pd.to_datetime(future["date"]).dt.normalize()
    future["weekday"] = future["date"].dt.day_name()
    future = add_business_edge_flags(future)
    future = future.merge(
        target_calendar.rename(columns={"日付": "date"}),
        on="date",
        how="left",
    )
    future["line_oa_flag"] = future["line_oa_flag"].fillna(0).astype(int)
    future["magitoku_after_flag"] = future["magitoku_after_flag"].fillna(0).astype(int)
    future["CPN名"] = "通常"
    future["cpn_factor"] = 1.0

    forecast = forecast_cv(future, factor_tables)

    # 診断列は forecast.py 側でも生成するが、旧ファイルとの混在や
    # デプロイ差し替え漏れがあってもバックテスト自体が落ちないよう、
    # ここでも不足列を再構築する。最終予測式は従来と同一。
    stage_formulas = {
        "stage_base_cv": ("base_cv", "cpn_factor"),
        "stage_unit_price_cv": ("stage_base_cv", "unit_price_factor"),
        "stage_weekday_cv": ("stage_unit_price_cv", "weekday_factor"),
        "stage_season_cv": ("stage_weekday_cv", "season_factor"),
        "stage_month_edge_cv": ("stage_season_cv", "month_edge_factor"),
        "stage_after_cv": ("stage_month_edge_cv", "after_factor"),
        "stage_line_cv": ("stage_after_cv", "line_factor"),
    }
    for stage_col, (left_col, factor_col) in stage_formulas.items():
        if stage_col not in forecast.columns:
            if left_col not in forecast.columns or factor_col not in forecast.columns:
                raise ValueError(
                    f"バックテスト診断列を再構築できません: {stage_col} "
                    f"(必要列: {left_col}, {factor_col})"
                )
            forecast[stage_col] = (
                pd.to_numeric(forecast[left_col], errors="coerce").fillna(0.0)
                * pd.to_numeric(forecast[factor_col], errors="coerce").fillna(1.0)
            )
    if "forecast_cv" not in forecast.columns:
        forecast["forecast_cv"] = forecast["stage_line_cv"]

    # 掲載有無は過去CVだけでは完全には予知できないため、
    # バックテストでは「未定」として過去の月度間掲載継続率を期待値に反映する。
    # 対象月の実績は一切使わず、予測基準日までの履歴だけで算出する。
    continuation_bt = calculate_media_continuation_table(
        work.loc[work["date"].le(cutoff_date)].copy(),
        cpn_master=master,
        cutoff_date=cutoff_date,
    )
    continuation_map = (
        continuation_bt.set_index("media")["continuation_rate"].to_dict()
        if not continuation_bt.empty else {}
    )
    global_cont_rate = (
        float(continuation_bt["global_continuation_rate"].iloc[0])
        if not continuation_bt.empty else 0.70
    )
    forecast["掲載継続率"] = forecast["media"].map(continuation_map).fillna(global_cont_rate).clip(0.0, 1.0)
    forecast["forecast_cv_before_publication"] = pd.to_numeric(forecast["forecast_cv"], errors="coerce").fillna(0.0)
    forecast["forecast_cv"] = forecast["forecast_cv_before_publication"] * forecast["掲載継続率"]

    forecast_daily = (
        forecast.groupby(["date", "media"], as_index=False)
        .agg(
            基礎CV=("stage_base_cv", "sum"),
            単価補正後CV=("stage_unit_price_cv", "sum"),
            曜日補正後CV=("stage_weekday_cv", "sum"),
            需要期補正後CV=("stage_season_cv", "sum"),
            月初月末補正後CV=("stage_month_edge_cv", "sum"),
            マジ得後補正後CV=("stage_after_cv", "sum"),
            LINE補正後CV=("stage_line_cv", "sum"),
            掲載判定前CV=("forecast_cv_before_publication", "sum"),
            予測掲載継続率=("掲載継続率", "mean"),
            forecast_cv=("forecast_cv", "sum"),
        )
    )

    target_date_set = set(pd.to_datetime(target_calendar["日付"], errors="coerce").dropna().dt.normalize())
    actual = work.loc[
        work["月度"].astype(str).eq(str(target_month))
        & work["CPN名"].isin(normal_labels)
        & work["date"].isin(target_date_set)
    ].copy()
    actual_daily = (
        actual.groupby(["date", "media"], as_index=False)
        .agg(actual_cv=("cv", "sum"))
    )

    daily_compare = forecast_daily.merge(actual_daily, on=["date", "media"], how="outer").fillna(0.0)
    daily_compare["差分"] = daily_compare["forecast_cv"] - daily_compare["actual_cv"]
    daily_compare["絶対誤差"] = daily_compare["差分"].abs()

    media_compare = (
        daily_compare.groupby("media", as_index=False)
        .agg(
            基礎CV=("基礎CV", "sum"),
            単価補正後CV=("単価補正後CV", "sum"),
            曜日補正後CV=("曜日補正後CV", "sum"),
            需要期補正後CV=("需要期補正後CV", "sum"),
            月初月末補正後CV=("月初月末補正後CV", "sum"),
            マジ得後補正後CV=("マジ得後補正後CV", "sum"),
            LINE補正後CV=("LINE補正後CV", "sum"),
            掲載判定前CV=("掲載判定前CV", "sum"),
            予測掲載継続率=("予測掲載継続率", "mean"),
            予測CV=("forecast_cv", "sum"),
            実績CV=("actual_cv", "sum"),
            絶対誤差=("絶対誤差", "sum"),
        )
    )
    # 各補正が予測総量を何CV動かしたか。暴走地点の特定に使う。
    media_compare["単価影響CV"] = media_compare["単価補正後CV"] - media_compare["基礎CV"]
    media_compare["曜日影響CV"] = media_compare["曜日補正後CV"] - media_compare["単価補正後CV"]
    media_compare["需要期影響CV"] = media_compare["需要期補正後CV"] - media_compare["曜日補正後CV"]
    media_compare["月初月末影響CV"] = media_compare["月初月末補正後CV"] - media_compare["需要期補正後CV"]
    media_compare["マジ得後影響CV"] = media_compare["マジ得後補正後CV"] - media_compare["月初月末補正後CV"]
    media_compare["LINE影響CV"] = media_compare["LINE補正後CV"] - media_compare["マジ得後補正後CV"]
    media_compare["差分"] = media_compare["予測CV"] - media_compare["実績CV"]
    media_compare["誤差率"] = np.where(
        media_compare["実績CV"].ne(0),
        media_compare["差分"] / media_compare["実績CV"],
        np.nan,
    )
    media_compare["掲載継続影響CV"] = media_compare["予測CV"] - media_compare["掲載判定前CV"]

    # バックテスト対象月度にローデータ上の行が存在した媒体を「掲載あり」とする。
    # CV=0でも行が存在すれば掲載ありとして扱い、CV発生有無と掲載有無を混同しない。
    target_published_media = set(actual["media"].dropna().astype(str).unique()) if not actual.empty else set()
    media_compare["対象月掲載実績"] = np.where(
        media_compare["media"].astype(str).isin(target_published_media),
        "掲載あり",
        "掲載なし",
    )

    # 予測の土台を監査できるよう、稼働率×稼働時CVの内訳を媒体別結果へ付与。
    diagnostics = calculate_normal_media_diagnostics(training, learning_months, calendar_dates=learning_calendar_dates)
    if not diagnostics.empty:
        diagnostics = diagnostics.rename(columns={
            "activity_rate": "学習稼働率",
            "active_daily_cv": "稼働時日平均CV",
            "expected_daily_cv": "基礎期待CV/日",
            "eligible_days": "学習対象日数",
            "active_days": "稼働日数",
        })
        media_compare = media_compare.merge(diagnostics, on="media", how="left")

    # 単価感応度の監査情報。予測にはtarget月実績を使わず、trainingだけから推定する。
    price_diag = calculate_unit_price_response_table(training, learning_months)
    if not price_diag.empty:
        price_diag = price_diag.rename(columns={
            "reference_unit_price": "学習基準単価",
            "latest_unit_price": "学習直近単価",
            "elasticity_raw": "単価感応度raw",
            "elasticity": "単価感応度",
            "sample_months": "単価学習月数",
            "price_variation": "単価変動幅",
            "fit_r2": "単価モデルR2",
        })
        media_compare = media_compare.merge(price_diag, on="media", how="left")
        media_compare["予測単価補正倍率"] = np.where(
            media_compare["基礎CV"].ne(0),
            media_compare["単価補正後CV"] / media_compare["基礎CV"],
            1.0,
        )

    # 対象月の実績単価は「説明用」だけに付ける。予測計算には一切利用しない。
    if not actual.empty:
        actual_price = actual.groupby("media", as_index=False).agg(
            _actual_cv=("cv", "sum"), _actual_cost=("cost", "sum")
        )
        actual_price["対象月実績単価_診断のみ"] = (
            actual_price["_actual_cost"] / actual_price["_actual_cv"].replace(0, pd.NA)
        )
        media_compare = media_compare.merge(
            actual_price[["media", "対象月実績単価_診断のみ"]], on="media", how="left"
        )

    total_forecast_before_publication = float(daily_compare["掲載判定前CV"].sum())
    total_forecast = float(daily_compare["forecast_cv"].sum())
    total_actual = float(daily_compare["actual_cv"].sum())

    # 営業向けバックテストは「対象月度に実際に掲載があった媒体」だけを自動抽出して評価する。
    published_mask = media_compare["media"].astype(str).isin(target_published_media)
    published_media_forecast = float(media_compare.loc[published_mask, "予測CV"].sum())
    published_media_actual = float(media_compare.loc[published_mask, "実績CV"].sum())
    published_media_diff = published_media_forecast - published_media_actual
    published_media_error_rate = (published_media_diff / published_media_actual) if published_media_actual else np.nan
    unpublished_media_forecast = float(media_compare.loc[~published_mask, "予測CV"].sum())

    published_daily = daily_compare.loc[daily_compare["media"].astype(str).isin(target_published_media)].copy()
    published_wape = (
        float(published_daily["絶対誤差"].sum() / published_media_actual)
        if published_media_actual else np.nan
    )

    # 全媒体ベースの値は分析用として残す。
    continued_media_raw_forecast = float(media_compare.loc[published_mask, "掲載判定前CV"].sum())
    publication_end_raw_forecast = total_forecast_before_publication - continued_media_raw_forecast
    continued_media_raw_error_rate = (continued_media_raw_forecast - published_media_actual) / published_media_actual if published_media_actual else np.nan
    continued_media_final_forecast = published_media_forecast
    inactive_media_final_forecast = unpublished_media_forecast
    continued_media_final_error_rate = published_media_error_rate
    total_diff = total_forecast - total_actual
    total_error_rate = total_diff / total_actual if total_actual else np.nan
    wape = float(daily_compare["絶対誤差"].sum() / total_actual) if total_actual else np.nan

    return {
        "target_month": str(target_month),
        "learning_months": learning_months,
        "learning_month_count": learning_month_count,
        "learning_score_table": learning_score_table,
        "cutoff_date": cutoff_date,
        "learning_start": learning_start,
        "learning_calendar_day_count": len(learning_calendar_dates),
        "target_start": target_start,
        "target_end": target_end,
        "target_day_count": int(target_dates.nunique()),
        "total_forecast_before_publication": total_forecast_before_publication,
        "total_forecast": total_forecast,
        "continued_media_raw_forecast": continued_media_raw_forecast,
        "continued_media_raw_error_rate": continued_media_raw_error_rate,
        "publication_end_raw_forecast": publication_end_raw_forecast,
        "continued_media_final_forecast": continued_media_final_forecast,
        "continued_media_final_error_rate": continued_media_final_error_rate,
        "inactive_media_final_forecast": inactive_media_final_forecast,
        "published_media_count": int(len(target_published_media)),
        "published_media_forecast": published_media_forecast,
        "published_media_actual": published_media_actual,
        "published_media_diff": published_media_diff,
        "published_media_error_rate": published_media_error_rate,
        "published_wape": published_wape,
        "unpublished_media_forecast": unpublished_media_forecast,
        "total_actual": total_actual,
        "total_diff": total_diff,
        "total_error_rate": total_error_rate,
        "wape": wape,
        "media_compare": media_compare,
        "daily_compare": daily_compare,
        "excluded": factor_tables.get("excluded", pd.DataFrame()),
        "inactive_media": factor_tables.get("inactive_media", pd.DataFrame()),
        "continuation_table": continuation_bt,
        "learning_audit": learning_audit,
        "learning_summary": learning_summary,
    }


def _prepare_af_history(uploaded_af_apply, uploaded_af_issue, af_code_master, cpn_master):
    """AF申込/発行を日次で結合し、CPN名・月度をTGと同じ日付基準で付与する。"""
    frames = []
    if uploaded_af_apply is not None:
        frames.append(load_af_data(uploaded_af_apply, af_code_master, "AF申込"))
    if uploaded_af_issue is not None:
        frames.append(load_af_data(uploaded_af_issue, af_code_master, "AF発行"))
    if not frames:
        return pd.DataFrame(columns=["date", "AF申込", "AF発行", "CPN名", "月度"])

    af = frames[0].copy()
    for frame in frames[1:]:
        af = af.merge(frame, on="date", how="outer")
    for col in ["AF申込", "AF発行"]:
        if col not in af.columns:
            af[col] = 0
        af[col] = pd.to_numeric(af[col], errors="coerce").fillna(0).astype(int)
    af["date"] = pd.to_datetime(af["date"], errors="coerce").dt.normalize()

    master_cols = [c for c in ["日付", "CPN名", "月度"] if c in cpn_master.columns]
    master_for_merge = cpn_master[master_cols].copy()
    master_for_merge["日付"] = pd.to_datetime(master_for_merge["日付"], errors="coerce").dt.normalize()
    master_for_merge = master_for_merge.drop_duplicates(subset=["日付"], keep="last")
    af = af.merge(master_for_merge, left_on="date", right_on="日付", how="left")
    af["CPN名"] = af.get("CPN名", pd.Series(index=af.index, dtype="object")).fillna("通常")
    af["月度"] = (
        af.get("月度", pd.Series(index=af.index, dtype="object"))
        .astype("string").str.strip().replace("", pd.NA).fillna("未設定")
    )
    return af.drop(columns=["日付"], errors="ignore").sort_values("date").reset_index(drop=True)


def _prepare_af_code_history(uploaded_af_apply, uploaded_af_issue, af_code_master, cpn_master):
    """AF申込/発行を日付×AFコード単位で結合し、月度を付与する。"""
    frames = []
    if uploaded_af_apply is not None:
        frames.append(load_af_code_data(uploaded_af_apply, af_code_master, "AF申込"))
    if uploaded_af_issue is not None:
        frames.append(load_af_code_data(uploaded_af_issue, af_code_master, "AF発行"))
    if not frames:
        return pd.DataFrame(columns=["date", "AFコード", "AF申込", "AF発行", "月度"])

    af = frames[0].copy()
    for frame in frames[1:]:
        af = af.merge(frame, on=["date", "AFコード"], how="outer")
    for col in ["AF申込", "AF発行"]:
        if col not in af.columns:
            af[col] = 0
        af[col] = pd.to_numeric(af[col], errors="coerce").fillna(0)

    af["date"] = pd.to_datetime(af["date"], errors="coerce").dt.normalize()
    master_cols = [c for c in ["日付", "月度"] if c in cpn_master.columns]
    if "日付" in master_cols:
        master = cpn_master[master_cols].copy()
        master["日付"] = pd.to_datetime(master["日付"], errors="coerce").dt.normalize()
        master = master.drop_duplicates(subset=["日付"], keep="last")
        af = af.merge(master, left_on="date", right_on="日付", how="left")
        af = af.drop(columns=["日付"], errors="ignore")
    if "月度" not in af.columns:
        af["月度"] = "未設定"
    else:
        af["月度"] = af["月度"].astype("string").str.strip().replace("", pd.NA).fillna("未設定")
    return af.sort_values(["date", "AFコード"]).reset_index(drop=True)


def _tg_daily_measurement(history_df):
    if history_df is None or history_df.empty:
        return pd.DataFrame(columns=["date", "TG申込", "TG発行", "CPN名", "月度"])
    work = history_df.copy()
    agg = (
        work.groupby("date", as_index=False)
        .agg(TG申込=("cv", "sum"), TG発行=("approved_cv", "sum"))
    )
    meta_cols = [c for c in ["date", "CPN名", "月度"] if c in work.columns]
    if len(meta_cols) > 1:
        meta = work[meta_cols].drop_duplicates(subset=["date"], keep="last")
        agg = agg.merge(meta, on="date", how="left")
    return agg


def _render_measurement_tabs(tg_df=None, af_df=None, af_code_df=None, key_prefix="measurement"):
    """TG/AF単独表示と差分を、共通の期間・月度フィルタで表示する。"""
    tg = _tg_daily_measurement(tg_df) if tg_df is not None else pd.DataFrame()
    af = af_df.copy() if af_df is not None else pd.DataFrame()
    af_code = af_code_df.copy() if af_code_df is not None else pd.DataFrame()
    if tg.empty and af.empty:
        return

    date_candidates = []
    for df in (tg, af):
        if not df.empty and "date" in df:
            d = pd.to_datetime(df["date"], errors="coerce").dropna()
            if not d.empty:
                date_candidates.extend([d.min().date(), d.max().date()])
    if not date_candidates:
        return

    st.sidebar.header("🔎 計測結果表示条件")
    min_date, max_date = min(date_candidates), max(date_candidates)
    filter_start = st.sidebar.date_input(
        "実績表示 開始", min_date, min_value=min_date, max_value=max_date,
        key=f"{key_prefix}_start",
    )
    filter_end = st.sidebar.date_input(
        "実績表示 終了", max_date, min_value=min_date, max_value=max_date,
        key=f"{key_prefix}_end",
    )
    # 月度は文字列順ではなく、年月として時系列順に並べる。
    # 例: 2025年3月度 → 2025年4月度 → ... → 2025年12月度
    import re

    def _month_sort_key(value):
        text = str(value).strip()
        match = re.search(r"(\d{4})年\s*(\d{1,2})月度?", text)
        if match:
            return (int(match.group(1)), int(match.group(2)), text)
        # 想定外の表記は最後に回す
        return (9999, 99, text)

    month_values = sorted({
        str(v) for df in (tg, af) if not df.empty and "月度" in df
        for v in df["月度"].dropna().astype(str).tolist() if str(v).strip() and str(v) != "未設定"
    }, key=_month_sort_key)

    def _sort_monthly_rows(df):
        """月度列を yyyy年m月度 の実年月順で並べる。"""
        if df.empty or "月度" not in df.columns:
            return df
        out = df.copy()
        out["_month_sort_key"] = out["月度"].map(_month_sort_key)
        out = out.sort_values("_month_sort_key", kind="stable").drop(columns="_month_sort_key")
        return out.reset_index(drop=True)
    selected_months = st.sidebar.multiselect(
        "実績表示 月度", month_values, default=month_values, key=f"{key_prefix}_months"
    ) if month_values else []

    def apply_filters(df):
        if df.empty:
            return df.copy()
        out = df.copy()
        out["date"] = pd.to_datetime(out["date"], errors="coerce")
        out = out[out["date"].dt.date.between(filter_start, filter_end)].copy()
        if selected_months and "月度" in out.columns:
            out = out[out["月度"].astype(str).isin(selected_months)].copy()
        return out

    tg_f = apply_filters(tg)
    af_f = apply_filters(af)
    af_code_f = apply_filters(af_code)
    tab_labels = []
    if not tg.empty:
        tab_labels.append("TG計測")
    if not af.empty:
        tab_labels.append("AF計測")
    if not tg.empty and not af.empty:
        tab_labels.append("TG・AF差分")
    tabs = st.tabs(tab_labels)
    idx = 0

    if not tg.empty:
        with tabs[idx]:
            idx += 1
            monthly = (
                tg_f.groupby("月度", as_index=False)[["TG申込", "TG発行"]].sum()
                if "月度" in tg_f.columns else tg_f[["TG申込", "TG発行"]].sum().to_frame().T
            )
            monthly = _sort_monthly_rows(monthly)
            monthly["承認率"] = (
                monthly["TG発行"]
                / monthly["TG申込"].replace(0, pd.NA)
            ).fillna(0) * 100
            st.dataframe(
                monthly,
                width="stretch",
                hide_index=True,
                column_config={
                    "承認率": st.column_config.NumberColumn(
                        "承認率",
                        format="%.1f%%",
                    ),
                },
            )
            st.caption(f"期間合計：TG申込 {tg_f['TG申込'].sum():,.0f}件 / TG発行 {tg_f['TG発行'].sum():,.0f}件")

    if not af.empty:
        with tabs[idx]:
            idx += 1
            monthly = (
                af_f.groupby("月度", as_index=False)[["AF申込", "AF発行"]].sum()
                if "月度" in af_f.columns else af_f[["AF申込", "AF発行"]].sum().to_frame().T
            )
            monthly = _sort_monthly_rows(monthly)
            monthly["承認率"] = (
                monthly["AF発行"]
                / monthly["AF申込"].replace(0, pd.NA)
            ).fillna(0) * 100
            st.dataframe(
                monthly,
                width="stretch",
                hide_index=True,
                column_config={
                    "承認率": st.column_config.NumberColumn(
                        "承認率",
                        format="%.1f%%",
                    ),
                },
            )
            st.caption(f"期間合計：AF申込 {af_f['AF申込'].sum():,.0f}件 / AF発行 {af_f['AF発行'].sum():,.0f}件")

            if not af_code_f.empty:
                st.subheader("AFコード別 月度集計")

                def _render_af_code_monthly_pivot(metric_name: str):
                    if metric_name not in af_code_f.columns:
                        return

                    source = af_code_f[["月度", "AFコード", metric_name]].copy()
                    source[metric_name] = pd.to_numeric(source[metric_name], errors="coerce").fillna(0)
                    source = source[
                        source["AFコード"].astype("string").fillna("").str.strip().ne("")
                    ].copy()
                    if source.empty:
                        return

                    pivot = source.pivot_table(
                        index="月度",
                        columns="AFコード",
                        values=metric_name,
                        aggfunc="sum",
                        fill_value=0,
                    )

                    # 月度はCPNマスタに付与された値をそのまま表示し、
                    # AFコード列はマスタ/実績側の並びに依存せず見やすく固定する。
                    pivot = pivot.reindex(sorted(pivot.columns.astype(str)), axis=1)
                    pivot = pivot.reset_index()
                    pivot.columns.name = None
                    pivot = _sort_monthly_rows(pivot)

                    value_cols = [c for c in pivot.columns if c != "月度"]
                    for col in value_cols:
                        numeric = pd.to_numeric(pivot[col], errors="coerce").fillna(0)
                        if (numeric % 1 == 0).all():
                            pivot[col] = numeric.astype(int)
                        else:
                            pivot[col] = numeric

                    st.markdown(f"**{metric_name}**")
                    st.dataframe(pivot, width="stretch", hide_index=True)

                _render_af_code_monthly_pivot("AF申込")
                _render_af_code_monthly_pivot("AF発行")

    if not tg.empty and not af.empty:
        with tabs[idx]:
            tg_month = tg_f.groupby("月度", as_index=False)[["TG申込", "TG発行"]].sum()
            af_month = af_f.groupby("月度", as_index=False)[["AF申込", "AF発行"]].sum()
            diff = tg_month.merge(af_month, on="月度", how="outer").fillna(0)
            diff = _sort_monthly_rows(diff)
            diff["申込差分(TG-AF)"] = diff["TG申込"] - diff["AF申込"]
            diff["発行差分(TG-AF)"] = diff["TG発行"] - diff["AF発行"]
            diff["TG承認率"] = (
                diff["TG発行"] / diff["TG申込"].replace(0, pd.NA)
            ).fillna(0)
            diff["AF承認率"] = (
                diff["AF発行"] / diff["AF申込"].replace(0, pd.NA)
            ).fillna(0)
            diff["承認率差分(TG-AF)"] = diff["TG承認率"] - diff["AF承認率"]
            diff["申込差分率"] = (diff["申込差分(TG-AF)"] / diff["AF申込"].replace(0, pd.NA)).fillna(0)
            diff["発行差分率"] = (diff["発行差分(TG-AF)"] / diff["AF発行"].replace(0, pd.NA)).fillna(0)
            display = diff.copy()
            display["TG承認率"] = display["TG承認率"].map(lambda x: f"{x:.1%}")
            display["AF承認率"] = display["AF承認率"].map(lambda x: f"{x:.1%}")
            display["承認率差分(TG-AF)"] = display["承認率差分(TG-AF)"].map(lambda x: f"{x:+.1%}")
            display["申込差分率"] = display["申込差分率"].map(lambda x: f"{x:+.1%}")
            display["発行差分率"] = display["発行差分率"].map(lambda x: f"{x:+.1%}")
            st.dataframe(display, width="stretch", hide_index=True)
            st.caption("差分は TG − AF。プラスならTG計測の方が多く、マイナスならAF計測の方が多い値です。")

# -----------------------
# ✅ UI
# -----------------------
st.set_page_config(page_title="AFプランニングツール", layout="wide")
st.title("📊 件数予測＆プランニングツール")
st.caption(
    "実績CSVと最新のCPNマスタをアップロードしてください。"
    "CPNマスタ内の『CPNマスタ』『媒体名マスタ』シートを使用します。"
    "ファイルはGitHubには保存されません。"
)
st.markdown(
    "📎 過去の成果データは[こちら](https://rak.box.com/s/90f38ar5w8lzijpbe1h2090fubwa0etw)"
    "からDL、マージしてご使用ください。"
)

col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader(
        "① TG計測 実績CSV",
        type=["csv"],
        help="従来どおりのTG計測実績です。未アップロードでもAF計測実績があれば起動できます。",
    )
    uploaded_af_apply = st.file_uploader(
        "③ AF計測 申込実績",
        type=["csv", "xlsx", "xlsm"],
        help="A列=日付（YYYYMMDD）、B列以降の1行目=AFコード、各セル=件数のAF申込実績。",
    )

with col2:
    uploaded_master = st.file_uploader(
        "② CPNマスタ",
        type=["xlsx", "xlsm"],
        help="『CPNマスタ』『媒体名マスタ』に加え、AF計測利用時は『AFコードマスタ』シートを使用します。",
    )
    uploaded_af_issue = st.file_uploader(
        "④ AF計測 発行実績",
        type=["csv", "xlsx", "xlsm"],
        help="A列=日付（YYYYMMDD）、B列以降の1行目=AFコード、各セル=件数のAF発行実績。",
    )

has_af_data = uploaded_af_apply is not None or uploaded_af_issue is not None
has_any_actual = uploaded_file is not None or has_af_data

if uploaded_master and has_any_actual:
    exclude_compensation = True
    if uploaded_file is not None:
        st.sidebar.header("⚙️ 実績データ条件")
        exclude_compensation = st.sidebar.toggle(
            "補填を除外",
            value=True,
            help="ON: TG実績データP列に文字列が入っている補填対象行を除外します。",
        )

    try:
        from data.loader import add_business_edge_flags
        from logic.factors import (
            calculate_dynamic_factor_tables,
            calculate_normal_month_base,
            calculate_selected_cpn_base,
            get_cpn_reference_periods,
            enforce_premium_media_cost,
            calculate_media_continuation_table,
        )
        from config.constants import RECENT_NORMAL_DAYS

        history_df = pd.DataFrame()
        tg_raw_max_date = pd.NaT
        if uploaded_file is not None:
            history_df = load_data(uploaded_file, exclude_compensation=exclude_compensation)
            history_df["date"] = pd.to_datetime(
                history_df["date"],
                errors="coerce",
            ).dt.normalize()
            # バックテスト可否は媒体/商品IDで絞り込む前のTGローデータ最終日で判定する。
            # 特定媒体に行がない日を「データ未完了」と誤判定しないため。
            tg_raw_max_date = history_df["date"].max()

        # 同一Excel内の「CPNマスタ」「媒体名マスタ」を1回だけ読み込み、
        # rerun時はキャッシュされた結果を再利用する。
        cpn_master, media_master, af_code_master = load_master_workbook_with_af(uploaded_master)

        required_master_columns = {"日付", "CPN名", "月度"}
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
        cpn_master["月度"] = (
            cpn_master["月度"]
            .astype("string")
            .str.strip()
        )
        cpn_master = cpn_master.dropna(
            subset=["日付", "CPN名"]
        )
        cpn_master["月度"] = (
            cpn_master["月度"]
            .replace("", pd.NA)
            .fillna("未設定")
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

        af_history_df = _prepare_af_history(
            uploaded_af_apply, uploaded_af_issue, af_code_master, cpn_master
        ) if has_af_data else pd.DataFrame()
        af_code_history_df = _prepare_af_code_history(
            uploaded_af_apply, uploaded_af_issue, af_code_master, cpn_master
        ) if has_af_data else pd.DataFrame()


    except Exception as exc:
        st.error(f"ファイルの読み込みに失敗しました: {exc}")
        st.stop()

    if uploaded_file is None:
        st.subheader("📊 計測結果")
        _render_measurement_tabs(tg_df=None, af_df=af_history_df, af_code_df=af_code_history_df, key_prefix="af_only")
        st.info("AF計測実績のみを読み込みました。TG実績がないため、従来の媒体別予測・最適化・提出用Excel生成は表示していません。")
        st.stop()

    history_df = history_df.merge(
        cpn_master[
            [
                "日付",
                "CPN名",
                "月度",
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
    history_df["月度"] = (
        history_df["月度"]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
        .fillna("未設定")
    )

    history_df = history_df.merge(
        media_master[["SID", "媒体名", "カテゴリ"]],
        on="SID",
        how="left",
    )
    history_df["raw_media"] = history_df["media"]
    mapped_media = (
        history_df["媒体名"]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
    )
    history_df["media"] = mapped_media.fillna(history_df["raw_media"])
    history_df["media_category"] = (
        history_df["カテゴリ"]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
        .fillna("未分類")
    )

    st.subheader("📊 計測結果")
    _render_measurement_tabs(
        tg_df=history_df,
        af_df=af_history_df if has_af_data else None,
        af_code_df=af_code_history_df if has_af_data else None,
        key_prefix="measurement_compare",
    )

    st.sidebar.header("商品ID選択")
    all_product_ids = sorted(
        [x for x in history_df["商品ID"].dropna().astype(str).unique() if str(x).strip() != ""]
    )
    # 商品IDは、存在する場合は 1 と 2 をデフォルト選択。
    # どちらも無い場合だけ、従来どおり先頭1件を初期値にする。
    default_product_ids = []
    for target_id in (1, 2):
        matched_id = next(
            (pid for pid in all_product_ids if str(pid).strip() == str(target_id)),
            next(
                (pid for pid in all_product_ids if pd.to_numeric(pid, errors="coerce") == target_id),
                None,
            ),
        )
        if matched_id is not None and matched_id not in default_product_ids:
            default_product_ids.append(matched_id)
    if not default_product_ids:
        default_product_ids = all_product_ids[:1]
    selected_product_ids = st.sidebar.multiselect(
        "商品ID",
        all_product_ids,
        default=default_product_ids,
    )
    if not selected_product_ids:
        st.stop()

    history_df = history_df[
        history_df["商品ID"].astype(str).isin(selected_product_ids)
    ].copy()

    st.sidebar.header("媒体選択")
    all_categories = sorted(history_df["media_category"].dropna().astype(str).unique())
    selected_categories = st.sidebar.multiselect(
        "カテゴリ",
        all_categories,
        default=all_categories,
    )
    if not selected_categories:
        st.stop()

    history_df = history_df[
        history_df["media_category"].isin(selected_categories)
    ].copy()

    all_media = sorted(history_df["media"].dropna().astype(str).unique())
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

    # 月度別過去実績分析はUI上部へ表示する。
    # 今回プランは後段で確定するため、ここでは表示位置だけ確保し、
    # 手動設定取得後にこのプレースホルダーへ描画する。
    history_analytics_placeholder = st.empty()

    # 定常 / マジ得の承認率を画面でも確認。
    try:
        _period_preview = _calculate_period_media_metrics(history_df)
        approval_preview = pd.DataFrame(
            [
                {
                    "媒体": media,
                    "定常承認率": _period_preview["normal_rate"].get(
                        media,
                        _period_preview["normal_rate_all"],
                    ),
                    "マジ得承認率": _period_preview["magi_rate"].get(
                        media,
                        _period_preview["magi_rate_all"],
                    ),
                    "定常単価": _period_preview["normal_unit"].get(
                        media,
                        _period_preview["normal_unit_all"],
                    ),
                    "マジ得単価": _period_preview["magi_unit"].get(
                        media,
                        _period_preview["magi_unit_all"],
                    ),
                }
                for media in selected_media
            ]
        )

        approval_preview["定常承認率"] = approval_preview["定常承認率"].map(
            lambda x: f"{x:.1%}"
        )
        approval_preview["マジ得承認率"] = approval_preview["マジ得承認率"].map(
            lambda x: f"{x:.1%}"
        )
        approval_preview["定常単価"] = approval_preview["定常単価"].map(
            lambda x: f"¥{x:,.0f}"
        )
        approval_preview["マジ得単価"] = approval_preview["マジ得単価"].map(
            lambda x: f"¥{x:,.0f}"
        )

        with st.expander("✅ 定常・マジ得の過去実績指標"):
            st.caption(
                "承認率 = 成果承認フラグYの件数 ÷ 全件数 / "
                "単価 = ローデータW列『グロス』の直近実績"
            )
            st.dataframe(
                approval_preview,
                width="stretch",
                hide_index=True,
            )

    except ValueError as approval_exc:
        st.warning(
            "定常・マジ得指標を算出できません。"
            f"{approval_exc}"
        )

    st.sidebar.header("📊 対象期間・施策")

    cpn_list = sorted(
        cpn_master["CPN名"]
        .dropna()
        .astype(str)
        .unique()
    )
    if not cpn_list:
        st.error("CPNマスタに選択可能な施策がありません。")
        st.stop()

    normal_labels = {"通常", "定常"}
    default_primary = next((x for x in ["定常", "通常"] if x in cpn_list), cpn_list[0])
    today = datetime.date.today()

    selected_cpn_1 = st.sidebar.selectbox(
        "施策①",
        cpn_list,
        index=cpn_list.index(default_primary),
        key="planning_cpn_1",
    )
    start_date_1 = st.sidebar.date_input(
        "施策① 開始",
        today,
        key="planning_start_1",
    )
    end_date_1 = st.sidebar.date_input(
        "施策① 終了",
        today + datetime.timedelta(days=7),
        key="planning_end_1",
    )

    use_second_period = st.sidebar.toggle(
        "施策②を追加",
        value=False,
        key="planning_use_second",
    )

    selected_cpn_2 = None
    start_date_2 = None
    end_date_2 = None
    if use_second_period:
        default_second = "マジ得" if "マジ得" in cpn_list else cpn_list[0]
        selected_cpn_2 = st.sidebar.selectbox(
            "施策②",
            cpn_list,
            index=cpn_list.index(default_second),
            key="planning_cpn_2",
        )
        start_date_2 = st.sidebar.date_input(
            "施策② 開始",
            end_date_1 + datetime.timedelta(days=1),
            key="planning_start_2",
        )
        end_date_2 = st.sidebar.date_input(
            "施策② 終了",
            end_date_1 + datetime.timedelta(days=7),
            key="planning_end_2",
        )

    planning_segments = [
        {"label": "施策①", "cpn": selected_cpn_1, "start": start_date_1, "end": end_date_1}
    ]
    if use_second_period:
        planning_segments.append(
            {"label": "施策②", "cpn": selected_cpn_2, "start": start_date_2, "end": end_date_2}
        )

    for seg in planning_segments:
        if seg["start"] > seg["end"]:
            st.error(f"{seg['label']}の開始日は終了日以前にしてください。")
            st.stop()

    if use_second_period:
        r1 = pd.date_range(start_date_1, end_date_1)
        r2 = pd.date_range(start_date_2, end_date_2)
        if len(r1.intersection(r2)) > 0:
            st.error("施策①と施策②の期間が重複しています。期間が重ならないように設定してください。")
            st.stop()

    # 定常学習はCPN月度数ではなく、予測開始日前の直近60暦日に固定する。
    # 月度日数・定常日数が不均一でも、同じ時間幅で現在の媒体力を評価する。
    needs_normal_learning = any(seg["cpn"] in normal_labels for seg in planning_segments)
    selected_learning_months = []
    normal_training_df = history_df
    normal_learning_calendar_dates = None
    normal_learning_start = None
    normal_learning_cutoff = None
    if needs_normal_learning:
        first_normal_start = min(
            pd.Timestamp(seg["start"]).normalize()
            for seg in planning_segments if seg["cpn"] in normal_labels
        )
        normal_learning_cutoff = first_normal_start - pd.Timedelta(days=1)
        requested_normal_learning_start = normal_learning_cutoff - pd.Timedelta(days=59)
        history_dates = pd.to_datetime(history_df["date"], errors="coerce").dt.normalize()
        available_history_dates = history_dates.loc[history_dates.le(normal_learning_cutoff)].dropna()
        if available_history_dates.empty:
            st.error("予測開始日より前の実績データがありません。")
            st.stop()
        earliest_history_date = pd.Timestamp(available_history_dates.min()).normalize()
        normal_learning_start = max(requested_normal_learning_start, earliest_history_date)
        normal_training_df = history_df.loc[
            history_dates.between(normal_learning_start, normal_learning_cutoff)
        ].copy()
        normal_learning_calendar_dates = (
            cpn_master.loc[
                pd.to_datetime(cpn_master["日付"], errors="coerce").dt.normalize()
                .between(normal_learning_start, normal_learning_cutoff)
                & cpn_master["CPN名"].astype(str).str.strip().isin(normal_labels),
                "日付",
            ]
            .pipe(pd.to_datetime, errors="coerce")
            .dropna().dt.normalize().drop_duplicates().sort_values().tolist()
        )
        selected_learning_months = (
            normal_training_df.loc[normal_training_df["CPN名"].isin(normal_labels), "月度"]
            .dropna().astype(str).loc[lambda x: x.ne("未設定")].drop_duplicates().tolist()
        )
        if not normal_learning_calendar_dates:
            st.error("予測開始日前の利用可能データ内にCPNマスタの定常日がありません。")
            st.stop()
        st.sidebar.caption(
            f"定常学習: {normal_learning_start.strftime('%Y/%m/%d')}〜"
            f"{normal_learning_cutoff.strftime('%Y/%m/%d')}（最大60日・利用可能実績を使用）"
        )

    segment_bases = []
    reference_descriptions = []

    for seg_idx, seg in enumerate(planning_segments, start=1):
        selected_cpn = seg["cpn"]
        if selected_cpn in normal_labels:
            base_pair_seg = _calculate_normal_month_base(
                normal_training_df,
                selected_learning_months,
                calendar_dates=normal_learning_calendar_dates,
            )
            if base_pair_seg.empty:
                st.error("予測開始日前の利用可能期間に定常実績がありません。")
                st.stop()
            reference_key_seg = ("normal_60days", str(normal_learning_start), str(normal_learning_cutoff))
            reference_descriptions.append(
                f"{seg['label']} {selected_cpn}: 定常学習 最大60日 "
                f"({normal_learning_start.strftime('%Y/%m/%d')}〜{normal_learning_cutoff.strftime('%Y/%m/%d')})"
            )
        else:
            available_periods = get_cpn_reference_periods(history_df, selected_cpn)
            if not available_periods:
                st.error(f"実績内に『{selected_cpn}』のキャンペーン期間がありません。")
                st.stop()

            period_options = {
                f"{p_start.strftime('%Y/%m/%d')} ～ {p_end.strftime('%Y/%m/%d')}": (p_start, p_end)
                for p_start, p_end in available_periods
            }
            selected_period_labels = st.sidebar.multiselect(
                f"{seg['label']} {selected_cpn} 参照期間",
                options=list(period_options.keys()),
                default=list(period_options.keys()),
                key=f"cpn_reference_periods_{seg_idx}",
            )
            if not selected_period_labels:
                st.warning(f"{seg['label']}の参照期間を1つ以上選択してください。")
                st.stop()
            selected_periods = [period_options[label] for label in selected_period_labels]
            base_pair_seg = calculate_selected_cpn_base(
                history_df,
                selected_cpn,
                selected_periods,
            )
            if base_pair_seg.empty:
                st.error(f"{seg['label']}の選択参照期間に対象媒体の実績がありません。")
                st.stop()
            reference_key_seg = (selected_cpn, tuple(selected_period_labels))
            reference_descriptions.append(
                f"{seg['label']} {selected_cpn}: " + " / ".join(selected_period_labels)
            )

        base_pair_seg = base_pair_seg[base_pair_seg["media"].isin(selected_media)].copy()
        if base_pair_seg.empty:
            st.error(f"{seg['label']}の学習実績に対象媒体がありません。")
            st.stop()
        segment_bases.append(
            {**seg, "base_pair": base_pair_seg, "reference_key": reference_key_seg}
        )

    st.subheader("📈 予測ベース")
    st.caption(
        "定常は選択月度の定常日平均、その他施策は選択した過去CPN期間の日平均を使用します。"
    )
    for desc in reference_descriptions:
        st.write(f"・{desc}")

    base_preview_rows = []
    for seg in segment_bases:
        tmp = (
            seg["base_pair"].groupby("media", as_index=False)
            .agg(base_cv=("base_cv", "sum"), cost=("cost", "sum"))
        )
        tmp["施策"] = seg["label"] + "：" + seg["cpn"]
        base_preview_rows.append(tmp)
    base_preview = pd.concat(base_preview_rows, ignore_index=True)
    base_preview = base_preview.rename(
        columns={"media": "媒体", "base_cv": "基礎CV/日", "cost": "基礎COST/日"}
    )
    base_preview["基礎CV/日"] = base_preview["基礎CV/日"].round(2)
    base_preview["基礎COST/日"] = base_preview["基礎COST/日"].round(0)
    st.dataframe(
        base_preview[["施策", "媒体", "基礎CV/日", "基礎COST/日"]],
        width="stretch",
        hide_index=True,
    )

    # 既存関数との互換用。提出用の代表施策は施策①とするが、
    # 日別施策はfuture_df側へ保持して複合期間を正しく予測する。
    selected_cpn = selected_cpn_1
    start_date = min(seg["start"] for seg in planning_segments)
    end_date = max(seg["end"] for seg in planning_segments)
    base_pair = pd.concat([seg["base_pair"] for seg in segment_bases], ignore_index=True)

    st.sidebar.header("🎯 最適化ロジック")

    opt_mode = st.sidebar.radio(
        "最適基準",
        ["単価最小", "CV最大"],
        index=0,
    )

    # ---------------------------------------------------------
    # 変動係数は同じ入力条件なら再計算しない。
    # Excel生成ボタン等によるStreamlit再実行でも再利用する。
    # ---------------------------------------------------------
    factor_cache_key = (
        tuple(selected_media),
        tuple(selected_product_ids),
        bool(exclude_compensation),
        len(history_df),
        str(history_df["date"].min()),
        str(history_df["date"].max()),
        float(pd.to_numeric(history_df["cv"], errors="coerce").fillna(0).sum()),
        float(pd.to_numeric(history_df["cost"], errors="coerce").fillna(0).sum()),
        (str(normal_learning_start), str(normal_learning_cutoff)) if needs_normal_learning else (),
    )

    if st.session_state.get("_factor_cache_key") == factor_cache_key:
        factor_tables = st.session_state["_factor_tables"]
    else:
        factor_tables = calculate_dynamic_factor_tables(
            normal_training_df if needs_normal_learning else history_df,
            selected_learning_months if needs_normal_learning else None,
            calendar_dates=normal_learning_calendar_dates if needs_normal_learning else None,
        )
        st.session_state["_factor_cache_key"] = factor_cache_key
        st.session_state["_factor_tables"] = factor_tables

    st.subheader("📐 実績から算出した変動係数")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["曜日", "月初・月末", "需要期", "LINE OA", "学習除外日", "休眠媒体"]
    )

    with tab1:
        st.dataframe(
            factor_tables["weekday"].round(3),
            width="stretch",
            hide_index=True,
        )

    with tab2:
        st.dataframe(
            factor_tables["month_edge"].round(3),
            width="stretch",
            hide_index=True,
        )

    with tab3:
        st.dataframe(
            factor_tables["season"].round(3),
            width="stretch",
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
                width="stretch",
                hide_index=True,
            )


    with tab5:
        excluded = factor_tables.get("excluded", pd.DataFrame())
        if excluded.empty:
            st.info("マジ得直後・異常値による学習除外日はありません。")
        else:
            preview_excluded = excluded.copy()
            if "date" in preview_excluded.columns:
                preview_excluded["date"] = pd.to_datetime(preview_excluded["date"]).dt.strftime("%Y/%m/%d")
            st.caption("定常の基礎値・変動係数の学習から除外した媒体×日です。予測対象実績そのものは削除しません。")
            st.dataframe(
                preview_excluded.round(3),
                width="stretch",
                hide_index=True,
            )

    with tab6:
        inactive_media = factor_tables.get("inactive_media", pd.DataFrame())
        if inactive_media.empty:
            st.info("直近30日CVなしで休眠判定された媒体はありません。")
        else:
            inactive_preview = inactive_media.copy()
            if "last_positive_date" in inactive_preview.columns:
                inactive_preview["last_positive_date"] = pd.to_datetime(
                    inactive_preview["last_positive_date"]
                ).dt.strftime("%Y/%m/%d")
            inactive_preview = inactive_preview.rename(
                columns={
                    "media": "媒体",
                    "last_positive_date": "最終CV日",
                    "recent_cv": "直近30日CV",
                    "reason": "判定理由",
                }
            )
            st.caption("過去にCV実績はあるものの、予測時点の直近30日でCVがない媒体です。定常予測の基礎値から除外します。")
            st.dataframe(
                inactive_preview.drop(columns=["lookback_days"], errors="ignore"),
                width="stretch",
                hide_index=True,
            )

    # ---------------------------------------------------------
    # 定常媒体の掲載予定。過去データだけでは翌月度の掲載終了を完全には予知できないため、
    # 営業が知っている情報を優先し、未定だけ過去の掲載継続率で期待値化する。
    # ---------------------------------------------------------
    publication_factor_map = {m: 1.0 for m in selected_media}
    publication_status_map = {m: "掲載予定" for m in selected_media}
    continuation_table_main = pd.DataFrame()
    if needs_normal_learning:
        continuation_history = history_df.loc[
            pd.to_datetime(history_df["date"], errors="coerce").dt.normalize().le(normal_learning_cutoff)
        ].copy()
        continuation_table_main = calculate_media_continuation_table(
            continuation_history,
            cpn_master=cpn_master,
            cutoff_date=normal_learning_cutoff,
        )
        cont_map_main = (
            continuation_table_main.set_index("media")["continuation_rate"].to_dict()
            if not continuation_table_main.empty else {}
        )
        global_cont_main = (
            float(continuation_table_main["global_continuation_rate"].iloc[0])
            if not continuation_table_main.empty else 0.70
        )

        st.subheader("📌 定常媒体の掲載予定")
        st.caption(
            "営業側で把握している掲載予定を優先します。掲載予定は100%、掲載なしは0件、未定だけ過去の掲載継続率で期待値化します。"
        )

        # 掲載状態はsession_stateへ保持し、表の再描画や一括変更でも選択を維持する。
        status_store = st.session_state.setdefault("_publication_status_values", {})
        for media in selected_media:
            status_store.setdefault(str(media), "未定")
        valid_statuses = {"掲載予定", "未定", "掲載なし"}
        status_store = {
            str(m): (status_store.get(str(m), "未定") if status_store.get(str(m), "未定") in valid_statuses else "未定")
            for m in selected_media
        }
        st.session_state["_publication_status_values"] = status_store
        st.session_state.setdefault("_publication_editor_version", 0)

        # ExcelなどのA列SID / B列媒体名をそのまま貼り付けて、掲載媒体を一括反映できる。
        # SIDを最優先で照合し、SIDが空欄・不一致の場合のみ媒体名で補完する。
        with st.expander("📋 掲載媒体をコピペで一括設定", expanded=False):
            st.caption(
                "ExcelのA列=SID、B列=媒体名をそのままコピーして貼り付けてください。"
                "記載のある媒体だけを『掲載予定』、それ以外を『掲載なし』にします。"
            )
            pasted_publication_text = st.text_area(
                "掲載媒体一覧",
                key="publication_media_paste_text",
                height=180,
                placeholder="123456\tモッピー\n234567\tハピタス\n345678\tLINEポイントクラブ",
                label_visibility="collapsed",
            )
            if st.button(
                "貼り付けた媒体を掲載予定に反映",
                use_container_width=True,
                key="publication_apply_pasted_list",
                type="primary",
            ):
                # 現在選択中の媒体についてSIDマスタを作る。
                media_sid_map = {}
                if "SID" in history_df.columns:
                    sid_src = history_df[["media", "SID"]].copy()
                    sid_src["media"] = sid_src["media"].astype(str).str.strip()
                    sid_src["SID"] = sid_src["SID"].map(_normalize_sid)
                    sid_src = sid_src[sid_src["media"].isin([str(m) for m in selected_media])]
                    sid_src = sid_src[sid_src["SID"] != ""]
                    for media, grp in sid_src.groupby("media", sort=False):
                        # 同一媒体に複数SIDがある場合も、どれか1つが一致すれば掲載予定にする。
                        media_sid_map[str(media)] = set(grp["SID"].astype(str))

                sid_to_media = {}
                for media, sid_values in media_sid_map.items():
                    for sid in sid_values:
                        sid_to_media.setdefault(sid, set()).add(media)
                name_to_media = {str(m).strip(): str(m) for m in selected_media}

                matched_media = set()
                unmatched_rows = []
                input_rows = 0
                for raw_line in (pasted_publication_text or "").splitlines():
                    line = raw_line.strip()
                    if not line:
                        continue
                    # Excel貼り付けはタブ区切り。念のためカンマ区切りにも対応。
                    parts = [x.strip() for x in (line.split("\t") if "\t" in line else line.split(","))]
                    sid = _normalize_sid(parts[0]) if parts else ""
                    media_name = parts[1].strip() if len(parts) >= 2 else ""
                    # ヘッダー行は読み飛ばす。
                    if sid.upper() == "SID" or media_name in {"媒体", "媒体名"}:
                        continue
                    input_rows += 1

                    row_matches = set(sid_to_media.get(sid, set())) if sid else set()
                    if not row_matches and media_name:
                        exact_media = name_to_media.get(media_name)
                        if exact_media:
                            row_matches.add(exact_media)
                    if row_matches:
                        matched_media.update(row_matches)
                    else:
                        unmatched_rows.append(line)

                if input_rows == 0:
                    st.warning("掲載媒体が入力されていません。SID・媒体名を貼り付けてください。")
                else:
                    st.session_state["_publication_status_values"] = {
                        str(m): ("掲載予定" if str(m) in matched_media else "掲載なし")
                        for m in selected_media
                    }
                    st.session_state["_publication_editor_version"] += 1
                    st.session_state["_publication_paste_result"] = {
                        "matched": len(matched_media),
                        "none": len(selected_media) - len(matched_media),
                        "unmatched": unmatched_rows,
                    }
                    st.rerun()

            paste_result = st.session_state.pop("_publication_paste_result", None)
            if paste_result:
                st.success(
                    f"{paste_result['matched']}媒体を掲載予定、"
                    f"{paste_result['none']}媒体を掲載なしに設定しました。"
                )
                if paste_result["unmatched"]:
                    st.warning(
                        f"一致しなかった入力が{len(paste_result['unmatched'])}件あります。"
                        "SIDまたは媒体名を確認してください。"
                    )
                    st.code("\n".join(paste_result["unmatched"][:30]))

        bulk_cols = st.columns(3)
        if bulk_cols[0].button("全て掲載予定", use_container_width=True, key="publication_all_planned"):
            st.session_state["_publication_status_values"] = {str(m): "掲載予定" for m in selected_media}
            st.session_state["_publication_editor_version"] += 1
            st.rerun()
        if bulk_cols[1].button("全て未定", use_container_width=True, key="publication_all_unknown"):
            st.session_state["_publication_status_values"] = {str(m): "未定" for m in selected_media}
            st.session_state["_publication_editor_version"] += 1
            st.rerun()
        if bulk_cols[2].button("全て掲載なし", use_container_width=True, key="publication_all_none"):
            st.session_state["_publication_status_values"] = {str(m): "掲載なし" for m in selected_media}
            st.session_state["_publication_editor_version"] += 1
            st.rerun()

        status_store = st.session_state["_publication_status_values"]
        publication_editor = pd.DataFrame({
            "媒体": selected_media,
            "掲載状態": [status_store.get(str(m), "未定") for m in selected_media],
        })
        editor_key = f"normal_publication_status_editor_{st.session_state['_publication_editor_version']}"
        edited_publication = st.data_editor(
            publication_editor,
            width="stretch",
            hide_index=True,
            disabled=["媒体"],
            column_config={
                "掲載状態": st.column_config.SelectboxColumn(
                    "掲載状態",
                    options=["掲載予定", "未定", "掲載なし"],
                    required=True,
                ),
            },
            key=editor_key,
        )

        # 編集結果を保存し、その場で予測係数へ反映する。
        updated_store = dict(status_store)
        for _, r in edited_publication.iterrows():
            media = str(r["媒体"])
            status = str(r["掲載状態"])
            updated_store[media] = status if status in valid_statuses else "未定"
        st.session_state["_publication_status_values"] = updated_store

        for media in selected_media:
            media = str(media)
            status = updated_store.get(media, "未定")
            cont = float(cont_map_main.get(media, global_cont_main))
            publication_status_map[media] = status
            publication_factor_map[media] = 1.0 if status == "掲載予定" else (0.0 if status == "掲載なし" else cont)

        status_counts = pd.Series([publication_status_map.get(str(m), "未定") for m in selected_media]).value_counts()
        st.caption(
            f"掲載予定 {int(status_counts.get('掲載予定', 0))}媒体 / "
            f"未定 {int(status_counts.get('未定', 0))}媒体 / "
            f"掲載なし {int(status_counts.get('掲載なし', 0))}媒体"
        )

        with st.expander("未定時の掲載継続率（分析用）"):
            continuation_preview = pd.DataFrame({
                "媒体": selected_media,
                "掲載継続率(%)": [float(cont_map_main.get(m, global_cont_main)) * 100.0 for m in selected_media],
            })
            continuation_preview["掲載継続率(%)"] = continuation_preview["掲載継続率(%)"].clip(0, 100).round(0).astype(int)
            st.dataframe(continuation_preview, width="stretch", hide_index=True)

    with st.expander("🧪 定常バックテスト", expanded=False):
        st.caption(
            "対象月度の実績を予測計算から隠し、CPNマスタ上の対象月度開始日前までのデータだけで当時の定常予測を再現します。"
            "対象月の実績は比較表示にだけ使用します。"
        )

        # 選択肢・月度順もローデータではなくCPNマスタを正とする。
        bt_master = cpn_master.copy()
        bt_master["日付"] = pd.to_datetime(bt_master["日付"], errors="coerce").dt.normalize()
        bt_master["月度"] = bt_master["月度"].astype("string").str.strip()
        bt_master["CPN名"] = bt_master["CPN名"].astype("string").str.strip()
        bt_month_labels = _get_cpn_normal_months(bt_master)
        # 実績比較ができ、かつCPNマスタ上の対象最終日までTGローデータが到達している
        # 「完了月度」だけをバックテスト候補にする。途中月度を完成実績として評価しない。
        history_months = set(
            history_df.loc[
                history_df["CPN名"].isin(normal_labels)
                & history_df["月度"].astype(str).ne("未設定"),
                "月度",
            ].astype(str)
        )
        bt_options = []
        incomplete_bt_months = []
        for m in bt_month_labels[1:]:
            if m not in history_months:
                continue
            m_dates = bt_master.loc[
                bt_master["月度"].astype(str).eq(str(m))
                & bt_master["CPN名"].isin(normal_labels),
                "日付",
            ].dropna()
            if m_dates.empty:
                continue
            m_end = pd.Timestamp(m_dates.max()).normalize()
            if pd.notna(tg_raw_max_date) and pd.Timestamp(tg_raw_max_date).normalize() >= m_end:
                bt_options.append(m)
            else:
                incomplete_bt_months.append((m, m_end))

        if not bt_options:
            st.info("バックテストには、CPNマスタ上の対象最終日まで実績が揃った月度が必要です。")
        else:
            if incomplete_bt_months:
                latest_incomplete, latest_end = incomplete_bt_months[-1]
                raw_last_text = pd.Timestamp(tg_raw_max_date).strftime("%Y/%m/%d") if pd.notna(tg_raw_max_date) else "-"
                st.caption(
                    f"未完了月度は候補から除外しています（ローデータ最終日: {raw_last_text}）。"
                )
            bt_target_month = st.selectbox(
                "検証する月度",
                options=bt_options,
                index=len(bt_options) - 1,
                key="normal_backtest_target_month",
                help="月度はCPNマスタ基準です。対象月度のCPNマスタ登録日だけを予測・実績比較します。",
            )

            try:
                bt_result = _run_normal_backtest(
                    history_df,
                    cpn_master,
                    bt_target_month,
                    auto_select_learning=False,
                )
                st.caption(
                    f"検証対象: CPNマスタ {bt_result['target_month']}（定常{bt_result['target_day_count']}日） ／ "
                    f"予測基準日: {bt_result['cutoff_date'].strftime('%Y/%m/%d')} ／ "
                    f"学習期間: {bt_result['learning_start'].strftime('%Y/%m/%d')}〜{bt_result['cutoff_date'].strftime('%Y/%m/%d')} "
                    f"（最大60日・定常{bt_result['learning_calendar_day_count']}日）"
                )

                st.caption(
                    f"対象月度に実際に掲載があった {bt_result.get('published_media_count', 0):,}媒体だけを自動抽出して評価しています。"
                )
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("予測CV", f"{bt_result.get('published_media_forecast', 0):,.0f}")
                m2.metric("実績CV", f"{bt_result.get('published_media_actual', 0):,.0f}")
                m3.metric(
                    "差分",
                    f"{bt_result.get('published_media_diff', 0):+,.0f}",
                    delta=(
                        f"{bt_result.get('published_media_error_rate'):+.1%}"
                        if pd.notna(bt_result.get("published_media_error_rate"))
                        else None
                    ),
                    delta_color="off",
                )
                m4.metric(
                    "日次WAPE",
                    f"{bt_result.get('published_wape'):.1%}" if pd.notna(bt_result.get("published_wape")) else "-",
                    help="対象月度に掲載があった媒体だけの、媒体×日ごとの絶対誤差合計 ÷ 実績CV合計。小さいほど予測精度が高い指標です。",
                )

                bt_display = bt_result["media_compare"].copy()
                bt_display["予測CV"] = bt_display["予測CV"].round(0).astype(int)
                bt_display["実績CV"] = bt_display["実績CV"].round(0).astype(int)
                bt_display["差分"] = bt_display["差分"].round(0).astype(int)
                bt_display["誤差率"] = bt_display["誤差率"].map(
                    lambda x: f"{x:+.1%}" if pd.notna(x) else "-"
                )
                diagnostic_cols = [
                    "media", "学習稼働率", "稼働時日平均CV", "基礎期待CV/日", "学習対象日数", "稼働日数",
                    "基礎CV", "曜日補正後CV", "曜日影響CV",
                    "需要期補正後CV", "需要期影響CV",
                    "月初月末補正後CV", "月初月末影響CV",
                    "マジ得後補正後CV", "マジ得後影響CV",
                    "LINE補正後CV", "LINE影響CV",
                    "掲載判定前CV", "予測掲載継続率", "掲載継続影響CV", "対象月掲載実績",
                    "予測CV", "実績CV", "差分", "誤差率",
                ]
                diagnostic_cols = [c for c in diagnostic_cols if c in bt_display.columns]
                bt_export = bt_display[diagnostic_cols].rename(columns={"media": "媒体"})
                # 文字列の判定列を数値変換すると空欄になるため、明示的に除外する。
                numeric_diag = [c for c in bt_export.columns if c not in {"媒体", "誤差率", "対象月掲載実績"}]
                for c in numeric_diag:
                    bt_export[c] = pd.to_numeric(bt_export[c], errors="coerce").round(2)

                # 営業向けは、対象月度に実際に掲載があった媒体だけを自動抽出する。
                bt_export_published = bt_export.loc[
                    bt_export["対象月掲載実績"].astype(str).eq("掲載あり")
                ].copy() if "対象月掲載実績" in bt_export.columns else bt_export.copy()

                bt_month_filename = str(bt_target_month).replace("/", "-").replace(" ", "")
                st.download_button(
                    "📥 バックテストCSVを保存",
                    data=bt_export_published.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"AFF定常予測_バックテスト_{bt_month_filename}.csv",
                    mime="text/csv",
                    key=f"download_backtest_media_{bt_month_filename}",
                    type="primary",
                )

                # 営業向けの評価表は、掲載媒体かつ判断に必要な項目だけを表示する。
                sales_cols = ["媒体", "予測CV", "実績CV", "差分", "誤差率"]
                sales_cols = [c for c in sales_cols if c in bt_export_published.columns]
                st.dataframe(
                    bt_export_published[sales_cols],
                    width="stretch",
                    hide_index=True,
                )

                # 掲載されなかった媒体への予測や全媒体ベースの評価は、分析用として普段は隠す。
                with st.expander("全媒体・掲載判定の詳細（分析用）"):
                    b1, b2, b3 = st.columns(3)
                    b1.metric("全媒体の最終予測", f"{bt_result.get('total_forecast', 0):,.0f}")
                    b2.metric("掲載媒体の予測", f"{bt_result.get('published_media_forecast', 0):,.0f}")
                    b3.metric(
                        "掲載なし媒体への予測",
                        f"{bt_result.get('unpublished_media_forecast', 0):,.0f}",
                        help="対象月度のローデータに登場しなかった媒体へ残っていた予測です。",
                    )
                    st.dataframe(
                        bt_export,
                        width="stretch",
                        hide_index=True,
                    )
                    st.download_button(
                        "📥 全媒体の診断CSVを保存",
                        data=bt_export.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"AFF定常予測_バックテスト診断_全媒体_{bt_month_filename}.csv",
                        mime="text/csv",
                        key=f"download_backtest_all_media_{bt_month_filename}",
                    )

                # 学習データそのものを監査する。予測値には影響しない診断表示。
                with st.expander("学習データ診断（分析用）"):
                    audit = bt_result.get("learning_audit", pd.DataFrame()).copy()
                    summary = bt_result.get("learning_summary", {})
                    if summary:
                        first_day = summary.get("最初の日")
                        last_day = summary.get("最後の日")
                        first_text = pd.Timestamp(first_day).strftime("%Y/%m/%d") if pd.notna(first_day) else "-"
                        last_text = pd.Timestamp(last_day).strftime("%Y/%m/%d") if pd.notna(last_day) else "-"
                        a1, a2, a3, a4 = st.columns(4)
                        a1.metric("学習実績日数", f"{summary.get('実績存在日数', 0):,}日")
                        a2.metric("学習CV合計", f"{summary.get('CV合計', 0):,.0f}")
                        a3.metric("学習媒体数", f"{summary.get('媒体数', 0):,}")
                        a4.metric("元データ行数", f"{summary.get('元データ行数', 0):,}")
                        st.caption(
                            f"実際に学習対象として抽出した期間: {first_text}〜{last_text} ／ "
                            f"異常値・マジ得直後の除外媒体日: {summary.get('除外媒体日数', 0):,} ／ "
                            f"休眠除外媒体: {summary.get('休眠除外媒体数', 0):,}"
                        )
                    if not audit.empty:
                        for c in ["CPN最初の日", "CPN最後の日"]:
                            audit[c] = pd.to_datetime(audit[c], errors="coerce").dt.strftime("%Y/%m/%d").fillna("-")
                        audit["CV合計"] = pd.to_numeric(audit["CV合計"], errors="coerce").fillna(0).round(0).astype(int)
                        st.dataframe(audit, width="stretch", hide_index=True)

                # ロジック検証用の詳細値は普段は隠し、必要なときだけ確認できるようにする。
                with st.expander("予測ロジック詳細（分析用）"):
                    st.dataframe(
                        bt_export,
                        width="stretch",
                        hide_index=True,
                    )

                inactive_bt = bt_result.get("inactive_media", pd.DataFrame())
                if not inactive_bt.empty:
                    st.caption(f"休眠判定で定常予測から除外: {len(inactive_bt):,}媒体")

                with st.expander("バックテスト日別明細"):
                    bt_daily = bt_result["daily_compare"].copy()
                    bt_daily["date"] = pd.to_datetime(bt_daily["date"]).dt.strftime("%Y/%m/%d")
                    bt_daily = bt_daily.rename(
                        columns={
                            "date": "日付",
                            "media": "媒体",
                            "forecast_cv": "予測CV",
                            "actual_cv": "実績CV",
                        }
                    )
                    st.dataframe(
                        bt_daily[["日付", "媒体", "予測CV", "実績CV", "差分"]].round(2),
                        width="stretch",
                        hide_index=True,
                    )
                    bt_daily_export = bt_daily[["日付", "媒体", "予測CV", "実績CV", "差分"]].round(2)
                    st.download_button(
                        "📥 日別明細をCSVで保存",
                        data=bt_daily_export.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"AFF定常予測_バックテスト日別_{bt_month_filename}.csv",
                        mime="text/csv",
                        key=f"download_backtest_daily_{bt_month_filename}",
                    )
            except Exception as bt_exc:
                st.warning(f"バックテストを実行できませんでした: {bt_exc}")


    # ---------------------------------------------------------
    # 予測 → 松竹梅 → 最適化 は一度計算したら session_state に保持。
    # Streamlitのボタン押下ではスクリプト全体が再実行されるが、
    # 入力条件が同じならここでは再計算しない。
    # ---------------------------------------------------------
    reference_key = tuple(
        (
            seg["label"],
            seg["cpn"],
            str(seg["start"]),
            str(seg["end"]),
            seg["reference_key"],
        )
        for seg in segment_bases
    )

    calc_key = (
        tuple(selected_media),
        tuple(selected_product_ids),
        bool(exclude_compensation),
        reference_key,
        opt_mode,
        len(base_pair),
        round(float(pd.to_numeric(base_pair["base_cv"], errors="coerce").fillna(0).sum()), 6),
        round(float(pd.to_numeric(base_pair["cost"], errors="coerce").fillna(0).sum()), 2),
        tuple((m, publication_status_map.get(m, "掲載予定"), round(float(publication_factor_map.get(m, 1.0)), 6)) for m in selected_media),
    )

    cached_calc = st.session_state.get("_planning_calc")

    if (
        cached_calc is not None
        and st.session_state.get("_planning_calc_key") == calc_key
    ):
        forecast_df = cached_calc["forecast_df"]
        sim_df = cached_calc["sim_df"]
        sim_summary = cached_calc["sim_summary"]
        opt_df = cached_calc["opt_df"]
        opt_summary = cached_calc["opt_summary"]

        sim_report_table = cached_calc.get("sim_report_table")
        opt_report_table = cached_calc.get("opt_report_table")

        if sim_report_table is None:
            sim_report_table = create_report_table(sim_summary)
            cached_calc["sim_report_table"] = sim_report_table

        if opt_report_table is None:
            opt_report_table = create_report_table(opt_summary)
            cached_calc["opt_report_table"] = opt_report_table

    else:
        future_parts = []
        for seg in segment_bases:
            future_dates = pd.date_range(start=seg["start"], end=seg["end"])
            part = pd.DataFrame({"date": future_dates}).merge(
                seg["base_pair"],
                how="cross",
            )
            part["weekday"] = part["date"].dt.day_name()
            part = add_business_edge_flags(part)

            # CPNマスタからは月度・特殊日フラグだけ取得。
            # 施策名そのものはUI選択を正とする。
            part = part.merge(
                cpn_master[
                    [
                        "日付",
                        "月度",
                        "line_oa_flag",
                        "magitoku_after_flag",
                    ]
                ],
                left_on="date",
                right_on="日付",
                how="left",
            )
            part["CPN名"] = seg["cpn"]
            part["planning_segment"] = seg["label"]
            part["月度"] = (
                part["月度"]
                .astype("string")
                .str.strip()
                .replace("", pd.NA)
                .fillna(pd.Series(part["date"].dt.strftime("%Y年%m月度"), index=part.index))
            )
            part["line_oa_flag"] = part["line_oa_flag"].fillna(0).astype(int)
            part["magitoku_after_flag"] = part["magitoku_after_flag"].fillna(0).astype(int)
            part["cpn_factor"] = 1.0
            future_parts.append(part)

        future_df = pd.concat(future_parts, ignore_index=True)
        forecast_df = forecast_cv(future_df, factor_tables)

        forecast_df = enforce_premium_media_cost(
            forecast_df
        )

        # 定常施策だけ掲載状態を反映。
        # 「未定」は過去継続率で期待値化し、掲載なしはCV/COSTとも0にする。
        forecast_df["掲載状態"] = forecast_df["media"].map(publication_status_map).fillna("掲載予定")
        forecast_df["掲載係数"] = forecast_df["media"].map(publication_factor_map).fillna(1.0)
        normal_pub_mask = forecast_df["CPN名"].astype(str).str.strip().isin(normal_labels)
        forecast_df["forecast_cv_before_publication"] = forecast_df["forecast_cv"]
        forecast_df.loc[normal_pub_mask, "forecast_cv"] = (
            pd.to_numeric(forecast_df.loc[normal_pub_mask, "forecast_cv"], errors="coerce").fillna(0.0)
            * pd.to_numeric(forecast_df.loc[normal_pub_mask, "掲載係数"], errors="coerce").fillna(1.0)
        )
        forecast_df.loc[normal_pub_mask, "cost"] = (
            pd.to_numeric(forecast_df.loc[normal_pub_mask, "cost"], errors="coerce").fillna(0.0)
            * pd.to_numeric(forecast_df.loc[normal_pub_mask, "掲載係数"], errors="coerce").fillna(1.0)
        )

        # 営業プランの計算定義へ合わせる。
        # Forecast(発生件数) × 承認率 = 発行見込み
        # 発行見込み × W列グロス単価 = 発行コスト
        _plan_period_metrics = _calculate_period_media_metrics(history_df)
        is_magi_plan = forecast_df["CPN名"].astype(str).str.strip().eq("マジ得")
        normal_rate_map = _plan_period_metrics["normal_rate"]
        magi_rate_map = _plan_period_metrics["magi_rate"]
        normal_unit_map = _plan_period_metrics["normal_unit"]
        magi_unit_map = _plan_period_metrics["magi_unit"]
        forecast_df["approval_rate"] = [
            (magi_rate_map.get(str(m), _plan_period_metrics["magi_rate_all"]) if magi
             else normal_rate_map.get(str(m), _plan_period_metrics["normal_rate_all"]))
            for m, magi in zip(forecast_df["media"], is_magi_plan)
        ]
        forecast_df["gross_unit"] = [
            (magi_unit_map.get(str(m), _plan_period_metrics["magi_unit_all"]) if magi
             else normal_unit_map.get(str(m), _plan_period_metrics["normal_unit_all"]))
            for m, magi in zip(forecast_df["media"], is_magi_plan)
        ]
        forecast_df["approval_rate"] = pd.to_numeric(forecast_df["approval_rate"], errors="coerce").fillna(0.0).clip(0, 1)
        forecast_df["gross_unit"] = pd.to_numeric(forecast_df["gross_unit"], errors="coerce").fillna(0.0)

        # simulate_plan用に日付表示形式を変換
        forecast_df = forecast_df.copy()
        forecast_df["date"] = format_date(
            forecast_df
        )

        sim_df = simulate_plan(
            forecast_df
        )

        # 営業ロジック列の後方互換・空データ対策。
        # approved_cv が無い旧形式/一部条件でも集計でKeyErrorにしない。
        sim_df = sim_df.copy()
        if "approved_cv" not in sim_df.columns:
            cv_s = pd.to_numeric(sim_df.get("cv", pd.Series(0.0, index=sim_df.index)), errors="coerce").fillna(0.0)
            rate_s = pd.to_numeric(sim_df.get("approval_rate", pd.Series(0.0, index=sim_df.index)), errors="coerce").fillna(0.0).clip(0, 1)
            sim_df["approved_cv"] = cv_s * rate_s
        if "cpa" not in sim_df.columns and "gross_unit" in sim_df.columns:
            sim_df["cpa"] = pd.to_numeric(sim_df["gross_unit"], errors="coerce").fillna(0.0)
        if "cost" not in sim_df.columns:
            unit_s = pd.to_numeric(sim_df.get("cpa", pd.Series(0.0, index=sim_df.index)), errors="coerce").fillna(0.0)
            sim_df["cost"] = pd.to_numeric(sim_df["approved_cv"], errors="coerce").fillna(0.0) * unit_s

        sim_group_cols = ["date", "media", "plan"]
        if "CPN名" in sim_df.columns:
            sim_group_cols.append("CPN名")
        sim_df = sim_df.copy()
        for _col in ["date", "media", "plan", "cv", "cost", "approved_cv", "cpa"]:
            if _col not in sim_df.columns:
                sim_df[_col] = pd.Series(dtype="object" if _col in {"date", "media", "plan"} else "float64")
        sim_df["_unit_weight"] = pd.to_numeric(sim_df["cpa"], errors="coerce").fillna(0) * pd.to_numeric(sim_df["cv"], errors="coerce").fillna(0)
        sim_summary = (
            sim_df.groupby(
                sim_group_cols,
                as_index=False,
            )
            .agg(
                cv=("cv", "sum"),
                cost=("cost", "sum"),
                approved_cv=("approved_cv", "sum"),
                _unit_weight=("_unit_weight", "sum"),
            )
        )
        sim_summary["cpa"] = (
            sim_summary["_unit_weight"]
            / sim_summary["cv"].replace(0, pd.NA)
        ).fillna(0)
        sim_summary = sim_summary.drop(columns=["_unit_weight"])

        sim_summary["date"] = format_date(
            sim_summary
        )

        opt_df = optimize_budget(
            sim_df,
            opt_mode,
        )

        opt_group_cols = ["date", "media", "plan"]
        if "CPN名" in opt_df.columns:
            opt_group_cols.append("CPN名")
        opt_df = opt_df.copy()
        for _col in ["date", "media", "plan", "cv", "cost", "approved_cv", "cpa"]:
            if _col not in opt_df.columns:
                opt_df[_col] = pd.Series(dtype="object" if _col in {"date", "media", "plan"} else "float64")
        opt_df["_unit_weight"] = pd.to_numeric(opt_df["cpa"], errors="coerce").fillna(0) * pd.to_numeric(opt_df["cv"], errors="coerce").fillna(0)
        opt_summary = (
            opt_df.groupby(
                opt_group_cols,
                as_index=False,
            )
            .agg(
                cv=("cv", "sum"),
                cost=("cost", "sum"),
                approved_cv=("approved_cv", "sum"),
                _unit_weight=("_unit_weight", "sum"),
            )
        )
        opt_summary["cpa"] = (
            opt_summary["_unit_weight"]
            / opt_summary["cv"].replace(0, pd.NA)
        ).fillna(0)
        opt_summary = opt_summary.drop(columns=["_unit_weight"])

        opt_summary["date"] = format_date(
            opt_summary
        )

        st.session_state["_planning_calc_key"] = calc_key
        # 表示用テーブルもここで1回だけ作って保存する。
        # Excelボタン操作や他ウィジェット操作で同じ表を作り直さない。
        sim_report_table = create_report_table(sim_summary)
        opt_report_table = create_report_table(opt_summary)

        st.session_state["_planning_calc"] = {
            "forecast_df": forecast_df,
            "sim_df": sim_df,
            "sim_summary": sim_summary,
            "opt_df": opt_df,
            "opt_summary": opt_summary,
            "sim_report_table": sim_report_table,
            "opt_report_table": opt_report_table,
        }

        # 条件が変わって再計算した場合、古い提出Excelは破棄
        st.session_state.pop("submission_excel", None)

    # 掲載状態を反映した定常予測の即時サマリー。
    # 掲載状態を変更するとcalc_keyが変わるため、Streamlitの再実行でここも即座に更新される。
    if needs_normal_learning and not forecast_df.empty:
        normal_pub_summary_mask = forecast_df["CPN名"].astype(str).str.strip().isin(normal_labels)
        normal_pub_rows = forecast_df.loc[normal_pub_summary_mask].copy()
        if not normal_pub_rows.empty:
            normal_pub_rows["forecast_cv_before_publication"] = pd.to_numeric(
                normal_pub_rows.get("forecast_cv_before_publication", normal_pub_rows["forecast_cv"]),
                errors="coerce",
            ).fillna(0.0)
            normal_pub_rows["forecast_cv"] = pd.to_numeric(normal_pub_rows["forecast_cv"], errors="coerce").fillna(0.0)
            pub_total_before = float(normal_pub_rows["forecast_cv_before_publication"].sum())
            pub_total_after = float(normal_pub_rows["forecast_cv"].sum())
            pub_impact = pub_total_after - pub_total_before

            st.subheader("📌 掲載状態反映後の定常予測")
            sm1, sm2, sm3 = st.columns(3)
            sm1.metric("掲載判定前CV", f"{pub_total_before:,.0f}")
            sm2.metric("掲載状態反映後CV", f"{pub_total_after:,.0f}")
            sm3.metric("掲載状態による増減", f"{pub_impact:+,.0f}")

            pub_media = (
                normal_pub_rows.groupby("media", as_index=False)
                .agg(
                    掲載判定前CV=("forecast_cv_before_publication", "sum"),
                    予測CV=("forecast_cv", "sum"),
                )
            )
            pub_media["掲載状態"] = pub_media["media"].map(publication_status_map).fillna("掲載予定")
            pub_media = pub_media.rename(columns={"media": "媒体"})
            pub_media["予測CV"] = pub_media["予測CV"].round(0).astype(int)
            pub_media["掲載判定前CV"] = pub_media["掲載判定前CV"].round(0).astype(int)
            pub_media = pub_media.sort_values(["予測CV", "掲載判定前CV"], ascending=False)
            st.dataframe(
                pub_media[["媒体", "掲載状態", "予測CV"]],
                width="stretch",
                hide_index=True,
            )
            with st.expander("掲載状態による補正詳細（分析用）"):
                detail = pub_media.copy()
                detail["増減CV"] = detail["予測CV"] - detail["掲載判定前CV"]
                st.dataframe(
                    detail[["媒体", "掲載状態", "掲載判定前CV", "予測CV", "増減CV"]],
                    width="stretch",
                    hide_index=True,
                )

    # 係数確認用の明細
    with st.expander("予測係数の確認"):
        factor_cols = [
            "date",
            "media",
            "商品ID",
            "base_cv",
            "cpn_factor",
            "unit_price_factor",
            "unit_price_reference",
            "unit_price_latest",
            "unit_price_elasticity",
            "weekday_factor",
            "season_factor",
            "month_edge_factor",
            "after_factor",
            "line_factor",
            "forecast_cv",
            "cost",
        ]

        available_factor_cols = [
            c for c in factor_cols
            if c in forecast_df.columns
        ]

        st.dataframe(
            forecast_df[available_factor_cols],
            width="stretch",
            hide_index=True,
        )

    st.subheader("📊 松竹梅")
    st.dataframe(
        sim_report_table,
        width="stretch",
    )

    st.subheader("🚀 最適プラン")
    st.dataframe(
        opt_report_table,
        width="stretch",
    )

    st.subheader("✍️ 手動設定")
    st.caption(
        "初期値は上の最適プランと過去実績から自動設定。"
        "グロス単価・承認率・採用件数を編集すると、費用は自動計算されます。"
        "承認件数と発行CPAは入力内容から自動計算します。"
    )

    render_manual_settings(
        opt_summary=opt_summary,
        history_df=history_df,
        selected_cpn=selected_cpn,
        calc_key=calc_key,
    )

    # 過去実績分析には、手動設定後の今回プランを「今回」として単価帯グラフへ追加する。
    current_manual_for_chart = _normalize_manual_settings(
        st.session_state.get(
            "_manual_settings",
            _build_manual_settings_defaults(
                opt_summary=opt_summary,
                history_df=history_df,
                selected_cpn=selected_cpn,
            ),
        )
    )

    # 「マジ得期間のみ」では、今回プランもマジ得指定期間分だけを表示する。
    # 手動設定値は最終版として維持し、媒体ごとの採用件数だけ
    # 今回予測に占めるマジ得CV比率で按分する。
    current_manual_magitoku_for_chart = pd.DataFrame()
    if (
        not current_manual_for_chart.empty
        and "CPN名" in opt_summary.columns
        and (opt_summary["CPN名"].astype(str).str.strip() == "マジ得").any()
    ):
        mix = opt_summary.copy()
        mix["media"] = mix["media"].astype(str)
        mix["cv"] = pd.to_numeric(mix["cv"], errors="coerce").fillna(0.0)
        mix["CPN名"] = mix["CPN名"].astype(str).str.strip()
        total_cv = mix.groupby("media")["cv"].sum()
        magi_cv = mix.loc[mix["CPN名"].eq("マジ得")].groupby("media")["cv"].sum()
        magi_ratio = (magi_cv / total_cv.replace(0, pd.NA)).fillna(0.0).clip(0, 1)

        current_manual_magitoku_for_chart = current_manual_for_chart.copy()
        current_manual_magitoku_for_chart["_magi_ratio"] = (
            current_manual_magitoku_for_chart["媒体名"].astype(str).map(magi_ratio).fillna(0.0)
        )
        current_manual_magitoku_for_chart["今回採用件数"] = (
            pd.to_numeric(
                current_manual_magitoku_for_chart["今回採用件数"], errors="coerce"
            ).fillna(0.0)
            * current_manual_magitoku_for_chart["_magi_ratio"]
        )
        current_manual_magitoku_for_chart = current_manual_magitoku_for_chart.drop(
            columns=["_magi_ratio"]
        )
        current_manual_magitoku_for_chart = _normalize_manual_settings(
            current_manual_magitoku_for_chart
        )
        current_manual_magitoku_for_chart = current_manual_magitoku_for_chart.loc[
            current_manual_magitoku_for_chart["今回採用件数"] > 0
        ].copy()

    try:
        with history_analytics_placeholder.container():
            render_history_analytics(
                history_df,
                current_plan_df=current_manual_for_chart,
                current_plan_magitoku_df=current_manual_magitoku_for_chart,
            )
    except ValueError as analytics_exc:
        with history_analytics_placeholder.container():
            st.warning(f"月度別実績を集計できません。{analytics_exc}")

    submission_month_label, submission_year, submission_month = _resolve_submission_cpn_month(
        cpn_master, start_date, end_date
    )
    submission_filename = (
        f"【提出用】楽天カード"
        f"{submission_year}年{submission_month}月度"
        f"プランニング.xlsx"
    )

    # ---------------------------------------------------------
    # 提出用Excel操作は独立Fragment。
    # ボタンを押しても予測・松竹梅・最適化は再実行しない。
    # Excel自体もクリックされるまで生成しない。
    # ---------------------------------------------------------
    manual_settings_for_export = _normalize_manual_settings(
        st.session_state.get(
            "_manual_settings",
            _build_manual_settings_defaults(
                opt_summary=opt_summary,
                history_df=history_df,
                selected_cpn=selected_cpn,
            ),
        )
    )

    render_submission_download(
        opt_summary=opt_summary,
        history_df=history_df,
        cpn_master=cpn_master,
        manual_settings=manual_settings_for_export,
        start_date=start_date,
        end_date=end_date,
        selected_cpn=selected_cpn,
        opt_mode=opt_mode,
        submission_filename=submission_filename,
    )
