import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

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
        c for c in result.columns if c not in ["media","plan","metric"]
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
st.caption("実績CSVと最新のCPNマスタをアップロードしてください。ファイルはGitHubには保存されません。")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("① 実績CSV", type=["csv"])
with col2:
    uploaded_master = st.file_uploader("② CPNマスタ", type=["xlsx", "xlsm"])

if uploaded_file and uploaded_master:
    try:
        from data.loader import add_business_edge_flags
        from logic.factors import (
            calculate_dynamic_factor_tables,
            calculate_transition_cpn_factor,
            enforce_premium_media_cost,
        )
        from config.constants import RECENT_NORMAL_DAYS

        history_df = load_data(uploaded_file)
        history_df["date"] = pd.to_datetime(history_df["date"])

        cpn_master = pd.read_excel(uploaded_master, engine="openpyxl")
        required_master_columns = {"日付", "CPN名"}
        missing_master = required_master_columns - set(cpn_master.columns)
        if missing_master:
            raise ValueError(f"CPNマスタに必要な列がありません: {', '.join(sorted(missing_master))}")

        cpn_master = cpn_master.copy()
        cpn_master["日付"] = pd.to_datetime(cpn_master["日付"], errors="coerce")
        cpn_master["CPN名"] = cpn_master["CPN名"].astype("string").str.strip()
        cpn_master = cpn_master.dropna(subset=["日付", "CPN名"])
        cpn_master = cpn_master.drop_duplicates(subset=["日付"], keep="last")
        if cpn_master.empty:
            raise ValueError("CPNマスタに有効な日付・CPN名がありません。")

        # 任意列。未登録なら補正なし。
        cpn_master["line_oa_flag"] = _truthy(cpn_master["LINE OA配信"]) if "LINE OA配信" in cpn_master else 0
        cpn_master["magitoku_after_flag"] = _truthy(cpn_master["マジ得後"]) if "マジ得後" in cpn_master else 0
    except Exception as exc:
        st.error(f"ファイルの読み込みに失敗しました: {exc}")
        st.stop()

    history_df = history_df.merge(
        cpn_master[["日付", "CPN名", "line_oa_flag", "magitoku_after_flag"]],
        left_on="date", right_on="日付", how="left"
    )
    history_df["CPN名"] = history_df["CPN名"].fillna("通常")
    history_df["line_oa_flag"] = history_df["line_oa_flag"].fillna(0).astype(int)
    history_df["magitoku_after_flag"] = history_df["magitoku_after_flag"].fillna(0).astype(int)

    st.sidebar.header("媒体選択")
    all_media = sorted(history_df["media"].unique())
    default_media = [m for m in all_media if "計測" not in m]
    selected_media = st.sidebar.multiselect("媒体", all_media, default=default_media)
    if not selected_media:
        st.stop()
    history_df = history_df[history_df["media"].isin(selected_media)].copy()

    st.sidebar.header("📊 CPN選択")
    cpn_list = sorted(cpn_master["CPN名"].dropna().astype(str).unique())
    selected_cpn = st.sidebar.selectbox("CPN", cpn_list)

    cpn_period = cpn_master[cpn_master["CPN名"] == selected_cpn]
    if not cpn_period.empty:
        st.sidebar.write(f"マスタ期間: {cpn_period['日付'].min().date()} ～ {cpn_period['日付'].max().date()}")

    today = datetime.date.today()
    start_date = st.sidebar.date_input("予測開始", today)
    end_date = st.sidebar.date_input("予測終了", today + datetime.timedelta(days=7))
    if start_date > end_date:
        st.error("予測期間の開始日は終了日以前にしてください。")
        st.stop()

    reference_date = pd.Timestamp(start_date)

    # 直近の定常日を最大28日分取得（欠損日を含めた単純28暦日ではない）
    recent_normal_dates = (
        history_df.loc[
            (history_df["CPN名"] == "通常") & (history_df["date"] < reference_date),
            "date",
        ]
        .drop_duplicates()
        .sort_values()
        .tail(RECENT_NORMAL_DAYS)
    )
    recent_normal = history_df[history_df["date"].isin(recent_normal_dates)].copy()
    if recent_normal.empty:
        st.error("予測開始日より前の直近定常データがありません。CPNマスタの『通常』登録を確認してください。")
        st.stop()

    # 実績に存在した媒体×商品IDだけを採用
    base_pair = _daily_pair_average(recent_normal)
    base_pair = base_pair[base_pair["media"].isin(selected_media)]
    if base_pair.empty:
        st.error("対象媒体の定常実績がありません。")
        st.stop()

    # 昨年同一CPNの「直前定常→CPN」切替係数
    cpn_factor_table = calculate_transition_cpn_factor(history_df, selected_cpn, reference_date)
    cpn_factor_map = cpn_factor_table.set_index("media")["cpn_factor"] if not cpn_factor_table.empty else pd.Series(dtype=float)

    st.subheader("📈 CPN実績係数")
    if cpn_factor_table.empty:
        st.warning("昨年同一CPN、またはその直前の定常実績が不足しているため、CPN実績係数は1.0で処理します。")
    else:
        display_factor = cpn_factor_table.copy()
        display_factor["normal_daily_cv"] = display_factor["normal_daily_cv"].round(2)
        display_factor["cpn_daily_cv"] = display_factor["cpn_daily_cv"].round(2)
        display_factor["cpn_factor"] = display_factor["cpn_factor"].round(3)
        display_factor = display_factor.rename(columns={
            "media": "媒体",
            "normal_daily_cv": "前年CPN直前の定常CV/日",
            "cpn_daily_cv": "前年同一CPNのCV/日",
            "cpn_factor": "CPN実績係数",
            "reference_cpn_start": "参照CPN開始",
            "reference_cpn_end": "参照CPN終了",
        })
        st.dataframe(display_factor, use_container_width=True, hide_index=True)

    factor_tables = calculate_dynamic_factor_tables(history_df)
    st.subheader("📐 実績から算出した変動係数")
    tab1, tab2, tab3, tab4 = st.tabs(["曜日", "月初・月末", "需要期", "LINE OA"])
    with tab1:
        st.dataframe(factor_tables["weekday"].round(3), use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(factor_tables["month_edge"].round(3), use_container_width=True, hide_index=True)
    with tab3:
        st.dataframe(factor_tables["season"].round(3), use_container_width=True, hide_index=True)
    with tab4:
        if factor_tables["line_oa"].empty:
            st.info("過去のLINE OA配信実績がないため、LINE OA係数は1.0です。")
        else:
            st.dataframe(factor_tables["line_oa"].round(3), use_container_width=True, hide_index=True)

    future_dates = pd.date_range(start=start_date, end=end_date)
    future_df = pd.DataFrame({"date": future_dates}).merge(base_pair, how="cross")
    future_df["weekday"] = future_df["date"].dt.day_name()
    future_df = add_business_edge_flags(future_df)

    future_df = future_df.merge(
        cpn_master[["日付", "CPN名", "line_oa_flag", "magitoku_after_flag"]],
        left_on="date", right_on="日付", how="left"
    )
    future_df["CPN名"] = future_df["CPN名"].fillna(selected_cpn)
    future_df["line_oa_flag"] = future_df["line_oa_flag"].fillna(0).astype(int)
    future_df["magitoku_after_flag"] = future_df["magitoku_after_flag"].fillna(0).astype(int)

    future_df["cpn_factor"] = future_df["media"].map(cpn_factor_map).fillna(1.0)

    # 曜日・月初月末・月別需要期・LINE OAは、アップロード実績から毎回算出。
    forecast_df = forecast_cv(future_df, factor_tables)
    forecast_df = enforce_premium_media_cost(forecast_df)

    # 係数確認用の明細
    with st.expander("予測係数の確認"):
        factor_cols = [
            "date", "media", "商品ID", "base_cv", "cpn_factor",
            "weekday_factor", "season_factor", "month_edge_factor", "after_factor",
            "line_factor", "forecast_cv", "cost",
        ]
        st.dataframe(forecast_df[factor_cols], use_container_width=True, hide_index=True)

    forecast_df["date"] = format_date(forecast_df)
    sim_df = simulate_plan(forecast_df)

    sim_summary = (
        sim_df.groupby(["date", "media", "plan"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    sim_summary["cpa"] = (sim_summary["cost"] / sim_summary["cv"]).replace([float("inf"), float("-inf")], 0).fillna(0)
    sim_summary["date"] = format_date(sim_summary)

    st.subheader("📊 松竹梅")
    st.dataframe(create_report_table(sim_summary), use_container_width=True)

    st.sidebar.header("🎯 最適化ロジック")
    opt_mode = st.sidebar.radio("最適基準", ["CPA最小", "CV最大"], index=0)
    opt_df = optimize_budget(sim_df, opt_mode)
    opt_summary = (
        opt_df.groupby(["date", "media", "plan"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    opt_summary["cpa"] = (opt_summary["cost"] / opt_summary["cv"]).replace([float("inf"), float("-inf")], 0).fillna(0)
    opt_summary["date"] = format_date(opt_summary)

    st.subheader("🚀 最適プラン")
    st.dataframe(create_report_table(opt_summary), use_container_width=True)

    target_cv = st.sidebar.number_input("目標", min_value=0, value=1000)
    gap = target_cv - forecast_df["forecast_cv"].sum()
    st.write(f"差分: {gap:.0f}")

    st.download_button(
        "Excel DL",
        data=to_excel_multi(sim_summary, opt_summary),
        file_name="af_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
