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
# ✅ UI
# -----------------------
st.set_page_config(page_title="AFプランニングツール", layout="wide")
st.title("📊 件数予測＆プランニングツール")

st.caption("実績CSVと最新のCPNマスタをアップロードしてください。アップロードしたファイルはGitHubには保存されません。")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("① 実績CSV", type=["csv"])
with col2:
    uploaded_master = st.file_uploader("② CPNマスタ", type=["xlsx", "xlsm"])

if uploaded_file and uploaded_master:
    try:
        history_df = load_data(uploaded_file)
        history_df["date"] = pd.to_datetime(history_df["date"])

        cpn_master = pd.read_excel(uploaded_master, engine="openpyxl")
        required_master_columns = {"日付", "CPN名"}
        missing_master = required_master_columns - set(cpn_master.columns)
        if missing_master:
            raise ValueError(
                f"CPNマスタに必要な列がありません: {', '.join(sorted(missing_master))}"
            )

        cpn_master = cpn_master.copy()
        cpn_master["日付"] = pd.to_datetime(cpn_master["日付"], errors="coerce")
        cpn_master["CPN名"] = cpn_master["CPN名"].astype("string").str.strip()
        cpn_master = cpn_master.dropna(subset=["日付", "CPN名"])
        cpn_master = cpn_master.drop_duplicates(subset=["日付"], keep="last")

        if cpn_master.empty:
            raise ValueError("CPNマスタに有効な日付・CPN名がありません。")
    except Exception as exc:
        st.error(f"ファイルの読み込みに失敗しました: {exc}")
        st.stop()

    history_df = history_df.merge(
        cpn_master,
        left_on="date",
        right_on="日付",
        how="left"
    )

    history_df["CPN名"] = history_df["CPN名"].fillna("通常")

    st.sidebar.header("媒体選択")

    all_media = sorted(history_df["media"].unique())
    default_media = [m for m in all_media if "計測" not in m]

    selected_media = st.sidebar.multiselect(
        "媒体",
        all_media,
        default=default_media
    )

    if not selected_media:
        st.stop()

    history_df = history_df[history_df["media"].isin(selected_media)]

    st.sidebar.header("📊 CPN選択")

    cpn_list = sorted(cpn_master["CPN名"].dropna().astype(str).unique())
    if not cpn_list:
        st.error("CPNマスタに選択可能なCPN名がありません。")
        st.stop()
    selected_cpn = st.sidebar.selectbox("CPN", cpn_list)

    cpn_period = cpn_master[cpn_master["CPN名"] == selected_cpn]

    cpn_start = cpn_period["日付"].min()
    cpn_end = cpn_period["日付"].max()

    st.sidebar.write(f"CPN期間: {cpn_start.date()} ~ {cpn_end.date()}")

    today = datetime.date.today()

    start_date = st.sidebar.date_input("開始", today)
    end_date = st.sidebar.date_input("終了", today + datetime.timedelta(days=7))

    if start_date > end_date:
        st.error("予測期間の開始日は終了日以前にしてください。")
        st.stop()

    st.sidebar.header("📊 学習方法")

    learning_mode = st.sidebar.radio(
        "学習方法",
        ["同CPN平均", "昨年同期間", "ハイブリッド", "手動指定"]
    )

    # -----------------------
    # ✅ 学習データ作成（修正済）
    # -----------------------
    if learning_mode == "同CPN平均":

        hist = history_df[history_df["CPN名"] == selected_cpn]

        if hist.empty:
            st.error("同CPNデータなし")
            st.stop()

        days = hist["date"].nunique()

        daily_cv = hist.groupby("media")["cv"].sum() / days
        daily_cost = hist.groupby("media")["cost"].sum() / days


    elif learning_mode == "昨年同期間":

        hist_start = pd.to_datetime(start_date) - pd.Timedelta(days=365)
        hist_end = pd.to_datetime(end_date) - pd.Timedelta(days=365)

        hist = history_df[
            (history_df["date"] >= hist_start) &
            (history_df["date"] <= hist_end)
        ]

        if hist.empty:
            st.error("昨年同期間データなし")
            st.stop()

        daily_cv = hist.groupby("media")["cv"].mean()
        daily_cost = hist.groupby("media")["cost"].mean()


    elif learning_mode == "ハイブリッド":

        cpn_hist = history_df[history_df["CPN名"] == selected_cpn]

        if cpn_hist.empty:
            st.error("同CPNデータなし")
            st.stop()

        cpn_days = cpn_hist["date"].nunique()

        cpn_daily_cv = cpn_hist.groupby("media")["cv"].sum() / cpn_days
        cpn_daily_cost = cpn_hist.groupby("media")["cost"].sum() / cpn_days

        hist_start = pd.to_datetime(start_date) - pd.Timedelta(days=365)
        hist_end = pd.to_datetime(end_date) - pd.Timedelta(days=365)

        last_year = history_df[
            (history_df["date"] >= hist_start) &
            (history_df["date"] <= hist_end)
        ]

        if last_year.empty:
            st.error("昨年データなし")
            st.stop()

        last_year_days = last_year["date"].nunique()
        last_year_cv = last_year.groupby("media")["cv"].sum() / last_year_days

        normal_df = history_df[history_df["CPN名"] != selected_cpn]

        normal_days = normal_df["date"].nunique()
        normal_cv = normal_df.groupby("media")["cv"].sum() / normal_days

        season_factor = (last_year_cv / normal_cv).replace([float("inf")], 1)

        daily_cv = cpn_daily_cv * season_factor
        daily_cost = cpn_daily_cost

        hist = cpn_hist


    else:

        st.sidebar.subheader("✏️ 学習期間指定")

        manual_start = st.sidebar.date_input("学習開始日")
        manual_end = st.sidebar.date_input("学習終了日")

        hist = history_df[
            (history_df["date"] >= pd.to_datetime(manual_start)) &
            (history_df["date"] <= pd.to_datetime(manual_end))
        ]

        if hist.empty:
            st.error("指定期間にデータなし")
            st.stop()

        daily_cv = hist.groupby("media")["cv"].mean()
        daily_cost = hist.groupby("media")["cost"].mean()

    # -----------------------
    # ✅ CPN係数
    # -----------------------
    normal_df = history_df[history_df["CPN名"] != selected_cpn]
    cpn_df = history_df[history_df["CPN名"] == selected_cpn]

    normal_days = normal_df["date"].nunique()
    cpn_days = cpn_df["date"].nunique()

    if normal_days == 0 or cpn_days == 0:
        cpn_factor = pd.Series(dtype="float64")
    else:
        normal_avg = normal_df.groupby("media")["cv"].sum() / normal_days
        cpn_avg = cpn_df.groupby("media")["cv"].sum() / cpn_days
        cpn_factor = (cpn_avg / normal_avg).replace([float("inf"), float("-inf")], 1).fillna(1)

    # -----------------------
    # ✅ future生成
    # -----------------------
    future_dates = pd.date_range(start=start_date, end=end_date)

    future_df = pd.DataFrame({"date": future_dates})
    future_df["weekday"] = future_df["date"].dt.day_name()

    future_df = future_df.merge(
        pd.DataFrame({"media": hist["media"].unique()}),
        how="cross"
    )

    future_df = future_df.merge(
        pd.DataFrame({"商品ID": hist["商品ID"].unique()}),
        how="cross"
    )

    future_df = future_df.merge(
        cpn_master,
        left_on="date",
        right_on="日付",
        how="left"
    )

    future_df["CPN名"] = future_df["CPN名"].fillna("通常")

    future_df["day"] = future_df["date"].dt.day
    future_df["is_month_start"] = future_df["day"].isin([1,2,3,4]).astype(int)
    future_df["is_month_end"] = future_df["day"].isin([28,29,30,31]).astype(int)
    future_df["magitoku_flag"] = 0
    future_df["line_oa_flag"] = 0


    # -----------------------
    # ✅ base + CPN適用（ここ超重要）
    # -----------------------
    future_df["base_cv"] = future_df["media"].map(daily_cv)
    future_df["cost"] = future_df["media"].map(daily_cost)
    future_df["cpn_factor"] = future_df["media"].map(cpn_factor).fillna(1)

    # -----------------------
    # ✅ 曜日係数
    # -----------------------
    weekday_factor = {
        "Monday":1.0,
        "Tuesday":1.05,
        "Wednesday":1.1,
        "Thursday":1.05,
        "Friday":1.2,
        "Saturday":1.25,
        "Sunday":1.2
    }
    future_df["weekday_factor"] = future_df["weekday"].map(weekday_factor).fillna(1)

    # -----------------------
    # ✅ 需要期係数（月）
    # -----------------------
    future_df["month"] = future_df["date"].dt.month

    season_factor = {
        1:0.9, 2:0.95, 3:1.2, 4:1.3, 5:1.1, 6:1.0,
        7:0.95, 8:0.9, 9:0.95, 10:1.0, 11:1.05, 12:1.1
    }
    future_df["season_factor"] = future_df["month"].map(season_factor).fillna(1)

    # -----------------------
    # ✅ 月初・月末
    # -----------------------
    future_df["month_factor"] = 1.0
    future_df.loc[future_df["is_month_start"]==1, "month_factor"] = 1.1
    future_df.loc[future_df["is_month_end"]==1, "month_factor"] = 1.15

    # -----------------------
    # ✅ CPN種類係数
    # -----------------------
    if "マジ得" in selected_cpn:
        future_df["cpn_type_factor"] = 1.4
    elif "JCB" in selected_cpn:
        future_df["cpn_type_factor"] = 1.2
    elif "みずほ" in selected_cpn:
        future_df["cpn_type_factor"] = 1.1
    else:
        future_df["cpn_type_factor"] = 1.0
    
    # -----------------------
    # ✅ マジ得後補正
    # -----------------------
    future_df["after_factor"] = 1.0
    
    if "マジ得" in selected_cpn:
        future_df.loc[
            future_df["date"] > pd.to_datetime(end_date) - pd.Timedelta(days=2),
            "after_factor"
        ] = 0.9

    # -----------------------
    # ✅ LINE OA（仮：金曜）
    # -----------------------
    future_df["line_factor"] = 1.0
    future_df.loc[future_df["weekday"]=="Friday", "line_factor"] = 1.2
    

    # -----------------------
    # ✅ 予測（ここ追加）
    # -----------------------
    forecast_df = forecast_cv(future_df, hist)
    
    # ✅ 係数すべて適用
    forecast_df["forecast_cv"] = (
        forecast_df["forecast_cv"]
        * future_df["season_factor"]
        * future_df["cpn_type_factor"]
        * future_df["after_factor"]
        * future_df["line_factor"]
    )

    forecast_df["date"] = format_date(forecast_df)

    sim_df = simulate_plan(forecast_df)

    sim_summary = (
        sim_df.groupby(["date","media","plan"])
        .agg({"cv":"sum","cost":"sum"})
        .reset_index()
    )

    sim_summary["cpa"] = sim_summary["cost"] / sim_summary["cv"]
    sim_summary["cpa"] = sim_summary["cpa"].replace([float("inf"), float("-inf")], 0).fillna(0)
    sim_summary["date"] = format_date(sim_summary)

    st.subheader("📊 松竹梅")
    st.dataframe(create_report_table(sim_summary), use_container_width=True)

    st.sidebar.header("🎯 最適化ロジック")

    opt_mode = st.sidebar.radio(
        "最適基準",
        ["CPA最小", "CV最大"],
        index=0
    )

    opt_df = optimize_budget(sim_df, opt_mode)

    opt_summary = (
        opt_df.groupby(["date","media","plan"])
        .agg({"cv":"sum","cost":"sum"})
        .reset_index()
    )

    opt_summary["cpa"] = opt_summary["cost"] / opt_summary["cv"]
    opt_summary["cpa"] = opt_summary["cpa"].replace([float("inf"), float("-inf")], 0).fillna(0)
    opt_summary["date"] = format_date(opt_summary)

    st.subheader("🚀 最適プラン")
    st.dataframe(create_report_table(opt_summary), use_container_width=True)

    target_cv = st.sidebar.number_input("目標", 1000)

    gap = target_cv - forecast_df["forecast_cv"].sum()
    st.write(f"差分: {gap:.0f}")

    st.download_button(
        "Excel DL",
        data=to_excel_multi(sim_summary, opt_summary),
        file_name="af_plan.xlsx"
    )