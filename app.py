import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

from openpyxl import Workbook
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
# ✅ 提出用Excel（最適プランのみ）
# -----------------------
def create_submission_excel(opt_summary, history_df, start_date, end_date, selected_cpn, opt_mode):
    """添付の提出用プランニング表を意識した帳票を、最適プランだけで生成する。"""
    output = BytesIO()

    plan = opt_summary.copy()
    plan["date"] = pd.to_datetime(plan["date"], errors="coerce")
    plan = plan.dropna(subset=["date", "media"])
    plan["cv"] = pd.to_numeric(plan["cv"], errors="coerce").fillna(0)
    plan["cost"] = pd.to_numeric(plan["cost"], errors="coerce").fillna(0)
    plan["cpa"] = (plan["cost"] / plan["cv"]).replace([float("inf"), float("-inf")], 0).fillna(0)

    dates = list(pd.date_range(start=pd.Timestamp(start_date), end=pd.Timestamp(end_date), freq="D"))
    media_list = list(dict.fromkeys(plan["media"].astype(str).tolist()))

    # 商品IDは提出表のSID欄へ表示。1媒体に複数ある場合は「 / 」で併記。
    sid_map = {}
    if "商品ID" in history_df.columns:
        sid_source = history_df[["media", "商品ID"]].dropna().copy()
        sid_source["media"] = sid_source["media"].astype(str)
        sid_source["商品ID"] = sid_source["商品ID"].astype(str)
        sid_map = (
            sid_source.groupby("media")["商品ID"]
            .agg(lambda x: " / ".join(dict.fromkeys(x.tolist())))
            .to_dict()
        )

    daily = (
        plan.groupby(["date", "media"], as_index=False)
        .agg(cv=("cv", "sum"), cost=("cost", "sum"))
    )
    daily["cpa"] = (daily["cost"] / daily["cv"]).replace([float("inf"), float("-inf")], 0).fillna(0)

    # 値引き用辞書
    cv_map = {(r.date.normalize(), str(r.media)): float(r.cv) for r in daily.itertuples()}
    cpa_map = {(r.date.normalize(), str(r.media)): float(r.cpa) for r in daily.itertuples()}
    total_by_date = daily.groupby("date", as_index=True)["cv"].sum().to_dict()

    wb_out = Workbook()
    ws = wb_out.active
    month_label = pd.Timestamp(start_date).strftime("%-m月") if hasattr(pd.Timestamp(start_date), "strftime") else "プラン"
    # Windows互換を考慮して %-m が使えない環境にも対応
    month_label = f"{pd.Timestamp(start_date).month}月"
    ws.title = f"{month_label}（最適プラン）"

    # 色（添付テンプレに近い淡色）
    pale_blue = "DDEBF7"
    pale_green = "E2F0D9"
    pale_green_2 = "F3F8EF"
    weekend_pink = "F4CCCC"
    total_yellow = "FFF2CC"
    white = "FFFFFF"
    grid = "B7B7B7"
    red = "FF3333"
    green = "00A651"
    blue = "4472C4"
    dark = "404040"

    thin = Side(style="thin", color=grid)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    fixed_cols = 4
    first_date_col = fixed_cols + 1
    total_col = first_date_col + len(dates)

    # 列幅
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 24
    for col in range(first_date_col, total_col):
        ws.column_dimensions[get_column_letter(col)].width = 8.5
    ws.column_dimensions[get_column_letter(total_col)].width = 12

    # タイトル・条件
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_col)
    ws.cell(1, 1, f"楽天カード {pd.Timestamp(start_date).year}年{pd.Timestamp(start_date).month}月 プランニング（最適プラン）")
    ws.cell(1, 1).font = Font(size=14, bold=True, color=dark)
    ws.cell(1, 1).alignment = left
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_col)
    ws.cell(2, 1, f"CPN：{selected_cpn}　／　最適基準：{opt_mode}　／　予測期間：{pd.Timestamp(start_date).strftime('%Y/%m/%d')}～{pd.Timestamp(end_date).strftime('%Y/%m/%d')}")
    ws.cell(2, 1).font = Font(size=9, color="666666")

    # 上部サマリ
    summary_start = 4
    ws.cell(summary_start, 4, "")
    for j, dt in enumerate(dates, start=first_date_col):
        c = ws.cell(summary_start, j, dt.strftime("%m/%d"))
        c.alignment = center
        c.font = Font(size=8, bold=True)
        c.fill = PatternFill("solid", fgColor=weekend_pink if dt.weekday() >= 5 else white)
        c.border = border
    tc = ws.cell(summary_start, total_col, "Total")
    tc.fill = PatternFill("solid", fgColor=total_yellow)
    tc.font = Font(size=8, bold=True)
    tc.alignment = center
    tc.border = border

    summary_rows = [("Target", blue), ("Actual", dark), ("GAP", red)]
    for idx, (label, font_color) in enumerate(summary_rows, start=summary_start + 1):
        c = ws.cell(idx, 4, label)
        c.fill = PatternFill("solid", fgColor=pale_blue)
        c.font = Font(size=8, bold=(label == "Target"), color=font_color)
        c.alignment = center
        c.border = border
        for j, dt in enumerate(dates, start=first_date_col):
            cell = ws.cell(idx, j)
            cell.border = border
            cell.alignment = center
            if dt.weekday() >= 5:
                cell.fill = PatternFill("solid", fgColor=weekend_pink)
            if label == "Target":
                cell.value = round(float(total_by_date.get(dt.normalize(), 0)))
                cell.font = Font(size=8, color=blue)
            elif label == "Actual":
                cell.value = None
            else:
                target_ref = f"{get_column_letter(j)}{summary_start + 1}"
                cell.value = f"=-{target_ref}"
                cell.font = Font(size=8, color=red)
                cell.number_format = '#,##0;[Red](#,##0)'
        total_cell = ws.cell(idx, total_col)
        total_cell.fill = PatternFill("solid", fgColor=total_yellow)
        total_cell.border = border
        total_cell.alignment = center
        if label == "Target":
            total_cell.value = f"=SUM({get_column_letter(first_date_col)}{idx}:{get_column_letter(total_col-1)}{idx})"
            total_cell.font = Font(size=8, color=blue, bold=True)
        elif label == "Actual":
            total_cell.value = "=0"
        else:
            total_cell.value = f"=SUM({get_column_letter(first_date_col)}{idx}:{get_column_letter(total_col-1)}{idx})"
            total_cell.font = Font(size=8, color=red, bold=True)
            total_cell.number_format = '#,##0;[Red](#,##0)'

    # メイン表ヘッダ
    header_date_row = 9
    header_weekday_row = 10
    for r in (header_date_row, header_weekday_row):
        for c in range(1, total_col + 1):
            ws.cell(r, c).border = border
            ws.cell(r, c).alignment = center
            ws.cell(r, c).font = Font(size=8, bold=True)

    headers = ["No.", "SID / 商品ID", "媒体名", "区分"]
    for c, value in enumerate(headers, start=1):
        ws.merge_cells(start_row=header_date_row, start_column=c, end_row=header_weekday_row, end_column=c)
        cell = ws.cell(header_date_row, c, value)
        cell.fill = PatternFill("solid", fgColor=pale_blue)
        cell.alignment = center
        cell.font = Font(size=8, bold=True)
        cell.border = border

    jp_weekdays = "月火水木金土日"
    for j, dt in enumerate(dates, start=first_date_col):
        fill = weekend_pink if dt.weekday() >= 5 else white
        ws.cell(header_date_row, j, dt.strftime("%m/%d"))
        ws.cell(header_weekday_row, j, jp_weekdays[dt.weekday()])
        for r in (header_date_row, header_weekday_row):
            ws.cell(r, j).fill = PatternFill("solid", fgColor=fill)
            ws.cell(r, j).font = Font(size=8, bold=True)
            ws.cell(r, j).alignment = center
            ws.cell(r, j).border = border

    ws.merge_cells(start_row=header_date_row, start_column=total_col, end_row=header_weekday_row, end_column=total_col)
    ws.cell(header_date_row, total_col, "Total")
    ws.cell(header_date_row, total_col).fill = PatternFill("solid", fgColor=total_yellow)
    ws.cell(header_date_row, total_col).font = Font(size=8, bold=True)
    ws.cell(header_date_row, total_col).alignment = center
    ws.cell(header_date_row, total_col).border = border

    # 媒体ごとの4行ブロック
    row = header_weekday_row + 1
    metric_defs = [
        ("Daily Target (Initiative)", blue),
        ("Actual", dark),
        ("Promotion Detail", green),
        ("GAP", red),
    ]
    for no, media in enumerate(media_list, start=1):
        start_row = row
        end_row = row + 3
        base_fill = pale_green if no % 2 == 0 else pale_green_2

        # No / SID / 媒体名は4行結合
        for col, value in [(1, no), (2, sid_map.get(media, "")), (3, media)]:
            ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
            c = ws.cell(start_row, col, value)
            c.alignment = center if col != 3 else left
            c.font = Font(size=8)
            c.fill = PatternFill("solid", fgColor=base_fill)
            c.border = border
            # merged範囲にも罫線・塗りを付与
            for rr in range(start_row, end_row + 1):
                ws.cell(rr, col).fill = PatternFill("solid", fgColor=base_fill)
                ws.cell(rr, col).border = border

        for offset, (label, font_color) in enumerate(metric_defs):
            rr = start_row + offset
            label_cell = ws.cell(rr, 4, label)
            label_cell.alignment = left
            label_cell.font = Font(size=8, color=font_color)
            label_cell.fill = PatternFill("solid", fgColor=base_fill)
            label_cell.border = border

            for j, dt in enumerate(dates, start=first_date_col):
                cell = ws.cell(rr, j)
                cell.alignment = center
                cell.border = border
                cell.fill = PatternFill("solid", fgColor=weekend_pink if dt.weekday() >= 5 else base_fill)
                key = (dt.normalize(), media)

                if label == "Daily Target (Initiative)":
                    cell.value = round(cv_map.get(key, 0))
                    cell.font = Font(size=8, color=blue)
                    cell.number_format = '#,##0'
                elif label == "Actual":
                    cell.value = None
                elif label == "Promotion Detail":
                    cell.value = round(cpa_map.get(key, 0)) if cv_map.get(key, 0) else 0
                    cell.font = Font(size=8, color=green)
                    cell.number_format = '¥#,##0'
                else:
                    target_ref = f"{get_column_letter(j)}{start_row}"
                    cell.value = f"=-{target_ref}"
                    cell.font = Font(size=8, color=red)
                    cell.number_format = '#,##0;[Red](#,##0)'

            total_cell = ws.cell(rr, total_col)
            total_cell.fill = PatternFill("solid", fgColor=total_yellow)
            total_cell.border = border
            total_cell.alignment = center
            total_cell.value = f"=SUM({get_column_letter(first_date_col)}{rr}:{get_column_letter(total_col-1)}{rr})"
            if label == "Promotion Detail":
                total_cell.number_format = '¥#,##0'
                total_cell.font = Font(size=8, color=green)
            elif label == "GAP":
                total_cell.number_format = '#,##0;[Red](#,##0)'
                total_cell.font = Font(size=8, color=red)
            elif label == "Daily Target (Initiative)":
                total_cell.font = Font(size=8, color=blue, bold=True)

        row += 4

    # 仕上げ
    ws.freeze_panes = f"{get_column_letter(first_date_col)}{header_weekday_row + 1}"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A{header_date_row}:C{row-1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"1:{header_weekday_row}"
    ws.print_area = f"A1:{get_column_letter(total_col)}{row-1}"

    # 行高
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    for rr in range(4, row):
        if rr not in (1, 2):
            ws.row_dimensions[rr].height = 18

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
            calculate_selected_cpn_base,
            get_cpn_reference_periods,
            enforce_premium_media_cost,
        )
        from config.constants import RECENT_NORMAL_DAYS

        history_df = load_data(uploaded_file)
        history_df["date"] = pd.to_datetime(history_df["date"], errors="coerce").dt.normalize()

        cpn_master = pd.read_excel(uploaded_master, engine="openpyxl")
        required_master_columns = {"日付", "CPN名"}
        missing_master = required_master_columns - set(cpn_master.columns)
        if missing_master:
            raise ValueError(f"CPNマスタに必要な列がありません: {', '.join(sorted(missing_master))}")

        cpn_master = cpn_master.copy()
        cpn_master["日付"] = pd.to_datetime(cpn_master["日付"], errors="coerce").dt.normalize()
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
    default_cpn_index = cpn_list.index("マジ得") if "マジ得" in cpn_list else 0
    selected_cpn = st.sidebar.selectbox("CPN", cpn_list, index=default_cpn_index)

    today = datetime.date.today()
    start_date = st.sidebar.date_input("予測開始", today)
    end_date = st.sidebar.date_input("予測終了", today + datetime.timedelta(days=7))
    if start_date > end_date:
        st.error("予測期間の開始日は終了日以前にしてください。")
        st.stop()

    normal_labels = {"通常", "定常"}
    is_normal_selected = selected_cpn in normal_labels

    if is_normal_selected:
        # 定常は予測期間を365日前へそのままずらした同期間を参照する。
        normal_reference_start = pd.Timestamp(start_date).normalize() - pd.Timedelta(days=365)
        normal_reference_end = pd.Timestamp(end_date).normalize() - pd.Timedelta(days=365)
        normal_reference = history_df[
            history_df["date"].between(normal_reference_start, normal_reference_end)
            & history_df["CPN名"].isin(normal_labels)
        ].copy()

        base_pair = _daily_pair_average(normal_reference)
        base_pair = base_pair[base_pair["media"].isin(selected_media)]
        if base_pair.empty:
            st.error(
                "365日前の同期間に定常実績がありません。"
                "実績CSVとCPNマスタの『定常／通常』登録を確認してください。"
            )
            st.stop()

        st.subheader("📈 前年同期間の定常実績")
        st.caption(
            f"参照期間: {normal_reference_start.date()} ～ {normal_reference_end.date()}（予測期間の365日前）"
        )
        normal_display = (
            normal_reference.groupby(["date", "media"], as_index=False)["cv"].sum()
            .groupby("media", as_index=False)["cv"].mean()
            .rename(columns={"media": "媒体", "cv": "前年同期間の定常CV/日"})
        )
        normal_display["前年同期間の定常CV/日"] = normal_display["前年同期間の定常CV/日"].round(2)
        st.dataframe(normal_display, use_container_width=True, hide_index=True)
    else:
        # 実績内に存在する同一CPNの連続期間を候補化し、複数選択できるようにする。
        available_periods = get_cpn_reference_periods(history_df, selected_cpn)
        if not available_periods:
            st.error(f"実績内に『{selected_cpn}』のキャンペーン期間がありません。")
            st.stop()

        period_options = {
            f"{start.strftime('%Y/%m/%d')} ～ {end.strftime('%Y/%m/%d')}": (start, end)
            for start, end in available_periods
        }
        selected_period_labels = st.sidebar.multiselect(
            "CPN参照期間（複数選択可）",
            options=list(period_options.keys()),
            default=list(period_options.keys()),
        )
        if not selected_period_labels:
            st.warning("CPN参照期間を1つ以上選択してください。")
            st.stop()

        selected_periods = [period_options[label] for label in selected_period_labels]
        base_pair = calculate_selected_cpn_base(history_df, selected_cpn, selected_periods)
        base_pair = base_pair[base_pair["media"].isin(selected_media)]
        if base_pair.empty:
            st.error("選択したCPN参照期間に対象媒体の実績がありません。")
            st.stop()

        total_reference_days = sum((end - start).days + 1 for start, end in selected_periods)
        st.subheader("📈 選択CPN期間の実績")
        st.caption(
            f"選択期間: {len(selected_periods)}期間 / 合計 {total_reference_days}日。"
            " 選択期間のCV・COST合計を合計日数で割った日平均を予測ベースに使用します。"
        )
        for label in selected_period_labels:
            st.write(f"・{label}")

        cpn_display = (
            base_pair.groupby("media", as_index=False)
            .agg(
                **{
                    "選択期間CV/日": ("base_cv", "sum"),
                    "選択期間COST/日": ("cost", "sum"),
                }
            )
            .rename(columns={"media": "媒体"})
        )
        cpn_display["選択期間CV/日"] = cpn_display["選択期間CV/日"].round(2)
        cpn_display["選択期間COST/日"] = cpn_display["選択期間COST/日"].round(0)
        st.dataframe(cpn_display, use_container_width=True, hide_index=True)

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

    # キャンペーン平均をbase_cvとして直接使用するため、CPN倍率は掛けない。
    future_df["cpn_factor"] = 1.0

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

    submission_excel = create_submission_excel(
        opt_summary=opt_summary,
        history_df=history_df,
        start_date=start_date,
        end_date=end_date,
        selected_cpn=selected_cpn,
        opt_mode=opt_mode,
    )
    submission_filename = (
        f"【提出用】楽天カード{pd.Timestamp(start_date).year}年"
        f"{pd.Timestamp(start_date).month}月プランニング.xlsx"
    )

    st.download_button(
        "📥 最適プランを提出用ExcelでDL",
        data=submission_excel,
        file_name=submission_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
