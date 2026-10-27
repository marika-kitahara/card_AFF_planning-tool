"""任意アップロードの競合還元額調査分析。予測ロジックには依存しない。"""
import datetime as dt
import math
import re
from io import BytesIO

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib.ticker import FuncFormatter, MultipleLocator
from openpyxl import load_workbook

DATE_SHEET = re.compile(r"^\d{6}$")
AMOUNT = re.compile(r"[-+]?\d[\d,]*(?:\.\d+)?")


@st.cache_data(show_spinner=False, max_entries=3)
def parse_competitor_workbook(content: bytes):
    wb = load_workbook(BytesIO(content), read_only=True, data_only=True)
    records = []
    warnings = []
    for name in wb.sheetnames:
        if not DATE_SHEET.fullmatch(name):
            continue
        try:
            date = dt.datetime.strptime(name, "%y%m%d").date()
        except ValueError:
            warnings.append(f"日付として認識できないシート: {name}")
            continue
        ws = wb[name]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header or str(header[0]).strip() != "媒体名":
            warnings.append(f"見出し形式が異なるシート: {name}")
            continue
        products = [(i, str(v).strip()) for i, v in enumerate(header) if i >= 3 and v is not None and str(v).strip()]
        for row in rows:
            if not row or row[0] is None:
                continue
            media = str(row[0]).strip()
            segment = str(row[1]).strip() if len(row) > 1 and row[1] is not None else "未設定"
            category = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
            if category != "単価":
                continue
            for idx, product in products:
                if idx >= len(row) or row[idx] is None:
                    continue
                raw = str(row[idx]).strip()
                match = AMOUNT.search(raw.replace("，", ","))
                if not match:
                    continue
                value = float(match.group().replace(",", ""))
                if not math.isfinite(value) or value < 0:
                    continue
                records.append((date, media, segment, product, value))
    wb.close()
    df = pd.DataFrame(records, columns=["調査日", "媒体", "セグメント", "商品", "還元額"])
    if not df.empty:
        df["調査日"] = pd.to_datetime(df["調査日"])
        df = df.drop_duplicates(["調査日", "媒体", "セグメント", "商品"], keep="last")
    return df, warnings


def show_competitor_analysis():
    with st.expander("📈 競合還元額分析（任意アップロード）", expanded=False):
        file = st.file_uploader("競合還元額調査Excel", type=["xlsx", "xlsm"], key="competitor_workbook", help="シート名がYYMMDDの日付シートを読み取ります。未アップロードでも既存機能はそのまま利用できます。")
        if file is None:
            st.info("調査Excelをアップロードすると、媒体別の最新還元額と推移を表示します。")
            return
        try:
            with st.spinner("日付シートを読み込み中…"):
                df, warnings = parse_competitor_workbook(file.getvalue())
        except Exception as exc:
            st.error(f"競合調査ファイルを読み取れませんでした: {exc}")
            return
        if df.empty:
            st.warning("対象の日付シートから単価データを取得できませんでした。")
            return
        if warnings:
            st.warning(f"読み取り対象外のシートが{len(warnings)}件あります。例: {', '.join(warnings[:3])}")
        st.caption(f"収録調査日: {df['調査日'].min():%Y/%m/%d} ～ {df['調査日'].max():%Y/%m/%d} / {df['調査日'].nunique():,}日")
        c1, c2, c3 = st.columns(3)
        with c1:
            product = st.selectbox("対象商品", sorted(df["商品"].unique()), index=next((i for i,p in enumerate(sorted(df['商品'].unique())) if '楽天カード' in p),0), key="competitor_product")
        with c2:
            segment_options = sorted(df["セグメント"].unique())
            segments = st.multiselect("セグメント", segment_options, default=segment_options, key="competitor_segments")
        with c3:
            period = st.selectbox("表示期間", [30,90,180,365,"全期間"], index=1, format_func=lambda x: f"過去{x}日" if isinstance(x,int) else x, key="competitor_period")
        base = df[(df["商品"] == product) & df["セグメント"].isin(segments)].copy()
        if base.empty:
            st.info("選択条件に該当するデータがありません。")
            return
        end = base["調査日"].max()
        if period != "全期間":
            base = base[base["調査日"] >= end-pd.Timedelta(days=period-1)]
        options = sorted(base["媒体"].unique())
        media = st.multiselect("表示媒体", options, default=options[:min(8,len(options))], key="competitor_media")
        selected = base[base["媒体"].isin(media)]
        if selected.empty:
            st.info("表示する媒体を選択してください。")
            return
        # 同一媒体に複数セグメントがある場合、混ぜずに別系列として表示する。
        selected = selected.copy()
        selected["系列"] = selected["媒体"] + " / " + selected["セグメント"]
        latest_date = selected["調査日"].max()
        latest = selected[selected["調査日"] == latest_date].sort_values("還元額", ascending=False)
        k1,k2,k3 = st.columns(3)
        k1.metric("最新調査日", latest_date.strftime("%Y/%m/%d"))
        k2.metric("最新最高還元額", f"{latest['還元額'].max():,.0f}円")
        k3.metric("期間内最高還元額", f"{selected['還元額'].max():,.0f}円")
        st.subheader("媒体別還元額推移")
        fig, ax = plt.subplots(figsize=(11,4.5))
        for label, group in selected.groupby("系列", sort=True):
            series = group.sort_values("調査日").drop_duplicates("調査日", keep="last")
            ax.step(series["調査日"], series["還元額"], where="post", label=label, linewidth=1.8)
        low, high = selected["還元額"].min(), selected["還元額"].max()
        lower = max(0, math.floor((low-500)/1000)*1000)
        upper = math.ceil((high+500)/1000)*1000
        if upper <= lower: upper = lower+1000
        ax.set_ylim(lower,upper)
        ax.yaxis.set_major_locator(MultipleLocator(1000))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda x,pos: f"{x:,.0f}円"))
        ax.grid(axis="y", alpha=.25)
        ax.legend(loc="upper left", bbox_to_anchor=(1.01,1), fontsize=8)
        fig.autofmt_xdate()
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
        st.caption("縦軸は1,000円刻み・範囲自動調整。調査日の値を次の調査日まで維持する階段表示です。未調査日の値を実測したものではありません。")
        st.subheader("最新調査日の還元額一覧")
        st.dataframe(latest[["媒体","セグメント","還元額"]], hide_index=True, width="stretch", column_config={"還元額":st.column_config.NumberColumn(format="%.1f 円")})
        st.caption("0円は元ファイル記載値のまま表示します。掲載なしとの区別は元データ側の定義確認が必要です。")
