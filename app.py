import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference

st.set_page_config(page_title="30分値リスケーリングアプリ", layout="wide")
st.title("30分値リスケーリングアプリ（最大・最小月のデマンドカーブ付き）")

st.sidebar.header("1. ファイルアップロード")
uploaded_file = st.sidebar.file_uploader("CSV または XLSX ファイルをアップロード", type=["csv", "xlsx"])
if uploaded_file is None:
    st.stop()

@st.cache_data
def load_sample(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

df_raw = load_sample(uploaded_file)
st.subheader("アップロードされたデータ")
st.dataframe(df_raw.head())

st.sidebar.header("2. 日付列選択")
possible_date_cols = [col for col in df_raw.columns if pd.to_datetime(df_raw[col], errors='coerce').notna().mean() > 0.5]
date_col = st.sidebar.selectbox("日付列を選択", options=df_raw.columns.tolist(), index=df_raw.columns.get_indexer([possible_date_cols[0]])[0] if possible_date_cols else 0)

df = df_raw.copy()
df[date_col] = pd.to_datetime(df[date_col])

expected_time_labels = [f"{h:02d}:{m:02d}:00" for h in range(24) for m in (0, 30)]
time_cols = [c for c in df.columns if str(c).strip() in expected_time_labels]

if not time_cols:
    st.error("時間帯列が見つかりません。形式を確認してください。")
    st.stop()

# 月集計と表示
df["__year_month"] = df[date_col].dt.to_period("M")
df["_row_total"] = df[time_cols].sum(axis=1)
monthly_original = df.groupby("__year_month")['_row_total'].sum().rename("元の月合計").to_frame()

st.subheader("各月の合計使用量と新しい目標使用量")
with st.form("monthly_form"):
    cols = st.columns(len(monthly_original))
    target_inputs = {}
    for (period, row), col in zip(monthly_original.iterrows(), cols):
        label = period.strftime("%Y/%-m")
        orig = row["元の月合計"]
        with col:
            st.markdown(f"**{label}**")
            st.markdown(f"<span style='font-size:0.65rem;'>元: {orig:.1f}</span>", unsafe_allow_html=True)
            val = st.text_input("", value="", key=str(period))
            try:
                target_inputs[str(period)] = float(val) if val.strip() else float(orig)
            except:
                target_inputs[str(period)] = float(orig)
    submitted = st.form_submit_button("スケーリング実行")

if not submitted:
    st.stop()

scaling = {pd.Period(k): v / monthly_original.loc[pd.Period(k), "元の月合計"] if monthly_original.loc[pd.Period(k), "元の月合計"] != 0 else 0 for k, v in target_inputs.items()}

# スケーリング
df_scaled = df.copy()
df_scaled["__scale_factor"] = df_scaled["__year_month"].map(scaling)
for col in time_cols:
    df_scaled[col] = df_scaled[col] * df_scaled["__scale_factor"].fillna(1.0)

# --- Excel 出力 ---
def to_excel_bytes_with_curve(df_rescaled, time_cols):
    output = BytesIO()
    df_rescaled["__year_month"] = df_rescaled[date_col].dt.to_period("M")
    monthly_sum = df_rescaled.groupby("__year_month")[time_cols].sum().sum(axis=1)
    max_month = monthly_sum.idxmax()
    min_month = monthly_sum.idxmin()

    max_curve = df_rescaled[df_rescaled["__year_month"] == max_month][time_cols].mean()
    min_curve = df_rescaled[df_rescaled["__year_month"] == min_month][time_cols].mean()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_rescaled.drop(columns=["__scale_factor"], errors="ignore").to_excel(writer, index=False, sheet_name="Rescaled")
        df_curve = pd.DataFrame({
            "時間帯": time_cols,
            f"{max_month.strftime('%Y-%m')}（最大）": max_curve.values,
            f"{min_month.strftime('%Y-%m')}（最小）": min_curve.values,
        })
        df_curve.to_excel(writer, index=False, sheet_name="DemandCurve")

    output.seek(0)
    wb = load_workbook(output)
    ws = wb["DemandCurve"]

    chart = LineChart()
    chart.title = "最大月・最小月のデマンドカーブ"
    chart.y_axis.title = "平均使用量"
    chart.x_axis.title = "時間帯"

    max_row = ws.max_row
    data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=max_row)
    cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"E{max_row + 2}")

    out_bytes = BytesIO()
    wb.save(out_bytes)
    return out_bytes.getvalue()

excel_bytes = to_excel_bytes_with_curve(df_scaled, time_cols)

st.subheader("結果ダウンロード")
st.download_button(
    label="Excel ファイルをダウンロード",
    data=excel_bytes,
    file_name="rescaled_with_curve.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("出力ファイルの生成が完了しました！")
