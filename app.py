import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="30分値リスケーリングアプリ", layout="wide")

st.title("サンプル30分値リスケーリングアプリ")
st.markdown(
    """
    目的：本来の30分値データがない場合でも、サンプルの30分値を月ごとの合計使用量に合わせてスケーリングし、
    同じ構造の新しい30分値データ（Excelファイル）を生成します。
    """
)

# --- ファイルアップロード ---
st.sidebar.header("1. サンプルファイルをアップロード")
uploaded_file = st.sidebar.file_uploader(
    "CSV または XLSX をアップロード", type=["csv", "xlsx"], accept_multiple_files=False
)

if uploaded_file is None:
    st.warning("まずはサンプルの30分値データ（CSV または XLSX）をアップロードしてください。")
    st.stop()

# --- 読み込み ---
@st.cache_data
def load_sample(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
    elif name.endswith((".xls", ".xlsx")):
        try:
            df = pd.read_excel(file)
        except ImportError as ie:
            raise RuntimeError(
                "Excel 読み込みに必要なライブラリ（openpyxl）が見つかりません。"
                " requirements.txt に `openpyxl` を追加して再デプロイしてください。"
            ) from ie
    else:
        raise ValueError("対応していないファイル形式です。CSVかXLSXをアップロードしてください。")
    return df

try:
    df_raw = load_sample(uploaded_file)
except Exception as e:
    st.error(f"ファイル読み込みでエラーが発生しました: {e}")
    st.stop()

st.subheader("1. アップロードされた元データの先頭（構造確認）")
st.dataframe(df_raw.head())

# --- 列指定 ---
st.sidebar.header("2. 列指定（自動検出がうまくいかなければ調整）")
possible_datetime = []
for col in df_raw.columns:
    try:
        parsed = pd.to_datetime(df_raw[col], errors="coerce")
        if parsed.notna().mean() > 0.5:
            possible_datetime.append(col)
    except Exception:
        continue

# 安全に int にキャストし、範囲チェックする
default_index = 0
if possible_datetime:
    try:
        idx_array = df_raw.columns.get_indexer([possible_datetime[0]])
        if len(idx_array) > 0:
            candidate = int(idx_array[0])  # numpy.int64 を明示的に int 化
            if 0 <= candidate < len(df_raw.columns):
                default_index = candidate
    except Exception:
        default_index = 0

datetime_col = st.sidebar.selectbox(
    "日時列を選択してください",
    options=df_raw.columns.tolist(),
    index=default_index,
)

df = df_raw.copy()
try:
    df[datetime_col] = pd.to_datetime(df[datetime_col])
except Exception as e:
    st.error(f"{datetime_col} を日時に変換できませんでした。形式を確認してください。詳細: {e}")
    st.stop()

numeric_cols = df.select_dtypes(include="number").columns.tolist()
if not numeric_cols:
    st.error("数値列が見つかりません。使用量列にあたる列を含んだデータをアップロードしてください。")
    st.stop()

usage_col = st.sidebar.selectbox(
    "使用量（スケーリング対象）となる列を選択してください", options=numeric_cols
)

if usage_col is None or datetime_col is None:
    st.error("日時列と使用量列を正しく指定してください。")
    st.stop()

# --- 月ごとの集計と入力フォーム ---
df["__year_month"] = df[datetime_col].dt.to_period("M")

monthly_original = (
    df.groupby("__year_month")[usage_col]
    .sum()
    .rename("元の月合計")
    .to_frame()
)
monthly_original["表示用月"] = monthly_original.index.to_timestamp()

st.subheader("2. 各月の元の合計使用量と新しい月合計の入力")
st.markdown("必要な月ごとの目標合計使用量を入力してください。元の合計がゼロの月はスケーリングできないのでご注意を。")

target_inputs = {}
with st.form("monthly_targets_form"):
    for period, row in monthly_original.iterrows():
        display_label = period.strftime("%Y-%m")
        original_val = row["元の月合計"]
        default = float(round(original_val, 6))
        target_inputs[display_label] = st.number_input(
            label=f"{display_label} の新しい月合計使用量",
            value=default,
            format="%.6f",
            help=f"元の合計: {original_val:.6f}",
            key=f"target_{display_label}",
            min_value=0.0,
        )
    submitted = st.form_submit_button("スケーリング実行")

if not submitted:
    st.info("上のフォームで新しい月合計使用量を入力して「スケーリング実行」を押してください。")
    st.stop()

# --- スケーリング処理 ---
scaling = {}
for display_label, target_value in target_inputs.items():
    period = pd.Period(display_label, freq="M")
    orig = monthly_original.loc[period, "元の月合計"]
    if orig == 0:
        scaling[period] = np.nan
    else:
        scaling[period] = target_value / orig

df_scaled = df.copy()


def apply_scale(row):
    period = row["__year_month"]
    factor = scaling.get(period, np.nan)
    if pd.isna(factor):
        return row[usage_col]
    return row[usage_col] * factor


df_scaled[usage_col] = df_scaled.apply(apply_scale, axis=1)

# --- 確認 ---
st.subheader("3. スケーリング後の簡易検証（各月の合計）")
scaled_monthly = (
    df_scaled.groupby("__year_month")[usage_col]
    .sum()
    .rename("スケーリング後合計")
    .to_frame()
)
compare = monthly_original.join(scaled_monthly)
compare["入力目標"] = [target_inputs[p.strftime("%Y-%m")] for p in compare.index]
compare["比率（実績/目標）"] = compare["スケーリング後合計"] / compare["入力目標"].replace({0: np.nan})
st.dataframe(compare.style.format({
    "元の月合計": "{:.6f}",
    "スケーリング後合計": "{:.6f}",
    "入力目標": "{:.6f}",
    "比率（実績/目標）": "{:.4f}"
}))

for period, row in compare.iterrows():
    if row["入力目標"] == 0 and row["元の月合計"] != 0:
        st.warning(f"{period.strftime('%Y-%m')} の入力目標が0です。元の値があるため、すべて0にスケーリングされます。")
    if row["元の月合計"] == 0:
        st.warning(f"{period.strftime('%Y-%m')} は元の合計が0のためスケーリングできていません（そのまま）。")

# --- 出力 ---
st.subheader("4. 出力ファイル設定とダウンロード")
output_name = st.text_input("出力ファイル名（拡張子 .xlsx を含めてください）", value="rescaled_30min.xlsx")
if not output_name.lower().endswith(".xlsx"):
    st.error("ファイル名は .xlsx で終わる必要があります。")
    st.stop()

to_export = df_scaled.copy()
to_export = to_export.drop(columns=["__year_month"])


def to_excel_bytes(df: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rescaled")
    return output.getvalue()


try:
    excel_bytes = to_excel_bytes(to_export)
except ImportError:
    st.error("Excel 書き出しに必要な `openpyxl` が見つかりません。requirements.txt に openpyxl を追加してください。")
    st.stop()

st.download_button(
    label="スケーリング結果をExcelでダウンロード",
    data=excel_bytes,
    file_name=output_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("処理が完了しました。上から結果ファイルをダウンロードしてください。")
