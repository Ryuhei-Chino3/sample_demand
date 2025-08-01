import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

st.set_page_config(page_title="30分値リスケーリング（全時間帯）", layout="wide")

st.title("サンプル30分値リスケーリングアプリ（0:00〜23:30 全列）")
st.markdown(
    """
    目的：サンプルの30分値データ（横に 0:00～23:30 の時間帯列を持つ形式）を、月ごとの合計使用量に合わせて
    **全ての時間帯列を同一比率でスケーリング**した新しいデータを出力します。
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

st.subheader("アップロードされた元データの先頭（構造確認）")
st.dataframe(df_raw.head())

# --- 日付/日時列指定 ---
st.sidebar.header("2. 日付列と時間帯列の指定")

# 自動候補：日付っぽい列
possible_date_cols = []
for col in df_raw.columns:
    try:
        parsed = pd.to_datetime(df_raw[col], errors="coerce")
        if parsed.notna().mean() > 0.5:
            possible_date_cols.append(col)
    except Exception:
        continue

default_date_idx = 0
if possible_date_cols:
    try:
        default_date_idx = int(df_raw.columns.get_indexer([possible_date_cols[0]])[0])
    except Exception:
        default_date_idx = 0

date_col = st.sidebar.selectbox(
    "日付／日時列を選択してください", options=df_raw.columns.tolist(), index=default_date_idx
)

df = df_raw.copy()
try:
    df[date_col] = pd.to_datetime(df[date_col])
except Exception as e:
    st.error(f"{date_col} を日時に変換できませんでした。形式を確認してください。詳細: {e}")
    st.stop()

# --- 時間帯列の自動検出（0:00〜23:30）と確認 ---
def is_time_label(col_name: str) -> bool:
    # 0:00, 0:30, 00:00, 23:30 などをカバー。時刻表記に余裕持たせる。
    pattern = r"^([01]?\d|2[0-3]):[0-5]\d$"
    return bool(re.match(pattern, str(col_name).strip()))

candidate_time_cols = [c for c in df.columns if is_time_label(c)]
st.sidebar.markdown("**自動検出された30分刻みの時間帯列（必要なら調整）**")
time_cols = st.sidebar.multiselect(
    "スケーリング対象とする時間帯列を選択（全部選ぶのが通常）",
    options=df.columns.tolist(),
    default=candidate_time_cols if candidate_time_cols else []
)

if not time_cols:
    st.error("0:00〜23:30 に相当する時間帯列が自動検出されませんでした。手動で選択してください。")
    st.stop()

# --- 月ごとの集計と入力 ---
df["__year_month"] = df[date_col].dt.to_period("M")

# 各行（日ごと／タイムスタンプごと）の合計（全時間帯列の合計）
df["_row_total"] = df[time_cols].sum(axis=1)

# 月合計：各行の合計を月ごとに足す
monthly_original = (
    df.groupby("__year_month")["_row_total"]
    .sum()
    .rename("元の月合計")
    .to_frame()
)
monthly_original["表示用月"] = monthly_original.index.to_timestamp()

st.subheader("各月の元の合計使用量と新しい月合計の入力")
st.markdown("各月ごとに目標とする合計使用量を入力してください（全時間帯の合計がその値になるようスケーリングされます）。")

target_inputs = {}
with st.form("monthly_targets_form"):
    for period, row in monthly_original.iterrows():
        label = period.strftime("%Y-%m")
        orig = row["元の月合計"]
        default = float(round(orig, 6))
        target_inputs[label] = st.number_input(
            label=f"{label} の新しい月合計使用量",
            value=default,
            format="%.6f",
            help=f"元の合計: {orig:.6f}",
            key=f"target_{label}",
            min_value=0.0,
        )
    submitted = st.form_submit_button("スケーリング実行")

if not submitted:
    st.info("月ごとの目標使用量を入力して「スケーリング実行」を押してください。")
    st.stop()

# --- スケーリング ---
scaling = {}
for label, target in target_inputs.items():
    period = pd.Period(label, freq="M")
    orig = monthly_original.loc[period, "元の月合計"]
    if orig == 0:
        scaling[period] = np.nan
    else:
        scaling[period] = target / orig

df_scaled = df.copy()

def scale_row(row):
    period = row["__year_month"]
    factor = scaling.get(period, np.nan)
    if pd.isna(factor):
        return row[time_cols]
    return row[time_cols] * factor

# 各時間帯列をスケーリング（全列一括）
scaled_times = df_scaled.apply(lambda r: scale_row(r), axis=1)
# DataFrame に戻して置き換え
for col in time_cols:
    df_scaled[col] = scaled_times.apply(lambda x: x[col])

# --- 結果表示 ---
st.subheader("スケーリング後の各月合計の検証")
scaled_monthly = (
    df_scaled.groupby("__year_month")[time_cols]
    .sum()
    .sum(axis=1)
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
        st.warning(f"{period.strftime('%Y-%m')} の入力目標が0です。元の値があるため、全て0になります。")
    if row["元の月合計"] == 0:
        st.warning(f"{period.strftime('%Y-%m')} は元の合計が0のためスケーリングされていません。")

# --- 出力 ---
st.subheader("出力ファイル設定とダウンロード")
output_name = st.text_input("出力ファイル名（.xlsx で終わる）", value="rescaled_30min_full.xlsx")
if not output_name.lower().endswith(".xlsx"):
    st.error("ファイル名は .xlsx で終わる必要があります。")
    st.stop()

# 補助列を外して元の構造を保つ
to_export = df_scaled.drop(columns=["__year_month", "_row_total"], errors="ignore")

def to_excel_bytes(df: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rescaled")
    return output.getvalue()

try:
    excel_bytes = to_excel_bytes(to_export)
except ImportError:
    st.error("Excel 出力に必要な openpyxl が見つかりません。requirements.txt に追加してください。")
    st.stop()

st.download_button(
    label="スケーリング結果を Excel ダウンロード",
    data=excel_bytes,
    file_name=output_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("処理完了。出力ファイルをダウンロードしてください。")
