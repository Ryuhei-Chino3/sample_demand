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

# 日付っぽい列を自動検出
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

# --- 時間帯列の自動検出（強化版） ---
# 典型的な 30 分刻みのラベルを明示リスト化（0:00, 0:30, ..., 23:30）
expected_time_labels = [f"{h}:{m:02d}" for h in range(24) for m in (0, 30)]

# 優先：明示的なラベルとの一致で候補を取る
candidate_time_cols = [c for c in df.columns if str(c) in expected_time_labels]

# フォールバック：柔軟に 0:00〜23:30 形式を拾う正規表現
if not candidate_time_cols:
    pattern = re.compile(r"^([01]?\d|2[0-3]):[0-5]\d$")
    candidate_time_cols = [c for c in df.columns if pattern.match(str(c).strip())]

st.sidebar.markdown("**自動検出された時間帯列（必要なら調整）**")
time_cols = st.sidebar.multiselect(
    "スケーリング対象とする時間帯列を選択（通常はすべて選択）",
    options=df.columns.tolist(),
    default=candidate_time_cols if candidate_time_cols else []
)

if not time_cols:
    st.error("時間帯列が選択されていません。0:00〜23:30 に相当する列を手動で選択してください。")
    st.stop()

# --- 月ごとの集計と目標入力 ---
df["__year_month"] = df[date_col].dt.to_period("M")

# 各行の全時間帯合計（対象列の合計）をとる
df["_row_total"] = df[time_cols].sum(axis=1)

# 月ごとの元の合計使用量
monthly_original = (
    df.groupby("__year_month")["_row_total"]
    .sum()
    .rename("元の月合計")
    .to_frame()
)
monthly_original["表示用月"] = monthly_original.index.to_timestamp()

st.subheader("各月の元の合計使用量と新しい月合計の入力")
st.markdown("各月ごとに目標とする合計使用量を入力してください。全時間帯列の合計がその値になるようにスケーリングします。")

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
# 各月の比率を計算
scaling = {}
for label, target in target_inputs.items():
    period = pd.Period(label, freq="M")
    orig = monthly_original.loc[period, "元の月合計"]
    if orig == 0:
        scaling[period] = np.nan
    else:
        scaling[period] = target / orig

# 行ごとに係数を割り当て
df_scaled = df.copy()
df_scaled["__scale_factor"] = df_scaled["__year_month"].map(scaling)

# 時間帯列を係数で一括スケーリング（NaNなら元のまま）
for col in time_cols:
    df_scaled[col] = df_scaled[col] * df_scaled["__scale_factor"].where(df_scaled["__scale_factor"].notna(), 1.0)

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

# 補助列を落として元の構造で出力
to_export = df_scaled.drop(columns=["__year_month", "_row_total", "__scale_factor"], errors="ignore")

def to_excel_bytes(df: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rescaled")
    return output.getvalue()

try:
    excel_bytes = to_excel_bytes(to_export)
except ImportError:
    st.error("Excel 出力に必要なライブラリ（openpyxl）が見つかりません。requirements.txt に追加してください。")
    st.stop()

st.download_button(
    label="スケーリング結果を Excel ダウンロード",
    data=excel_bytes,
    file_name=output_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("処理完了。出力ファイルをダウンロードしてください。")
