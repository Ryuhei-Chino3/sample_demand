import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

st.set_page_config(page_title="30分値リスケーリング（全時間帯）", layout="wide")

st.title("サンプル30分値リスケーリングアプリ（全時間帯）")
st.markdown(
    """
    目的：サンプルの30分値データ（横に時間帯列を持つ形式）を、月ごとの合計使用量に合わせて
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
        return pd.read_csv(file)
    elif name.endswith((".xls", ".xlsx")):
        try:
            return pd.read_excel(file)
        except ImportError as ie:
            raise RuntimeError(
                "Excel 読み込みに必要なライブラリ（openpyxl）が見つかりません。"
                " requirements.txt に `openpyxl` を追加して再デプロイしてください。"
            ) from ie
    else:
        raise ValueError("対応していないファイル形式です。CSVかXLSXをアップロードしてください。")

try:
    df_raw = load_sample(uploaded_file)
except Exception as e:
    st.error(f"ファイル読み込みでエラーが発生しました: {e}")
    st.stop()

st.subheader("アップロードされた元データの先頭（構造確認）")
st.dataframe(df_raw.head())

# --- 日付/日時列指定 ---
st.sidebar.header("2. 日付列と時間帯列の指定")

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

# --- 時間帯列の自動検出（00:00:00 形式優先、フォールバックあり） ---
expected_time_labels_hms = [f"{h:02d}:{m:02d}:00" for h in range(24) for m in (0, 30)]
candidate_time_cols = [c for c in df.columns if str(c).strip() in expected_time_labels_hms]
if not candidate_time_cols:
    pattern = re.compile(r"^([01]?\d|2[0-3]):[0-5]\d(:00)?$")
    candidate_time_cols = [c for c in df.columns if pattern.match(str(c).strip())]

st.sidebar.markdown("**自動検出された時間帯列（必要なら調整）**")
time_cols = st.sidebar.multiselect(
    "スケーリング対象とする時間帯列を選択（通常はすべて選択）",
    options=df.columns.tolist(),
    default=candidate_time_cols if candidate_time_cols else []
)

if not time_cols:
    st.error("時間帯列が選択されていません。手動で 00:00:00 ～ 23:30:00 相当の列を選択してください。")
    st.stop()

# --- 月ごとの集計 ---
df["__year_month"] = df[date_col].dt.to_period("M")
df["_row_total"] = df[time_cols].sum(axis=1)

monthly_original = (
    df.groupby("__year_month")["_row_total"]
    .sum()
    .rename("元の月合計")
    .to_frame()
)
monthly_original["表示用月"] = monthly_original.index.to_timestamp()

# --- 目標入力（表形式を優先、なければカード形式） ---
target_inputs = {}
st.subheader("各月の元の合計使用量と新しい月合計の入力")
st.markdown("各月ごとに目標とする合計使用量を指定してください。選択された全時間帯列の合計がその値になるようスケーリングされます。")

use_table = False
edited_df = None
# 優先：experimental_data_editor または data_editor
if hasattr(st, "experimental_data_editor"):
    use_table = True
    editable = monthly_original.reset_index()
    editable["表示用月"] = editable["__year_month"].dt.strftime("%Y-%m")
    editable = editable.rename(columns={"元の月合計": "元の月合計使用量"})
    editable["入力目標"] = editable["元の月合計使用量"]
    edited_df = st.experimental_data_editor(
        editable[["__year_month", "表示用月", "元の月合計使用量", "入力目標"]],
        num_rows="fixed",
        use_container_width=True,
    )
elif hasattr(st, "data_editor"):
    use_table = True
    editable = monthly_original.reset_index()
    editable["表示用月"] = editable["__year_month"].dt.strftime("%Y-%m")
    editable = editable.rename(columns={"元の月合計": "元の月合計使用量"})
    editable["入力目標"] = editable["元の月合計使用量"]
    edited_df = st.data_editor(
        editable[["__year_month", "表示用月", "元の月合計使用量", "入力目標"]],
        num_rows="fixed",
        use_container_width=True,
    )

if use_table and edited_df is not None:
    for _, row in edited_df.iterrows():
        label = row["表示用月"]  # "YYYY-MM"
        try:
            target_value = float(row["入力目標"])
        except Exception:
            target_value = float(row["元の月合計使用量"])
        target_inputs[label] = target_value
else:
    # フォールバック：横スクロールカード入力（number_input）
    st.markdown(
        """
        <style>
        .horizontal-inputs {
            display: flex;
            gap: 12px;
            overflow-x: auto;
            padding: 6px 0;
        }
        .monthly-box {
            min-width: 160px;
            flex: 0 0 auto;
            background: #f1f5f9;
            padding: 8px;
            border-radius: 8px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        }
        .monthly-label {
            font-weight: 600;
            margin-bottom: 4px;
            font-size: 0.9rem;
        }
        .small-meta {
            font-size: 0.65rem;
            color: #555;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown('<div class="horizontal-inputs">', unsafe_allow_html=True)
    for period, row in monthly_original.iterrows():
        label = period.strftime("%Y-%m")
        orig = row["元の月合計"]
        default = float(round(orig, 6))
        st.markdown(f'<div class="monthly-box">', unsafe_allow_html=True)
        st.markdown(f'<div class="monthly-label">{label}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="small-meta">元の合計: {orig:.6f}</div>', unsafe_allow_html=True)
        # number_input を直接使う（ラベル空にして見た目をコンパクトに）
        target_value = st.number_input(
            label="",
            value=default,
            format="%.6f",
            key=f"target_{label}",
            min_value=0.0,
            help=f"{label} の目標（月合計）",
        )
        st.markdown("</div>", unsafe_allow_html=True)
        target_inputs[label] = float(target_value)
    st.markdown('</div>', unsafe_allow_html=True)

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
df_scaled["__scale_factor"] = df_scaled["__year_month"].map(scaling)

for col in time_cols:
    df_scaled[col] = df_scaled[col] * df_scaled["__scale_factor"].where(df_scaled["__scale_factor"].notna(), 1.0)

# --- 検証表示 ---
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
