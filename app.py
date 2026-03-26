import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import smtplib
from email.message import EmailMessage
import os

st.set_page_config(page_title="支払照合ツール", layout="wide")
st.title("支払照合ツール（社内運用版）")
st.markdown("""
<style>

/* ===== 全体背景 ===== */
.stApp {
    background-color: #F5F7FA;
}

/* ===== 入力欄（テキスト・数値） ===== */
input, textarea {
    background-color: #FFFFFF !important;
    color: #000000 !important;
    border: 2px solid #007BFF !important;
    border-radius: 6px !important;
}

/* ===== ファイルアップロード ===== */
section[data-testid="stFileUploader"] {
    background-color: #FFFFFF;
    padding: 10px;
    border-radius: 8px;
    border: 2px solid #007BFF;
}

/* ===== ボタン ===== */
button {
    background-color: #007BFF !important;
    color: white !important;
    border-radius: 8px !important;
    font-weight: bold;
    border: none;
}

/* ===== ボタン（ホバー） ===== */
button:hover {
    background-color: #0056b3 !important;
    color: white !important;
}

/* ===== テーブル ===== */
[data-testid="stDataFrame"] {
    background-color: #FFFFFF;
}

/* ===== 成功メッセージ ===== */
.stAlert-success {
    background-color: #E6F4EA !important;
    color: #1E7E34 !important;
}

/* ===== エラーメッセージ ===== */
.stAlert-error {
    background-color: #FDECEA !important;
    color: #C0392B !important;
}

</style>
""", unsafe_allow_html=True)

# ===== 状態管理 =====
if "calculated" not in st.session_state:
    st.session_state["calculated"] = False

# ===== 入力 =====
target = st.number_input("目標金額", min_value=0)

file_orders = st.file_uploader("注文書整理ファイル", type=["xlsx"])
file_paid = st.file_uploader("入金一覧ファイル", type=["xlsx"])

# ★メール入力はここだけに統一
recipient = st.text_input("送信先メールアドレス")

# ===== データ整形 =====
def clean_data(df_orders, df_paid):
    col_pono = df_orders.columns[7]
    col_amount = df_orders.columns[8]

    df_orders[col_pono] = df_orders[col_pono].astype(str).str.strip()
    df_paid.iloc[:, 2] = df_paid.iloc[:, 2].astype(str).str.strip()

    df_orders[col_amount] = df_orders[col_amount].astype(str).str.replace(",", "")
    df_orders[col_amount] = pd.to_numeric(df_orders[col_amount], errors="coerce")

    return df_orders, df_paid

# ===== 高速組み合わせ =====
def find_best_combination(records, target):
    records = sorted(records, key=lambda x: x[1], reverse=True)

    best_combo = []
    best_total = 0

    total = 0
    current = []

    for pono, amount in records:
        if total + amount <= target:
            current.append((pono, amount))
            total += amount

    best_combo = current
    best_total = total

    for i in range(len(records)):
        temp = []
        total = 0

        for j in range(len(records)):
            if i == j:
                continue

            pono, amount = records[j]
            if total + amount <= target:
                temp.append((pono, amount))
                total += amount

        if abs(target - total) < abs(target - best_total):
            best_combo = temp
            best_total = total

    return best_combo

# ===== メール =====
def send_mail(file_path, recipient):

    SMTP_USER = os.getenv("SMTP_USER")
    SMTP_PASS = os.getenv("SMTP_PASS")

    if not SMTP_USER or not SMTP_PASS:
        return False, "SMTP設定なし"

    try:
        msg = EmailMessage()
        msg["Subject"] = "支払データ送付"
        msg["From"] = SMTP_USER
        msg["To"] = recipient

        msg.set_content("支払データを送付します。")

        with open(file_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="xlsx", filename=file_path)

        with smtplib.SMTP("smtp.sina.net", 587) as smtp:
            smtp.starttls()
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)

        return True, "成功"

    except Exception as e:
        return False, str(e)

# ===== 計算 =====
if st.button("計算", key="calc_button"):

    if file_orders is None or file_paid is None:
        st.error("ファイルをアップしてください")
        st.stop()

    with st.spinner("処理中..."):

        df_orders = pd.read_excel(file_orders)
        df_paid = pd.read_excel(file_paid)

        df_orders, df_paid = clean_data(df_orders, df_paid)

        col_pono = df_orders.columns[7]
        col_amount = df_orders.columns[8]

        paid_list = df_paid.iloc[:, 2].dropna().tolist()
        df_unpaid = df_orders[~df_orders[col_pono].isin(paid_list)]

        records = [
            (row[col_pono], row[col_amount])
            for _, row in df_unpaid.iterrows()
            if pd.notna(row[col_amount]) and row[col_amount] <= target
        ]

        best_combo = find_best_combination(records, target)

        st.session_state["result"] = pd.DataFrame(best_combo, columns=["PONO", "金額"])
        best_pono = [x[0] for x in best_combo]

        df_orders["候補"] = df_orders[col_pono].isin(best_pono)

        file_path = "支払チェック結果.xlsx"
        df_orders.to_excel(file_path, index=False)

        wb = load_workbook(file_path)
        ws = wb.active

        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        col_index = df_orders.columns.get_loc("候補") + 1

        for row in range(2, ws.max_row+1):
            if ws.cell(row=row, column=col_index).value:
                for col in range(1, ws.max_column+1):
                    ws.cell(row=row, column=col).fill = fill

        wb.save(file_path)

        st.session_state["file_path"] = file_path
        st.session_state["calculated"] = True

# ===== 計算後UI =====
if st.session_state["calculated"]:

    st.success("計算完了")

    st.dataframe(st.session_state["result"])

    st.download_button(
        "Excelダウンロード",
        open(st.session_state["file_path"], "rb"),
        file_name="支払チェック結果.xlsx"
    )

    # ===== メール送信 =====
    if st.button("OK（メール送信）", key="send_button"):

        success, message = send_mail(st.session_state["file_path"], recipient)

        if success:
            st.success("メール送信完了")
        else:
            st.error(message)

    # ===== 再計算 =====
    if st.button("再計算", key="recalc_button"):
        st.session_state["calculated"] = False
        st.experimental_rerun()



