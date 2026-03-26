import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import smtplib
from email.message import EmailMessage
import os
import json

st.set_page_config(page_title="支払照合ツール", layout="wide")
st.title("支払照合ツール（社内運用版）")

# ===== CSS（見やすさ改善） =====
st.markdown("""
<style>
.stApp { background-color: #F5F7FA; }
input { background-color: #FFFFFF !important; color: #000000 !important; border: 2px solid #007BFF !important; }
button { background-color: #007BFF !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ===== 設定保存 =====
CONFIG_FILE = "config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"user": "", "pass": ""}

def save_config(user, pw):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"user": user, "pass": pw}, f)

config = load_config()

# ===== メール設定UI =====
st.markdown("### 📧 メール設定")

smtp_user = st.text_input("送信元メールアドレス", value=config["user"])
smtp_pass = st.text_input("メールパスワード", type="password", value=config["pass"])

if st.button("設定を保存", key="save_config"):
    save_config(smtp_user, smtp_pass)
    st.success("保存しました")

# ===== 状態管理 =====
if "calculated" not in st.session_state:
    st.session_state["calculated"] = False

# ===== 入力 =====
st.markdown("### 💰 金額入力")
target = st.number_input("目標金額", min_value=0)

st.markdown("### 📂 データアップロード")
file_orders = st.file_uploader("注文書整理ファイル", type=["xlsx"])
file_paid = st.file_uploader("入金一覧ファイル", type=["xlsx"])

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

# ===== 組み合わせ =====
def find_best_combination(records, target):
    records = sorted(records, key=lambda x: x[1], reverse=True)

    best_combo = []
    total = 0

    for pono, amount in records:
        if total + amount <= target:
            best_combo.append((pono, amount))
            total += amount

    return best_combo

# ===== メール送信 =====
def send_mail(file_path, recipient, smtp_user, smtp_pass):

    if not smtp_user or not smtp_pass:
        return False, "メール設定未入力"

    try:
        msg = EmailMessage()
        msg["Subject"] = "支払データ送付"
        msg["From"] = smtp_user
        msg["To"] = recipient

        msg.set_content("支払データを送付します。")

        with open(file_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="xlsx",
                filename=file_path
            )

        with smtplib.SMTP("smtp.sina.net", 587) as smtp:
            smtp.starttls()
            smtp.login(smtp_user, smtp_pass)
            smtp.send_message(msg)

        return True, "成功"

    except Exception as e:
        return False, str(e)

# ===== 計算 =====
if st.button("計算", key="calc_button"):

    if file_orders is None or file_paid is None:
        st.error("ファイル不足")
        st.stop()

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

    file_path = "結果.xlsx"
    df_orders.to_excel(file_path, index=False)

    wb = load_workbook(file_path)
    ws = wb.active

    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row in range(2, ws.max_row+1):
        if ws.cell(row=row, column=10).value:
            for col in range(1, ws.max_column+1):
                ws.cell(row=row, column=col).fill = fill

    wb.save(file_path)

    st.session_state["file_path"] = file_path
    st.session_state["calculated"] = True

# ===== 結果表示 =====
if st.session_state["calculated"]:

    st.success("計算完了")
    st.dataframe(st.session_state["result"])

    st.download_button(
        "Excelダウンロード",
        open(st.session_state["file_path"], "rb"),
        file_name="結果.xlsx"
    )

    if st.button("OK（メール送信）", key="send_button"):

        success, message = send_mail(
            st.session_state["file_path"],
            recipient,
            smtp_user,
            smtp_pass
        )

        if success:
            st.success("メール送信完了")
        else:
            st.error(message)

    if st.button("再計算", key="recalc_button"):
        st.session_state["calculated"] = False
        st.experimental_rerun()
):
     


