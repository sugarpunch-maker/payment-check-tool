import subprocess
import sys
import time
import webbrowser
import socket
import os

PORT = 8501

# ===============================
# app.py のパス取得（exe対応）
# ===============================
def get_app_path():
    if getattr(sys, 'frozen', False):
        # exe化された場合
        return os.path.join(sys._MEIPASS, "app.py")
    else:
        # 通常実行
        return os.path.join(os.path.dirname(__file__), "app.py")

# ===============================
# ポート確認
# ===============================
def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

# ===============================
# メイン処理
# ===============================
if __name__ == "__main__":

    app_path = get_app_path()

    # すでに起動している場合はスキップ
    if not is_port_open(PORT):

        subprocess.Popen([
            sys.executable,
            "-m",
            "streamlit",
            "run",
            app_path,
            "--server.port=8501",
            "--server.headless=true",
            "--browser.gatherUsageStats=false"
        ])

        # 起動待機（最大10秒）
        for _ in range(10):
            time.sleep(1)
            if is_port_open(PORT):
                break

    # ブラウザは1回だけ開く
    webbrowser.open(f"http://localhost:{PORT}")

