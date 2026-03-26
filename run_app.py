import os
import webbrowser
import subprocess
import time
import sys

# 二重起動防止フラグ
LOCK_FILE = "app.lock"

if not os.path.exists(LOCK_FILE):

    # ロック作成
    with open(LOCK_FILE, "w") as f:
        f.write("running")

    # Streamlit起動
    subprocess.Popen([
        sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", "8501"
    ])

    time.sleep(5)

    webbrowser.open("http://127.0.0.1:8501")
