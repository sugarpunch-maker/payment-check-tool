import subprocess
import sys
import webbrowser
import time

# Streamlit起動
subprocess.Popen([
    sys.executable, "-m", "streamlit", "run", "app.py",
    "--server.port=8501",
    "--server.headless=true"
])

# 少し待つ
time.sleep(3)

# ブラウザ起動（1回だけ）
webbrowser.open("http://localhost:8501")

