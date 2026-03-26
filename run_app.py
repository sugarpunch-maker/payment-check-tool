import webbrowser
import subprocess
import time
import sys
import os

base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
app_path = os.path.join(base_dir, "app.py")

# Streamlit起動
subprocess.Popen([
    sys.executable, "-m", "streamlit", "run", app_path, "--server.port", "8501"
])

# 起動待ち
time.sleep(5)

# 1回だけブラウザ起動
webbrowser.open("http://127.0.0.1:8501")
