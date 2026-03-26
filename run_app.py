import os
import webbrowser
import subprocess
import time

os.environ["SMTP_USER"] = "あなたのメール"
os.environ["SMTP_PASS"] = "あなたのパスワード"

subprocess.Popen([
    "streamlit", "run", "app.py", "--server.port", "8501"
])

time.sleep(5)

webbrowser.open("http://localhost:8501")
