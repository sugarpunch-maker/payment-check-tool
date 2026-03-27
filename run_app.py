import subprocess
import webbrowser
import time
import socket

PORT = 8501

def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

# Streamlit起動（ローカル環境）
subprocess.Popen([
    "streamlit",
    "run",
    "app.py",
    "--server.port=8501"
])

# 起動待ち
for _ in range(15):
    time.sleep(1)
    if is_port_open(PORT):
        break

# ブラウザ起動
webbrowser.open(f"http://localhost:{PORT}")


