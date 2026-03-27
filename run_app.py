import subprocess
import sys
import time
import webbrowser
import socket
import os

PORT = 8501

def get_app_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "app.py")
    return "app.py"

def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

if __name__ == "__main__":

    app_path = get_app_path()

    # Streamlit起動
    process = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", app_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )

    # 起動待機
    for _ in range(15):
        time.sleep(1)
        if is_port_open(PORT):
            webbrowser.open(f"http://localhost:{PORT}")
            break


