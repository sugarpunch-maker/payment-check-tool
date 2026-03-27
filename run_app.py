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
    return os.path.join(os.path.dirname(__file__), "app.py")

def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

if __name__ == "__main__":

    app_path = get_app_path()

    # Streamlit起動（ログを出す）
    process = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", app_path,
         "--server.port=8501",
         "--server.headless=true"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )

    # 起動確認（最大30秒待つ）
    started = False
    for _ in range(30):
        time.sleep(1)
        if is_port_open(PORT):
            started = True
            break

    if started:
        webbrowser.open(f"http://localhost:{PORT}")
    else:
        print("Streamlitが起動していません")
        print("ログ出力:")
        try:
            out, err = process.communicate(timeout=5)
            print(out)
            print(err)
        except:
            pass




