import webbrowser
import subprocess
import time
import sys
import os

base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
app_path = os.path.join(base_dir, "app.py")

subprocess.Popen([
    sys.executable, "-m", "streamlit", "run", app_path
])

time.sleep(5)
webbrowser.open("http://localhost:8501")
