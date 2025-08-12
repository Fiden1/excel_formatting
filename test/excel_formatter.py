import os
import requests
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox

APP_NAME = "Excel Formatter"

if getattr(sys, 'frozen', False):
    INSTALL_DIR = os.path.dirname(sys.executable)
else:
    INSTALL_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_DIR = os.path.join(INSTALL_DIR, "config")
VERSION_FILE = os.path.join(CONFIG_DIR, "version.txt")



VERSION_JSON_URL = "https://raw.githubusercontent.com/fiden1/excel_formatting/main/test/config/version.json"

def get_local_version():
    if os.path.exists(VERSION_FILE):
        return open(VERSION_FILE).read().strip()
    return "0.0.0"


def save_local_version(version):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(VERSION_FILE, "w") as f:
        f.write(version)

def get_online_version():
    try:
        response = requests.get(VERSION_JSON_URL, timeout=5)
        response.raise_for_status()
        data = response.json()
        return data["version"], data["download_url"]
    except Exception as e:
        messagebox.showerror("Update Error", f"Failed to check for updates:\n{e}")
        return None, None

def download_with_progress(url, dest, progress_bar, status_label):
    try:
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            total_length = int(r.headers.get('content-length', 0))
            downloaded = 0
            with open(dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total_length:
                            percent = (downloaded / total_length) * 100
                            progress_bar["value"] = percent
                            status_label.config(text=f"Downloading update... {percent:.1f}%")
                            root.update_idletasks()
        status_label.config(text="Update downloaded successfully.")
    except Exception as e:
        messagebox.showerror("Download Error", f"Failed to download update:\n{e}")
        sys.exit(1)

def save_local_version(version):
    with open(VERSION_FILE, "w") as f:
        f.write(version)

def run_app():
    app_path = os.path.join(INSTALL_DIR, "app.exe")
    if not os.path.exists(app_path):
        messagebox.showerror("Error", "app.exe not found!")
        sys.exit(1)
    subprocess.Popen([app_path])
    sys.exit(0)

def update_and_launch(local_ver, online_ver, download_url):
    global root
    root = tk.Tk()
    root.title(f"{APP_NAME} Updater")
    root.geometry("400x150")
    root.resizable(False, False)

    tk.Label(root, text=f"Updating {APP_NAME}", font=("Arial", 14, "bold")).pack(pady=10)
    status_label = tk.Label(root, text=f"Updating from {local_ver} to {online_ver}...", font=("Arial", 10))
    status_label.pack(pady=5)

    progress_bar = ttk.Progressbar(root, length=300, mode="determinate")
    progress_bar.pack(pady=5)

    root.update()
    exe_path = os.path.join(INSTALL_DIR, "app.exe")
    download_with_progress(download_url, exe_path, progress_bar, status_label)
    save_local_version(online_ver)

    root.destroy()
    run_app()

def main():
    local_ver = get_local_version()
    online_ver, download_url = get_online_version()

    if online_ver and download_url:
        if online_ver != local_ver:
            update_and_launch(local_ver, online_ver, download_url)
        else:
            run_app()
    else:
        run_app()

if __name__ == "__main__":
    main()
