from __future__ import annotations

import os
import socket
import subprocess
import sys
import threading
import time
import tempfile
import urllib.request
import webbrowser
import ctypes
import traceback
from pathlib import Path

APP_NAME = "Revisor de Bitácoras"
SERVER_FLAG = "--run-server"


def resource_path(relative: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base / relative


def available_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def log_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "RevisaBitacoras"
    try:
        base.mkdir(parents=True, exist_ok=True)
    except OSError:
        base = Path(tempfile.gettempdir()) / "RevisaBitacoras"
        base.mkdir(parents=True, exist_ok=True)
    return base / "app.log"


def write_log(message: str) -> None:
    with log_path().open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")


def run_server(port: int) -> None:
    from streamlit.web import bootstrap

    app_path = resource_path("app.py")
    os.chdir(resource_path("."))
    write_log(f"Iniciando servidor en el puerto {port}; app={app_path}")
    flag_options = {
        "global_developmentMode": False,
        "server_address": "127.0.0.1",
        "server_port": port,
        "server_headless": True,
        "server_fileWatcherType": "none",
        "browser_gatherUsageStats": False,
        "server_maxUploadSize": 500,
    }
    bootstrap.load_config_options(flag_options=flag_options)
    bootstrap.run(
        str(app_path),
        False,
        [],
        flag_options,
    )


def child_command(port: int) -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, SERVER_FLAG, str(port)]
    return [sys.executable, str(Path(__file__).resolve()), SERVER_FLAG, str(port)]


def wait_until_ready(url: str, process: subprocess.Popen[bytes]) -> bool:
    health_url = f"{url}/_stcore/health"
    for _ in range(120):
        if process.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(health_url, timeout=1) as response:
                if response.status == 200:
                    return True
        except Exception:
            time.sleep(0.25)
    return False


def tray_image():
    from PIL import Image, ImageDraw

    image = Image.new("RGBA", (64, 64), "#0B6E4F")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((3, 3, 61, 61), radius=14, fill="#0B6E4F")
    draw.line((17, 33, 28, 44, 49, 20), fill="white", width=8, joint="curve")
    return image


def show_error(message: str) -> None:
    ctypes.windll.user32.MessageBoxW(None, message, APP_NAME, 0x10)


def launch_desktop() -> None:
    import pystray

    port = available_port()
    url = f"http://127.0.0.1:{port}"
    log_file = log_path().open("ab")
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        child_command(port),
        stdout=log_file,
        stderr=subprocess.STDOUT,
        creationflags=creation_flags,
    )
    closing = threading.Event()

    def stop_process() -> None:
        if closing.is_set():
            return
        closing.set()
        try:
            if process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
        finally:
            log_file.close()

    def open_app(icon=None, item=None) -> None:
        webbrowser.open(url)

    def exit_app(icon, item=None) -> None:
        stop_process()
        icon.stop()

    icon = pystray.Icon(
        "RevisaBitacoras",
        tray_image(),
        APP_NAME,
        menu=pystray.Menu(
            pystray.MenuItem("Abrir en el navegador", open_app, default=True),
            pystray.MenuItem("Cerrar aplicación", exit_app),
        ),
    )

    def start_and_monitor() -> None:
        ready = wait_until_ready(url, process)
        if ready:
            open_app()
            try:
                icon.notify(
                    "La aplicación está activa. Usa este icono para abrirla o cerrarla.",
                    APP_NAME,
                )
            except Exception:
                pass
            while not closing.is_set() and process.poll() is None:
                time.sleep(1)
            if not closing.is_set():
                show_error(
                    f"La aplicación se detuvo inesperadamente.\n\nRegistro:\n{log_path()}"
                )
                stop_process()
                icon.stop()
        else:
            show_error(
                f"No fue posible iniciar el sistema.\n\nRevisa el registro:\n{log_path()}"
            )
            stop_process()
            icon.stop()

    threading.Thread(target=start_and_monitor, daemon=True).start()
    try:
        icon.run()
    finally:
        stop_process()


def main() -> None:
    try:
        write_log(f"Arranque: argumentos={sys.argv[1:]}")
        if len(sys.argv) == 3 and sys.argv[1] == SERVER_FLAG:
            run_server(int(sys.argv[2]))
        else:
            launch_desktop()
    except Exception:
        details = traceback.format_exc()
        write_log(details)
        if SERVER_FLAG not in sys.argv:
            show_error(
                f"No fue posible iniciar el sistema.\n\nRevisa el registro:\n{log_path()}"
            )
        raise


if __name__ == "__main__":
    main()
