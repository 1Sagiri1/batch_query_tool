import atexit
import socket
import subprocess
import sys
import time
from pathlib import Path

import webview


BASE_DIR = Path(__file__).resolve().parent
BRIDGE_HOST = "127.0.0.1"
BRIDGE_PORT = 8765


def is_port_open(host: str, port: int, timeout: float = 0.4) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def wait_bridge_ready(host: str, port: int, timeout_sec: float = 12.0) -> bool:
    start = time.time()
    while time.time() - start <= timeout_sec:
        if is_port_open(host, port):
            return True
        time.sleep(0.2)
    return False


def stop_proc(proc: subprocess.Popen | None) -> None:
    if not proc:
        return
    if proc.poll() is not None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=3)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def start_bridge() -> subprocess.Popen:
    cmd = [sys.executable, str(BASE_DIR / "src" / "bridge_server.py")]
    proc = subprocess.Popen(
        cmd,
        cwd=str(BASE_DIR),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if not wait_bridge_ready(BRIDGE_HOST, BRIDGE_PORT):
        stop_proc(proc)
        raise RuntimeError("桥接服务启动失败，请检查 Python 环境与依赖。")
    return proc


def main() -> None:
    bridge_proc = start_bridge()
    atexit.register(lambda: stop_proc(bridge_proc))

    html_path = (BASE_DIR / "index.html").resolve()
    html_uri = html_path.as_uri()
    window = webview.create_window("批量查询工作台", html_uri, width=1320, height=900)
    window.events.closed += lambda: stop_proc(bridge_proc)
    webview.start(debug=False)


if __name__ == "__main__":
    main()
