"""
updater.py — Self-update mechanism for Sales Analyzer.

Checks GitHub Releases API, downloads new .exe, and launches an apply-update
batch script that replaces the current executable without blocking the UI.
"""

import os
import sys
import subprocess
import threading
import logging
import time
from pathlib import Path
from typing import Optional, Callable

import requests
from packaging.version import Version

log = logging.getLogger(__name__)

GITHUB_REPO = "iulian-mormeci/Sales_Analyzer"
GITHUB_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))


def get_local_version() -> str:
    try:
        return (BASE_DIR / "VERSION").read_text().strip()
    except Exception:
        return "0.0.0"


def check_for_update(
    on_update_available: Callable[[str, str], None],
    on_error: Optional[Callable[[str], None]] = None,
) -> None:
    """
    Runs in a background thread.
    Calls on_update_available(latest_version, download_url) if a newer release exists.
    """
    def _run():
        try:
            resp = requests.get(GITHUB_API, timeout=8)
            try:
                resp.raise_for_status()
            except requests.HTTPError as http_exc:
                if resp.status_code == 404:
                    log.info("Nessuna release trovata su GitHub, skip update check.")
                else:
                    log.warning(f"Update check HTTP error: {http_exc}")
                    if on_error:
                        on_error(str(http_exc))
                return

            data = resp.json()

            tag = data.get("tag_name", "")
            remote_ver = tag.lstrip("v").strip()
            local_ver = get_local_version()

            log.info(f"Update check — local={local_ver} remote={remote_ver}")

            if not remote_ver:
                return

            if Version(remote_ver) > Version(local_ver):
                # Find .exe asset
                assets = data.get("assets", [])
                url = ""
                for asset in assets:
                    name = asset.get("name", "")
                    if name.lower().endswith(".exe"):
                        url = asset.get("browser_download_url", "")
                        break

                on_update_available(remote_ver, url)
        except (requests.ConnectionError, requests.Timeout):
            pass  # No internet — silently ignore
        except Exception as exc:
            log.warning(f"Update check failed: {exc}")
            if on_error:
                on_error(str(exc))

    threading.Thread(target=_run, daemon=True).start()


def download_update(
    download_url: str,
    on_progress: Callable[[int, int], None],
    on_complete: Callable[[Path], None],
    on_error: Callable[[str], None],
) -> None:
    """
    Downloads the new .exe in a background thread.
    on_progress(bytes_downloaded, total_bytes)
    on_complete(path_to_new_exe)
    """
    def _run():
        try:
            dest = _get_new_exe_path()
            log.info(f"Downloading update from {download_url} → {dest}")

            resp = requests.get(download_url, stream=True, timeout=60)
            resp.raise_for_status()

            total = int(resp.headers.get("content-length", 0))
            downloaded = 0

            with open(dest, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        fh.write(chunk)
                        downloaded += len(chunk)
                        on_progress(downloaded, total)

            log.info(f"Download complete: {dest}")
            on_complete(dest)

        except Exception as exc:
            log.error(f"Download failed: {exc}")
            on_error(str(exc))

    threading.Thread(target=_run, daemon=True).start()


def apply_update(new_exe_path: Path) -> None:
    """
    Writes apply_update.bat, launches it, then exits the current app.
    The batch script waits 2 s, replaces the old .exe, launches the new one,
    and deletes itself.
    """
    current_exe = Path(sys.executable)
    if not getattr(sys, "frozen", False):
        # Running as plain Python — no real .exe to replace
        log.info("Not running as frozen exe; skipping apply-update batch.")
        return

    bat_path = current_exe.parent / "apply_update.bat"
    bat_content = f"""@echo off
timeout /t 2 /nobreak >nul
move /Y "{new_exe_path}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
"""
    bat_path.write_text(bat_content, encoding="utf-8")
    log.info(f"Launching {bat_path}")
    subprocess.Popen(["cmd", "/c", str(bat_path)], creationflags=subprocess.CREATE_NO_WINDOW)
    sys.exit(0)


def _get_new_exe_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "SalesAnalyzer_new.exe"
    return BASE_DIR / "SalesAnalyzer_new.exe"
