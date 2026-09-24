"""
Automated GitHub Releases Update Checker & Manager for SLAM Jobcard Auto-Filler
Checks GitHub Releases for new versions, compares semantic versions, and downloads updates.
"""
import os
import sys
import re
import json
import urllib.request
import urllib.error
import subprocess
import webbrowser
from typing import Optional, Dict, Any, Callable

try:
    from config import APP_VERSION, GITHUB_RELEASES_API, GITHUB_RELEASES_URL
except ImportError:
    APP_VERSION = "1.0.0"
    GITHUB_RELEASES_API = "https://api.github.com/repos/Redpheonix024/slamautofillbot/releases/latest"
    GITHUB_RELEASES_URL = "https://github.com/Redpheonix024/slamautofillbot/releases"

def parse_version(version_str: str) -> tuple:
    """
    Converts a version string (e.g. 'v1.2.3', '1.0.0', 'v2.1-beta') into a tuple of integers.
    Examples:
        'v1.0.0'  -> (1, 0, 0)
        '1.2.3'   -> (1, 2, 3)
        'v2'      -> (2,)
    """
    if not version_str:
        return (0, 0, 0)
    numbers = re.findall(r"\d+", str(version_str))
    if not numbers:
        return (0, 0, 0)
    return tuple(int(n) for n in numbers)

def is_newer_version(latest_str: str, current_str: str) -> bool:
    """
    Compares latest_str against current_str.
    Returns True if latest_str is strictly newer than current_str.
    """
    v_latest = parse_version(latest_str)
    v_current = parse_version(current_str)

    # Pad with zeros to equal length if needed e.g. (1, 2) vs (1, 2, 0)
    max_len = max(len(v_latest), len(v_current))
    v_latest_padded = v_latest + (0,) * (max_len - len(v_latest))
    v_current_padded = v_current + (0,) * (max_len - len(v_current))

    return v_latest_padded > v_current_padded

def check_for_updates(
    current_version: str = APP_VERSION,
    api_url: str = GITHUB_RELEASES_API,
    timeout: int = 6
) -> Dict[str, Any]:
    """
    Queries GitHub API to check if a newer release exists.
    Returns a dict with release metadata and update status.
    """
    result: Dict[str, Any] = {
        "success": False,
        "update_available": False,
        "current_version": current_version,
        "latest_version": current_version,
        "release_name": "",
        "release_notes": "",
        "release_url": GITHUB_RELEASES_URL,
        "download_url": "",
        "asset_name": "",
        "asset_size_mb": 0.0,
        "published_at": "",
        "error": None
    }

    try:
        req = urllib.request.Request(
            api_url,
            headers={
                "User-Agent": f"SLAM-Auto-Filler/{current_version} (Windows)",
                "Accept": "application/vnd.github.v3+json"
            }
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            status = resp.status
            if status != 200:
                result["error"] = f"GitHub API returned HTTP {status}"
                return result

            data = json.loads(resp.read().decode("utf-8"))

        tag_name = data.get("tag_name", "").strip()
        release_name = data.get("name", tag_name).strip()
        body = data.get("body", "").strip()
        html_url = data.get("html_url", GITHUB_RELEASES_URL)
        published_at = data.get("published_at", "")

        result["latest_version"] = tag_name.lstrip("v")
        result["release_name"] = release_name
        result["release_notes"] = body
        result["release_url"] = html_url
        result["published_at"] = published_at

        # Search assets for SLAM_Auto_Filler.exe or any .exe
        assets = data.get("assets", [])
        exe_asset = None
        for a in assets:
            a_name = a.get("name", "").lower()
            if a_name == "slam_auto_filler.exe":
                exe_asset = a
                break
            elif a_name.endswith(".exe") and not exe_asset:
                exe_asset = a

        if exe_asset:
            result["download_url"] = exe_asset.get("browser_download_url", "")
            result["asset_name"] = exe_asset.get("name", "SLAM_Auto_Filler.exe")
            size_b = exe_asset.get("size", 0)
            result["asset_size_mb"] = round(size_b / (1024 * 1024), 2)
        else:
            # Fallback to web release URL if no asset directly attached
            result["download_url"] = html_url

        result["update_available"] = is_newer_version(tag_name, current_version)
        result["success"] = True

    except urllib.error.HTTPError as e:
        if e.code == 404:
            # No releases created yet on GitHub
            result["success"] = True
            result["update_available"] = False
            result["error"] = "No releases found on GitHub repository yet."
        else:
            result["error"] = f"HTTP Error {e.code}: {e.reason}"
    except urllib.error.URLError as e:
        result["error"] = f"Network connection failed: {e.reason}"
    except Exception as e:
        result["error"] = str(e)

    return result

def download_update_asset(
    download_url: str,
    target_path: str,
    progress_callback: Optional[Callable[[int, int, float], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None
) -> bool:
    """
    Downloads the update asset from download_url and saves it to target_path.
    progress_callback(bytes_downloaded, total_bytes, percent_complete)
    cancel_check() -> bool returns True if download should abort.
    """
    req = urllib.request.Request(
        download_url,
        headers={"User-Agent": "SLAM-Auto-Filler-Updater (Windows)"}
    )

    temp_path = target_path + ".tmp"
    try:
        with urllib.request.urlopen(req, timeout=30) as response, open(temp_path, "wb") as out_file:
            total_size = int(response.headers.get("content-length", 0))
            downloaded = 0
            chunk_size = 64 * 1024  # 64 KB chunks

            while True:
                if cancel_check and cancel_check():
                    out_file.close()
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    return False

                chunk = response.read(chunk_size)
                if not chunk:
                    break

                out_file.write(chunk)
                downloaded += len(chunk)

                if progress_callback:
                    pct = (downloaded / total_size * 100.0) if total_size > 0 else 0.0
                    progress_callback(downloaded, total_size, pct)

        # Replace target file with completed download
        if os.path.exists(target_path):
            try:
                os.remove(target_path)
            except Exception:
                pass
        os.replace(temp_path, target_path)
        return True

    except Exception as e:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        raise e

def launch_updated_executable(exe_path: str):
    """
    Launches the new executable in a detached process and exits the current application.
    """
    if not os.path.exists(exe_path):
        raise FileNotFoundError(f"File not found: {exe_path}")

    # Launch detached process on Windows
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen([exe_path], creationflags=DETACHED_PROCESS, close_fds=True)
    sys.exit(0)
