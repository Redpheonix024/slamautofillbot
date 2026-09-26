"""
SLAM Jobcard Auto-Filler Configuration
"""
import os
import sys
import json
from datetime import datetime

# Application Version & GitHub Releases
APP_VERSION = "1.0.4"
GITHUB_REPO = "Redpheonix024/slamautofillbot"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"

BASE_URL = "https://slam.indianrail.gov.in/SLAM"
LOGIN_URL = f"{BASE_URL}/Login.aspx"
JOBCARD_CREATE_URL = f"{BASE_URL}/JobSchedules/JobcardCreate.aspx"

def get_settings_file_path() -> str:
    """
    Returns the persistent path for user_settings.json.
    When running as a compiled .exe, it stores it next to the executable,
    or in %APPDATA%\\SLAMAutoFiller\\ if the directory is read-only.
    """
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        test_file = os.path.join(exe_dir, ".test_write.tmp")
        try:
            with open(test_file, "w") as f:
                f.write("ok")
            os.remove(test_file)
            return os.path.join(exe_dir, "user_settings.json")
        except Exception:
            appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
            folder = os.path.join(appdata, "SLAMAutoFiller")
            os.makedirs(folder, exist_ok=True)
            return os.path.join(folder, "user_settings.json")
    else:
        return os.path.join(os.path.dirname(__file__), "user_settings.json")

SETTINGS_FILE = get_settings_file_path()

def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_settings(settings: dict):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving settings: {e}")

# Default credentials (can be overridden via user_settings.json or environment variables)
DEFAULT_USERNAME = os.getenv("SLAM_USERNAME", "manickarajap")
DEFAULT_PASSWORD = os.getenv("SLAM_PASSWORD", "SSE!1ppio")

def get_credentials():
    s = load_settings()
    username = s.get("username") or DEFAULT_USERNAME
    password = s.get("password") or DEFAULT_PASSWORD
    return username, password

def save_credentials(username: str, password: str):
    s = load_settings()
    s["username"] = username.strip()
    s["password"] = password.strip()
    save_settings(s)

def get_daily_otp() -> str:
    """Returns today's cached OTP if recorded today, else empty string."""
    s = load_settings()
    today_str = datetime.now().strftime("%Y-%m-%d")
    if s.get("otp_date") == today_str and s.get("otp"):
        return str(s.get("otp")).strip()
    return ""

def set_daily_otp(otp: str):
    """Saves OTP with today's date so it can be reused for the entire day."""
    s = load_settings()
    s["otp"] = str(otp).strip()
    s["otp_date"] = datetime.now().strftime("%Y-%m-%d")
    save_settings(s)

# Valid SLAM Section Codes and their internal DOM checkbox IDs
SECTION_IDS = {
    "BRS": 58,
    "E3": 46,
    "E3A": 48,
    "E4": 49,
    "E5A": 50,
    "E5B": 51,
    "E8": 53,
    "LAB": 55,
    "M1": 56,
    "M2": 57,
    "PPIO": 60,
    "QAI": 45
}

# Jobcard Form Default Settings
DEFAULT_JOBCARD_TYPE = "Schedule Attention" # value "2"
DEFAULT_EQT_REQUIRED = "Jobcard Without Eq" # value "2"
DEFAULT_OBS_TYPE = "Testing Remarks"        # value "2"
DEFAULT_JOB_KIND = "Minor"                  # value "2"

JOBCARD_TYPE_VALUES = {
    "Inspection/Testing": "1",
    "Schedule Attention": "2",
    "Post Schedule Attention": "3",
    "Repair Attention": "4",
    "Shunting": "5",
    "Loco Lifting": "6",
    "Wheel Turning": "7",
    "Loco Trail Run": "8",
    "Out Station Repair Attention": "9",
    "QAC": "10",
    "Water Leakage": "11"
}

OBS_TYPE_VALUES = {
    "Driver Remarks": "1",
    "Testing Remarks": "2",
    "Officer Remarks": "3",
    "Special Instructions": "4",
    "Investigation Remarks": "5",
    "Seasonal Remarks": "6",
    "Before Test Remarks": "7",
    "Super Check": "11",
    "Out Station Remarks": "13",
    "DDS Remarks": "14"
}

JOB_KIND_VALUES = {
    "Major": "1",
    "Minor": "2"
}
