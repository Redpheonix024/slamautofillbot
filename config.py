"""
SLAM Jobcard Auto-Filler Configuration
"""
import os
import sys
import json
from datetime import datetime

# Application Version & GitHub Releases
APP_VERSION = "1.1.0"
GITHUB_REPO = "Redpheonix024/slamautofillbot"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"

BASE_URL = "https://slam.indianrail.gov.in/SLAM"
LOGIN_URL = f"{BASE_URL}/Login.aspx"
JOBCARD_CREATE_URL = f"{BASE_URL}/JobSchedules/JobcardCreate.aspx"
EQUIPMENT_ADD_URL = f"{BASE_URL}/EquipmentRegister/LocoEquipmentAdd.aspx"

# Loco Equipment Default Google Sheet Settings
DEFAULT_EQ_SHEET_ID = "1hzj67LKBl15xVTc7E43APChoSzMZfNAB"
DEFAULT_EQ_SHEET_GID = "1093714984"

# SLAM Shop Floor Sub-Menu Mapping (parentSubMenu21_SubMenuID92..112)
SHOPFLOOR_MODULES = {
    "Add Loco Eq./Eq. Fitment": "parentSubMenu21_SubMenuID93",
    "Update Loco Equipment": "parentSubMenu5_SubMenuID19",
    "Add Equipment History": "parentSubMenu21_SubMenuID92",
    "Equipment Register": "parentSubMenu21_SubMenuID99",
    "Loco Inventory": "parentSubMenu21_SubMenuID106",
    "Loco Inventory Tree": "parentSubMenu21_SubMenuID107",
    "Section Loco Inventory": "parentSubMenu21_SubMenuID111",
    "Equipment History": "parentSubMenu21_SubMenuID98",
    "Equipment Change Register": "parentSubMenu21_SubMenuID96",
    "Equipment Failure Register": "parentSubMenu21_SubMenuID97",
    "Equipment Removal (bulk)": "parentSubMenu21_SubMenuID100",
    "Equipment Warranty": "parentSubMenu21_SubMenuID103",
    "Find Equipment/Equipment Transfer": "parentSubMenu21_SubMenuID105",
    "Attention Paid On Loco": "parentSubMenu21_SubMenuID94",
    "Minor Equipment Repair": "parentSubMenu21_SubMenuID108",
    "Repair Sheets/Eq.overhaul": "parentSubMenu21_SubMenuID110",
    "Repair Sheet History": "parentSubMenu21_SubMenuID109",
    "Unit Exchange Spares": "parentSubMenu21_SubMenuID112",
    "Cracks Register": "parentSubMenu21_SubMenuID95",
}

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

# All 12 Sections in logical shed sequence
ALL_SECTIONS_LIST = [
    "M1", "M2", "E3", "E3A", "E4", "E5A", "E5B", "BRS", "E8", "QAI", "LAB", "PPIO"
]

# Default rule for 'All Sec': All sections except E8, QAI (Q&I), LAB, PPIO
DEFAULT_ALL_SEC_EXCLUDED = ["E8", "QAI", "LAB", "PPIO"]
DEFAULT_ALL_SEC_INCLUDED = [s for s in ALL_SECTIONS_LIST if s not in DEFAULT_ALL_SEC_EXCLUDED]

def get_all_sec_preferences() -> list:
    """
    Returns the list of sections assigned when 'All Sec' is detected.
    Defaults to all sections except E8, QAI (Q&I), LAB, PPIO.
    """
    s = load_settings()
    custom = s.get("all_sec_included")
    if isinstance(custom, list) and custom:
        valid = [sec for sec in custom if sec in SECTION_IDS]
        if valid:
            return valid
    return list(DEFAULT_ALL_SEC_INCLUDED)

def save_all_sec_preferences(sections: list):
    """Saves user-configured sections for 'All Sec' to persistent settings."""
    s = load_settings()
    s["all_sec_included"] = [sec for sec in sections if sec in SECTION_IDS]
    save_settings(s)


# Browser Automation Settings & Detection
BROWSER_CHOICE_AUTO = "auto"
BROWSER_CHOICE_EDGE = "msedge"
BROWSER_CHOICE_CHROME = "chrome"
BROWSER_CHOICE_BRAVE = "brave"
BROWSER_CHOICE_CHROMIUM = "chromium"
BROWSER_CHOICE_CUSTOM = "custom"

BROWSER_CHOICES = [
    BROWSER_CHOICE_AUTO,
    BROWSER_CHOICE_EDGE,
    BROWSER_CHOICE_CHROME,
    BROWSER_CHOICE_BRAVE,
    BROWSER_CHOICE_CHROMIUM,
    BROWSER_CHOICE_CUSTOM,
]

BROWSER_DISPLAY_NAMES = {
    BROWSER_CHOICE_AUTO: "Auto-Detect / System Default",
    BROWSER_CHOICE_EDGE: "Microsoft Edge",
    BROWSER_CHOICE_CHROME: "Google Chrome",
    BROWSER_CHOICE_BRAVE: "Brave Browser",
    BROWSER_CHOICE_CHROMIUM: "Playwright Chromium",
    BROWSER_CHOICE_CUSTOM: "Custom Browser Executable",
}

def get_browser_preference() -> dict:
    """
    Returns dict:
      {
         'choice': 'auto' | 'msedge' | 'chrome' | 'brave' | 'chromium' | 'custom',
         'custom_path': str
      }
    """
    s = load_settings()
    choice = s.get("browser_choice", BROWSER_CHOICE_AUTO)
    if choice not in BROWSER_DISPLAY_NAMES:
        choice = BROWSER_CHOICE_AUTO
    return {
        "choice": choice,
        "custom_path": s.get("custom_browser_path", "")
    }

def save_browser_preference(choice: str, custom_path: str = ""):
    """Saves browser preferences persistently in user_settings.json."""
    s = load_settings()
    if choice in BROWSER_DISPLAY_NAMES:
        s["browser_choice"] = choice
    if custom_path is not None:
        s["custom_browser_path"] = custom_path.strip()
    save_settings(s)

def detect_system_browsers() -> dict:
    """
    Detects installed Chromium-based browsers on Windows and identifies the system default browser.
    Returns:
        {
            'windows_default': ('chrome' | 'msedge' | 'brave' | 'firefox' | 'unknown', 'Display Name') or None,
            'browsers': {
                'msedge': {'label': 'Microsoft Edge', 'installed': bool, 'path': Optional[str]},
                'chrome': {'label': 'Google Chrome', 'installed': bool, 'path': Optional[str]},
                'brave': {'label': 'Brave Browser', 'installed': bool, 'path': Optional[str]},
            }
        }
    """
    res = {
        "windows_default": None,
        "browsers": {
            "msedge": {"label": "Microsoft Edge", "installed": False, "path": None},
            "chrome": {"label": "Google Chrome", "installed": False, "path": None},
            "brave": {"label": "Brave Browser", "installed": False, "path": None},
        }
    }

    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"
            ) as k:
                prog_id, _ = winreg.QueryValueEx(k, "ProgId")
                prog_lower = str(prog_id).lower()
                if "chrome" in prog_lower:
                    res["windows_default"] = ("chrome", "Google Chrome")
                elif "edge" in prog_lower or "msedge" in prog_lower:
                    res["windows_default"] = ("msedge", "Microsoft Edge")
                elif "brave" in prog_lower:
                    res["windows_default"] = ("brave", "Brave Browser")
                elif "firefox" in prog_lower:
                    res["windows_default"] = ("firefox", "Mozilla Firefox")
                else:
                    res["windows_default"] = ("unknown", str(prog_id))
        except Exception:
            pass

        candidates = {
            "msedge": [
                r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
                r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
                r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe",
            ],
            "chrome": [
                r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
                r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
                r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
            ],
            "brave": [
                r"%ProgramFiles%\BraveSoftware\Brave-Browser\Application\brave.exe",
                r"%ProgramFiles(x86)%\BraveSoftware\Brave-Browser\Application\brave.exe",
                r"%LocalAppData%\BraveSoftware\Brave-Browser\Application\brave.exe",
            ]
        }

        for key, paths in candidates.items():
            found = None
            for p in paths:
                expanded = os.path.expandvars(p)
                if os.path.isfile(expanded):
                    found = expanded
                    break
            res["browsers"][key]["installed"] = (found is not None)
            res["browsers"][key]["path"] = found

    return res


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
