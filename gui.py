"""
SLAM Jobcard Auto-Filler - Desktop GUI Application
Tailored for Indian Railways ELS/RPME Royapuram Shed
Supports:
  - Preset Settings: Schedule Attention | Jobcard Without Eq | Testing Remarks | Minor
  - Account Credentials (Username & Password) Editing & Saving
  - Daily OTP Management (Cached once per day so no repeated prompts)
  - Desktop Excel Auto-Detection & Quick Picker
  - Section Breakdown & Validation
  - Live Browser Staging Progress
  - Mandatory Interactive Verification Dialog before Saving
"""
import os
import sys

# Register Windows AppUserModelID at module level before any Tkinter windows are created
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("IndianRailways.SLAM.JobcardAutoFiller.v1.1.Equipment")
    except Exception:
        pass

import threading
import queue
import collections
from datetime import datetime
from typing import Optional, Any, List, Dict, Set
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
import tempfile

import config
import updater
import excel_parser
from excel_parser import parse_excel_bookings
from slam_bot import SLAMBot
from PIL import Image, ImageTk


def get_base_dir() -> str:
    """Returns the application working/base directory."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(filename: str) -> str:
    """
    Returns the absolute path to a resource/asset file.
    Handles running in dev mode as well as PyInstaller frozen onefile mode (_MEIPASS).
    """
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", "")
        p = os.path.join(base, filename)
        if os.path.exists(p):
            return p
        p_exe = os.path.join(os.path.dirname(sys.executable), filename)
        if os.path.exists(p_exe):
            return p_exe
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def force_bring_to_front(win):
    """
    Forcefully brings a Tkinter window or dialog to the foreground,
    bypassing Windows OS Foreground Lockout restrictions by attaching to
    the active foreground thread and pulsing window z-order.
    Works reliably across Windows 7, 8, 10, 11, and terminal environments.
    """
    if not win:
        return
    try:
        win.update_idletasks()
        if hasattr(win, "state") and win.state() == "iconic":
            win.deiconify()
        win.lift()

        if sys.platform == "win32":
            try:
                import ctypes
                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32

                wid = win.winfo_id()
                hwnd = user32.GetAncestor(wid, 2) or user32.GetParent(wid) or wid

                # If minimized at OS level, restore it
                SW_RESTORE = 9
                if user32.IsIconic(hwnd):
                    user32.ShowWindow(hwnd, SW_RESTORE)

                # ASFW_ANY (-1) allows any process to set foreground window
                user32.AllowSetForegroundWindow(-1)

                fg_hwnd = user32.GetForegroundWindow()
                if fg_hwnd and fg_hwnd != hwnd:
                    fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)
                    cur_tid = kernel32.GetCurrentThreadId()
                    attached = False
                    if fg_tid != cur_tid and fg_tid != 0:
                        attached = bool(user32.AttachThreadInput(cur_tid, fg_tid, True))

                    user32.BringWindowToTop(hwnd)
                    user32.SetForegroundWindow(hwnd)

                    if attached:
                        user32.AttachThreadInput(cur_tid, fg_tid, False)
                else:
                    user32.BringWindowToTop(hwnd)
                    user32.SetForegroundWindow(hwnd)

                # Momentary Z-order pulse to pop in front of browsers without locking permanently topmost
                HWND_TOPMOST = -1
                HWND_NOTOPMOST = -2
                SWP_NOMOVE = 0x0002
                SWP_NOSIZE = 0x0001
                SWP_SHOWWINDOW = 0x0040
                user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
                user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
            except Exception:
                pass

        try:
            win.focus_force()
        except Exception:
            pass
    except Exception:
        pass


def notify_user_attention(win=None):
    """Brings the window to the front and sounds an alert chime."""
    if win:
        force_bring_to_front(win)
    if sys.platform == "win32":
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except Exception:
            pass


class UpdateDialog(tk.Toplevel):
    """
    Dedicated dialog for reviewing release details, changelog,
    and downloading/installing the latest version.
    """
    def __init__(self, parent, update_info: dict, gui_ref: Optional[Any] = None):
        super().__init__(parent)
        self.gui_ref = gui_ref
        if self.gui_ref:
            self.gui_ref._active_modal = self
        self.transient(parent)
        latest_v = update_info.get("latest_version", "--")
        self.title(f"🚀 SLAM Auto-Filler Update Available - v{latest_v}")
        self.geometry("660x520")
        self.minsize(580, 440)
        self.update_info = update_info
        self._download_thread = None
        self._cancel_download = False
        self._downloaded_file = ""

        self._build_ui()

        # Center on parent window
        self.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self.grab_set()
        force_bring_to_front(self)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        banner = tk.Frame(self, bg="#1a365d", height=60)
        banner.pack(fill=tk.X, side=tk.TOP)

        latest_v = self.update_info.get("latest_version", "--")
        curr_v = self.update_info.get("current_version", config.APP_VERSION)
        rel_name = self.update_info.get("release_name", f"Release v{latest_v}")

        tk.Label(
            banner,
            text=f"🚀 New Update Available: v{latest_v}",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#1a365d"
        ).pack(anchor=tk.W, padx=16, pady=(8, 2))

        tk.Label(
            banner,
            text=f"Current installed: v{curr_v}  •  {rel_name}",
            font=("Segoe UI", 9),
            fg="#90cdf4",
            bg="#1a365d"
        ).pack(anchor=tk.W, padx=16, pady=(0, 8))

        content = ttk.Frame(self, padding=12)
        content.pack(fill=tk.BOTH, expand=True)

        # Release Notes / Changelog
        notes_frame = ttk.LabelFrame(content, text=" Release Notes & Changelog ", padding=8)
        notes_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        notes_txt = tk.Text(notes_frame, wrap=tk.WORD, font=("Segoe UI", 9), height=10, relief=tk.SOLID, borderwidth=1)
        notes_sb = ttk.Scrollbar(notes_frame, orient=tk.VERTICAL, command=notes_txt.yview)
        notes_txt.configure(yscrollcommand=notes_sb.set)

        notes_body = self.update_info.get("release_notes", "") or "No release notes provided."
        notes_txt.insert(tk.END, notes_body)
        notes_txt.config(state=tk.DISABLED)

        notes_sb.pack(side=tk.RIGHT, fill=tk.Y)
        notes_txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Download / Asset Status Frame
        status_frame = ttk.LabelFrame(content, text=" Download Status ", padding=8)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        asset_name = self.update_info.get("asset_name", "SLAM_Auto_Filler.exe")
        asset_size = self.update_info.get("asset_size_mb", 0.0)
        size_str = f"({asset_size:.1f} MB)" if asset_size > 0 else ""

        self.asset_lbl = ttk.Label(
            status_frame,
            text=f"Package: {asset_name} {size_str}",
            font=("Segoe UI", 9, "bold")
        )
        self.asset_lbl.pack(anchor=tk.W, pady=(0, 4))

        self.progress_bar = ttk.Progressbar(status_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 4))

        self.status_detail_lbl = ttk.Label(
            status_frame,
            text="Ready to download and install.",
            font=("Segoe UI", 8)
        )
        self.status_detail_lbl.pack(anchor=tk.W)

        # Buttons Frame
        btn_frame = ttk.Frame(content)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.action_btn = ttk.Button(
            btn_frame,
            text="📥 Download & Update",
            command=self._start_download
        )
        self.action_btn.pack(side=tk.LEFT, padx=(0, 8))

        github_btn = ttk.Button(
            btn_frame,
            text="🌐 View on GitHub",
            command=lambda: webbrowser.open(self.update_info.get("release_url", config.GITHUB_RELEASES_URL))
        )
        github_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.close_btn = ttk.Button(
            btn_frame,
            text="Cancel",
            command=self._on_close
        )
        self.close_btn.pack(side=tk.RIGHT)

    def _start_download(self):
        download_url = self.update_info.get("download_url")
        if not download_url or not download_url.startswith("http"):
            webbrowser.open(self.update_info.get("release_url", config.GITHUB_RELEASES_URL))
            return

        latest_v = self.update_info.get("latest_version", "new")
        current_exe = updater.get_current_executable_path()
        self._target_exe = current_exe

        # Stage the update file next to the current exe, or in %TEMP% if directory is read-only
        staged_file = current_exe + ".update.exe"
        try:
            with open(staged_file, "wb") as f:
                pass
            os.remove(staged_file)
        except Exception:
            staged_file = os.path.join(tempfile.gettempdir(), f"SLAM_Auto_Filler_v{latest_v}.update.exe")

        self._downloaded_file = staged_file

        self.action_btn.config(state=tk.DISABLED, text="Downloading...")
        self.progress_bar["value"] = 0
        self.status_detail_lbl.config(text="Downloading update package from GitHub...")

        def _do_download():
            try:
                def _prog(dl, total, pct):
                    def _update_ui():
                        self.progress_bar["value"] = pct
                        mb_dl = dl / (1024 * 1024)
                        mb_tot = total / (1024 * 1024)
                        self.status_detail_lbl.config(
                            text=f"Downloading: {mb_dl:.1f} MB / {mb_tot:.1f} MB ({pct:.1f}%)"
                        )
                    self.after(0, _update_ui)

                success = updater.download_update_asset(
                    download_url,
                    staged_file,
                    progress_callback=_prog,
                    cancel_check=lambda: self._cancel_download
                )

                if success:
                    def _on_success():
                        self.progress_bar["value"] = 100
                        exe_name = os.path.basename(self._target_exe)
                        self.status_detail_lbl.config(
                            text=f"✅ Download complete!\nReady to replace {exe_name} in-place with v{latest_v}."
                        )
                        self.action_btn.config(
                            state=tk.NORMAL,
                            text="🚀 Replace & Restart Now",
                            command=self._launch_downloaded
                        )
                        self.close_btn.config(text="Later")
                    self.after(0, _on_success)
                else:
                    def _on_abort():
                        self.status_detail_lbl.config(text="Download cancelled.")
                        self.action_btn.config(state=tk.NORMAL, text="📥 Download & Update")
                    self.after(0, _on_abort)

            except Exception as e:
                def _on_err(err_str=str(e)):
                    self.status_detail_lbl.config(text=f"❌ Download failed: {err_str}")
                    self.action_btn.config(state=tk.NORMAL, text="Retry Download")
                self.after(0, _on_err)

        self._download_thread = threading.Thread(target=_do_download, daemon=True)
        self._download_thread.start()

    def _launch_downloaded(self):
        if self._downloaded_file and os.path.exists(self._downloaded_file):
            try:
                # Stop worker before launching updater so background threads don't linger
                if hasattr(self, "gui_ref") and self.gui_ref and hasattr(self.gui_ref, "worker"):
                    try:
                        self.gui_ref.worker.stop_worker()
                    except Exception:
                        pass
                updater.replace_and_restart(self._downloaded_file, self._target_exe)
            except Exception as e:
                messagebox.showerror("Update Error", f"Could not replace application:\n{e}")

    def _on_close(self):
        self._cancel_download = True
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

class VerificationDialog(tk.Toplevel):
    """
    Dedicated Verification Modal that pops up once all observations are staged in SLAM.
    Shows the staged table, section breakdown, and requires explicit user confirmation
    to either 'Save & Submit to Section' or 'Stop / Keep Staged'.
    """
    def __init__(self, parent, staged_rows, meta, gui_ref: Optional[Any] = None):
        super().__init__(parent)
        self.gui_ref = gui_ref
        if self.gui_ref:
            self.gui_ref._active_modal = self
        self.transient(parent)
        self.title(f"⚠️ Review & Verify Before Saving - Loco {meta.get('loco_number', '--')} ({meta.get('schedule', '--')})")
        self.geometry("940x630")
        self.minsize(820, 520)
        self.result = False

        self._build_ui(staged_rows, meta)

        # Center on parent window
        self.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self.grab_set()
        force_bring_to_front(self)
        self.protocol("WM_DELETE_WINDOW", self._on_stop)

    def _build_ui(self, staged_rows, meta):
        # 1. Header Banner
        banner = tk.Frame(self, bg="#2b6cb0", height=50)
        banner.pack(fill=tk.X, side=tk.TOP)

        tk.Label(
            banner,
            text="📋 VERIFICATION STAGE: Staged Observations in SLAM Portal",
            font=("Segoe UI", 13, "bold"),
            fg="white",
            bg="#2b6cb0"
        ).pack(side=tk.LEFT, padx=16, pady=10)

        tk.Label(
            banner,
            text="Please review carefully before saving to sections",
            font=("Segoe UI", 9, "italic"),
            fg="#bee3f8",
            bg="#2b6cb0"
        ).pack(side=tk.RIGHT, padx=16, pady=12)

        # Main content container
        content = ttk.Frame(self, padding=12)
        content.pack(fill=tk.BOTH, expand=True)

        # 2. Metadata Info Card
        info_frame = ttk.LabelFrame(content, text=" ℹ️ Locomotive & Form Configuration ", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        row1 = ttk.Frame(info_frame)
        row1.pack(fill=tk.X, pady=2)

        ttk.Label(row1, text="Locomotive:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(row1, text=meta.get("loco_number", "--"), font=("Segoe UI", 10, "bold"), fg="#2b6cb0").pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row1, text="Schedule:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(row1, text=meta.get("schedule", "--"), font=("Segoe UI", 9, "bold"), fg="#2f855a").pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row1, text="Date:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row1, text=meta.get("date", "--"), font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(row1, text="Total Staged Rows:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(row1, text=str(len(staged_rows)), font=("Segoe UI", 10, "bold"), fg="#d69e2e").pack(side=tk.LEFT)

        row2 = ttk.Frame(info_frame)
        row2.pack(fill=tk.X, pady=4)

        ttk.Label(row2, text="Jobcard Type:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row2, text=meta.get("jobcard_type", config.DEFAULT_JOBCARD_TYPE), font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Eqt Required:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row2, text=meta.get("eqt_required", config.DEFAULT_EQT_REQUIRED), font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Source:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row2, text=meta.get("obs_type", config.DEFAULT_OBS_TYPE), font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="Kind:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(row2, text=meta.get("job_kind", config.DEFAULT_JOB_KIND), font=("Segoe UI", 8)).pack(side=tk.LEFT)

        # Section-wise breakdown count
        sec_counts = collections.Counter(r.get("section", "") for r in staged_rows if r.get("section"))
        if sec_counts:
            sec_text = "  |  ".join(f"{s}: {c}" for s, c in sec_counts.most_common())
            row3 = ttk.Frame(info_frame)
            row3.pack(fill=tk.X, pady=(4, 0))
            ttk.Label(row3, text="Section Counts:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
            tk.Label(row3, text=sec_text, font=("Segoe UI", 8, "bold"), fg="#2c5282", bg="#ebf8ff", padx=6, pady=2).pack(side=tk.LEFT)

        # 3. Staged Rows Table
        table_frame = ttk.LabelFrame(content, text=" 🔍 Staged Observations Table (#tbldata1) ", padding=6)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        cols = ("sno", "job_no", "section", "desc", "obs_type", "kind")
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=10)

        tree.heading("sno", text="S.No", anchor=tk.CENTER)
        tree.heading("job_no", text="Job No", anchor=tk.CENTER)
        tree.heading("section", text="Section", anchor=tk.CENTER)
        tree.heading("desc", text="Observation / Booking Description", anchor=tk.W)
        tree.heading("obs_type", text="Observation Type", anchor=tk.CENTER)
        tree.heading("kind", text="Kind", anchor=tk.CENTER)

        tree.column("sno", width=45, stretch=False, anchor=tk.CENTER)
        tree.column("job_no", width=65, stretch=False, anchor=tk.CENTER)
        tree.column("section", width=80, stretch=False, anchor=tk.CENTER)
        tree.column("desc", width=490, stretch=True)
        tree.column("obs_type", width=140, stretch=False, anchor=tk.CENTER)
        tree.column("kind", width=70, stretch=False, anchor=tk.CENTER)

        tree.tag_configure("even", background="#f7fafc")
        tree.tag_configure("odd", background="#ffffff")

        for idx, row in enumerate(staged_rows, start=1):
            tag = "even" if idx % 2 == 0 else "odd"
            tree.insert("", tk.END, values=(
                row.get("sno", str(idx)),
                row.get("job_no", "--"),
                row.get("section", "--"),
                row.get("desc", ""),
                row.get("obs_type", config.DEFAULT_OBS_TYPE),
                row.get("kind", config.DEFAULT_JOB_KIND)
            ), tags=(tag,))

        scroll_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll_y.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 4. Action Buttons Bar
        btn_bar = tk.Frame(content)
        btn_bar.pack(fill=tk.X, pady=(4, 0))

        save_btn = tk.Button(
            btn_bar,
            text="✅ YES, PROCEED & SAVE TO SECTIONS",
            font=("Segoe UI", 11, "bold"),
            bg="#22543d",
            fg="white",
            activebackground="#1c4532",
            activeforeground="white",
            padx=20,
            pady=10,
            cursor="hand2",
            command=self._on_save
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 16))

        stop_btn = tk.Button(
            btn_bar,
            text="⏹ STOP & DO NOT SAVE",
            font=("Segoe UI", 10, "bold"),
            bg="#742a2a",
            fg="white",
            activebackground="#521e1e",
            activeforeground="white",
            padx=16,
            pady=10,
            cursor="hand2",
            command=self._on_stop
        )
        stop_btn.pack(side=tk.LEFT)

        hint_lbl = tk.Label(
            btn_bar,
            text="Clicking Save submits to SLAM and automatically sends to sections.\nClicking Stop aborts without saving (observations stay staged in portal).",
            font=("Segoe UI", 8),
            fg="#4a5568",
            justify=tk.LEFT
        )
        hint_lbl.pack(side=tk.RIGHT)

    def _on_save(self):
        self.result = True
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _on_stop(self):
        self.result = False
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
class PreferencesDialog(tk.Toplevel):
    """
    Comprehensive Preferences Modal Dialog:
    - Tab 1: 🌐 Web Browser Configuration (Default Browser, Detection, Custom Exe, Test Launch, Options)
    - Tab 2: ⭐ 'All Sec' Section Mapping Preferences (Configurable target sections with presets)
    """
    def __init__(self, parent, gui_ref: Optional[Any] = None, initial_tab: int = 0):
        super().__init__(parent)
        self.gui_ref = gui_ref
        if self.gui_ref:
            self.gui_ref._active_modal = self
        self.transient(parent)
        self.title("⚙️ Application Preferences - SLAM Auto-Filler")
        self.geometry("680x640")
        self.minsize(620, 560)
        self.result = False

        self._initial_tab = initial_tab
        self._check_vars: Dict[str, tk.BooleanVar] = {}

        # Browser settings state
        self._browser_choice_var = tk.StringVar(value=config.BROWSER_CHOICE_AUTO)
        self._custom_path_var = tk.StringVar(value="")
        self._headless_pref_var = tk.BooleanVar(value=False)
        self._keep_browser_pref_var = tk.BooleanVar(value=True)

        self._load_current_settings()
        self._build_ui()

        self.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self.grab_set()
        force_bring_to_front(self)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _load_current_settings(self):
        b_pref = config.get_browser_preference()
        self._browser_choice_var.set(b_pref.get("choice", config.BROWSER_CHOICE_AUTO))
        self._custom_path_var.set(b_pref.get("custom_path", ""))

        if self.gui_ref:
            self._headless_pref_var.set(self.gui_ref.headless_var.get())
            self._keep_browser_pref_var.set(self.gui_ref.keep_browser_var.get())

    def _on_cancel(self):
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _build_ui(self):
        # 1. Header Banner
        banner = tk.Frame(self, bg="#1a365d", height=60)
        banner.pack(fill=tk.X, side=tk.TOP)

        tk.Label(
            banner,
            text="⚙️ Application Preferences",
            font=("Segoe UI", 13, "bold"),
            fg="white",
            bg="#1a365d"
        ).pack(anchor=tk.W, padx=16, pady=(10, 2))

        tk.Label(
            banner,
            text="Configure default web browser, automation options, and 'All Sec' section mapping rules.",
            font=("Segoe UI", 9),
            fg="#90cdf4",
            bg="#1a365d"
        ).pack(anchor=tk.W, padx=16, pady=(0, 10))

        # Main Notebook (Tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=(10, 4))

        # Tab 1: Browser Settings
        tab_browser = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_browser, text=" 🌐 Web Browser ")

        # Tab 2: 'All Sec' Mapping
        tab_all_sec = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab_all_sec, text=" ⭐ 'All Sec' Mapping ")

        self._build_browser_tab(tab_browser)
        self._build_all_sec_tab(tab_all_sec)

        try:
            self.notebook.select(self._initial_tab)
        except Exception:
            pass

        # Bottom Button Bar
        btn_bar = ttk.Frame(self, padding=(12, 6, 12, 10))
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        save_btn = tk.Button(
            btn_bar,
            text="💾 Save Preferences",
            font=("Segoe UI", 10, "bold"),
            bg="#2f855a",
            fg="white",
            activebackground="#276749",
            activeforeground="white",
            padx=16,
            pady=6,
            cursor="hand2",
            command=self._on_save
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 8))

        reset_btn = tk.Button(
            btn_bar,
            text="↺ Reset All Defaults",
            font=("Segoe UI", 9),
            bg="#edf2f7",
            activebackground="#e2e8f0",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._apply_all_defaults
        )
        reset_btn.pack(side=tk.LEFT, padx=(0, 8))

        cancel_btn = tk.Button(
            btn_bar,
            text="Cancel",
            font=("Segoe UI", 9),
            bg="#e2e8f0",
            activebackground="#cbd5e0",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._on_cancel
        )
        cancel_btn.pack(side=tk.LEFT)

    def _build_browser_tab(self, parent):
        detected = config.detect_system_browsers()
        edge_info = detected["browsers"].get("msedge", {})
        chrome_info = detected["browsers"].get("chrome", {})
        brave_info = detected["browsers"].get("brave", {})
        win_def = detected.get("windows_default")
        win_def_name = win_def[1] if win_def else "Standard Windows Browser"

        # 1. Browser Selection Frame
        sel_frame = ttk.LabelFrame(parent, text=" 🎯 Default Web Browser Selection ", padding=8)
        sel_frame.pack(fill=tk.X, pady=(0, 8))

        options = [
            (
                config.BROWSER_CHOICE_AUTO,
                f"Auto-Detect / System Default (Recommended)",
                f"Prioritizes your Windows default ({win_def_name}) if supported, or falls back to Edge → Chrome."
            ),
            (
                config.BROWSER_CHOICE_EDGE,
                f"Microsoft Edge  {'[Installed]' if edge_info.get('installed') else '[Not Detected]'}",
                f"Uses Microsoft Edge (pre-installed on Windows 10/11)."
            ),
            (
                config.BROWSER_CHOICE_CHROME,
                f"Google Chrome  {'[Installed]' if chrome_info.get('installed') else '[Not Detected]'}",
                f"Uses Google Chrome browser."
            ),
            (
                config.BROWSER_CHOICE_BRAVE,
                f"Brave Browser  {'[Installed]' if brave_info.get('installed') else '[Not Detected]'}",
                f"Uses Brave Browser (Chromium-based privacy browser)."
            ),
            (
                config.BROWSER_CHOICE_CHROMIUM,
                f"Playwright Chromium  [Bundled Engine]",
                f"Uses Playwright's built-in standalone Chromium build."
            ),
            (
                config.BROWSER_CHOICE_CUSTOM,
                f"Custom Browser Executable (.exe)...",
                f"Specify a custom Chromium-based browser executable."
            ),
        ]

        for val, title, desc in options:
            rb_frame = ttk.Frame(sel_frame)
            rb_frame.pack(fill=tk.X, pady=2)

            rb = ttk.Radiobutton(
                rb_frame,
                text=title,
                value=val,
                variable=self._browser_choice_var,
                command=self._on_browser_choice_change
            )
            rb.pack(anchor=tk.W)

            sub_lbl = tk.Label(
                rb_frame,
                text=f"    {desc}",
                font=("Segoe UI", 8),
                fg="#718096"
            )
            sub_lbl.pack(anchor=tk.W)

        # Custom Path Entry
        self.custom_path_frame = ttk.Frame(sel_frame)
        self.custom_path_frame.pack(fill=tk.X, pady=(4, 2), padx=(22, 0))

        ttk.Label(self.custom_path_frame, text="Path:", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 6))
        self.custom_entry = ttk.Entry(self.custom_path_frame, textvariable=self._custom_path_var, font=("Segoe UI", 9))
        self.custom_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.browse_btn = ttk.Button(self.custom_path_frame, text="Browse...", command=self._browse_custom_browser)
        self.browse_btn.pack(side=tk.LEFT)

        self._on_browser_choice_change()

        # 2. System Browser Detection Card
        det_frame = ttk.LabelFrame(parent, text=" 🔍 Detected Browsers on This System ", padding=8)
        det_frame.pack(fill=tk.X, pady=(0, 8))

        det_grid = ttk.Frame(det_frame)
        det_grid.pack(fill=tk.X)

        tk.Label(det_grid, text="🖥️ Windows Default:", font=("Segoe UI", 8, "bold"), fg="#2d3748").grid(row=0, column=0, sticky=tk.W, padx=4, pady=1)
        tk.Label(det_grid, text=f"{win_def_name}", font=("Segoe UI", 8), fg="#2b6cb0").grid(row=0, column=1, sticky=tk.W, padx=4, pady=1)

        tk.Label(det_grid, text="🌐 Microsoft Edge:", font=("Segoe UI", 8, "bold"), fg="#2d3748").grid(row=1, column=0, sticky=tk.W, padx=4, pady=1)
        edge_status = f"✅ Installed ({edge_info.get('path')})" if edge_info.get("installed") else "❌ Not Found"
        tk.Label(det_grid, text=edge_status, font=("Segoe UI", 8), fg="#276749" if edge_info.get("installed") else "#c53030").grid(row=1, column=1, sticky=tk.W, padx=4, pady=1)

        tk.Label(det_grid, text="🌐 Google Chrome:", font=("Segoe UI", 8, "bold"), fg="#2d3748").grid(row=2, column=0, sticky=tk.W, padx=4, pady=1)
        chrome_status = f"✅ Installed ({chrome_info.get('path')})" if chrome_info.get("installed") else "❌ Not Found"
        tk.Label(det_grid, text=chrome_status, font=("Segoe UI", 8), fg="#276749" if chrome_info.get("installed") else "#c53030").grid(row=2, column=1, sticky=tk.W, padx=4, pady=1)

        tk.Label(det_grid, text="🦁 Brave Browser:", font=("Segoe UI", 8, "bold"), fg="#2d3748").grid(row=3, column=0, sticky=tk.W, padx=4, pady=1)
        brave_status = f"✅ Installed ({brave_info.get('path')})" if brave_info.get("installed") else "❌ Not Found"
        tk.Label(det_grid, text=brave_status, font=("Segoe UI", 8), fg="#276749" if brave_info.get("installed") else "#718096").grid(row=3, column=1, sticky=tk.W, padx=4, pady=1)

        # 3. Launch Diagnostics & Test Action
        diag_frame = ttk.LabelFrame(parent, text=" ⚡ Diagnostics & Test ", padding=8)
        diag_frame.pack(fill=tk.X)

        test_row = ttk.Frame(diag_frame)
        test_row.pack(fill=tk.X)

        self.test_btn = tk.Button(
            test_row,
            text="🚀 Test Launch Selected Browser",
            font=("Segoe UI", 9, "bold"),
            bg="#ebf8ff",
            fg="#2b6cb0",
            activebackground="#bee3f8",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._test_launch_browser
        )
        self.test_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.test_status_lbl = tk.Label(
            test_row,
            text="Click to test that the selected browser can launch without errors.",
            font=("Segoe UI", 8, "italic"),
            fg="#4a5568"
        )
        self.test_status_lbl.pack(side=tk.LEFT)

    def _browse_custom_browser(self):
        f = filedialog.askopenfilename(
            parent=self,
            title="Select Browser Executable",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if f:
            self._custom_path_var.set(f)
            self._browser_choice_var.set(config.BROWSER_CHOICE_CUSTOM)
            self._on_browser_choice_change()

    def _on_browser_choice_change(self):
        is_custom = (self._browser_choice_var.get() == config.BROWSER_CHOICE_CUSTOM)
        state = tk.NORMAL if is_custom else tk.DISABLED
        self.custom_entry.config(state=state)
        self.browse_btn.config(state=state)

    def _test_launch_browser(self):
        choice = self._browser_choice_var.get()
        custom_path = self._custom_path_var.get().strip()

        if choice == config.BROWSER_CHOICE_CUSTOM and (not custom_path or not os.path.isfile(custom_path)):
            messagebox.showwarning(
                "Custom Path Required",
                "Please browse or specify a valid .exe path for your Custom Browser before testing.",
                parent=self
            )
            return

        self.test_btn.config(state=tk.DISABLED, text="⏳ Testing Launch...")
        self.test_status_lbl.config(text="Launching browser...", fg="#2b6cb0")

        def _worker():
            try:
                from slam_bot import SLAMBot
                bot = SLAMBot(
                    headless=self._headless_pref_var.get(),
                    browser_choice=choice,
                    custom_browser_path=custom_path
                )
                bot.start_browser()
                import time
                time.sleep(2.0)
                bot_launched_name = bot.current_browser_launched or "Browser"
                bot.stop()
                try:
                    if bot.page:
                        bot.page.close()
                    if bot.context:
                        bot.context.close()
                    if bot.browser:
                        bot.browser.close()
                except Exception:
                    pass

                def _success():
                    self.test_btn.config(state=tk.NORMAL, text="🚀 Test Launch Selected Browser")
                    self.test_status_lbl.config(
                        text=f"✅ Success! {bot_launched_name} launched and closed cleanly.",
                        fg="#22543d"
                    )
                    messagebox.showinfo(
                        "Browser Test Succeeded",
                        f"🎉 Success!\n\n{bot_launched_name} launched and closed cleanly without errors.",
                        parent=self
                    )
                self.after(0, _success)
            except Exception as e:
                def _fail(err=e):
                    self.test_btn.config(state=tk.NORMAL, text="🚀 Test Launch Selected Browser")
                    self.test_status_lbl.config(
                        text=f"❌ Launch failed: {str(err)[:50]}...",
                        fg="#9b2c2c"
                    )
                    messagebox.showerror(
                        "Browser Test Failed",
                        f"Could not launch browser:\n\n{err}",
                        parent=self
                    )
                self.after(0, _fail)

        threading.Thread(target=_worker, daemon=True).start()

    def _build_all_sec_tab(self, parent):
        # 1. Rule Info Card
        info_card = ttk.LabelFrame(parent, text=" ℹ️ Rule Explanation ", padding=8)
        info_card.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            info_card,
            text="When an observation row in Excel specifies Section as 'All Sec':\n"
                 "The bot will automatically stage that observation to all checked sections below.\n"
                 "Default Rule: Send to all sections EXCEPT E8, Q&I, LAB, PPIO.",
            font=("Segoe UI", 9),
            fg="#2d3748",
            justify=tk.LEFT
        ).pack(anchor=tk.W)

        # 2. Checkbox Grid Frame
        grid_frame = ttk.LabelFrame(parent, text=" 🎯 Select Target Sections for 'All Sec' ", padding=8)
        grid_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        current_saved = set(config.get_all_sec_preferences())

        sec_descriptions = {
            "M1": "M1 (Mechanical 1 - TM / Bogie)",
            "M2": "M2 (Mechanical 2 - Pneumatic / Mech)",
            "E3": "E3 (Electrical 3 - Control & Aux)",
            "E3A": "E3A (Electrical 3A - Cab Equipment)",
            "E4": "E4 (Electrical 4 - Panto / Roof / VCB)",
            "E5A": "E5A (Electrical 5A - Electronics)",
            "E5B": "E5B (Electrical 5B - Inverter / Converter)",
            "BRS": "BRS (Brake System)",
            "E8": "E8 (Speedo / Electronic Testing) [Excluded by default]",
            "QAI": "QAI / Q&I (Quality Assurance) [Excluded by default]",
            "LAB": "LAB (Chemical & Oil Lab) [Excluded by default]",
            "PPIO": "PPIO (Planning / Progress) [Excluded by default]"
        }

        col_left = ttk.Frame(grid_frame)
        col_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        col_right = ttk.Frame(grid_frame)
        col_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(8, 0))

        all_ordered = config.ALL_SECTIONS_LIST
        midpoint = len(all_ordered) // 2

        for idx, sec in enumerate(all_ordered):
            parent_col = col_left if idx < midpoint else col_right
            var = tk.BooleanVar(value=(sec in current_saved))
            self._check_vars[sec] = var
            var.trace_add("write", lambda *a: self._update_summary())

            cb_text = sec_descriptions.get(sec, sec)
            cb = ttk.Checkbutton(parent_col, text=cb_text, variable=var)
            cb.pack(anchor=tk.W, pady=2)

        # 3. Quick Presets Toolbar
        presets_frame = ttk.Frame(parent)
        presets_frame.pack(fill=tk.X, pady=(0, 6))

        default_btn = tk.Button(
            presets_frame,
            text="⚡ Default (All except E8, Q&I, LAB, PPIO)",
            font=("Segoe UI", 8, "bold"),
            bg="#ebf8ff",
            fg="#2b6cb0",
            activebackground="#bee3f8",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._apply_default_preset
        )
        default_btn.pack(side=tk.LEFT, padx=(0, 6))

        all_btn = ttk.Button(presets_frame, text="☑ Select All", command=self._select_all)
        all_btn.pack(side=tk.LEFT, padx=(0, 6))

        none_btn = ttk.Button(presets_frame, text="☐ Clear All", command=self._clear_all)
        none_btn.pack(side=tk.LEFT)

        # 4. Live Summary Banner
        self.summary_frame = tk.Frame(parent, bg="#f7fafc", relief=tk.SOLID, borderwidth=1, padx=8, pady=4)
        self.summary_frame.pack(fill=tk.X)

        self.summary_lbl = tk.Label(
            self.summary_frame,
            text="",
            font=("Segoe UI", 9, "bold"),
            bg="#f7fafc",
            fg="#2d3748"
        )
        self.summary_lbl.pack(anchor=tk.W)

        self._update_summary()

    def _apply_default_preset(self):
        default_secs = set(config.DEFAULT_ALL_SEC_INCLUDED)
        for sec, var in self._check_vars.items():
            var.set(sec in default_secs)

    def _select_all(self):
        for var in self._check_vars.values():
            var.set(True)

    def _clear_all(self):
        for var in self._check_vars.values():
            var.set(False)

    def _get_selected_sections(self) -> List[str]:
        return [sec for sec in config.ALL_SECTIONS_LIST if self._check_vars.get(sec) and self._check_vars[sec].get()]

    def _update_summary(self):
        selected = self._get_selected_sections()
        excluded = [s for s in config.ALL_SECTIONS_LIST if s not in selected]
        sel_text = ", ".join(selected) if selected else "None"
        excl_text = ", ".join(excluded) if excluded else "None"
        self.summary_lbl.config(
            text=f"Selected ({len(selected)}/12): {sel_text}\nExcluded ({len(excluded)}): {excl_text}"
        )

    def _apply_all_defaults(self):
        self._browser_choice_var.set(config.BROWSER_CHOICE_AUTO)
        self._custom_path_var.set("")
        self._on_browser_choice_change()
        self._apply_default_preset()

    def _on_save(self):
        choice = self._browser_choice_var.get()
        custom_path = self._custom_path_var.get().strip()

        if choice == config.BROWSER_CHOICE_CUSTOM and (not custom_path or not os.path.isfile(custom_path)):
            if not messagebox.askyesno(
                "Invalid Custom Browser Path",
                "The custom browser path does not point to an existing .exe file.\n\nSave anyway (will fall back to Auto-Detect)?",
                parent=self
            ):
                return

        config.save_browser_preference(choice, custom_path)

        selected = self._get_selected_sections()
        if not selected:
            if not messagebox.askyesno(
                "No Sections Selected",
                "You have not selected any sections for 'All Sec'.\n\nDo you want to save anyway? (Observations with 'All Sec' will be skipped).",
                parent=self
            ):
                return

        config.save_all_sec_preferences(selected)

        if self.gui_ref:
            self.gui_ref._sync_preferences()
            b_name = config.BROWSER_DISPLAY_NAMES.get(choice, choice)
            self.gui_ref.log(f"Preferences saved: Browser = {b_name}, 'All Sec' = {len(selected)} sections ({', '.join(selected)})")
            self.gui_ref._active_modal = None

        messagebox.showinfo(
            "Preferences Saved",
            f"Successfully saved preferences!\n\n"
            f"• Default Browser: {config.BROWSER_DISPLAY_NAMES.get(choice, choice)}\n"
            f"• 'All Sec' Target Sections ({len(selected)}):\n  {', '.join(selected) if selected else 'None'}",
            parent=self
        )
        self.destroy()

class AllSecPreferencesDialog(PreferencesDialog):
    """Backwards-compatible alias for opening Preferences directly to 'All Sec' tab."""
    def __init__(self, parent, gui_ref: Optional[Any] = None):
        super().__init__(parent, gui_ref=gui_ref, initial_tab=1)

class EditBookingDialog(tk.Toplevel):
    """
    Modal dialog allowing the user to view and edit a booking's defect message and section.
    """
    def __init__(self, parent, booking_data: dict, is_new: bool = False, gui_ref: Optional[Any] = None):
        super().__init__(parent)
        self.gui_ref = gui_ref
        if self.gui_ref:
            self.gui_ref._active_modal = self
        self.transient(parent)
        self.title("➕ Add New Booking" if is_new else f"✏️ Edit Booking Message - S.No {booking_data.get('sno', '--')}")
        self.geometry("680x480")
        self.minsize(560, 420)
        self.result = None
        self.booking_data = booking_data
        self.is_new = is_new

        self._build_ui()

        self.update_idletasks()
        try:
            px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except Exception:
            pass

        self.grab_set()
        force_bring_to_front(self)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_cancel(self):
        self.result = None
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def _build_ui(self):
        container = ttk.Frame(self, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        header_text = "➕ Add New Jobcard Booking" if self.is_new else f"✏️ Edit Defect / Booking Message (S.No {self.booking_data.get('sno', '--')})"
        tk.Label(container, text=header_text, font=("Segoe UI", 12, "bold"), fg="#1a365d").pack(anchor=tk.W, pady=(0, 10))

        # S.No, Obs Type & Kind row
        row1 = ttk.Frame(container)
        row1.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(row1, text="S.No:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.sno_var = tk.StringVar(value=str(self.booking_data.get("sno", "")))
        sno_entry = ttk.Entry(row1, textvariable=self.sno_var, width=6, font=("Segoe UI", 9))
        sno_entry.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(row1, text="Source (Obs Type):", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.obs_type_var = tk.StringVar(value=self.booking_data.get("obs_type", config.DEFAULT_OBS_TYPE))
        obs_cb = ttk.Combobox(row1, textvariable=self.obs_type_var, values=list(config.OBS_TYPE_VALUES.keys()), state="readonly", width=20)
        obs_cb.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(row1, text="Kind:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.kind_var = tk.StringVar(value=self.booking_data.get("kind", config.DEFAULT_JOB_KIND))
        kind_cb = ttk.Combobox(row1, textvariable=self.kind_var, values=list(config.JOB_KIND_VALUES.keys()), state="readonly", width=10)
        kind_cb.pack(side=tk.LEFT)

        # Description / Message Box
        ttk.Label(container, text="Defect / Observation Description (Message text in SLAM):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(4, 2))

        text_frame = ttk.Frame(container)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.desc_text = tk.Text(text_frame, height=5, font=("Segoe UI", 10), wrap="word")
        self.desc_text.insert("1.0", self.booking_data.get("description", ""))
        desc_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.desc_text.yview)
        self.desc_text.configure(yscrollcommand=desc_scroll.set)
        self.desc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        desc_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.desc_text.focus_set()

        # Section row
        ttk.Label(container, text="Section(s) in SLAM (e.g. M1, M2/BRS, E3A & E5A):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        sec_frame = ttk.Frame(container)
        sec_frame.pack(fill=tk.X, pady=(0, 6))

        self.sec_var = tk.StringVar(value=self.booking_data.get("raw_section", ""))
        sec_entry = ttk.Entry(sec_frame, textvariable=self.sec_var, font=("Segoe UI", 10))
        sec_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.sec_var.trace_add("write", self._update_mapped_preview)

        # Quick Section Buttons
        chips_frame = ttk.LabelFrame(container, text=" Quick Section Buttons (Click to Toggle / Add) ", padding=4)
        chips_frame.pack(fill=tk.X, pady=(0, 10))

        common_secs = ["M1", "M2", "E3", "E3A", "E4", "E5A", "E5B", "E8", "BRS", "LAB", "PPIO", "QAI"]
        c_grid = ttk.Frame(chips_frame)
        c_grid.pack(fill=tk.X)
        for s in common_secs:
            btn = tk.Button(
                c_grid,
                text=s,
                font=("Segoe UI", 8, "bold"),
                bg="#edf2f7",
                activebackground="#cbd5e0",
                padx=5,
                pady=2,
                cursor="hand2",
                command=lambda sec=s: self._toggle_section_chip(sec)
            )
            btn.pack(side=tk.LEFT, padx=2, pady=2)

        all_sec_btn = tk.Button(
            c_grid,
            text="⭐ All Sec",
            font=("Segoe UI", 8, "bold"),
            bg="#feebc8",
            fg="#7b341e",
            activebackground="#fbd38d",
            padx=7,
            pady=2,
            cursor="hand2",
            command=self._set_all_sec
        )
        all_sec_btn.pack(side=tk.LEFT, padx=(5, 2), pady=2)

        # Mapped Section Preview
        preview_frame = ttk.Frame(container)
        preview_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(preview_frame, text="Mapped SLAM Codes:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.mapped_lbl = tk.Label(preview_frame, text="--", font=("Segoe UI", 8, "bold"), fg="#2b6cb0", bg="#ebf8ff", padx=6, pady=1)
        self.mapped_lbl.pack(side=tk.LEFT)
        self._update_mapped_preview()

        # Action Buttons
        btn_bar = ttk.Frame(container)
        btn_bar.pack(fill=tk.X)

        save_btn = tk.Button(
            btn_bar,
            text="💾 Save Changes",
            font=("Segoe UI", 10, "bold"),
            bg="#2f855a",
            fg="white",
            activebackground="#276749",
            activeforeground="white",
            padx=16,
            pady=6,
            cursor="hand2",
            command=self._on_save
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 8))

        cancel_btn = tk.Button(
            btn_bar,
            text="Cancel",
            font=("Segoe UI", 9),
            bg="#e2e8f0",
            activebackground="#cbd5e0",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._on_cancel
        )
        cancel_btn.pack(side=tk.LEFT)

        self.bind("<Control-Return>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self._on_cancel())

    def _set_all_sec(self):
        curr = self.sec_var.get().strip().upper()
        if "ALL SEC" in curr:
            self.sec_var.set("")
        else:
            self.sec_var.set("All Sec")

    def _toggle_section_chip(self, sec: str):
        from excel_parser import normalize_sections
        curr = self.sec_var.get().strip()
        existing = normalize_sections(curr)
        if sec in existing:
            existing.remove(sec)
        else:
            existing.append(sec)
        self.sec_var.set(" / ".join(existing) if existing else "")

    def _update_mapped_preview(self, *args):
        from excel_parser import normalize_sections
        secs = normalize_sections(self.sec_var.get())
        if secs:
            self.mapped_lbl.config(text=", ".join(secs), fg="#22543d", bg="#c6f6d5")
        else:
            self.mapped_lbl.config(text="⚠️ Unrecognized Section", fg="#c53030", bg="#fff5f5")

    def _on_save(self):
        from excel_parser import normalize_sections
        desc = self.desc_text.get("1.0", tk.END).strip()
        raw_sec = self.sec_var.get().strip()
        sno_str = self.sno_var.get().strip()
        try:
            sno = int(sno_str)
        except ValueError:
            sno = self.booking_data.get("sno", 1)

        if not desc:
            messagebox.showwarning("Missing Description", "Please enter a defect / observation description!")
            return

        mapped_secs = normalize_sections(raw_sec)
        if not mapped_secs:
            confirm = messagebox.askyesno(
                "Unrecognized Section",
                f"The section '{raw_sec}' is not recognized as a valid SLAM section code.\n\nDo you want to save it anyway?"
            )
            if not confirm:
                return

        self.result = {
            "sno": sno,
            "description": desc,
            "raw_section": raw_sec,
            "sections": mapped_secs,
            "category": self.booking_data.get("category", "General"),
            "obs_type": self.obs_type_var.get(),
            "kind": self.kind_var.get()
        }
        if hasattr(self, "gui_ref") and self.gui_ref:
            self.gui_ref._active_modal = None
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

class SLAMBotWorker(threading.Thread):
    """
    Dedicated persistent worker thread that owns the Playwright automation session.
    Because Playwright sync_api uses greenlet, ALL calls to Playwright/browser/page
    MUST execute within this exact same thread.
    This thread stays alive across multiple runs, allowing the browser session to remain
    open and reused smoothly without 'cannot switch to a different thread' errors.
    """
    def __init__(self, gui: "SLAMAutoFillerGUI"):
        super().__init__(name="SLAMBotWorker", daemon=True)
        self.gui = gui
        self.task_queue: queue.Queue = queue.Queue()
        self.bot: Optional[SLAMBot] = None
        self._running = True

    def _ensure_bot(self, params: dict):
        """Initializes or updates the persistent SLAMBot instance on the worker thread."""
        b_choice = params.get("browser_choice")
        c_path = params.get("custom_browser_path")
        if self.bot is None:
            self.bot = SLAMBot(
                username=params["username"],
                password=params["password"],
                jobcard_type=params.get("jobcard_type", config.DEFAULT_JOBCARD_TYPE),
                eqt_required=params.get("eqt_required", config.DEFAULT_EQT_REQUIRED),
                obs_type=params.get("obs_type", config.DEFAULT_OBS_TYPE),
                job_kind=params.get("job_kind", config.DEFAULT_JOB_KIND),
                headless=params.get("headless", False),
                browser_choice=b_choice,
                custom_browser_path=c_path,
                log_callback=self.gui.log,
                otp_callback=self.gui._prompt_otp,
                progress_callback=self.gui._update_progress,
                verify_before_save=True,
                verify_callback=self.gui._prompt_verify
            )
        else:
            self.bot.username = params["username"]
            self.bot.password = params["password"]
            self.bot.jobcard_type = params.get("jobcard_type", self.bot.jobcard_type)
            self.bot.eqt_required = params.get("eqt_required", self.bot.eqt_required)
            self.bot.obs_type = params.get("obs_type", self.bot.obs_type)
            self.bot.job_kind = params.get("job_kind", self.bot.job_kind)
            self.bot.headless = params.get("headless", False)
            self.bot.log_callback = self.gui.log
            self.bot.otp_callback = self.gui._prompt_otp
            self.bot.progress_callback = self.gui._update_progress
            self.bot.verify_before_save = True
            self.bot.verify_callback = self.gui._prompt_verify
            self.bot._stop_requested = False

            old_b_choice = getattr(self.bot, "browser_choice", None)
            if old_b_choice != b_choice:
                self.gui.log(f"Browser preference changed ({old_b_choice} -> {b_choice}). Closing previous browser session...")
                try:
                    if self.bot.page:
                        self.bot.page.close()
                    if self.bot.context:
                        self.bot.context.close()
                    if self.bot.browser:
                        self.bot.browser.close()
                except Exception:
                    pass
                self.bot.page = None
                self.bot.context = None
                self.bot.browser = None
                self.bot.is_logged_in = False

            self.bot.browser_choice = b_choice
            self.bot.custom_browser_path = c_path

            if self.bot.browser and self.bot.browser.is_connected() and self.bot.page and not self.bot.page.is_closed():
                self.gui.log("Reusing existing open browser session...")
            else:
                self.gui.log("Opening new browser window for automation...")

    def run(self):
        while self._running:
            try:
                task = self.task_queue.get(timeout=0.5)
            except queue.Empty:
                continue

            action = task.get("action")

            if action == "PROCESS_EXCEL":
                excel_path = task.get("excel_path", "")
                parsed_data = task.get("parsed_data")
                params = task["params"]
                try:
                    self._ensure_bot(params)

                    # Update loco number in parsed data if edited by user
                    if self.gui.parsed_data:
                        self.gui.parsed_data["loco_number"] = params["loco_number"]

                    res = self.bot.process_excel(excel_path=excel_path, parsed_data=parsed_data)

                    # Refresh daily OTP display if it was saved during login
                    self.gui.root.after(0, lambda: (
                        self.gui.daily_otp_var.set(config.get_daily_otp()),
                        self.gui._update_otp_status()
                    ))

                    if res.get("status") == "completed":
                        def _done_success(r=res):
                            self.gui.status_lbl.config(text="Completed Successfully!")
                            self.gui.progress_bar["value"] = 100
                            notify_user_attention(self.gui.root)
                            messagebox.showinfo(
                                "SLAM Jobcards Created Successfully",
                                f"🎉 Success!\n\nAll {r['entered']}/{r['total']} jobcards for Locomotive {r['loco_number']} have been created, saved, and dispatched in the SLAM portal.",
                                parent=self.gui.root
                            )
                        self.gui.root.after(0, _done_success)
                    elif res.get("status") == "verification_deferred":
                        def _done_deferred(r=res):
                            self.gui.status_lbl.config(text="Save deferred by user.")
                            notify_user_attention(self.gui.root)
                            messagebox.showinfo(
                                "Save Deferred",
                                f"Observations for Locomotive {r.get('loco_number')} remain staged in the SLAM portal.\n\nNo save was committed.",
                                parent=self.gui.root
                            )
                        self.gui.root.after(0, _done_deferred)
                    elif res.get("status") == "cancelled":
                        self.gui.root.after(0, lambda: self.gui.status_lbl.config(text="Stopped by user."))
                    else:
                        self.gui.root.after(0, lambda: self.gui.status_lbl.config(text="Failed to complete."))

                except Exception as e:
                    self.gui.log(f"FATAL ERROR: {e}")
                    def _done_error(err=e):
                        self.gui.status_lbl.config(text="Error occurred.")
                        notify_user_attention(self.gui.root)
                        messagebox.showerror("Execution Error", f"An error occurred during auto-fill:\n\n{err}", parent=self.gui.root)
                    self.gui.root.after(0, _done_error)
                finally:
                    if not params.get("keep_browser", True):
                        if self.bot:
                            try:
                                self.bot.close()
                            except Exception:
                                pass
                            self.bot = None
                    else:
                        self.gui.log("Browser session remains open for your review.")
                    self.gui.root.after(0, self.gui._reset_buttons)
                    self.task_queue.task_done()

            elif action == "PROCESS_EQUIPMENT":
                eq_rows = task.get("rows", [])
                params = task["params"]
                try:
                    self._ensure_bot(params)
                    res = self.bot.process_loco_equipment(
                        rows=eq_rows,
                        select_for=params.get("select_for", "Home"),
                        target_locos=params.get("target_locos"),
                        auto_save=params.get("auto_save", True)
                    )

                    self.gui.root.after(0, lambda: (
                        self.gui.daily_otp_var.set(config.get_daily_otp()),
                        self.gui._update_otp_status()
                    ))

                    if res.get("status") == "completed":
                        def _done_eq_success(r=res):
                            self.gui.status_lbl.config(text="Equipment Fitment Completed!")
                            self.gui.progress_bar["value"] = 100
                            notify_user_attention(self.gui.root)
                            messagebox.showinfo(
                                "SLAM Loco Equipment Fitment Complete",
                                f"🔧 Loco Equipment Automation Finished!\n\n"
                                f"• Locomotives Processed: {r.get('locos_processed', 0)}/{r.get('total_locos', 0)}\n"
                                f"• Equipment Popups Filled: {r.get('filled', 0)}\n"
                                f"• Equipment Saved: {r.get('saved', 0)}\n"
                                f"• Already Fitted / Skipped: {r.get('skipped', 0)}\n"
                                f"• Unmatched / Errors: {r.get('failed', 0)}",
                                parent=self.gui.root
                            )
                        self.gui.root.after(0, _done_eq_success)
                    elif res.get("status") == "cancelled":
                        self.gui.root.after(0, lambda: self.gui.status_lbl.config(text="Equipment fill stopped by user."))
                    else:
                        def _done_eq_fail(r=res):
                            self.gui.status_lbl.config(text="Equipment fill failed.")
                            notify_user_attention(self.gui.root)
                            messagebox.showwarning(
                                "Equipment Fitment Incomplete",
                                f"Could not complete equipment fitment:\n\n{r.get('error', 'Check activity log for details.')}",
                                parent=self.gui.root
                            )
                        self.gui.root.after(0, _done_eq_fail)
                except Exception as e:
                    self.gui.log(f"FATAL EQUIPMENT ERROR: {e}")
                    def _done_eq_err(err=e):
                        self.gui.status_lbl.config(text="Error occurred.")
                        notify_user_attention(self.gui.root)
                        messagebox.showerror("Equipment Automation Error", f"An error occurred during equipment auto-fill:\n\n{err}", parent=self.gui.root)
                    self.gui.root.after(0, _done_eq_err)
                finally:
                    if not params.get("keep_browser", True):
                        if self.bot:
                            try:
                                self.bot.close()
                            except Exception:
                                pass
                            self.bot = None
                    else:
                        self.gui.log("Browser session remains open for your review.")
                    self.gui.root.after(0, self.gui._reset_buttons)
                    self.task_queue.task_done()

            elif action == "OPEN_SHOPFLOOR":
                module_name = task.get("module_name", "Add Loco Eq./Eq. Fitment")
                params = task["params"]
                try:
                    self._ensure_bot(params)
                    ok = self.bot.open_shopfloor_module(module_name=module_name)
                    self.gui.root.after(0, lambda: (
                        self.gui.daily_otp_var.set(config.get_daily_otp()),
                        self.gui._update_otp_status()
                    ))
                    if ok:
                        self.gui.root.after(0, lambda m=module_name: self.gui.status_lbl.config(text=f"Opened Shop Floor: {m}"))
                    else:
                        self.gui.root.after(0, lambda m=module_name: self.gui.status_lbl.config(text=f"Failed to open {m}"))
                except Exception as e:
                    self.gui.log(f"Error opening Shop Floor module '{module_name}': {e}")
                finally:
                    self.gui.root.after(0, self.gui._reset_buttons)
                    self.task_queue.task_done()

            elif action == "CLOSE_BROWSER":
                if self.bot:
                    try:
                        self.bot.close()
                    except Exception:
                        pass
                    self.gui.log("Browser session closed by user.")
                    self.gui.root.after(0, lambda: messagebox.showinfo("Browser Closed", "The browser session has been closed."))
                else:
                    self.gui.root.after(0, lambda: messagebox.showinfo("Browser Status", "No active browser session is currently open."))
                self.task_queue.task_done()

            elif action == "SHUTDOWN":
                if self.bot:
                    try:
                        self.bot.close()
                    except Exception:
                        pass
                    self.bot = None
                self._running = False
                self.task_queue.task_done()
                break

    def stop_current_job(self):
        """Signals the bot to stop without touching Playwright objects from another thread."""
        if self.bot:
            self.bot.stop()

    def stop_worker(self):
        """Signals the worker loop to shut down cleanly."""
        self._running = False
        try:
            self.task_queue.put({"action": "SHUTDOWN"})
        except Exception:
            pass

class SLAMAutoFillerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"SLAM Jobcard & Loco Equipment Auto-Filler v{config.APP_VERSION} - Indian Railways (ELS/RPME)")
        self.root.geometry("1120x860")
        self.root.minsize(960, 700)

        # Apply ttk theme
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
            self.style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[14, 5])
            self.style.map(
                "TNotebook.Tab",
                background=[("selected", "#2b6cb0"), ("!selected", "#e2e8f0")],
                foreground=[("selected", "white"), ("!selected", "#1a365d")]
            )
        except Exception:
            pass

        self.current_excel_path = ""
        self.parsed_data = None
        self.update_info = None
        self._active_modal: Optional[tk.Toplevel] = None
        self._inline_editor_frame: Optional[tk.Frame] = None
        self._inline_chip_bar: Optional[tk.Frame] = None
        self._inline_entry: Optional[tk.Widget] = None
        self._inline_item_id: Optional[str] = None
        self._inline_col: Optional[str] = None
        self._inline_orig_val: str = ""
        self._inline_clicking_chip: bool = False
        self._direct_chip_btns: Dict[str, tk.Button] = {}
        self._syncing_direct_bar: bool = False

        # Loco Equipment state
        self.current_eq_excel_path: str = ""
        self.eq_parsed_data: Optional[Dict[str, Any]] = None
        self._eq_inline_frame: Optional[tk.Frame] = None
        self._eq_inline_widget: Optional[tk.Widget] = None
        self._eq_inline_item_id: Optional[str] = None
        self._eq_inline_col: Optional[str] = None
        self._eq_syncing_direct_bar: bool = False

        # Load credentials from config / user_settings
        saved_u, saved_p = config.get_credentials()
        self.username_var = tk.StringVar(value=saved_u)
        self.password_var = tk.StringVar(value=saved_p)
        self.show_pwd_var = tk.BooleanVar(value=False)
        self.daily_otp_var = tk.StringVar(value=config.get_daily_otp())
        self.keep_browser_var = tk.BooleanVar(value=True)

        # Start persistent single-threaded automation worker
        self.worker = SLAMBotWorker(self)
        self.worker.start()

        self._apply_window_icon()
        self._build_ui()
        self._sync_preferences()
        self._update_otp_status()
        self._scan_desktop_files()
        self._auto_detect_equipment_file()
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

        # Bind window focus and mapping events to auto-raise active dialogs when taskbar is clicked
        self.root.bind("<FocusIn>", self._on_root_focus)
        self.root.bind("<Map>", self._on_root_map)

        # Re-apply window icon after Tkinter event loop starts and window is mapped to Windows DWM
        self.root.after(100, self._apply_window_icon)
        self.root.after(500, self._apply_window_icon)

        # Check for updates in background
        self.root.after(1500, self._check_update_startup)

    def _on_root_focus(self, event=None):
        if event and event.widget != self.root:
            return
        if hasattr(self, "_active_modal") and self._active_modal:
            try:
                if self._active_modal.winfo_exists():
                    force_bring_to_front(self._active_modal)
                    self._active_modal.focus_force()
            except Exception:
                pass

    def _on_root_map(self, event=None):
        if event and event.widget != self.root:
            return
        if hasattr(self, "_active_modal") and self._active_modal:
            try:
                if self._active_modal.winfo_exists():
                    self.root.after(60, self._on_root_focus)
            except Exception:
                pass

    def _apply_window_icon(self):
        """Applies application icon to Tkinter window and taskbar."""
        ico_p = get_resource_path("app_icon.ico")
        png_p = get_resource_path("app_icon.png")

        # 1. Native Windows Win32 Icon configuration (Taskbar, Titlebar & Class)
        if sys.platform == "win32" and os.path.exists(ico_p):
            try:
                self.root.iconbitmap(default=ico_p)
            except Exception:
                try:
                    self.root.iconbitmap(ico_p)
                except Exception:
                    pass

            # Explicitly send WM_SETICON and SetClassLongPtr to HWNDs
            try:
                import ctypes
                user32 = ctypes.windll.user32
                self.root.update_idletasks()
                wid = self.root.winfo_id()
                hwnd = user32.GetParent(wid)
                if not hwnd:
                    hwnd = wid
                root_ancestor = user32.GetAncestor(wid, 2)  # GA_ROOT = 2

                # Ensure WS_EX_APPWINDOW is set on container/top-level HWNDs
                # so Windows Taskbar manages restore & foreground activation reliably on all systems
                GWL_EXSTYLE = -20
                WS_EX_APPWINDOW = 0x00040000
                for h in set([hwnd, root_ancestor]):
                    try:
                        exstyle = user32.GetWindowLongW(h, GWL_EXSTYLE)
                        if not (exstyle & WS_EX_APPWINDOW):
                            user32.SetWindowLongW(h, GWL_EXSTYLE, exstyle | WS_EX_APPWINDOW)
                    except Exception:
                        pass

                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x00000010
                WM_SETICON = 0x0080
                ICON_SMALL = 0
                ICON_BIG = 1

                # Load 32x32 icon for Taskbar and Alt-Tab
                hicon_big = user32.LoadImageW(None, ico_p, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
                # Load 16x16 icon for Titlebar
                hicon_small = user32.LoadImageW(None, ico_p, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)

                target_hwnds = set(filter(None, [hwnd, wid, root_ancestor]))
                for h in target_hwnds:
                    if hicon_big:
                        user32.SendMessageW(h, WM_SETICON, ICON_BIG, hicon_big)
                    if hicon_small:
                        user32.SendMessageW(h, WM_SETICON, ICON_SMALL, hicon_small)

                # Set window class icon via SetClassLongPtr
                if ctypes.sizeof(ctypes.c_void_p) == 8:
                    SetClassLong = user32.SetClassLongPtrW
                else:
                    SetClassLong = user32.SetClassLongW
                SetClassLong.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
                SetClassLong.restype = ctypes.c_void_p

                GCLP_HICON = -14
                GCLP_HICONSM = -34
                for h in target_hwnds:
                    if hicon_big:
                        SetClassLong(h, GCLP_HICON, hicon_big)
                    if hicon_small:
                        SetClassLong(h, GCLP_HICONSM, hicon_small)

                self._hicons = (hicon_big, hicon_small)
            except Exception:
                pass
        else:
            # Linux / Non-Windows fallback
            if os.path.exists(png_p):
                try:
                    img = Image.open(png_p).resize((32, 32), Image.Resampling.LANCZOS)
                    self._taskbar_icon = ImageTk.PhotoImage(img)
                    self.root.iconphoto(True, self._taskbar_icon)
                except Exception:
                    pass

    def _build_ui(self):
        # ── Header Banner ───────────────────────────────────────────────────────
        header_frame = tk.Frame(self.root, bg="#1a365d", height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        # Header Logo Image
        png_p = get_resource_path("app_icon.png")
        if os.path.exists(png_p):
            try:
                logo_img = Image.open(png_p).resize((52, 52), Image.Resampling.LANCZOS)
                self.header_logo = ImageTk.PhotoImage(logo_img)
                logo_lbl = tk.Label(header_frame, image=self.header_logo, bg="#1a365d")
                logo_lbl.pack(side=tk.LEFT, padx=(14, 2), pady=9)
            except Exception:
                pass

        title_lbl = tk.Label(
            header_frame,
            text="Indian Railways - SLAM Jobcard & Equipment Auto-Filler",
            font=("Segoe UI", 16, "bold"),
            fg="white",
            bg="#1a365d"
        )
        title_lbl.pack(side=tk.LEFT, padx=(6, 16), pady=10)

        subtitle_frame = tk.Frame(header_frame, bg="#1a365d")
        subtitle_frame.pack(side=tk.RIGHT, padx=16, pady=10)

        # Version badge & Manual update check
        sub_top = tk.Frame(subtitle_frame, bg="#1a365d")
        sub_top.pack(anchor=tk.E, pady=(0, 2))

        self.update_btn = tk.Button(
            sub_top,
            text="🔄 Check Updates",
            font=("Segoe UI", 8),
            fg="#1a365d",
            bg="#e2e8f0",
            activebackground="#cbd5e0",
            relief=tk.FLAT,
            padx=6,
            pady=1,
            cursor="hand2",
            command=self._on_manual_check_update
        )
        self.update_btn.pack(side=tk.RIGHT, padx=(6, 0))

        ver_lbl = tk.Label(
            sub_top,
            text=f"v{config.APP_VERSION}",
            font=("Segoe UI", 8, "bold"),
            fg="#e2e8f0",
            bg="#2d3748",
            padx=5,
            pady=1
        )
        ver_lbl.pack(side=tk.RIGHT)

        tk.Label(
            subtitle_frame,
            text="ELS/RPME Royapuram Shed Edition",
            font=("Segoe UI", 10, "bold"),
            fg="#90cdf4",
            bg="#1a365d"
        ).pack(anchor=tk.E)

        tk.Label(
            subtitle_frame,
            text="⚡ Jobcards + Shop Floor Loco Equipment Fitment Suite",
            font=("Segoe UI", 8),
            fg="#cbd5e0",
            bg="#1a365d"
        ).pack(anchor=tk.E)

        # ── Update Alert Banner (Hidden by default, shown when update found) ──────
        self.update_banner = tk.Frame(self.root, bg="#2b6cb0", pady=6, padx=14)
        self.update_banner_lbl = tk.Label(
            self.update_banner,
            text="",
            font=("Segoe UI", 9, "bold"),
            fg="white",
            bg="#2b6cb0"
        )
        self.update_banner_lbl.pack(side=tk.LEFT)

        close_banner_btn = tk.Button(
            self.update_banner,
            text="✕",
            font=("Segoe UI", 9, "bold"),
            fg="white",
            bg="#2b6cb0",
            activebackground="#2c5282",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=self.update_banner.pack_forget
        )
        close_banner_btn.pack(side=tk.RIGHT, padx=(10, 0))

        view_rel_btn = tk.Button(
            self.update_banner,
            text="Release Notes",
            font=("Segoe UI", 8),
            fg="#1a365d",
            bg="#ebf8ff",
            relief=tk.FLAT,
            cursor="hand2",
            padx=6,
            pady=1,
            command=self._open_release_notes
        )
        view_rel_btn.pack(side=tk.RIGHT, padx=4)

        update_now_btn = tk.Button(
            self.update_banner,
            text="⚡ Update Now",
            font=("Segoe UI", 8, "bold"),
            fg="#1a365d",
            bg="#feebc8",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=1,
            command=self._open_update_dialog
        )
        update_now_btn.pack(side=tk.RIGHT, padx=4)

        # ── Main Content Container ──────────────────────────────────────────────
        main_container = ttk.Frame(self.root, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)
        self.main_container = main_container

        # 1. Credentials & Daily OTP Frame (Shared by Jobcard & Equipment Tabs)
        cred_frame = ttk.LabelFrame(main_container, text=" 🔐 Portal Credentials & Daily OTP ", padding=6)
        cred_frame.pack(fill=tk.X, pady=(0, 6))

        c_row1 = ttk.Frame(cred_frame)
        c_row1.pack(fill=tk.X)

        ttk.Label(c_row1, text="Username:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.user_entry = ttk.Entry(c_row1, textvariable=self.username_var, width=15, font=("Segoe UI", 9))
        self.user_entry.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(c_row1, text="Password:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.pwd_entry = ttk.Entry(c_row1, textvariable=self.password_var, width=14, font=("Segoe UI", 9), show="*")
        self.pwd_entry.pack(side=tk.LEFT, padx=(0, 4))

        self.show_pwd_btn = ttk.Button(c_row1, text="👁️", width=3, command=self._toggle_show_password)
        self.show_pwd_btn.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(c_row1, text="Today's OTP:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.otp_entry = ttk.Entry(c_row1, textvariable=self.daily_otp_var, width=10, font=("Segoe UI", 9, "bold"))
        self.otp_entry.pack(side=tk.LEFT, padx=(0, 4))

        self.otp_status_lbl = tk.Label(c_row1, text="Checking...", font=("Segoe UI", 8, "bold"), padx=4, pady=1)
        self.otp_status_lbl.pack(side=tk.LEFT, padx=(0, 12))

        save_cred_btn = ttk.Button(c_row1, text="💾 Save Credentials", command=self._save_credentials_click)
        save_cred_btn.pack(side=tk.LEFT, padx=(0, 4))

        clear_otp_btn = ttk.Button(c_row1, text="↺ Clear OTP", command=self._clear_otp_click)
        clear_otp_btn.pack(side=tk.LEFT)

        # ── Mode Notebook (Tab 1: Jobcard Auto-Filler | Tab 2: Loco Equipment & Shop Floor) ──
        self.mode_notebook = ttk.Notebook(main_container)
        self.mode_notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        self.jobcard_tab = ttk.Frame(self.mode_notebook, padding=4)
        self.equipment_tab = ttk.Frame(self.mode_notebook, padding=4)

        self.mode_notebook.add(self.jobcard_tab, text=" 📋 Jobcard Auto-Filler ")
        self.mode_notebook.add(self.equipment_tab, text=" 🔧 Loco Equipment & Shop Floor ")

        # ═══════════════════════════════════════════════════════════════════════
        # TAB 1: JOBCARD AUTO-FILLER
        # ═══════════════════════════════════════════════════════════════════════

        # 2. File Selection Frame
        file_frame = ttk.LabelFrame(self.jobcard_tab, text=" 📁 Select Bookings Excel File ", padding=6)
        file_frame.pack(fill=tk.X, pady=(0, 6))

        # Quick detected files dropdown
        quick_frame = ttk.Frame(file_frame)
        quick_frame.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(quick_frame, text="Quick Select:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))
        self.detected_files_cb = ttk.Combobox(quick_frame, state="readonly", width=58, font=("Segoe UI", 9))
        self.detected_files_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.detected_files_cb.bind("<<ComboboxSelected>>", self._on_detected_file_selected)

        scan_btn = ttk.Button(quick_frame, text="🔄 Scan", command=self._scan_desktop_files)
        scan_btn.pack(side=tk.LEFT)

        # Manual path entry
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X)

        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, font=("Segoe UI", 9))
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        browse_btn = ttk.Button(path_frame, text="Browse Excel...", command=self._browse_file)
        browse_btn.pack(side=tk.LEFT, padx=3)

        load_btn = ttk.Button(path_frame, text="Reload", command=self._reload_file)
        load_btn.pack(side=tk.LEFT, padx=3)

        # File & Locomotive Info Card
        info_frame = ttk.Frame(file_frame)
        info_frame.pack(fill=tk.X, pady=(6, 0))

        ttk.Label(info_frame, text="Locomotive:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.loco_var = tk.StringVar(value="--")
        self.loco_entry = ttk.Entry(info_frame, textvariable=self.loco_var, width=10, font=("Segoe UI", 10, "bold"))
        self.loco_entry.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(info_frame, text="Schedule:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.sched_var = tk.StringVar(value="--")
        self.sched_lbl = tk.Label(info_frame, textvariable=self.sched_var, font=("Segoe UI", 9, "bold"), fg="#2b6cb0", bg="#ebf8ff", padx=6, pady=2)
        self.sched_lbl.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(info_frame, text="Date:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.date_var = tk.StringVar(value="--")
        ttk.Label(info_frame, textvariable=self.date_var, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(info_frame, text="Total Bookings:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.count_var = tk.StringVar(value="0")
        self.count_lbl = tk.Label(info_frame, textvariable=self.count_var, font=("Segoe UI", 10, "bold"), fg="#2f855a", bg="#f0fff4", padx=8, pady=2)
        self.count_lbl.pack(side=tk.LEFT, padx=(0, 14))

        self.sec_breakdown_var = tk.StringVar(value="Sections: --")
        self.sec_breakdown_lbl = tk.Label(info_frame, textvariable=self.sec_breakdown_var, font=("Segoe UI", 8, "italic"), fg="#4a5568")
        self.sec_breakdown_lbl.pack(side=tk.LEFT)

        # 3. Jobcard Settings Frame (Exact User Presets)
        settings_frame = ttk.LabelFrame(self.jobcard_tab, text=" ⚙️ Jobcard Creation Settings (Pre-Configured) ", padding=6)
        settings_frame.pack(fill=tk.X, pady=(0, 6))

        s_row1 = ttk.Frame(settings_frame)
        s_row1.pack(fill=tk.X, pady=2)

        ttk.Label(s_row1, text="Jobcard Type:", width=16, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.jobcard_type_var = tk.StringVar(value=config.DEFAULT_JOBCARD_TYPE)
        jobcard_type_cb = ttk.Combobox(s_row1, textvariable=self.jobcard_type_var, values=list(config.JOBCARD_TYPE_VALUES.keys()), state="readonly", width=22)
        jobcard_type_cb.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(s_row1, text="Eqt Required:", width=14, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.eqt_req_var = tk.StringVar(value=config.DEFAULT_EQT_REQUIRED)
        eqt_cb = ttk.Combobox(s_row1, textvariable=self.eqt_req_var, values=["Jobcard Without Eq", "Jobcard With Eq"], state="readonly", width=20)
        eqt_cb.pack(side=tk.LEFT, padx=(0, 20))

        self.headless_var = tk.BooleanVar(value=False)
        headless_chk = ttk.Checkbutton(s_row1, text="Headless (Hide Browser)", variable=self.headless_var)
        headless_chk.pack(side=tk.LEFT, padx=(10, 8))

        keep_browser_chk = ttk.Checkbutton(s_row1, text="Keep Browser Open", variable=self.keep_browser_var)
        keep_browser_chk.pack(side=tk.LEFT, padx=4)

        s_row2 = ttk.Frame(settings_frame)
        s_row2.pack(fill=tk.X, pady=(4, 2))

        ttk.Label(s_row2, text="Source (Obs Type):", width=16, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.obs_type_var = tk.StringVar(value=config.DEFAULT_OBS_TYPE)
        obs_type_cb = ttk.Combobox(s_row2, textvariable=self.obs_type_var, values=list(config.OBS_TYPE_VALUES.keys()), state="readonly", width=20)
        obs_type_cb.pack(side=tk.LEFT, padx=(0, 6))

        apply_obs_btn = ttk.Button(s_row2, text="⚡ Apply to All Entries", command=self._apply_obs_type_to_all)
        apply_obs_btn.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(s_row2, text="Kind:", width=6, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.job_kind_var = tk.StringVar(value=config.DEFAULT_JOB_KIND)
        job_kind_cb = ttk.Combobox(s_row2, textvariable=self.job_kind_var, values=list(config.JOB_KIND_VALUES.keys()), state="readonly", width=12)
        job_kind_cb.pack(side=tk.LEFT, padx=(0, 14))

        reset_btn = ttk.Button(s_row2, text="↺ Reset Defaults", command=self._reset_settings_defaults)
        reset_btn.pack(side=tk.LEFT, padx=4)

        # 3b. Browser Selection & Preferences Toolbar
        s_row3 = ttk.Frame(settings_frame)
        s_row3.pack(fill=tk.X, pady=(6, 2))

        ttk.Label(s_row3, text="🌐 Browser:", width=11, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.browser_var = tk.StringVar(value="")
        browser_names = [config.BROWSER_DISPLAY_NAMES[k] for k in config.BROWSER_CHOICES]
        self.browser_cb = ttk.Combobox(
            s_row3,
            textvariable=self.browser_var,
            values=browser_names,
            state="readonly",
            width=26,
            font=("Segoe UI", 9)
        )
        self.browser_cb.pack(side=tk.LEFT, padx=(0, 14))
        self.browser_cb.bind("<<ComboboxSelected>>", self._on_browser_combobox_changed)

        ttk.Label(s_row3, text="⭐ 'All Sec':", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.all_sec_badge_var = tk.StringVar(value="")
        self.all_sec_badge_lbl = tk.Label(
            s_row3,
            textvariable=self.all_sec_badge_var,
            font=("Segoe UI", 8, "bold"),
            fg="#22543d",
            bg="#c6f6d5",
            padx=8,
            pady=2
        )
        self.all_sec_badge_lbl.pack(side=tk.LEFT, padx=(0, 10))

        pref_btn = tk.Button(
            s_row3,
            text="⚙️ Preferences...",
            font=("Segoe UI", 8, "bold"),
            bg="#edf2f7",
            fg="#2b6cb0",
            activebackground="#e2e8f0",
            padx=10,
            pady=2,
            cursor="hand2",
            command=self._open_preferences
        )
        pref_btn.pack(side=tk.LEFT, padx=(0, 6))

        all_sec_btn = tk.Button(
            s_row3,
            text="⭐ 'All Sec' Rules...",
            font=("Segoe UI", 8),
            bg="#edf2f7",
            fg="#4a5568",
            activebackground="#e2e8f0",
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._open_all_sec_preferences
        )
        all_sec_btn.pack(side=tk.LEFT)

        # 4. Bookings Preview & Direct Inline Editor Table
        table_frame = ttk.LabelFrame(
            self.jobcard_tab,
            text=" 📋 Bookings to be Staged into SLAM (Click Any Cell or Use the Bar Below to Edit Defects, Sections, Source & Kind Directly) ",
            padding=6
        )
        table_frame.pack(fill=tk.BOTH, expand=True)

        tbl_container = ttk.Frame(table_frame)
        tbl_container.pack(fill=tk.BOTH, expand=True)

        cols = ("sno", "desc", "section", "target_sections", "obs_type", "kind")
        self.tree = ttk.Treeview(tbl_container, columns=cols, show="headings", height=5)
        self.tree.heading("sno", text="✏️ S.No", anchor=tk.CENTER)
        self.tree.heading("desc", text="✏️ Booking / Defect Description", anchor=tk.W)
        self.tree.heading("section", text="✏️ Excel Section", anchor=tk.CENTER)
        self.tree.heading("target_sections", text="✏️ Mapped SLAM Section(s)", anchor=tk.CENTER)
        self.tree.heading("obs_type", text="✏️ Source (Obs Type)", anchor=tk.CENTER)
        self.tree.heading("kind", text="✏️ Kind", anchor=tk.CENTER)

        self.tree.column("sno", width=52, stretch=False, anchor=tk.CENTER)
        self.tree.column("desc", width=410, stretch=True, anchor=tk.W)
        self.tree.column("section", width=105, stretch=False, anchor=tk.CENTER)
        self.tree.column("target_sections", width=160, stretch=False, anchor=tk.CENTER)
        self.tree.column("obs_type", width=135, stretch=False, anchor=tk.CENTER)
        self.tree.column("kind", width=70, stretch=False, anchor=tk.CENTER)

        self.tree.tag_configure("even", background="#f7fafc")
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("warn", background="#fff5f5", foreground="#c53030")

        scroll_y = ttk.Scrollbar(tbl_container, orient=tk.VERTICAL, command=self.tree.yview)
        def _on_tree_yscroll(*args):
            self._close_inline_editor(commit=True)
            scroll_y.set(*args)
        self.tree.configure(yscrollcommand=_on_tree_yscroll)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Direct Row Editor Panel (Always visible for instant popup-free editing of Defect, Source, Kind & Section)
        direct_edit_bar = tk.Frame(table_frame, bg="#f7fafc", relief=tk.SOLID, borderwidth=1, padx=6, pady=4)
        direct_edit_bar.pack(fill=tk.X, pady=(5, 0))

        # Row 1: Defect Description, Source (Obs Type), and Kind
        direct_row1 = tk.Frame(direct_edit_bar, bg="#f7fafc")
        direct_row1.pack(fill=tk.X, pady=(0, 3))

        self.direct_row_badge = tk.Label(
            direct_row1,
            text="No Row",
            font=("Segoe UI", 8, "bold"),
            bg="#e2e8f0",
            fg="#2d3748",
            padx=6,
            pady=1
        )
        self.direct_row_badge.pack(side=tk.LEFT, padx=(0, 6))

        tk.Label(
            direct_row1,
            text="✏️ Defect:",
            font=("Segoe UI", 8, "bold"),
            bg="#f7fafc",
            fg="#1a365d"
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.direct_desc_var = tk.StringVar(value="")
        self.direct_desc_entry = ttk.Entry(
            direct_row1,
            textvariable=self.direct_desc_var,
            font=("Segoe UI", 9)
        )
        self.direct_desc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.direct_desc_var.trace_add("write", self._on_direct_desc_live_change)
        self.direct_desc_entry.bind("<Return>", self._on_direct_desc_commit)
        self.direct_desc_entry.bind("<KP_Enter>", self._on_direct_desc_commit)

        tk.Label(
            direct_row1,
            text="Source:",
            font=("Segoe UI", 8, "bold"),
            bg="#f7fafc",
            fg="#1a365d"
        ).pack(side=tk.LEFT, padx=(0, 3))

        self.direct_obs_var = tk.StringVar(value=config.DEFAULT_OBS_TYPE)
        self.direct_obs_cb = ttk.Combobox(
            direct_row1,
            textvariable=self.direct_obs_var,
            values=list(config.OBS_TYPE_VALUES.keys()),
            state="readonly",
            width=17,
            font=("Segoe UI", 8, "bold")
        )
        self.direct_obs_cb.pack(side=tk.LEFT, padx=(0, 8))
        self.direct_obs_cb.bind("<<ComboboxSelected>>", self._on_direct_obs_changed)

        tk.Label(
            direct_row1,
            text="Kind:",
            font=("Segoe UI", 8, "bold"),
            bg="#f7fafc",
            fg="#1a365d"
        ).pack(side=tk.LEFT, padx=(0, 3))

        self.direct_kind_var = tk.StringVar(value=config.DEFAULT_JOB_KIND)
        self.direct_kind_cb = ttk.Combobox(
            direct_row1,
            textvariable=self.direct_kind_var,
            values=list(config.JOB_KIND_VALUES.keys()),
            state="readonly",
            width=8,
            font=("Segoe UI", 8, "bold")
        )
        self.direct_kind_cb.pack(side=tk.LEFT)
        self.direct_kind_cb.bind("<<ComboboxSelected>>", self._on_direct_kind_changed)

        # Row 2: Direct Section Text + 1-Click Section Toggle Chips
        direct_row2 = tk.Frame(direct_edit_bar, bg="#f7fafc")
        direct_row2.pack(fill=tk.X)

        tk.Label(
            direct_row2,
            text="🎯 Section:",
            font=("Segoe UI", 8, "bold"),
            bg="#f7fafc",
            fg="#1a365d"
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.direct_sec_var = tk.StringVar(value="")
        self.direct_sec_entry = ttk.Entry(direct_row2, textvariable=self.direct_sec_var, width=14, font=("Segoe UI", 9, "bold"))
        self.direct_sec_entry.pack(side=tk.LEFT, padx=(0, 4))
        self.direct_sec_entry.bind("<Return>", self._on_direct_sec_entry_commit)
        self.direct_sec_entry.bind("<KP_Enter>", self._on_direct_sec_entry_commit)

        apply_sec_btn = tk.Button(
            direct_row2,
            text="✓ Set",
            font=("Segoe UI", 8, "bold"),
            bg="#2b6cb0",
            fg="white",
            activebackground="#2c5282",
            activeforeground="white",
            padx=6,
            pady=1,
            cursor="hand2",
            command=self._on_direct_sec_entry_commit
        )
        apply_sec_btn.pack(side=tk.LEFT, padx=(0, 8))

        tk.Label(
            direct_row2,
            text="Toggle:",
            font=("Segoe UI", 8, "bold"),
            bg="#f7fafc",
            fg="#4a5568"
        ).pack(side=tk.LEFT, padx=(0, 3))

        self._direct_chip_btns = {}
        for s in config.ALL_SECTIONS_LIST:
            b = tk.Button(
                direct_row2,
                text=s,
                font=("Segoe UI", 8, "bold"),
                bg="#edf2f7",
                fg="#2d3748",
                activebackground="#cbd5e0",
                padx=4,
                pady=1,
                cursor="hand2",
                command=lambda sec=s: self._on_direct_chip_click(sec)
            )
            b.pack(side=tk.LEFT, padx=1)
            self._direct_chip_btns[s] = b

        self._direct_all_sec_btn = tk.Button(
            direct_row2,
            text="⭐ All Sec",
            font=("Segoe UI", 8, "bold"),
            bg="#feebc8",
            fg="#7b341e",
            activebackground="#fbd38d",
            padx=5,
            pady=1,
            cursor="hand2",
            command=self._on_direct_all_sec_click
        )
        self._direct_all_sec_btn.pack(side=tk.LEFT, padx=(4, 2))

        # Toolbar below the table
        tbl_toolbar = ttk.Frame(table_frame)
        tbl_toolbar.pack(fill=tk.X, pady=(5, 0))

        edit_btn = tk.Button(
            tbl_toolbar,
            text="✏️ Edit Defect Inline",
            font=("Segoe UI", 9, "bold"),
            bg="#2b6cb0",
            fg="white",
            activebackground="#2c5282",
            activeforeground="white",
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._on_edit_booking_click
        )
        edit_btn.pack(side=tk.LEFT, padx=(0, 6))

        edit_sec_btn = tk.Button(
            tbl_toolbar,
            text="🎯 Edit Section Inline",
            font=("Segoe UI", 8, "bold"),
            bg="#edf2f7",
            fg="#2b6cb0",
            activebackground="#cbd5e0",
            padx=8,
            pady=3,
            cursor="hand2",
            command=self._edit_selected_section_inline
        )
        edit_sec_btn.pack(side=tk.LEFT, padx=(0, 6))

        add_btn = ttk.Button(tbl_toolbar, text="➕ Add Booking Inline", command=self._on_add_booking_click)
        add_btn.pack(side=tk.LEFT, padx=(0, 6))

        del_btn = ttk.Button(tbl_toolbar, text="🗑️ Delete Selected", command=self._on_delete_booking_click)
        del_btn.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(
            tbl_toolbar,
            text="💡 Tip: Click ANY cell (Defect, Section, Source, Kind, S.No) in the table or use the editor bar above to edit directly without popups!",
            font=("Segoe UI", 8, "italic"),
            fg="#4a5568"
        ).pack(side=tk.LEFT)

        # Bindings
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Button-1>", self._on_tree_single_click, add="+")
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Return>", lambda e: self._on_edit_booking_click())

        # Right-click context menu
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="✏️ Edit Defect Description Inline", command=self._on_edit_booking_click)
        self.tree_menu.add_command(label="🎯 Edit Section Inline", command=self._edit_selected_section_inline)
        self.quick_sec_menu = tk.Menu(self.tree_menu, tearoff=0)
        self.quick_sec_menu.add_command(label="⭐ All Sec", command=lambda: self._quick_set_selected_section("All Sec"))
        self.quick_sec_menu.add_separator()
        for s in config.ALL_SECTIONS_LIST:
            self.quick_sec_menu.add_command(label=f"Set only {s}", command=lambda sec=s: self._quick_set_selected_section(sec))
        self.quick_sec_menu.add_separator()
        for s in config.ALL_SECTIONS_LIST:
            self.quick_sec_menu.add_command(label=f"Toggle +/{s}", command=lambda sec=s: self._on_direct_chip_click(sec))
        self.tree_menu.add_cascade(label="🎯 Quick Set / Toggle Section", menu=self.quick_sec_menu)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label="➕ Add New Booking Inline", command=self._on_add_booking_click)
        self.tree_menu.add_command(label="🗑️ Delete Selected Booking", command=self._on_delete_booking_click)
        self.tree.bind("<Button-3>", self._show_tree_context_menu)

        # ═══════════════════════════════════════════════════════════════════════
        # TAB 2: LOCO EQUIPMENT & SHOP FLOOR
        # ═══════════════════════════════════════════════════════════════════════
        self._build_equipment_tab()

        # 5. Progress and Controls Bar (Shared)
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=(0, 6))

        self.start_btn = tk.Button(
            action_frame,
            text="🚀 START JOBCARD AUTO-FILL",
            font=("Segoe UI", 10, "bold"),
            bg="#2f855a",
            fg="white",
            activebackground="#276749",
            activeforeground="white",
            padx=14,
            pady=7,
            cursor="hand2",
            command=self._start_autofill
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.start_eq_btn = tk.Button(
            action_frame,
            text="🔧 START EQUIPMENT AUTO-FILL",
            font=("Segoe UI", 10, "bold"),
            bg="#2b6cb0",
            fg="white",
            activebackground="#2c5282",
            activeforeground="white",
            padx=14,
            pady=7,
            cursor="hand2",
            command=self._start_equipment_autofill
        )
        self.start_eq_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = tk.Button(
            action_frame,
            text="⏹ Stop Process",
            font=("Segoe UI", 9, "bold"),
            bg="#c53030",
            fg="white",
            activebackground="#9b2c2c",
            activeforeground="white",
            padx=10,
            pady=7,
            state=tk.DISABLED,
            cursor="hand2",
            command=self._stop_autofill
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.close_browser_btn = tk.Button(
            action_frame,
            text="🌐 Close Browser",
            font=("Segoe UI", 9, "bold"),
            bg="#4a5568",
            fg="white",
            activebackground="#2d3748",
            activeforeground="white",
            padx=10,
            pady=7,
            cursor="hand2",
            command=self._close_browser_manually
        )
        self.close_browser_btn.pack(side=tk.LEFT, padx=(0, 12))

        self.progress_bar = ttk.Progressbar(action_frame, orient=tk.HORIZONTAL, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.status_lbl = ttk.Label(action_frame, text="Ready", font=("Segoe UI", 9, "italic"))
        self.status_lbl.pack(side=tk.RIGHT)

        # 6. Real-Time Activity Log
        log_frame = ttk.LabelFrame(main_container, text=" 📝 Activity & Execution Log ", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=6, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4")
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        clear_log_btn = ttk.Button(log_frame, text="Clear Log", command=self._clear_log)
        clear_log_btn.pack(side=tk.RIGHT, anchor=tk.NE, padx=4)

    def _toggle_show_password(self):
        show = self.show_pwd_var.get()
        if show:
            self.pwd_entry.config(show="*")
            self.show_pwd_btn.config(text="👁️")
            self.show_pwd_var.set(False)
        else:
            self.pwd_entry.config(show="")
            self.show_pwd_btn.config(text="🔒")
            self.show_pwd_var.set(True)

    def _update_otp_status(self):
        cached = config.get_daily_otp()
        if cached:
            today_str = datetime.now().strftime("%d-%m-%Y")
            self.otp_status_lbl.config(text=f"[Active Today: {today_str}]", fg="#22543d", bg="#c6f6d5")
        else:
            self.otp_status_lbl.config(text="[Not Set (Will Prompt)]", fg="#744210", bg="#fefcbf")

    def _save_credentials_click(self):
        u = self.username_var.get().strip()
        p = self.password_var.get().strip()
        otp = self.daily_otp_var.get().strip()
        if not u or not p:
            messagebox.showwarning("Missing Credentials", "Username and Password cannot be empty!")
            return
        config.save_credentials(u, p)
        if otp:
            config.set_daily_otp(otp)
        self._update_otp_status()
        self.log(f"Credentials saved! Username: '{u}', Password: [updated], Today's OTP: '{otp or 'none'}'")
        messagebox.showinfo("Saved", "Credentials and Daily OTP saved successfully!")

    def _clear_otp_click(self):
        config.set_daily_otp("")
        self.daily_otp_var.set("")
        self._update_otp_status()
        self.log("Daily OTP cleared. The bot will prompt for a fresh OTP on next login.")

    def _clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def _reset_settings_defaults(self):
        self.jobcard_type_var.set(config.DEFAULT_JOBCARD_TYPE)
        self.eqt_req_var.set(config.DEFAULT_EQT_REQUIRED)
        self.obs_type_var.set(config.DEFAULT_OBS_TYPE)
        self.job_kind_var.set(config.DEFAULT_JOB_KIND)
        self.headless_var.set(False)
        self.keep_browser_var.set(True)
        config.save_browser_preference(config.BROWSER_CHOICE_AUTO, "")
        config.save_all_sec_preferences(config.DEFAULT_ALL_SEC_INCLUDED)
        self._sync_preferences()
        self.log("Settings reset to RPME defaults (Browser: Auto-Detect | All Sec excludes E8, Q&I, LAB, PPIO).")

    def _get_selected_browser_key(self) -> str:
        sel = self.browser_var.get()
        for k, v in config.BROWSER_DISPLAY_NAMES.items():
            if v == sel:
                return k
        return config.BROWSER_CHOICE_AUTO

    def _on_browser_combobox_changed(self, event=None):
        k = self._get_selected_browser_key()
        config.save_browser_preference(k)
        name = config.BROWSER_DISPLAY_NAMES.get(k, k)
        self.log(f"Default browser updated to: {name}")

    def _open_preferences(self):
        """Opens the comprehensive Preferences dialog on the Web Browser tab."""
        PreferencesDialog(self.root, gui_ref=self, initial_tab=0)

    def _open_all_sec_preferences(self):
        """Opens the modal dialog allowing the user to select target sections for 'All Sec'."""
        PreferencesDialog(self.root, gui_ref=self, initial_tab=1)

    def _sync_preferences(self):
        """Synchronizes GUI badges, dropdowns, and active bookings with user_settings.json."""
        b_pref = config.get_browser_preference()
        b_choice = b_pref.get("choice", config.BROWSER_CHOICE_AUTO)
        b_name = config.BROWSER_DISPLAY_NAMES.get(b_choice, "Auto-Detect / System Default")
        if hasattr(self, "browser_var"):
            self.browser_var.set(b_name)
        self._update_all_sec_badge()
        self._reapply_all_sec_preferences()

    def _update_all_sec_badge(self):
        """Updates the badge in the settings bar displaying current target sections for All Sec."""
        secs = config.get_all_sec_preferences()
        self.all_sec_badge_var.set(f"Targets ({len(secs)}): {', '.join(secs)}")

    def _reapply_all_sec_preferences(self):
        """
        When 'All Sec' preferences are updated, refreshes any currently loaded bookings
        that have raw_section set to 'All Sec'.
        """
        self._update_all_sec_badge()
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return

        from excel_parser import is_all_sec_string
        updated_count = 0
        all_sec_pref = config.get_all_sec_preferences()

        for b in self.parsed_data["bookings"]:
            raw = b.get("raw_section", "")
            if is_all_sec_string(raw):
                b["sections"] = list(all_sec_pref)
                updated_count += 1

        if updated_count > 0:
            self._refresh_bookings_table()
            self.log(f"Re-applied 'All Sec' preferences: Updated {updated_count} booking(s) to [{', '.join(all_sec_pref)}].")

    def _close_browser_manually(self):
        self.worker.task_queue.put({"action": "CLOSE_BROWSER"})

    def _on_app_close(self):
        self.worker.task_queue.put({"action": "SHUTDOWN"})
        self.root.after(150, self.root.destroy)

    def log(self, message: str):
        """Thread-safe logging into the GUI text widget with timestamps."""
        def _append():
            try:
                ts = datetime.now().strftime("%H:%M:%S")
                self.log_text.insert(tk.END, f"[{ts}] {message}\n")
                self.log_text.see(tk.END)
            except Exception:
                pass
        try:
            self.root.after(0, _append)
        except Exception:
            pass

    def _scan_desktop_files(self):
        """Scans Desktop, Telegram Desktop, and Downloads for .xlsx booking files."""
        scan_roots = [
            (os.path.expanduser(r"~\Desktop"), ""),
            (os.path.expanduser(r"~\Downloads\Telegram Desktop"), "Telegram"),
            (os.path.expanduser(r"~\Downloads"), "Downloads"),
        ]

        candidates = []
        for root_dir, prefix in scan_roots:
            if not os.path.exists(root_dir):
                continue
            try:
                for entry in os.scandir(root_dir):
                    try:
                        if entry.is_file() and entry.name.endswith(".xlsx") and not entry.name.startswith("~$"):
                            if any(kw in entry.name.upper() for kw in ["BOOKING", "LBR", "QAI", "CHECKING"]):
                                disp = f"[{prefix}] {entry.name}" if prefix else entry.name
                                candidates.append((entry.stat().st_mtime, entry.path, disp))
                        elif entry.is_dir() and not entry.name.startswith(".") and not prefix:
                            # Only scan immediate subdirectories on Desktop to keep scan fast
                            for sub in os.scandir(entry.path):
                                if sub.is_file() and sub.name.endswith(".xlsx") and not sub.name.startswith("~$"):
                                    if any(kw in sub.name.upper() for kw in ["BOOKING", "LBR", "QAI", "CHECKING"]):
                                        display_name = f"[{entry.name}] {sub.name}"
                                        candidates.append((sub.stat().st_mtime, sub.path, display_name))
                    except Exception:
                        pass
            except Exception:
                pass

        candidates.sort(reverse=True)
        self._desktop_files = {disp: full_p for _, full_p, disp in candidates}

        if candidates:
            vals = [disp for _, _, disp in candidates]
            self.detected_files_cb["values"] = vals
            self.detected_files_cb.current(0)
            self._on_detected_file_selected()
        else:
            self.detected_files_cb["values"] = ["(No matching Excel files detected)"]

    def _on_detected_file_selected(self, event=None):
        selected_name = self.detected_files_cb.get()
        if hasattr(self, "_desktop_files") and selected_name in self._desktop_files:
            file_path = self._desktop_files[selected_name]
            self.file_path_var.set(file_path)
            self._load_file(file_path)

    def _browse_file(self):
        desktop = os.path.expanduser(r"~\Desktop")
        f = filedialog.askopenfilename(
            initialdir=desktop,
            title="Select SLAM Bookings Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if f:
            self.file_path_var.set(f)
            self._load_file(f)

    def _reload_file(self):
        path = self.file_path_var.get().strip()
        if path:
            self._load_file(path)

    def _load_file(self, file_path: str):
        try:
            self.parsed_data = parse_excel_bookings(file_path)
            self.current_excel_path = file_path

            self.loco_var.set(self.parsed_data["loco_number"])
            self.sched_var.set(self.parsed_data["schedule"] or "UnSch / QAI")
            self.date_var.set(self.parsed_data["date"] or "--")

            self._refresh_bookings_table()

            self.status_lbl.config(text=f"Loaded {self.parsed_data['bookings_count']} bookings from {os.path.basename(file_path)}")
            self.log(f"Loaded file: {file_path} (Loco: {self.parsed_data['loco_number']}, Schedule: {self.parsed_data['schedule']}, Total: {self.parsed_data['bookings_count']})")
        except Exception as e:
            messagebox.showerror("Error Parsing Excel", f"Failed to load Excel file:\n{e}")
            self.log(f"Error loading Excel: {e}")

    def _get_row_values_tuple(self, idx: int, booking: dict) -> tuple:
        """Returns the 6-column Treeview values tuple for a booking row."""
        secs = booking.get("sections", [])
        sec_display = ", ".join(secs) if secs else "⚠️ Unrecognized"
        obs_type_val = booking.get("obs_type") or self.obs_type_var.get()
        kind_val = booking.get("kind") or self.job_kind_var.get()
        return (
            booking.get("sno", idx + 1),
            booking.get("description", ""),
            booking.get("raw_section", ""),
            sec_display,
            obs_type_val,
            kind_val
        )

    def _refresh_row_in_tree(self, idx: int):
        """Updates a single row in-place in the Treeview from self.parsed_data['bookings'][idx]."""
        if not self.parsed_data:
            return
        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return
        booking = bookings[idx]
        item_id = f"row_{idx}"
        if self.tree.exists(item_id):
            secs = booking.get("sections", [])
            tag = "warn" if not secs else ("even" if idx % 2 == 0 else "odd")
            self.tree.item(
                item_id,
                values=self._get_row_values_tuple(idx, booking),
                tags=(tag,)
            )

    def _refresh_bookings_table(self):
        """Refreshes the Treeview table, total count, and section breakdown from self.parsed_data."""
        self._close_inline_editor(commit=False)
        if not self.parsed_data:
            return

        prev_sel = self.tree.selection()

        for item in self.tree.get_children():
            self.tree.delete(item)

        bookings = self.parsed_data.get("bookings", [])
        self.count_var.set(str(len(bookings)))
        self.parsed_data["bookings_count"] = len(bookings)

        for idx, b in enumerate(bookings):
            secs = b.get("sections", [])
            tag = "warn" if not secs else ("even" if idx % 2 == 0 else "odd")
            self.tree.insert(
                "",
                tk.END,
                iid=f"row_{idx}",
                values=self._get_row_values_tuple(idx, b),
                tags=(tag,)
            )

        self._update_section_breakdown_summary()

        # Restore previous selection or select first row so Direct Editor Bar is ready
        restored = [iid for iid in prev_sel if self.tree.exists(iid)]
        if restored:
            self.tree.selection_set(restored)
        elif bookings and self.tree.exists("row_0"):
            self.tree.selection_set("row_0")

        self._sync_direct_section_bar()

    def _update_section_breakdown_summary(self):
        """Updates the top section breakdown summary label from self.parsed_data."""
        if not self.parsed_data:
            self.sec_breakdown_var.set("Sections: --")
            return
        sec_counts = collections.Counter()
        for b in self.parsed_data.get("bookings", []):
            sec_counts.update(b.get("sections", []))
        if sec_counts:
            sec_summary = "  |  ".join(f"{s}: {c}" for s, c in sec_counts.most_common())
            self.sec_breakdown_var.set(f"Sections: {sec_summary}")
        else:
            self.sec_breakdown_var.set("Sections: None detected")

    def _get_selected_row_indices(self) -> List[int]:
        """Returns sorted list of booking indices currently selected in the Treeview."""
        indices = []
        for item_id in self.tree.selection():
            try:
                idx = int(item_id.replace("row_", ""))
                indices.append(idx)
            except ValueError:
                try:
                    indices.append(self.tree.index(item_id))
                except Exception:
                    pass
        return sorted(set(indices))

    def _on_tree_select(self, event=None):
        """Called whenever Treeview selection changes; syncs the Direct Row Editor Panel."""
        self._sync_direct_section_bar()

    def _sync_direct_section_bar(self, idx: Optional[int] = None):
        """Updates the Direct Row Editor Panel (Defect, Source, Kind, Section & Chips) for the selected row."""
        if self._syncing_direct_bar:
            return
        self._syncing_direct_bar = True
        try:
            if not self.parsed_data or not self.parsed_data.get("bookings"):
                self.direct_row_badge.config(text="No Row", bg="#e2e8f0", fg="#2d3748")
                self.direct_desc_var.set("")
                self.direct_sec_var.set("")
                for s, btn in self._direct_chip_btns.items():
                    btn.config(bg="#edf2f7", fg="#2d3748")
                self._direct_all_sec_btn.config(bg="#feebc8", fg="#7b341e")
                return

            if idx is None:
                sel_indices = self._get_selected_row_indices()
                if not sel_indices:
                    self.direct_row_badge.config(text="Select Row", bg="#e2e8f0", fg="#2d3748")
                    self.direct_desc_var.set("")
                    self.direct_sec_var.set("")
                    for s, btn in self._direct_chip_btns.items():
                        btn.config(bg="#edf2f7", fg="#2d3748")
                    self._direct_all_sec_btn.config(bg="#feebc8", fg="#7b341e")
                    return
                idx = sel_indices[0]
            else:
                sel_indices = [idx]

            bookings = self.parsed_data.get("bookings", [])
            if idx < 0 or idx >= len(bookings):
                return

            booking = bookings[idx]
            sno = booking.get("sno", idx + 1)
            if len(sel_indices) > 1:
                self.direct_row_badge.config(text=f"{len(sel_indices)} Rows", bg="#bee3f8", fg="#1a365d")
            else:
                self.direct_row_badge.config(text=f"S.No {sno}", bg="#bee3f8", fg="#1a365d")

            self.direct_desc_var.set(booking.get("description", ""))
            self.direct_obs_var.set(booking.get("obs_type") or self.obs_type_var.get())
            self.direct_kind_var.set(booking.get("kind") or self.job_kind_var.get())

            raw_sec = booking.get("raw_section", "")
            active_secs = set(booking.get("sections", []))
            self.direct_sec_var.set(raw_sec)

            for s, btn in self._direct_chip_btns.items():
                if s in active_secs:
                    btn.config(bg="#2f855a", fg="white", activebackground="#276749", activeforeground="white")
                else:
                    btn.config(bg="#edf2f7", fg="#2d3748", activebackground="#cbd5e0", activeforeground="#1a202c")

            from excel_parser import is_all_sec_string
            if is_all_sec_string(raw_sec):
                self._direct_all_sec_btn.config(bg="#dd6b20", fg="white", activebackground="#c05621", activeforeground="white")
            else:
                self._direct_all_sec_btn.config(bg="#feebc8", fg="#7b341e", activebackground="#fbd38d", activeforeground="#7b341e")
        finally:
            self._syncing_direct_bar = False

    def _on_direct_desc_live_change(self, *args):
        """Live-updates the selected booking's defect description as the user types in the Direct Row Editor bar."""
        if self._syncing_direct_bar:
            return
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            return
        idx = sel_indices[0]
        bookings = self.parsed_data.get("bookings", [])
        if 0 <= idx < len(bookings):
            new_desc = self.direct_desc_var.get()
            bookings[idx]["description"] = new_desc
            self._refresh_row_in_tree(idx)
            if self._inline_item_id == f"row_{idx}" and self._inline_col == "#2" and hasattr(self, "_inline_desc_var"):
                if self._inline_desc_var.get() != new_desc:
                    self._inline_desc_var.set(new_desc)

    def _on_direct_desc_commit(self, event=None):
        """Logs the updated defect description when Enter is pressed in the Direct Row Editor bar."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            return
        idx = sel_indices[0]
        bookings = self.parsed_data.get("bookings", [])
        if 0 <= idx < len(bookings):
            new_desc = self.direct_desc_var.get().strip()
            bookings[idx]["description"] = new_desc
            self._refresh_row_in_tree(idx)
            self.log(f"Updated S.No {bookings[idx].get('sno', idx + 1)} Defect -> '{new_desc[:60]}'")

    def _on_direct_obs_changed(self, event=None):
        """Immediately updates Source (Obs Type) for all selected booking rows from the Direct Editor bar."""
        if self._syncing_direct_bar or not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            return
        new_obs = self.direct_obs_var.get()
        for idx in sel_indices:
            if 0 <= idx < len(self.parsed_data["bookings"]):
                self.parsed_data["bookings"][idx]["obs_type"] = new_obs
                self._refresh_row_in_tree(idx)
                self.log(f"Updated S.No {self.parsed_data['bookings'][idx].get('sno', idx + 1)} Source -> '{new_obs}'")

    def _on_direct_kind_changed(self, event=None):
        """Immediately updates Kind (Minor / Major) for all selected booking rows from the Direct Editor bar."""
        if self._syncing_direct_bar or not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            return
        new_kind = self.direct_kind_var.get()
        for idx in sel_indices:
            if 0 <= idx < len(self.parsed_data["bookings"]):
                self.parsed_data["bookings"][idx]["kind"] = new_kind
                self._refresh_row_in_tree(idx)
                self.log(f"Updated S.No {self.parsed_data['bookings'][idx].get('sno', idx + 1)} Kind -> '{new_kind}'")

    def _update_booking_section(self, idx: int, new_raw_sec: str, log_change: bool = True):
        """Updates a booking's raw_section and mapped sections in-place and refreshes the table row."""
        if not self.parsed_data or idx < 0 or idx >= len(self.parsed_data.get("bookings", [])):
            return
        from excel_parser import normalize_sections
        booking = self.parsed_data["bookings"][idx]
        raw_clean = new_raw_sec.strip()
        mapped_secs = normalize_sections(raw_clean)
        booking["raw_section"] = raw_clean
        booking["sections"] = mapped_secs

        self._refresh_row_in_tree(idx)
        self._update_section_breakdown_summary()
        self._sync_direct_section_bar(idx)

        if log_change:
            sec_log = ", ".join(mapped_secs) if mapped_secs else "⚠️ Unrecognized"
            self.log(f"Updated S.No {booking.get('sno', idx + 1)} section -> '{raw_clean}' (Mapped: {sec_log})")

    def _on_direct_sec_entry_commit(self, event=None):
        """Commits the text typed into the Direct Section Quick-Editor Bar to the selected row(s)."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            sel_indices = [0]
            if self.tree.exists("row_0"):
                self.tree.selection_set("row_0")

        new_raw = self.direct_sec_var.get().strip()
        for idx in sel_indices:
            self._update_booking_section(idx, new_raw, log_change=True)

    def _on_direct_chip_click(self, sec_code: str):
        """Toggles a SLAM section chip directly on the selected booking row(s)."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            messagebox.showinfo("No Bookings", "Please load an Excel file or click 'Add Booking Inline' first.")
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            sel_indices = [0]
            if self.tree.exists("row_0"):
                self.tree.selection_set("row_0")

        from excel_parser import normalize_sections
        for idx in sel_indices:
            booking = self.parsed_data["bookings"][idx]
            existing = list(booking.get("sections") or normalize_sections(booking.get("raw_section", "")))
            if sec_code in existing:
                existing.remove(sec_code)
            else:
                existing.append(sec_code)
            ordered = [s for s in config.ALL_SECTIONS_LIST if s in existing]
            for s in existing:
                if s not in ordered:
                    ordered.append(s)
            new_raw = " / ".join(ordered) if ordered else ""
            self._update_booking_section(idx, new_raw, log_change=True)
            if self._inline_item_id == f"row_{idx}" and hasattr(self, "_inline_sec_var"):
                self._inline_sec_var.set(new_raw)

    def _on_direct_all_sec_click(self):
        """Toggles 'All Sec' directly on the selected booking row(s)."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            messagebox.showinfo("No Bookings", "Please load an Excel file or click 'Add Booking Inline' first.")
            return
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            sel_indices = [0]
            if self.tree.exists("row_0"):
                self.tree.selection_set("row_0")

        from excel_parser import is_all_sec_string
        for idx in sel_indices:
            booking = self.parsed_data["bookings"][idx]
            if is_all_sec_string(booking.get("raw_section", "")):
                new_raw = ""
            else:
                new_raw = "All Sec"
            self._update_booking_section(idx, new_raw, log_change=True)
            if self._inline_item_id == f"row_{idx}" and hasattr(self, "_inline_sec_var"):
                self._inline_sec_var.set(new_raw)

    def _quick_set_selected_section(self, raw_sec: str):
        """Sets the exact section string on all selected rows via context menu."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return
        for idx in self._get_selected_row_indices():
            self._update_booking_section(idx, raw_sec, log_change=True)

    def _edit_selected_section_inline(self):
        """Opens the inline cell editor on the Excel Section column of the selected row."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            messagebox.showinfo("No Bookings", "Please load an Excel file or click 'Add Booking Inline' first.")
            return
        sel = self.tree.selection()
        item_id = sel[0] if sel else "row_0"
        if self.tree.exists(item_id):
            self.tree.selection_set(item_id)
            self.tree.see(item_id)
            self.root.update_idletasks()
            self._open_inline_section_editor(item_id, "#3")

    def _open_inline_cell(self, item_id: str, col: str):
        """Dispatches inline cell editing for any column (#1 S.No, #2 Defect, #3/#4 Section, #5 Source, #6 Kind)."""
        if col == "#1":
            self._open_inline_sno_editor(item_id, "#1")
        elif col == "#2":
            self._open_inline_desc_editor(item_id, "#2")
        elif col in ("#3", "#4"):
            self._open_inline_section_editor(item_id, col)
        elif col == "#5":
            self._open_inline_obs_editor(item_id, "#5")
        elif col == "#6":
            self._open_inline_kind_editor(item_id, "#6")

    def _on_tree_single_click(self, event):
        """Opens direct inline editor when clicking on any cell (#1 S.No, #2 Defect, #3/#4 Section, #5 Source, #6 Kind)."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            self._close_inline_editor(commit=True)
            return

        item_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item_id or not col:
            self._close_inline_editor(commit=True)
            return

        if self._inline_item_id == item_id and self._inline_col == col and self._inline_editor_frame:
            return

        self._close_inline_editor(commit=True)
        self.tree.selection_set(item_id)
        self.tree.focus(item_id)
        self.root.after(15, lambda i=item_id, c=col: self._open_inline_cell(i, c))

    def _on_tree_double_click(self, event):
        """Handles double-click: opens the direct inline editor on whichever cell was double-clicked (no popup)."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        item_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item_id or not col:
            return
        self._open_inline_cell(item_id, col)

    def _open_inline_sno_editor(self, item_id: str, col: str = "#1"):
        """Places a direct inline Entry over the S.No cell (#1)."""
        if not self.parsed_data or not self.tree.exists(item_id):
            return
        self._close_inline_editor(commit=True)

        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return

        bbox = self.tree.bbox(item_id, col)
        if not bbox:
            self.tree.see(item_id)
            self.root.update_idletasks()
            bbox = self.tree.bbox(item_id, col)
        if not bbox:
            return

        x, y, width, height = bbox
        booking = bookings[idx]
        orig_sno = str(booking.get("sno", idx + 1))

        self._inline_item_id = item_id
        self._inline_col = col
        self._inline_orig_val = orig_sno

        editor_frame = tk.Frame(self.tree, bg="#2b6cb0", bd=0)
        editor_frame.place(x=x, y=y, width=max(width, 45), height=height)
        self._inline_editor_frame = editor_frame

        self._inline_sno_var = tk.StringVar(value=orig_sno)
        entry = tk.Entry(
            editor_frame,
            textvariable=self._inline_sno_var,
            font=("Segoe UI", 9, "bold"),
            justify=tk.CENTER,
            bg="#fffff0",
            fg="#1a365d",
            relief=tk.FLAT,
            insertbackground="#1a365d"
        )
        entry.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        entry.select_range(0, tk.END)
        entry.focus_set()
        self._inline_entry = entry

        entry.bind("<Return>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<KP_Enter>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<Escape>", lambda e: self._close_inline_editor(commit=False))
        entry.bind("<Down>", lambda e, r=idx: self._inline_move_row(r + 1, "#1"))
        entry.bind("<Up>", lambda e, r=idx: self._inline_move_row(r - 1, "#1"))
        entry.bind("<Tab>", lambda e, r=idx: self._inline_move_cell(r, "#2"))
        entry.bind("<FocusOut>", self._on_inline_focus_out)

    def _open_inline_desc_editor(self, item_id: str, col: str = "#2"):
        """Places a direct inline Entry over the Booking / Defect Description cell (#2) for popup-free editing."""
        if not self.parsed_data or not self.tree.exists(item_id):
            return
        if self._inline_editor_frame and self._inline_item_id == item_id and self._inline_col == col:
            if self._inline_entry:
                try:
                    self._inline_entry.focus_set()
                except Exception:
                    pass
            return

        self._close_inline_editor(commit=True)

        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return

        bbox = self.tree.bbox(item_id, col)
        if not bbox:
            self.tree.see(item_id)
            self.root.update_idletasks()
            bbox = self.tree.bbox(item_id, col)
        if not bbox:
            return

        x, y, width, height = bbox
        booking = bookings[idx]
        orig_desc = booking.get("description", "")

        self._inline_item_id = item_id
        self._inline_col = col
        self._inline_orig_val = orig_desc

        editor_frame = tk.Frame(self.tree, bg="#2b6cb0", bd=0)
        editor_frame.place(x=x, y=y, width=max(width, 220), height=height)
        self._inline_editor_frame = editor_frame

        self._inline_desc_var = tk.StringVar(value=orig_desc)
        entry = tk.Entry(
            editor_frame,
            textvariable=self._inline_desc_var,
            font=("Segoe UI", 9, "bold"),
            justify=tk.LEFT,
            bg="#fffff0",
            fg="#1a202c",
            relief=tk.FLAT,
            insertbackground="#1a365d"
        )
        entry.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        entry.focus_set()
        entry.icursor(tk.END)
        self._inline_entry = entry

        def _on_inline_desc_change(*args):
            new_text = self._inline_desc_var.get()
            booking["description"] = new_text
            self._refresh_row_in_tree(idx)
            self._sync_direct_section_bar(idx)

        self._inline_desc_var.trace_add("write", _on_inline_desc_change)

        entry.bind("<Return>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<KP_Enter>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<Escape>", lambda e: self._close_inline_editor(commit=False))
        entry.bind("<Down>", lambda e, r=idx: self._inline_move_row(r + 1, "#2"))
        entry.bind("<Up>", lambda e, r=idx: self._inline_move_row(r - 1, "#2"))
        entry.bind("<Tab>", lambda e, r=idx: self._inline_move_cell(r, "#3"))
        entry.bind("<FocusOut>", self._on_inline_focus_out)

    def _open_inline_section_editor(self, item_id: str, col: str = "#3"):
        """Places a direct inline Entry and floating Section Chip Picker over a Treeview Section cell."""
        if not self.parsed_data or not self.tree.exists(item_id):
            return
        if self._inline_editor_frame and self._inline_item_id == item_id and self._inline_col == col:
            if self._inline_entry:
                try:
                    self._inline_entry.focus_set()
                except Exception:
                    pass
            return

        self._close_inline_editor(commit=True)

        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return

        bbox = self.tree.bbox(item_id, col)
        if not bbox:
            self.tree.see(item_id)
            self.root.update_idletasks()
            bbox = self.tree.bbox(item_id, col)
        if not bbox:
            return

        x, y, width, height = bbox
        booking = bookings[idx]
        raw_val = booking.get("raw_section", "")

        self._inline_item_id = item_id
        self._inline_col = col
        self._inline_orig_val = raw_val
        self._inline_clicking_chip = False

        # 1. Cell Overlay Entry
        editor_frame = tk.Frame(self.tree, bg="#2b6cb0", bd=0)
        editor_frame.place(x=x, y=y, width=max(width, 105), height=height)
        self._inline_editor_frame = editor_frame

        self._inline_sec_var = tk.StringVar(value=raw_val)
        entry = tk.Entry(
            editor_frame,
            textvariable=self._inline_sec_var,
            font=("Segoe UI", 9, "bold"),
            justify=tk.CENTER,
            bg="#fffff0",
            fg="#1a365d",
            relief=tk.FLAT,
            insertbackground="#1a365d"
        )
        entry.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        entry.select_range(0, tk.END)
        entry.focus_set()
        self._inline_entry = entry

        # 2. Floating Section Quick-Toggle Chip Bar attached to the cell
        chip_bar = tk.Frame(self.tree, bg="#ebf8ff", highlightbackground="#2b6cb0", highlightthickness=1, padx=4, pady=3)
        self._inline_chip_bar = chip_bar

        def _mark_chip_click(event=None):
            self._inline_clicking_chip = True

        chip_bar.bind("<ButtonPress-1>", _mark_chip_click)

        inline_chip_btns: Dict[str, tk.Button] = {}

        def _refresh_inline_chips(*args):
            from excel_parser import normalize_sections, is_all_sec_string
            curr_text = self._inline_sec_var.get().strip()
            active = set(normalize_sections(curr_text))
            for s_code, b_widget in inline_chip_btns.items():
                if s_code in active:
                    b_widget.config(bg="#2f855a", fg="white", activebackground="#276749", activeforeground="white")
                else:
                    b_widget.config(bg="#ffffff", fg="#2d3748", activebackground="#cbd5e0", activeforeground="#1a202c")
            if is_all_sec_string(curr_text):
                all_btn.config(bg="#dd6b20", fg="white")
            else:
                all_btn.config(bg="#feebc8", fg="#7b341e")
            # Live preview in table
            self._update_booking_section(idx, curr_text, log_change=False)

        def _toggle_inline_chip(sec_code: str):
            self._inline_clicking_chip = True
            from excel_parser import normalize_sections
            curr_text = self._inline_sec_var.get().strip()
            existing = list(normalize_sections(curr_text))
            if sec_code in existing:
                existing.remove(sec_code)
            else:
                existing.append(sec_code)
            ordered = [s for s in config.ALL_SECTIONS_LIST if s in existing]
            for s in existing:
                if s not in ordered:
                    ordered.append(s)
            new_text = " / ".join(ordered) if ordered else ""
            self._inline_sec_var.set(new_text)
            if self._inline_entry and self._inline_entry.winfo_exists():
                self._inline_entry.focus_set()
                self._inline_entry.icursor(tk.END)
            self.root.after(60, lambda: setattr(self, "_inline_clicking_chip", False))

        def _toggle_inline_all_sec():
            self._inline_clicking_chip = True
            from excel_parser import is_all_sec_string
            if is_all_sec_string(self._inline_sec_var.get()):
                self._inline_sec_var.set("")
            else:
                self._inline_sec_var.set("All Sec")
            if self._inline_entry and self._inline_entry.winfo_exists():
                self._inline_entry.focus_set()
                self._inline_entry.icursor(tk.END)
            self.root.after(60, lambda: setattr(self, "_inline_clicking_chip", False))

        for s in config.ALL_SECTIONS_LIST:
            b = tk.Button(
                chip_bar,
                text=s,
                font=("Segoe UI", 8, "bold"),
                bg="#ffffff",
                fg="#2d3748",
                padx=4,
                pady=1,
                relief=tk.RAISED,
                bd=1,
                cursor="hand2",
                takefocus=0,
                command=lambda sec=s: _toggle_inline_chip(sec)
            )
            b.bind("<ButtonPress-1>", _mark_chip_click)
            b.pack(side=tk.LEFT, padx=1)
            inline_chip_btns[s] = b

        all_btn = tk.Button(
            chip_bar,
            text="⭐ All Sec",
            font=("Segoe UI", 8, "bold"),
            bg="#feebc8",
            fg="#7b341e",
            padx=5,
            pady=1,
            relief=tk.RAISED,
            bd=1,
            cursor="hand2",
            takefocus=0,
            command=_toggle_inline_all_sec
        )
        all_btn.bind("<ButtonPress-1>", _mark_chip_click)
        all_btn.pack(side=tk.LEFT, padx=(4, 2))

        ok_btn = tk.Button(
            chip_bar,
            text="✓ OK",
            font=("Segoe UI", 8, "bold"),
            bg="#2b6cb0",
            fg="white",
            activebackground="#2c5282",
            activeforeground="white",
            padx=6,
            pady=1,
            relief=tk.FLAT,
            cursor="hand2",
            takefocus=0,
            command=lambda: self._close_inline_editor(commit=True)
        )
        ok_btn.bind("<ButtonPress-1>", _mark_chip_click)
        ok_btn.pack(side=tk.LEFT, padx=(4, 1))

        chip_bar.update_idletasks()
        req_w = chip_bar.winfo_reqwidth()
        req_h = chip_bar.winfo_reqheight()
        tw = self.tree.winfo_width()
        th = self.tree.winfo_height()
        bar_x = max(4, min(x - 100, tw - req_w - 4))
        bar_y = (y + height + 1) if (y + height + req_h + 2 <= th) else max(2, y - req_h - 1)
        chip_bar.place(x=bar_x, y=bar_y)

        _refresh_inline_chips()
        self._inline_sec_var.trace_add("write", _refresh_inline_chips)

        # Keyboard navigation & commit bindings
        entry.bind("<Return>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<KP_Enter>", lambda e: self._close_inline_editor(commit=True))
        entry.bind("<Escape>", lambda e: self._close_inline_editor(commit=False))
        entry.bind("<Down>", lambda e, r=idx, c=col: self._inline_move_row(r + 1, c))
        entry.bind("<Up>", lambda e, r=idx, c=col: self._inline_move_row(r - 1, c))
        entry.bind("<Tab>", lambda e, r=idx: self._inline_move_cell(r, "#5"))
        entry.bind("<FocusOut>", self._on_inline_focus_out)

    def _open_inline_obs_editor(self, item_id: str, col: str = "#5"):
        """Places a direct inline Combobox over the Source (Obs Type) cell (#5)."""
        if not self.parsed_data or not self.tree.exists(item_id):
            return
        self._close_inline_editor(commit=True)

        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return

        bbox = self.tree.bbox(item_id, col)
        if not bbox:
            self.tree.see(item_id)
            self.root.update_idletasks()
            bbox = self.tree.bbox(item_id, col)
        if not bbox:
            return
        x, y, width, height = bbox
        booking = bookings[idx]
        curr_obs = booking.get("obs_type", self.obs_type_var.get())

        self._inline_item_id = item_id
        self._inline_col = col
        self._inline_orig_val = curr_obs

        editor_frame = tk.Frame(self.tree, bg="#2b6cb0", bd=0)
        editor_frame.place(x=x, y=y, width=max(width, 130), height=height)
        self._inline_editor_frame = editor_frame

        obs_var = tk.StringVar(value=curr_obs)
        cb = ttk.Combobox(
            editor_frame,
            textvariable=obs_var,
            values=list(config.OBS_TYPE_VALUES.keys()),
            state="readonly",
            font=("Segoe UI", 8, "bold")
        )
        cb.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        cb.focus_set()
        self._inline_entry = cb

        def _on_obs_selected(event=None):
            new_obs = obs_var.get()
            booking["obs_type"] = new_obs
            self._refresh_row_in_tree(idx)
            self._sync_direct_section_bar(idx)
            self.log(f"Updated S.No {booking.get('sno', idx + 1)} Source -> '{new_obs}'")
            self._close_inline_editor(commit=True)

        cb.bind("<<ComboboxSelected>>", _on_obs_selected)
        cb.bind("<Return>", _on_obs_selected)
        cb.bind("<Escape>", lambda e: self._close_inline_editor(commit=False))
        cb.bind("<Down>", lambda e, r=idx: self._inline_move_row(r + 1, "#5"))
        cb.bind("<Up>", lambda e, r=idx: self._inline_move_row(r - 1, "#5"))
        cb.bind("<Tab>", lambda e, r=idx: self._inline_move_cell(r, "#6"))
        cb.bind("<FocusOut>", self._on_inline_focus_out)

    def _open_inline_kind_editor(self, item_id: str, col: str = "#6"):
        """Places a direct inline Combobox over the Kind cell (#6)."""
        if not self.parsed_data or not self.tree.exists(item_id):
            return
        self._close_inline_editor(commit=True)

        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        bookings = self.parsed_data.get("bookings", [])
        if idx < 0 or idx >= len(bookings):
            return

        bbox = self.tree.bbox(item_id, col)
        if not bbox:
            self.tree.see(item_id)
            self.root.update_idletasks()
            bbox = self.tree.bbox(item_id, col)
        if not bbox:
            return
        x, y, width, height = bbox
        booking = bookings[idx]
        curr_kind = booking.get("kind", self.job_kind_var.get())

        self._inline_item_id = item_id
        self._inline_col = col
        self._inline_orig_val = curr_kind

        editor_frame = tk.Frame(self.tree, bg="#2b6cb0", bd=0)
        editor_frame.place(x=x, y=y, width=max(width, 68), height=height)
        self._inline_editor_frame = editor_frame

        kind_var = tk.StringVar(value=curr_kind)
        cb = ttk.Combobox(
            editor_frame,
            textvariable=kind_var,
            values=list(config.JOB_KIND_VALUES.keys()),
            state="readonly",
            font=("Segoe UI", 8, "bold")
        )
        cb.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        cb.focus_set()
        self._inline_entry = cb

        def _on_kind_selected(event=None):
            new_kind = kind_var.get()
            booking["kind"] = new_kind
            self._refresh_row_in_tree(idx)
            self._sync_direct_section_bar(idx)
            self.log(f"Updated S.No {booking.get('sno', idx + 1)} Kind -> '{new_kind}'")
            self._close_inline_editor(commit=True)

        cb.bind("<<ComboboxSelected>>", _on_kind_selected)
        cb.bind("<Return>", _on_kind_selected)
        cb.bind("<Escape>", lambda e: self._close_inline_editor(commit=False))
        cb.bind("<Down>", lambda e, r=idx: self._inline_move_row(r + 1, "#6"))
        cb.bind("<Up>", lambda e, r=idx: self._inline_move_row(r - 1, "#6"))
        cb.bind("<Tab>", lambda e, r=idx: self._inline_move_row(r + 1, "#2"))
        cb.bind("<FocusOut>", self._on_inline_focus_out)

    def _inline_move_cell(self, row_idx: int, next_col: str):
        """Commits current inline edit and moves inline editor to another column in the same row."""
        self._close_inline_editor(commit=True)
        if not self.parsed_data:
            return "break"
        bookings = self.parsed_data.get("bookings", [])
        if 0 <= row_idx < len(bookings):
            item_id = f"row_{row_idx}"
            if self.tree.exists(item_id):
                self.tree.selection_set(item_id)
                self.tree.focus(item_id)
                self.tree.see(item_id)
                self.root.after(20, lambda: self._open_inline_cell(item_id, next_col))
        return "break"

    def _inline_move_row(self, next_idx: int, col: str):
        """Commits current inline edit and moves inline editor to next/previous row in the same column."""
        self._close_inline_editor(commit=True)
        if not self.parsed_data:
            return "break"
        bookings = self.parsed_data.get("bookings", [])
        if 0 <= next_idx < len(bookings):
            next_id = f"row_{next_idx}"
            if self.tree.exists(next_id):
                self.tree.selection_set(next_id)
                self.tree.focus(next_id)
                self.tree.see(next_id)
                self.root.after(20, lambda: self._open_inline_cell(next_id, col))
        return "break"

    def _on_inline_focus_out(self, event=None):
        """Delayed focus check so clicking inline section chips does not prematurely close the editor."""
        self.root.after(100, self._check_inline_focus)

    def _check_inline_focus(self):
        if self._inline_clicking_chip:
            return
        if not self._inline_editor_frame:
            return
        try:
            focused = self.root.focus_get()
            if focused is not None:
                w = focused
                while w is not None:
                    if w == self._inline_editor_frame or w == self._inline_chip_bar:
                        return
                    w = getattr(w, "master", None)
        except Exception:
            pass
        self._close_inline_editor(commit=True)

    def _close_inline_editor(self, commit: bool = True):
        """Commits or reverts active inline cell edit and destroys overlay widgets."""
        if not self._inline_editor_frame and not self._inline_chip_bar:
            return

        item_id = self._inline_item_id
        col = self._inline_col
        orig_val = self._inline_orig_val

        self._inline_item_id = None
        self._inline_col = None

        if item_id and self.parsed_data:
            try:
                idx = int(item_id.replace("row_", ""))
                bookings = self.parsed_data.get("bookings", [])
                if 0 <= idx < len(bookings):
                    booking = bookings[idx]
                    if col == "#1" and hasattr(self, "_inline_sno_var"):
                        if commit:
                            raw_sno = self._inline_sno_var.get().strip()
                            try:
                                new_sno = int(raw_sno)
                            except ValueError:
                                new_sno = raw_sno or (idx + 1)
                            booking["sno"] = new_sno
                            if str(new_sno) != str(orig_val):
                                self.log(f"Updated row #{idx + 1} S.No -> {new_sno}")
                        else:
                            try:
                                booking["sno"] = int(orig_val)
                            except ValueError:
                                booking["sno"] = orig_val
                        self._refresh_row_in_tree(idx)
                        self._sync_direct_section_bar(idx)
                    elif col == "#2" and hasattr(self, "_inline_desc_var"):
                        if commit:
                            new_desc = self._inline_desc_var.get().strip()
                            booking["description"] = new_desc
                            if new_desc != orig_val.strip():
                                self.log(f"Updated S.No {booking.get('sno', idx + 1)} Defect -> '{new_desc[:60]}'")
                        else:
                            booking["description"] = orig_val
                        self._refresh_row_in_tree(idx)
                        self._sync_direct_section_bar(idx)
                    elif col in ("#3", "#4") and hasattr(self, "_inline_sec_var"):
                        if commit:
                            new_val = self._inline_sec_var.get().strip()
                            changed = (new_val != orig_val.strip())
                            self._update_booking_section(idx, new_val, log_change=changed)
                        else:
                            self._update_booking_section(idx, orig_val, log_change=False)
            except Exception:
                pass

        if self._inline_editor_frame:
            try:
                self._inline_editor_frame.destroy()
            except Exception:
                pass
            self._inline_editor_frame = None

        if self._inline_chip_bar:
            try:
                self._inline_chip_bar.destroy()
            except Exception:
                pass
            self._inline_chip_bar = None

        self._inline_entry = None
        self._inline_clicking_chip = False

    def _apply_obs_type_to_all(self):
        """Sets the selected Source (Obs Type) onto all bookings currently loaded."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            messagebox.showinfo("No Bookings", "No bookings are currently loaded to update.")
            return
        selected_obs = self.obs_type_var.get()
        count = 0
        for b in self.parsed_data["bookings"]:
            b["obs_type"] = selected_obs
            count += 1
        self._refresh_bookings_table()
        self.log(f"Updated Source (Obs Type) to '{selected_obs}' for all {count} entries.")
        messagebox.showinfo("Source Updated", f"Applied '{selected_obs}' to all {count} booking entries.")

    def _show_tree_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree_menu.post(event.x_root, event.y_root)

    def _on_edit_booking_click(self):
        """Opens the direct inline Defect Description editor on the selected row (no popup window)."""
        if not self.parsed_data or not self.parsed_data.get("bookings"):
            messagebox.showinfo("No Bookings", "Please load an Excel file or click 'Add Booking Inline' first.")
            return
        sel = self.tree.selection()
        item_id = sel[0] if sel else "row_0"
        if self.tree.exists(item_id):
            self.tree.selection_set(item_id)
            self.tree.focus(item_id)
            self.tree.see(item_id)
            self.root.update_idletasks()
            self._open_inline_desc_editor(item_id, "#2")

    def _on_add_booking_click(self):
        """Appends a new booking row directly into the table and opens the inline Defect editor (no popup window)."""
        self._close_inline_editor(commit=True)
        if not self.parsed_data:
            loco_no = self.loco_var.get().strip()
            if loco_no == "--":
                loco_no = ""
            self.parsed_data = {
                "file_path": self.current_excel_path or "",
                "loco_number": loco_no,
                "schedule": self.sched_var.get() if self.sched_var.get() != "--" else "UnSch / QAI",
                "date": datetime.now().strftime("%d.%m.%Y"),
                "bookings": [],
                "bookings_count": 0
            }

        bookings = self.parsed_data.setdefault("bookings", [])
        next_sno = len(bookings) + 1
        new_booking = {
            "sno": next_sno,
            "description": "",
            "raw_section": "M1",
            "sections": ["M1"],
            "category": "General",
            "obs_type": self.obs_type_var.get(),
            "kind": self.job_kind_var.get()
        }
        bookings.append(new_booking)
        self._refresh_bookings_table()

        new_idx = len(bookings) - 1
        new_item_id = f"row_{new_idx}"
        if self.tree.exists(new_item_id):
            self.tree.selection_set(new_item_id)
            self.tree.focus(new_item_id)
            self.tree.see(new_item_id)
            self.root.update_idletasks()
            self._open_inline_desc_editor(new_item_id, "#2")

        self.log(f"Added new booking row S.No {next_sno} inline — type defect description directly in the table.")

    def _on_delete_booking_click(self):
        """Deletes the selected booking row(s) directly from the table and re-numbers S.No."""
        self._close_inline_editor(commit=False)
        sel_indices = self._get_selected_row_indices()
        if not sel_indices:
            messagebox.showinfo("Select Booking", "Please select a booking row in the table to delete.")
            return

        if not self.parsed_data or not self.parsed_data.get("bookings"):
            return

        bookings = self.parsed_data["bookings"]
        for idx in sorted(sel_indices, reverse=True):
            if 0 <= idx < len(bookings):
                removed = bookings.pop(idx)
                self.log(f"Deleted booking S.No {removed.get('sno', idx + 1)}: '{removed.get('description', '')[:50]}'")

        for i, b in enumerate(bookings, start=1):
            b["sno"] = i

        self._refresh_bookings_table()

    def _prompt_otp(self) -> str:
        """Prompts the user for OTP via a custom modal dialog in the main GUI thread."""
        result = [None]
        event = threading.Event()

        def _ask():
            dlg = tk.Toplevel(self.root)
            self._active_modal = dlg
            dlg.transient(self.root)
            dlg.title("SLAM Portal Security Verification")
            dlg.geometry("450x240")
            dlg.resizable(False, False)

            # Center on root
            try:
                rx = self.root.winfo_rootx() + (self.root.winfo_width() - 450) // 2
                ry = self.root.winfo_rooty() + (self.root.winfo_height() - 240) // 2
                dlg.geometry(f"+{max(0, rx)}+{max(0, ry)}")
            except Exception:
                pass

            dlg.grab_set()
            notify_user_attention(dlg)

            tk.Label(dlg, text="🔑 SLAM Security Verification", font=("Segoe UI", 12, "bold"), fg="#1a365d").pack(pady=(14, 4))
            tk.Label(
                dlg,
                text="Enter the OTP received on your registered Mobile / WhatsApp:\n(Sent to 8810431331 / SSE)\n\nNote: OTP is valid for the whole day and will be saved.",
                font=("Segoe UI", 9),
                fg="#4a5568",
                justify=tk.CENTER
            ).pack(pady=(0, 8))

            otp_var = tk.StringVar(value=self.daily_otp_var.get())
            entry = ttk.Entry(dlg, textvariable=otp_var, font=("Segoe UI", 16, "bold"), justify="center", width=12)
            entry.pack(pady=4)
            entry.focus_set()

            def _submit(event=None):
                val = otp_var.get().strip()
                result[0] = val
                if val:
                    self.daily_otp_var.set(val)
                    self._update_otp_status()
                self._active_modal = None
                try:
                    dlg.grab_release()
                except Exception:
                    pass
                dlg.destroy()
                event_set_safe()

            def _cancel():
                result[0] = ""
                self._active_modal = None
                try:
                    dlg.grab_release()
                except Exception:
                    pass
                dlg.destroy()
                event_set_safe()

            def event_set_safe():
                if not event.is_set():
                    event.set()

            entry.bind("<Return>", _submit)
            dlg.protocol("WM_DELETE_WINDOW", _cancel)

            btn_frame = ttk.Frame(dlg)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="Submit & Save for Today", command=_submit).pack(side=tk.LEFT, padx=6)
            ttk.Button(btn_frame, text="Cancel", command=_cancel).pack(side=tk.LEFT, padx=6)

        self.root.after(0, _ask)
        event.wait()
        return result[0] or ""

    def _prompt_verify(self, staged_rows, meta) -> bool:
        """
        Opens the dedicated Verification Window displaying all staged observations.
        Blocks until the user clicks either 'PROCEED & SAVE' or 'STOP & DO NOT SAVE'.
        """
        event = threading.Event()
        result = [False]

        def _ask():
            dlg = VerificationDialog(self.root, staged_rows, meta, gui_ref=self)
            self._active_modal = dlg
            notify_user_attention(dlg)
            self.root.wait_window(dlg)
            self._active_modal = None
            result[0] = dlg.result
            event.set()

        self.root.after(0, _ask)
        event.wait()
        return result[0]

    def _start_autofill(self):
        self._close_inline_editor(commit=True)
        if not self.parsed_data or self.parsed_data["bookings_count"] == 0:
            messagebox.showwarning("No Data", "Please select a valid Excel file containing bookings first!")
            return

        loco_no = self.loco_var.get().strip()
        if not loco_no or loco_no == "--":
            messagebox.showwarning("Missing Loco", "Please enter a valid locomotive number!")
            return

        # Ensure credentials are saved
        u = self.username_var.get().strip()
        p = self.password_var.get().strip()
        if not u or not p:
            messagebox.showwarning("Missing Credentials", "Please enter a valid Username and Password!")
            return
        config.save_credentials(u, p)

        confirm = messagebox.askyesno(
            "Confirm SLAM Auto-Fill",
            f"Are you sure you want to start auto-filling {self.parsed_data['bookings_count']} jobcards for Locomotive {loco_no}?\n\n"
            f"• Username: {u}\n"
            f"• Schedule: {self.sched_var.get()}\n"
            f"• Jobcard Type: {self.jobcard_type_var.get()}\n"
            f"• Eqt Required: {self.eqt_req_var.get()}\n"
            f"• Source: {self.obs_type_var.get()}\n"
            f"• Kind: {self.job_kind_var.get()}\n\n"
            f"Note: The bot will pause after staging all observations so you can review them in a full table before saving."
        )
        if not confirm:
            return

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.status_lbl.config(text="Starting automation...")

        # Submit task to persistent worker thread
        self.worker.task_queue.put({
            "action": "PROCESS_EXCEL",
            "excel_path": self.current_excel_path,
            "parsed_data": self.parsed_data,
            "params": {
                "username": u,
                "password": p,
                "loco_number": self.loco_var.get().strip(),
                "jobcard_type": self.jobcard_type_var.get(),
                "eqt_required": self.eqt_req_var.get(),
                "obs_type": self.obs_type_var.get(),
                "job_kind": self.job_kind_var.get(),
                "headless": self.headless_var.get(),
                "keep_browser": self.keep_browser_var.get(),
                "browser_choice": self._get_selected_browser_key(),
                "custom_browser_path": config.get_browser_preference().get("custom_path", "")
            }
        })

    def _stop_autofill(self):
        self.worker.stop_current_job()
        self.status_lbl.config(text="Stopping...")
        self.log("Stopping process as requested by user...")

    def _update_progress(self, current: int, total: int, msg: str):
        def _update():
            try:
                pct = int((current / total) * 100) if total > 0 else 0
                self.progress_bar["value"] = pct
                self.status_lbl.config(text=f"{msg} ({pct}%)")
            except Exception:
                pass
        try:
            self.root.after(0, _update)
        except Exception:
            pass

    def _reset_buttons(self):
        self.start_btn.config(state=tk.NORMAL)
        if hasattr(self, "start_eq_btn"):
            self.start_eq_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2: LOCO EQUIPMENT & SHOP FLOOR
    # ══════════════════════════════════════════════════════════════════════════
    def _build_equipment_tab(self):
        # 1. Equipment Data Source Loader Frame
        eq_file_frame = ttk.LabelFrame(
            self.equipment_tab,
            text=" 📂 1. Load Equipment Data (Formatted Template, Raw Section Workbook, or Google Sheet) ",
            padding=6
        )
        eq_file_frame.pack(fill=tk.X, pady=(0, 5))

        ef_row1 = ttk.Frame(eq_file_frame)
        ef_row1.pack(fill=tk.X, pady=(0, 3))

        ttk.Label(ef_row1, text="Excel File:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.eq_file_var = tk.StringVar(value="")
        eq_entry = ttk.Entry(ef_row1, textvariable=self.eq_file_var, width=44, font=("Segoe UI", 9))
        eq_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        ttk.Button(ef_row1, text="📂 Browse Excel...", command=self._browse_eq_file).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(ef_row1, text="🔄 Reload", command=self._reload_eq_file).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(ef_row1, text="Sheet:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_sheet_var = tk.StringVar(value="")
        self.eq_sheet_cb = ttk.Combobox(ef_row1, textvariable=self.eq_sheet_var, state="readonly", width=22)
        self.eq_sheet_cb.pack(side=tk.LEFT, padx=(0, 8))
        self.eq_sheet_cb.bind("<<ComboboxSelected>>", self._on_eq_sheet_selected)

        export_eq_btn = tk.Button(
            ef_row1,
            text="💾 Export Formatted Excel",
            font=("Segoe UI", 8, "bold"),
            fg="#1a365d",
            bg="#ebf8ff",
            activebackground="#bee3f8",
            relief=tk.GROOVE,
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._export_eq_excel
        )
        export_eq_btn.pack(side=tk.RIGHT)

        ef_row2 = ttk.Frame(eq_file_frame)
        ef_row2.pack(fill=tk.X, pady=(2, 0))

        ttk.Label(ef_row2, text="Google Sheet ID / URL:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        self.eq_gsheet_id_var = tk.StringVar(value=config.DEFAULT_EQ_SHEET_ID)
        ttk.Entry(ef_row2, textvariable=self.eq_gsheet_id_var, width=38, font=("Consolas", 8)).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(ef_row2, text="GID:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_gsheet_gid_var = tk.StringVar(value=config.DEFAULT_EQ_SHEET_GID)
        ttk.Entry(ef_row2, textvariable=self.eq_gsheet_gid_var, width=10, font=("Consolas", 8)).pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(ef_row2, text="☁️ Fetch Google Sheet", command=self._load_eq_gsheet).pack(side=tk.LEFT, padx=(0, 10))

        self.eq_summary_var = tk.StringVar(value="Rows: 0 | Locos: --")
        tk.Label(
            ef_row2,
            textvariable=self.eq_summary_var,
            font=("Segoe UI", 9, "bold"),
            fg="#2f855a",
            bg="#f0fff4",
            padx=8,
            pady=1
        ).pack(side=tk.RIGHT)

        # 2. Default Fill-In Controls + Shop Floor Quick-Launch Bar
        ctrl_frame = ttk.LabelFrame(
            self.equipment_tab,
            text=" ⚡ 2. Default Fill-In Controls (For Raw Sheets) & Shop Floor Navigation ",
            padding=6
        )
        ctrl_frame.pack(fill=tk.X, pady=(0, 5))

        cf_row1 = ttk.Frame(ctrl_frame)
        cf_row1.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(cf_row1, text="Default Schedule:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_def_sched_var = tk.StringVar(value="IA")
        ttk.Combobox(
            cf_row1,
            textvariable=self.eq_def_sched_var,
            values=["IA", "IB", "IC", "TOH", "IOH", "POH", "UNSCH", "COMM"],
            width=8
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(cf_row1, text="Default Prov. Date:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_def_date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ttk.Entry(cf_row1, textvariable=self.eq_def_date_var, width=12, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(cf_row1, text="Default Make:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_def_make_var = tk.StringVar(value="")
        ttk.Entry(cf_row1, textvariable=self.eq_def_make_var, width=12, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(0, 10))

        self.eq_auto_num_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cf_row1,
            text="Auto-Generate Blank Eq.Number ({Loco}-POS{Pos})",
            variable=self.eq_auto_num_var
        ).pack(side=tk.LEFT, padx=(0, 10))

        apply_def_btn = tk.Button(
            cf_row1,
            text="⚡ Apply Defaults to Blank Cells",
            font=("Segoe UI", 8, "bold"),
            fg="white",
            bg="#2b6cb0",
            activebackground="#2c5282",
            relief=tk.FLAT,
            padx=8,
            pady=2,
            cursor="hand2",
            command=self._apply_eq_defaults_to_blanks
        )
        apply_def_btn.pack(side=tk.LEFT, padx=4)

        cf_row2 = ttk.Frame(ctrl_frame)
        cf_row2.pack(fill=tk.X, pady=(2, 0))

        ttk.Label(cf_row2, text="Target Loco Filter:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_loco_filter_var = tk.StringVar(value="ALL LOCOS")
        self.eq_loco_filter_cb = ttk.Combobox(
            cf_row2,
            textvariable=self.eq_loco_filter_var,
            values=["ALL LOCOS"],
            width=14
        )
        self.eq_loco_filter_cb.pack(side=tk.LEFT, padx=(0, 10))
        self.eq_loco_filter_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_eq_table())

        ttk.Label(cf_row2, text="Select For:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.eq_select_for_var = tk.StringVar(value="Home")
        ttk.Combobox(
            cf_row2,
            textvariable=self.eq_select_for_var,
            values=["Home", "Foreign"],
            state="readonly",
            width=8
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.eq_auto_save_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            cf_row2,
            text="Auto-Click Save in Popup (Uncheck for Dry-Run Staging Only)",
            variable=self.eq_auto_save_var
        ).pack(side=tk.LEFT, padx=(0, 14))

        # Shop Floor Quick-Launch Buttons
        sf_frame = ttk.Frame(cf_row2)
        sf_frame.pack(side=tk.RIGHT)
        ttk.Label(sf_frame, text="🏭 Shop Floor:", font=("Segoe UI", 8, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.shopfloor_mod_var = tk.StringVar(value="Add Loco Eq./Eq. Fitment")
        sf_cb = ttk.Combobox(
            sf_frame,
            textvariable=self.shopfloor_mod_var,
            values=list(config.SHOPFLOOR_MODULES.keys()),
            state="readonly",
            width=22
        )
        sf_cb.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(sf_frame, text="🚀 Open in SLAM", command=self._open_shopfloor_click).pack(side=tk.LEFT)

        # 3. Equipment Fitment Table (15 Columns, Inline Cell Editor)
        eq_table_frame = ttk.LabelFrame(
            self.equipment_tab,
            text=" 🔧 3. Loco Equipment Fitment Rows (Single-Click Any Cell to Edit Directly) ",
            padding=6
        )
        eq_table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 4))

        cols = (
            "sno", "loco", "equipment", "position", "make", "number", "loc_number",
            "prov_date", "schedule", "model", "doc", "dom", "dor", "warranty", "remarks"
        )
        self.eq_tree = ttk.Treeview(eq_table_frame, columns=cols, show="headings", height=6, selectmode="browse")

        headings = {
            "sno": ("#", 34, tk.CENTER),
            "loco": ("Loco", 58, tk.CENTER),
            "equipment": ("Equipment Name (SLAM)", 155, tk.W),
            "position": ("Pos", 42, tk.CENTER),
            "make": ("Eq.Make", 95, tk.W),
            "number": ("Eq.Number", 105, tk.W),
            "loc_number": ("Local Eq.No", 95, tk.W),
            "prov_date": ("Prov.Date", 82, tk.CENTER),
            "schedule": ("Sched", 55, tk.CENTER),
            "model": ("Model", 70, tk.W),
            "doc": ("DoC", 75, tk.CENTER),
            "dom": ("DoM", 75, tk.CENTER),
            "dor": ("DoR", 75, tk.CENTER),
            "warranty": ("Warranty", 70, tk.CENTER),
            "remarks": ("Remarks", 130, tk.W),
        }
        for col_id, (title, width, anchor) in headings.items():
            self.eq_tree.heading(col_id, text=title)
            self.eq_tree.column(col_id, width=width, anchor=anchor, stretch=(col_id in ("equipment", "remarks")))

        eq_yscroll = ttk.Scrollbar(eq_table_frame, orient=tk.VERTICAL, command=self.eq_tree.yview)
        eq_xscroll = ttk.Scrollbar(eq_table_frame, orient=tk.HORIZONTAL, command=self.eq_tree.xview)
        self.eq_tree.configure(yscrollcommand=eq_yscroll.set, xscrollcommand=eq_xscroll.set)

        self.eq_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        eq_xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        eq_yscroll.place(relx=1.0, rely=0.0, relheight=0.88, anchor=tk.NE)

        self.eq_tree.bind("<<TreeviewSelect>>", self._on_eq_tree_select)
        self.eq_tree.bind("<ButtonRelease-1>", self._on_eq_tree_single_click)

        # Direct Inline Editor Bar below Equipment Tree
        eq_edit_bar = tk.Frame(eq_table_frame, bg="#edf2f7", bd=1, relief=tk.SOLID, padx=6, pady=4)
        eq_edit_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))

        eb_r1 = tk.Frame(eq_edit_bar, bg="#edf2f7")
        eb_r1.pack(fill=tk.X, pady=(0, 2))

        tk.Label(eb_r1, text="Loco:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_loco_var = tk.StringVar()
        self.eq_inline_loco_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_loco_var, width=8)
        self.eq_inline_loco_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Equipment:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_name_var = tk.StringVar()
        self.eq_inline_name_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_name_var, width=20)
        self.eq_inline_name_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Pos:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_pos_var = tk.StringVar()
        self.eq_inline_pos_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_pos_var, width=5)
        self.eq_inline_pos_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Make:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_make_var = tk.StringVar()
        self.eq_inline_make_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_make_var, width=12)
        self.eq_inline_make_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Eq.No:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_num_var = tk.StringVar()
        self.eq_inline_num_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_num_var, width=12)
        self.eq_inline_num_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Local No:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_locnum_var = tk.StringVar()
        self.eq_inline_locnum_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_locnum_var, width=10)
        self.eq_inline_locnum_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Prov.Date:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_pdate_var = tk.StringVar()
        self.eq_inline_pdate_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_pdate_var, width=10)
        self.eq_inline_pdate_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Sched:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_sched_var = tk.StringVar()
        self.eq_inline_sched_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_sched_var, width=6)
        self.eq_inline_sched_ent.pack(side=tk.LEFT, padx=(2, 6))

        tk.Label(eb_r1, text="Remarks:", font=("Segoe UI", 8, "bold"), bg="#edf2f7").pack(side=tk.LEFT)
        self.eq_inline_rem_var = tk.StringVar()
        self.eq_inline_rem_ent = ttk.Entry(eb_r1, textvariable=self.eq_inline_rem_var, width=14)
        self.eq_inline_rem_ent.pack(side=tk.LEFT, padx=(2, 6))

        for ent in (
            self.eq_inline_loco_ent, self.eq_inline_name_ent, self.eq_inline_pos_ent,
            self.eq_inline_make_ent, self.eq_inline_num_ent, self.eq_inline_locnum_ent,
            self.eq_inline_pdate_ent, self.eq_inline_sched_ent, self.eq_inline_rem_ent
        ):
            ent.bind("<KeyRelease>", self._on_eq_direct_bar_changed)

        ttk.Button(eb_r1, text="➕ Add Row", command=self._add_eq_row_inline).pack(side=tk.LEFT, padx=2)
        ttk.Button(eb_r1, text="🗑️ Delete", command=self._delete_eq_row_click).pack(side=tk.LEFT, padx=2)

    # ── Equipment Helper Methods ──────────────────────────────────────────────
    def _auto_detect_equipment_file(self):
        base_dir = get_base_dir()
        candidates = [
            os.path.join(base_dir, "TM_Temp_Sensor_Sample.xlsx"),
            os.path.join(base_dir, "TM_Speed_Sensor_Sample.xlsx"),
            os.path.join(base_dir, "TM in & out details.xlsx"),
            os.path.join(base_dir, "equipment fill templete.xlsx"),
        ]
        for c in candidates:
            if os.path.exists(c):
                self._load_eq_excel(c)
                return

    def _browse_eq_file(self):
        path = filedialog.askopenfilename(
            title="Select Equipment Excel File (Template or Raw Section File)",
            initialdir=get_base_dir(),
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if path:
            self._load_eq_excel(path)

    def _reload_eq_file(self):
        path = self.eq_file_var.get().strip()
        if path and os.path.exists(path):
            sheet = self.eq_sheet_var.get().strip() or None
            self._load_eq_excel(path, sheet_name=sheet)
        else:
            messagebox.showwarning("No File", "Please select a valid Excel file first.")

    def _load_eq_excel(self, path: str, sheet_name: str = None):
        try:
            self._close_eq_inline_editor(commit=False)
            self.current_eq_excel_path = path
            self.eq_file_var.set(path)
            data = excel_parser.parse_equipment_excel(path, sheet_name=sheet_name)
            self.eq_parsed_data = data

            sheets = data.get("sheet_names", [])
            self.eq_sheet_cb["values"] = sheets
            active_sheet = data.get("sheet_name") or (sheets[0] if sheets else "")
            self.eq_sheet_var.set(active_sheet)

            locos = data.get("locos", [])
            self.eq_loco_filter_cb["values"] = ["ALL LOCOS"] + locos
            self.eq_loco_filter_var.set("ALL LOCOS")

            self._refresh_eq_table()
            self.log(
                f"Loaded Equipment File '{os.path.basename(path)}' [Sheet: '{active_sheet.strip()}']: "
                f"{data.get('total_rows', 0)} fitment rows across {len(locos)} locos ({', '.join(locos[:8])})."
            )
        except Exception as e:
            self.log(f"Error loading equipment Excel '{path}': {e}")
            messagebox.showerror("Equipment Parse Error", f"Failed to parse equipment Excel:\n\n{e}")

    def _on_eq_sheet_selected(self, event=None):
        path = self.eq_file_var.get().strip()
        if not path or not os.path.exists(path):
            return
        sel = self.eq_sheet_var.get()
        if sel:
            self._load_eq_excel(path, sheet_name=sel)

    def _load_eq_gsheet(self):
        sheet_id = self.eq_gsheet_id_var.get().strip()
        gid = self.eq_gsheet_gid_var.get().strip() or "0"
        if not sheet_id:
            messagebox.showwarning("Missing Sheet ID", "Please enter a Google Sheet ID or URL.")
            return
        try:
            self._close_eq_inline_editor(commit=False)
            self.log(f"Fetching Google Sheet '{sheet_id}' (gid={gid})...")
            data = excel_parser.parse_equipment_google_sheet(sheet_id, gid)
            self.eq_parsed_data = data
            locos = data.get("locos", [])
            self.eq_loco_filter_cb["values"] = ["ALL LOCOS"] + locos
            self.eq_loco_filter_var.set("ALL LOCOS")
            self._refresh_eq_table()
            self.log(f"Loaded {data.get('rows_count', 0)} equipment rows from Google Sheet across {len(locos)} locos.")
        except Exception as e:
            self.log(f"Google Sheet fetch error: {e}")
            messagebox.showerror("Google Sheet Error", f"Could not load Google Sheet:\n\n{e}")

    def _export_eq_excel(self):
        self._close_eq_inline_editor(commit=True)
        if not self.eq_parsed_data or not self.eq_parsed_data.get("rows"):
            messagebox.showwarning("No Data", "No equipment rows available to export.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Export Formatted SLAM Equipment Excel",
            initialdir=get_base_dir(),
            initialfile="formatted_equipment_fitment.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out_path:
            return
        try:
            saved = excel_parser.export_equipment_to_excel(self.eq_parsed_data["rows"], out_path)
            self.log(f"Exported {len(self.eq_parsed_data['rows'])} formatted equipment rows to '{saved}'.")
            messagebox.showinfo("Export Complete", f"Formatted equipment template exported to:\n{saved}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export Excel:\n\n{e}")

    def _apply_eq_defaults_to_blanks(self):
        self._close_eq_inline_editor(commit=True)
        if not self.eq_parsed_data or not self.eq_parsed_data.get("rows"):
            messagebox.showwarning("No Rows", "Please load an equipment file first.")
            return
        def_sched = self.eq_def_sched_var.get().strip()
        def_date = self.eq_def_date_var.get().strip()
        def_make = excel_parser.normalize_make_alias(self.eq_def_make_var.get().strip())
        auto_num = self.eq_auto_num_var.get()

        updated_count = 0
        for r in self.eq_parsed_data["rows"]:
            changed = False
            if def_sched and not str(r.get("schedule", "")).strip():
                r["schedule"] = def_sched
                changed = True
            if def_date and not str(r.get("provision_date", "")).strip():
                r["provision_date"] = def_date
                changed = True
            if def_make and not str(r.get("make", "")).strip():
                r["make"] = def_make
                changed = True
            if auto_num:
                loco = str(r.get("loco_no", "")).strip()
                pos = str(r.get("position", "")).strip() or "1"
                if not str(r.get("eq_number", "")).strip() and loco:
                    r["eq_number"] = f"{loco}-POS{pos}"
                    changed = True
                if not str(r.get("loc_eq_number", "")).strip() and r.get("eq_number"):
                    r["loc_eq_number"] = str(r["eq_number"])
                    changed = True
            if changed:
                updated_count += 1

        self._refresh_eq_table()
        self.log(f"Applied default fill-in values to {updated_count} equipment rows.")

    def _get_filtered_eq_rows_with_indices(self):
        if not self.eq_parsed_data or not self.eq_parsed_data.get("rows"):
            return []
        loco_filter = self.eq_loco_filter_var.get().strip()
        res = []
        for idx, r in enumerate(self.eq_parsed_data["rows"]):
            if loco_filter and loco_filter != "ALL LOCOS":
                if str(r.get("loco_no", "")).strip() != loco_filter:
                    continue
            res.append((idx, r))
        return res

    def _refresh_eq_table(self):
        self._close_eq_inline_editor(commit=False)
        for item in self.eq_tree.get_children():
            self.eq_tree.delete(item)

        filtered = self._get_filtered_eq_rows_with_indices()
        for disp_idx, (real_idx, r) in enumerate(filtered, start=1):
            self.eq_tree.insert(
                "",
                tk.END,
                iid=str(real_idx),
                values=(
                    disp_idx,
                    r.get("loco_no", ""),
                    r.get("eq_name", ""),
                    r.get("position", ""),
                    r.get("make", ""),
                    r.get("eq_number", ""),
                    r.get("loc_eq_number", ""),
                    r.get("provision_date", ""),
                    r.get("schedule", ""),
                    r.get("model", ""),
                    r.get("doc", ""),
                    r.get("dom", ""),
                    r.get("receive_date", ""),
                    r.get("warranty", ""),
                    r.get("remarks", ""),
                )
            )

        all_rows = self.eq_parsed_data.get("rows", []) if self.eq_parsed_data else []
        locos = sorted({str(r.get("loco_no", "")).strip() for r in all_rows if str(r.get("loco_no", "")).strip()})
        if self.eq_parsed_data is not None:
            self.eq_parsed_data["rows_count"] = len(all_rows)
            self.eq_parsed_data["total_rows"] = len(all_rows)
            self.eq_parsed_data["locos"] = locos
        self.eq_summary_var.set(f"Showing {len(filtered)}/{len(all_rows)} Rows | Locos: {len(locos)}")

    def _on_eq_tree_select(self, event=None):
        sel = self.eq_tree.selection()
        if not sel or not self.eq_parsed_data:
            return
        idx = int(sel[0])
        rows = self.eq_parsed_data.get("rows", [])
        if idx < 0 or idx >= len(rows):
            return
        r = rows[idx]
        self._updating_eq_inline_bar = True
        try:
            self.eq_inline_loco_var.set(r.get("loco_no", ""))
            self.eq_inline_name_var.set(r.get("eq_name", ""))
            self.eq_inline_pos_var.set(r.get("position", ""))
            self.eq_inline_make_var.set(r.get("make", ""))
            self.eq_inline_num_var.set(r.get("eq_number", ""))
            self.eq_inline_locnum_var.set(r.get("loc_eq_number", ""))
            self.eq_inline_pdate_var.set(r.get("provision_date", ""))
            self.eq_inline_sched_var.set(r.get("schedule", ""))
            self.eq_inline_rem_var.set(r.get("remarks", ""))
        finally:
            self._updating_eq_inline_bar = False

    def _on_eq_direct_bar_changed(self, event=None):
        if getattr(self, "_updating_eq_inline_bar", False):
            return
        sel = self.eq_tree.selection()
        if not sel or not self.eq_parsed_data:
            return
        idx = int(sel[0])
        rows = self.eq_parsed_data.get("rows", [])
        if idx < 0 or idx >= len(rows):
            return
        r = rows[idx]
        r["loco_no"] = self.eq_inline_loco_var.get().strip()
        r["eq_name"] = self.eq_inline_name_var.get().strip()
        r["eq_raw"] = r["eq_name"]
        r["position"] = self.eq_inline_pos_var.get().strip()
        r["make"] = self.eq_inline_make_var.get().strip()
        r["eq_number"] = self.eq_inline_num_var.get().strip()
        r["loc_eq_number"] = self.eq_inline_locnum_var.get().strip()
        r["provision_date"] = self.eq_inline_pdate_var.get().strip()
        r["schedule"] = self.eq_inline_sched_var.get().strip()
        r["remarks"] = self.eq_inline_rem_var.get().strip()

        cur_vals = list(self.eq_tree.item(sel[0], "values"))
        disp_sno = cur_vals[0] if cur_vals else idx + 1
        self.eq_tree.item(
            sel[0],
            values=(
                disp_sno,
                r.get("loco_no", ""),
                r.get("eq_name", ""),
                r.get("position", ""),
                r.get("make", ""),
                r.get("eq_number", ""),
                r.get("loc_eq_number", ""),
                r.get("provision_date", ""),
                r.get("schedule", ""),
                r.get("model", ""),
                r.get("doc", ""),
                r.get("dom", ""),
                r.get("receive_date", ""),
                r.get("warranty", ""),
                r.get("remarks", ""),
            )
        )

    def _on_eq_tree_single_click(self, event):
        region = self.eq_tree.identify("region", event.x, event.y)
        if region != "cell":
            self._close_eq_inline_editor(commit=True)
            return
        row_id = self.eq_tree.identify_row(event.y)
        col_id = self.eq_tree.identify_column(event.x)
        if not row_id or col_id == "#1":
            self._close_eq_inline_editor(commit=True)
            return
        self._open_eq_inline_editor(row_id, col_id)

    def _open_eq_inline_editor(self, row_id: str, col_id: str):
        self._close_eq_inline_editor(commit=True)
        if not self.eq_parsed_data:
            return
        bbox = self.eq_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox
        col_map = {
            "#2": "loco_no",
            "#3": "eq_name",
            "#4": "position",
            "#5": "make",
            "#6": "eq_number",
            "#7": "loc_eq_number",
            "#8": "provision_date",
            "#9": "schedule",
            "#10": "model",
            "#11": "doc",
            "#12": "dom",
            "#13": "receive_date",
            "#14": "warranty",
            "#15": "remarks",
        }
        field_key = col_map.get(col_id)
        if not field_key:
            return
        idx = int(row_id)
        rows = self.eq_parsed_data.get("rows", [])
        if idx < 0 or idx >= len(rows):
            return
        cur_val = str(rows[idx].get(field_key, ""))

        var = tk.StringVar(value=cur_val)
        editor = tk.Entry(
            self.eq_tree,
            textvariable=var,
            font=("Segoe UI", 9),
            bg="#ffffcc",
            fg="#1a202c",
            relief=tk.SOLID,
            bd=1
        )
        editor.place(x=x, y=y, width=max(width, 65), height=height)
        editor.focus_set()
        editor.select_range(0, tk.END)

        self._eq_inline_editor = editor
        self._eq_inline_editor_info = (row_id, field_key, var)

        editor.bind("<Return>", lambda e: self._close_eq_inline_editor(commit=True))
        editor.bind("<Escape>", lambda e: self._close_eq_inline_editor(commit=False))
        editor.bind("<FocusOut>", lambda e: self._close_eq_inline_editor(commit=True))

    def _close_eq_inline_editor(self, commit: bool = True):
        editor = getattr(self, "_eq_inline_editor", None)
        info = getattr(self, "_eq_inline_editor_info", None)
        self._eq_inline_editor = None
        self._eq_inline_editor_info = None
        if not editor:
            return
        if commit and info and self.eq_parsed_data:
            row_id, field_key, var = info
            idx = int(row_id)
            rows = self.eq_parsed_data.get("rows", [])
            if 0 <= idx < len(rows):
                rows[idx][field_key] = var.get().strip()
                if field_key == "eq_name":
                    rows[idx]["eq_raw"] = rows[idx]["eq_name"]
                if self.eq_tree.exists(row_id):
                    cur_vals = list(self.eq_tree.item(row_id, "values"))
                    disp_sno = cur_vals[0] if cur_vals else idx + 1
                    r = rows[idx]
                    self.eq_tree.item(
                        row_id,
                        values=(
                            disp_sno,
                            r.get("loco_no", ""),
                            r.get("eq_name", ""),
                            r.get("position", ""),
                            r.get("make", ""),
                            r.get("eq_number", ""),
                            r.get("loc_eq_number", ""),
                            r.get("provision_date", ""),
                            r.get("schedule", ""),
                            r.get("model", ""),
                            r.get("doc", ""),
                            r.get("dom", ""),
                            r.get("receive_date", ""),
                            r.get("warranty", ""),
                            r.get("remarks", ""),
                        )
                    )
                    self._on_eq_tree_select()
        try:
            editor.destroy()
        except Exception:
            pass

    def _add_eq_row_inline(self):
        self._close_eq_inline_editor(commit=True)
        if not self.eq_parsed_data:
            self.eq_parsed_data = {"rows": [], "locos": [], "rows_count": 0, "sheet_names": []}
        eq_nm = self.eq_inline_name_var.get().strip() or "TM TEMP SENSOR"
        new_row = {
            "sno": len(self.eq_parsed_data.get("rows", [])) + 1,
            "loco_no": self.eq_inline_loco_var.get().strip() or "30045",
            "eq_raw": eq_nm,
            "eq_name": eq_nm,
            "position": self.eq_inline_pos_var.get().strip() or "1",
            "make": self.eq_inline_make_var.get().strip() or "FLORICON",
            "eq_number": self.eq_inline_num_var.get().strip() or "",
            "loc_eq_number": self.eq_inline_locnum_var.get().strip() or "",
            "provision_date": self.eq_inline_pdate_var.get().strip() or datetime.now().strftime("%d/%m/%Y"),
            "schedule": self.eq_inline_sched_var.get().strip() or "IA",
            "model": "",
            "doc": "",
            "dom": "",
            "receive_date": "",
            "warranty": "NO",
            "warranty_upto": "",
            "amc": "NO",
            "amc_upto": "",
            "remarks": self.eq_inline_rem_var.get().strip(),
        }
        self.eq_parsed_data["rows"].append(new_row)
        new_idx = len(self.eq_parsed_data["rows"]) - 1
        self._refresh_eq_table()
        if self.eq_tree.exists(str(new_idx)):
            self.eq_tree.selection_set(str(new_idx))
            self.eq_tree.see(str(new_idx))

    def _delete_eq_row_click(self):
        self._close_eq_inline_editor(commit=False)
        sel = self.eq_tree.selection()
        if not sel or not self.eq_parsed_data:
            return
        idx = int(sel[0])
        rows = self.eq_parsed_data.get("rows", [])
        if 0 <= idx < len(rows):
            rows.pop(idx)
            self._refresh_eq_table()

    def _open_shopfloor_click(self):
        u = self.username_var.get().strip()
        p = self.password_var.get().strip()
        if not u or not p:
            messagebox.showwarning("Missing Credentials", "Please enter SLAM Username and Password first.")
            return
        config.save_credentials(u, p)
        mod_name = self.shopfloor_mod_var.get().strip()
        self.status_lbl.config(text=f"Opening {mod_name}...")
        self.worker.task_queue.put({
            "action": "OPEN_SHOPFLOOR",
            "module_name": mod_name,
            "params": {
                "username": u,
                "password": p,
                "headless": self.headless_var.get(),
                "browser_choice": self._get_selected_browser_key(),
                "custom_browser_path": config.get_browser_preference().get("custom_path", ""),
            }
        })

    def _start_equipment_autofill(self):
        self._close_eq_inline_editor(commit=True)
        filtered = [r for _, r in self._get_filtered_eq_rows_with_indices()]
        if not filtered:
            messagebox.showwarning("No Equipment Rows", "Please load or add equipment rows before starting!")
            return

        u = self.username_var.get().strip()
        p = self.password_var.get().strip()
        if not u or not p:
            messagebox.showwarning("Missing Credentials", "Please enter a valid Username and Password!")
            return
        config.save_credentials(u, p)

        locos = sorted({str(r.get("loco_no", "")).strip() for r in filtered if str(r.get("loco_no", "")).strip()})
        auto_save = self.eq_auto_save_var.get()
        mode_str = "LIVE SAVE (Clicks Save in Popup)" if auto_save else "DRY-RUN STAGING ONLY (Fills Popup Without Saving)"

        confirm = messagebox.askyesno(
            "Confirm Loco Equipment Auto-Fill",
            f"Ready to process {len(filtered)} equipment fitment row(s) across {len(locos)} locomotive(s):\n"
            f"• Locos: {', '.join(locos[:10])}\n"
            f"• Select For: {self.eq_select_for_var.get()}\n"
            f"• Mode: {mode_str}\n\n"
            f"Proceed?"
        )
        if not confirm:
            return

        self.start_btn.config(state=tk.DISABLED)
        self.start_eq_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.status_lbl.config(text="Starting Equipment Auto-Fill...")

        self.worker.task_queue.put({
            "action": "PROCESS_EQUIPMENT",
            "rows": filtered,
            "params": {
                "username": u,
                "password": p,
                "select_for": self.eq_select_for_var.get(),
                "auto_save": auto_save,
                "close_popup_on_dry_run": True,
                "headless": self.headless_var.get(),
                "keep_browser": self.keep_browser_var.get(),
                "browser_choice": self._get_selected_browser_key(),
                "custom_browser_path": config.get_browser_preference().get("custom_path", ""),
            }
        })

    # ── Update Checking & Notification Methods ────────────────────────────────
    def _check_update_startup(self):
        """Silently checks for updates in the background on startup."""
        threading.Thread(target=self._run_update_check_silent, daemon=True).start()

    def _run_update_check_silent(self):
        try:
            res = updater.check_for_updates(config.APP_VERSION)
            if res.get("update_available"):
                self.update_info = res
                self.root.after(0, lambda: self._show_update_banner(res))
        except Exception:
            pass

    def _show_update_banner(self, info: dict):
        latest = info.get("latest_version", "--")
        self.update_banner_lbl.config(
            text=f"🚀 A new version of SLAM Auto-Filler is available: v{latest} (Current: v{config.APP_VERSION})"
        )
        if hasattr(self, "main_container"):
            self.update_banner.pack(fill=tk.X, side=tk.TOP, before=self.main_container)

    def _on_manual_check_update(self):
        self.update_btn.config(text="⏳ Checking...", state=tk.DISABLED)
        def _check():
            res = updater.check_for_updates(config.APP_VERSION)
            def _done():
                self.update_btn.config(text="🔄 Check Updates", state=tk.NORMAL)
                if res.get("update_available"):
                    self.update_info = res
                    self._show_update_banner(res)
                    self._open_update_dialog()
                elif res.get("success"):
                    messagebox.showinfo(
                        "Up to Date",
                        f"You are running the latest version of SLAM Auto-Filler (v{config.APP_VERSION}).\n\nNo newer version found on GitHub."
                    )
                else:
                    messagebox.showwarning(
                        "Update Check",
                        f"Could not check for updates:\n\n{res.get('error', 'Unknown network error')}"
                    )
            self.root.after(0, _done)
        threading.Thread(target=_check, daemon=True).start()

    def _open_release_notes(self):
        if self.update_info and self.update_info.get("release_url"):
            webbrowser.open(self.update_info["release_url"])

    def _open_update_dialog(self):
        if self.update_info:
            UpdateDialog(self.root, self.update_info, gui_ref=self)

def main():
    root = tk.Tk()
    app = SLAMAutoFillerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
