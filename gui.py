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
import threading
import queue
import collections
from datetime import datetime
from typing import Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser

import config
import updater
from excel_parser import parse_excel_bookings
from slam_bot import SLAMBot

class UpdateDialog(tk.Toplevel):
    """
    Dedicated dialog for reviewing release details, changelog,
    and downloading/installing the latest version.
    """
    def __init__(self, parent, update_info: dict):
        super().__init__(parent)
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
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.exists(downloads_dir):
            downloads_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))

        dest_file = os.path.join(downloads_dir, f"SLAM_Auto_Filler_v{latest_v}.exe")
        self._downloaded_file = dest_file

        self.action_btn.config(state=tk.DISABLED, text="Downloading...")
        self.progress_bar["value"] = 0
        self.status_detail_lbl.config(text="Starting download from GitHub...")

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
                    dest_file,
                    progress_callback=_prog,
                    cancel_check=lambda: self._cancel_download
                )

                if success:
                    def _on_success():
                        self.progress_bar["value"] = 100
                        self.status_detail_lbl.config(
                            text=f"✅ Download complete!\nSaved to: {dest_file}"
                        )
                        self.action_btn.config(
                            state=tk.NORMAL,
                            text="🚀 Launch New Version",
                            command=self._launch_downloaded
                        )
                        self.close_btn.config(text="Close")
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
                updater.launch_updated_executable(self._downloaded_file)
            except Exception as e:
                messagebox.showerror("Launch Error", f"Could not launch new version:\n{e}")

    def _on_close(self):
        self._cancel_download = True
        self.destroy()

class VerificationDialog(tk.Toplevel):
    """
    Dedicated Verification Modal that pops up once all observations are staged in SLAM.
    Shows the staged table, section breakdown, and requires explicit user confirmation
    to either 'Save & Submit to Section' or 'Stop / Keep Staged'.
    """
    def __init__(self, parent, staged_rows, meta):
        super().__init__(parent)
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
        self.destroy()

    def _on_stop(self):
        self.result = False
        self.destroy()

class EditBookingDialog(tk.Toplevel):
    """
    Modal dialog allowing the user to view and edit a booking's defect message and section.
    """
    def __init__(self, parent, booking_data: dict, is_new: bool = False):
        super().__init__(parent)
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
        self.protocol("WM_DELETE_WINDOW", self.destroy)

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
                padx=6,
                pady=2,
                cursor="hand2",
                command=lambda sec=s: self._toggle_section_chip(sec)
            )
            btn.pack(side=tk.LEFT, padx=3, pady=2)

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
            command=self.destroy
        )
        cancel_btn.pack(side=tk.LEFT)

        self.bind("<Control-Return>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self.destroy())

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
                    if self.bot is None:
                        self.bot = SLAMBot(
                            username=params["username"],
                            password=params["password"],
                            jobcard_type=params["jobcard_type"],
                            eqt_required=params["eqt_required"],
                            obs_type=params["obs_type"],
                            job_kind=params["job_kind"],
                            headless=params["headless"],
                            log_callback=self.gui.log,
                            otp_callback=self.gui._prompt_otp,
                            progress_callback=self.gui._update_progress,
                            verify_before_save=True,
                            verify_callback=self.gui._prompt_verify
                        )
                    else:
                        self.bot.username = params["username"]
                        self.bot.password = params["password"]
                        self.bot.jobcard_type = params["jobcard_type"]
                        self.bot.eqt_required = params["eqt_required"]
                        self.bot.obs_type = params["obs_type"]
                        self.bot.job_kind = params["job_kind"]
                        self.bot.headless = params["headless"]
                        self.bot.log_callback = self.gui.log
                        self.bot.otp_callback = self.gui._prompt_otp
                        self.bot.progress_callback = self.gui._update_progress
                        self.bot.verify_before_save = True
                        self.bot.verify_callback = self.gui._prompt_verify
                        self.bot._stop_requested = False

                        if self.bot.browser and self.bot.browser.is_connected() and self.bot.page and not self.bot.page.is_closed():
                            self.gui.log("Reusing existing open browser session...")
                        else:
                            self.gui.log("Opening new browser window for auto-fill...")

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
                            messagebox.showinfo(
                                "SLAM Jobcards Created Successfully",
                                f"🎉 Success!\n\nAll {r['entered']}/{r['total']} jobcards for Locomotive {r['loco_number']} have been created, saved, and dispatched in the SLAM portal."
                            )
                        self.gui.root.after(0, _done_success)
                    elif res.get("status") == "verification_deferred":
                        def _done_deferred(r=res):
                            self.gui.status_lbl.config(text="Save deferred by user.")
                            messagebox.showinfo(
                                "Save Deferred",
                                f"Observations for Locomotive {r.get('loco_number')} remain staged in the SLAM portal.\n\nNo save was committed."
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
                        messagebox.showerror("Execution Error", f"An error occurred during auto-fill:\n\n{err}")
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

class SLAMAutoFillerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"SLAM Jobcard Auto-Filler v{config.APP_VERSION} - Indian Railways (ELS/RPME)")
        self.root.geometry("1040x840")
        self.root.minsize(940, 680)

        # Apply ttk theme
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.current_excel_path = ""
        self.parsed_data = None
        self.update_info = None

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

        self._build_ui()
        self._update_otp_status()
        self._scan_desktop_files()
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

        # Check for updates in background
        self.root.after(1500, self._check_update_startup)

    def _build_ui(self):
        # ── Header Banner ───────────────────────────────────────────────────────
        header_frame = tk.Frame(self.root, bg="#1a365d", height=70)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        title_lbl = tk.Label(
            header_frame,
            text="🚆 Indian Railways - SLAM Jobcard Auto-Filler",
            font=("Segoe UI", 16, "bold"),
            fg="white",
            bg="#1a365d"
        )
        title_lbl.pack(side=tk.LEFT, padx=16, pady=10)

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
            text="⚡ Preset: Schedule Attention | Without Eq | Testing Remarks | Minor",
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

        # 1. Credentials & Daily OTP Frame
        cred_frame = ttk.LabelFrame(main_container, text=" 🔐 Portal Credentials & Daily OTP ", padding=8)
        cred_frame.pack(fill=tk.X, pady=(0, 8))

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

        # 2. File Selection Frame
        file_frame = ttk.LabelFrame(main_container, text=" 📁 Select Bookings Excel File ", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

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
        info_frame.pack(fill=tk.X, pady=(8, 0))

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
        settings_frame = ttk.LabelFrame(main_container, text=" ⚙️ Jobcard Creation Settings (Pre-Configured) ", padding=8)
        settings_frame.pack(fill=tk.X, pady=(0, 8))

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

        # 4. Bookings Preview Table
        table_frame = ttk.LabelFrame(main_container, text=" 📋 Bookings to be Staged into SLAM ", padding=6)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        tbl_container = ttk.Frame(table_frame)
        tbl_container.pack(fill=tk.BOTH, expand=True)

        cols = ("sno", "desc", "section", "target_sections", "obs_type")
        self.tree = ttk.Treeview(tbl_container, columns=cols, show="headings", height=7)
        self.tree.heading("sno", text="S.No", anchor=tk.CENTER)
        self.tree.heading("desc", text="Booking / Defect Description", anchor=tk.W)
        self.tree.heading("section", text="Excel Section", anchor=tk.CENTER)
        self.tree.heading("target_sections", text="Mapped SLAM Section(s)", anchor=tk.CENTER)
        self.tree.heading("obs_type", text="Source (Obs Type)", anchor=tk.CENTER)

        self.tree.column("sno", width=45, stretch=False, anchor=tk.CENTER)
        self.tree.column("desc", width=480, stretch=True)
        self.tree.column("section", width=95, stretch=False, anchor=tk.CENTER)
        self.tree.column("target_sections", width=150, stretch=False, anchor=tk.CENTER)
        self.tree.column("obs_type", width=140, stretch=False, anchor=tk.CENTER)

        self.tree.tag_configure("even", background="#f7fafc")
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("warn", background="#fff5f5", foreground="#c53030")

        scroll_y = ttk.Scrollbar(tbl_container, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # Toolbar below the table
        tbl_toolbar = ttk.Frame(table_frame)
        tbl_toolbar.pack(fill=tk.X, pady=(6, 0))

        edit_btn = tk.Button(
            tbl_toolbar,
            text="✏️ Edit Selected Message",
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

        add_btn = ttk.Button(tbl_toolbar, text="➕ Add Booking", command=self._on_add_booking_click)
        add_btn.pack(side=tk.LEFT, padx=(0, 6))

        del_btn = ttk.Button(tbl_toolbar, text="🗑️ Delete Selected", command=self._on_delete_booking_click)
        del_btn.pack(side=tk.LEFT, padx=(0, 14))

        tk.Label(
            tbl_toolbar,
            text="💡 Tip: Double-click any row to edit its defect message or section.",
            font=("Segoe UI", 8, "italic"),
            fg="#4a5568"
        ).pack(side=tk.LEFT)

        # Bindings
        self.tree.bind("<Double-1>", lambda e: self._on_edit_booking_click())
        self.tree.bind("<Return>", lambda e: self._on_edit_booking_click())

        # Right-click context menu
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="✏️ Edit Message / Booking", command=self._on_edit_booking_click)
        self.tree_menu.add_command(label="➕ Add New Booking", command=self._on_add_booking_click)
        self.tree_menu.add_command(label="🗑️ Delete Selected Booking", command=self._on_delete_booking_click)
        self.tree.bind("<Button-3>", self._show_tree_context_menu)

        # 5. Progress and Controls Bar
        action_frame = ttk.Frame(main_container)
        action_frame.pack(fill=tk.X, pady=(0, 6))

        self.start_btn = tk.Button(
            action_frame,
            text="🚀 START AUTO-FILL JOBCARDS",
            font=("Segoe UI", 11, "bold"),
            bg="#2f855a",
            fg="white",
            activebackground="#276749",
            activeforeground="white",
            padx=16,
            pady=8,
            cursor="hand2",
            command=self._start_autofill
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = tk.Button(
            action_frame,
            text="⏹ Stop Process",
            font=("Segoe UI", 10, "bold"),
            bg="#c53030",
            fg="white",
            activebackground="#9b2c2c",
            activeforeground="white",
            padx=12,
            pady=8,
            state=tk.DISABLED,
            cursor="hand2",
            command=self._stop_autofill
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.close_browser_btn = tk.Button(
            action_frame,
            text="🌐 Close Browser",
            font=("Segoe UI", 9, "bold"),
            bg="#4a5568",
            fg="white",
            activebackground="#2d3748",
            activeforeground="white",
            padx=10,
            pady=8,
            cursor="hand2",
            command=self._close_browser_manually
        )
        self.close_browser_btn.pack(side=tk.LEFT, padx=(0, 16))

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
        self.log("Settings reset to RPME defaults (Schedule Attention | Without Eq | Testing Remarks | Minor).")

    def _close_browser_manually(self):
        self.worker.task_queue.put({"action": "CLOSE_BROWSER"})

    def _on_app_close(self):
        self.worker.task_queue.put({"action": "SHUTDOWN"})
        self.root.after(150, self.root.destroy)

    def log(self, message: str):
        """Thread-safe logging into the GUI text widget with timestamps."""
        def _append():
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{ts}] {message}\n")
            self.log_text.see(tk.END)
        self.root.after(0, _append)

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

    def _refresh_bookings_table(self):
        """Refreshes the Treeview table, total count, and section breakdown from self.parsed_data."""
        if not self.parsed_data:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        bookings = self.parsed_data.get("bookings", [])
        self.count_var.set(str(len(bookings)))
        self.parsed_data["bookings_count"] = len(bookings)

        sec_counts = collections.Counter()
        for idx, b in enumerate(bookings):
            secs = b.get("sections", [])
            sec_counts.update(secs)
            sec_display = ", ".join(secs) if secs else "⚠️ Unrecognized"
            tag = "warn" if not secs else ("even" if idx % 2 == 0 else "odd")
            obs_type_val = b.get("obs_type", self.obs_type_var.get())
            self.tree.insert(
                "",
                tk.END,
                iid=f"row_{idx}",
                values=(b.get("sno", idx + 1), b.get("description", ""), b.get("raw_section", ""), sec_display, obs_type_val),
                tags=(tag,)
            )

        if sec_counts:
            sec_summary = "  |  ".join(f"{s}: {c}" for s, c in sec_counts.most_common())
            self.sec_breakdown_var.set(f"Sections: {sec_summary}")
        else:
            self.sec_breakdown_var.set("Sections: None detected")

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
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select Booking", "Please select a booking row in the table to edit.")
            return

        item_id = sel[0]
        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        if not self.parsed_data or idx >= len(self.parsed_data.get("bookings", [])):
            return

        booking = self.parsed_data["bookings"][idx]
        dlg = EditBookingDialog(self.root, booking, is_new=False)
        self.root.wait_window(dlg)

        if dlg.result:
            self.parsed_data["bookings"][idx] = dlg.result
            self._refresh_bookings_table()
            if self.tree.exists(f"row_{idx}"):
                self.tree.selection_set(f"row_{idx}")
                self.tree.see(f"row_{idx}")
            self.log(f"Updated S.No {dlg.result['sno']} (Source: {dlg.result.get('obs_type')}): '{dlg.result['description'][:50]}...'")

    def _on_add_booking_click(self):
        if not self.parsed_data:
            messagebox.showwarning("No File", "Please select or load an Excel file first.")
            return

        next_sno = len(self.parsed_data.get("bookings", [])) + 1
        new_booking = {
            "sno": next_sno,
            "description": "",
            "raw_section": "M1",
            "sections": ["M1"],
            "category": "General",
            "obs_type": self.obs_type_var.get(),
            "kind": self.job_kind_var.get()
        }
        dlg = EditBookingDialog(self.root, new_booking, is_new=True)
        self.root.wait_window(dlg)

        if dlg.result:
            self.parsed_data["bookings"].append(dlg.result)
            self._refresh_bookings_table()
            new_idx = len(self.parsed_data["bookings"]) - 1
            if self.tree.exists(f"row_{new_idx}"):
                self.tree.selection_set(f"row_{new_idx}")
                self.tree.see(f"row_{new_idx}")
            self.log(f"Added new booking S.No {dlg.result['sno']}: '{dlg.result['description'][:50]}...'")

    def _on_delete_booking_click(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select Booking", "Please select a booking row in the table to delete.")
            return

        item_id = sel[0]
        try:
            idx = int(item_id.replace("row_", ""))
        except ValueError:
            idx = self.tree.index(item_id)

        if not self.parsed_data or idx >= len(self.parsed_data.get("bookings", [])):
            return

        booking = self.parsed_data["bookings"][idx]
        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete S.No {booking.get('sno')}?\n\n'{booking.get('description', '')[:70]}...'"
        )
        if confirm:
            self.parsed_data["bookings"].pop(idx)
            for i, b in enumerate(self.parsed_data["bookings"], start=1):
                b["sno"] = i
            self._refresh_bookings_table()
            self.log(f"Deleted booking S.No {booking.get('sno')}. Total remaining: {len(self.parsed_data['bookings'])}")

    def _prompt_otp(self) -> str:
        """Prompts the user for OTP via a custom modal dialog in the main GUI thread."""
        result = [None]
        event = threading.Event()

        def _ask():
            dlg = tk.Toplevel(self.root)
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
                dlg.destroy()
                event_set_safe()

            def _cancel():
                result[0] = ""
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
            dlg = VerificationDialog(self.root, staged_rows, meta)
            self.root.wait_window(dlg)
            result[0] = dlg.result
            event.set()

        self.root.after(0, _ask)
        event.wait()
        return result[0]

    def _start_autofill(self):
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
                "keep_browser": self.keep_browser_var.get()
            }
        })

    def _stop_autofill(self):
        self.worker.stop_current_job()
        self.status_lbl.config(text="Stopping...")
        self.log("Stopping process as requested by user...")

    def _update_progress(self, current: int, total: int, msg: str):
        def _update():
            pct = int((current / total) * 100) if total > 0 else 0
            self.progress_bar["value"] = pct
            self.status_lbl.config(text=f"{msg} ({pct}%)")
        self.root.after(0, _update)

    def _reset_buttons(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

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
            UpdateDialog(self.root, self.update_info)

def main():
    root = tk.Tk()
    app = SLAMAutoFillerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
