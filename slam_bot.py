"""
SLAM Portal Jobcard Automation Bot
Uses Playwright to automate Indian Railways SLAM Portal Jobcard creation.
"""
import time
import re
import os
import sys
from typing import List, Dict, Any, Optional, Callable
from playwright.sync_api import sync_playwright, Page, BrowserContext, Browser, TimeoutError as PlaywrightTimeoutError

import config
from excel_parser import parse_excel_bookings

class SLAMBot:
    def __init__(
        self,
        username: Optional[str] = None,
        password: Optional[str] = None,
        jobcard_type: str = config.DEFAULT_JOBCARD_TYPE,
        eqt_required: str = config.DEFAULT_EQT_REQUIRED,
        obs_type: str = config.DEFAULT_OBS_TYPE,
        job_kind: str = config.DEFAULT_JOB_KIND,
        headless: bool = False,
        log_callback: Optional[Callable[[str], None]] = None,
        otp_callback: Optional[Callable[[], str]] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        verify_before_save: bool = True,
        verify_callback: Optional[Callable[..., bool]] = None
    ):
        cfg_u, cfg_p = config.get_credentials()
        self.username = username if username is not None else cfg_u
        self.password = password if password is not None else cfg_p
        self.jobcard_type = jobcard_type
        self.eqt_required = eqt_required
        self.obs_type = obs_type
        self.job_kind = job_kind
        self.headless = headless
        self.log_callback = log_callback
        self.otp_callback = otp_callback
        self.progress_callback = progress_callback
        self.verify_before_save = verify_before_save
        self.verify_callback = verify_callback

        self.playwright = None
        self.browser = None
        self.context = None
        self.page: Optional[Page] = None
        self.is_logged_in = False
        self._stop_requested = False

    def log(self, message: str):
        print(f"[SLAM Bot] {message}", flush=True)
        if self.log_callback:
            try:
                self.log_callback(message)
            except Exception:
                pass

    def stop(self):
        """Signal the bot to stop execution."""
        self._stop_requested = True
        self.log("Stop requested by user.")

    def _solve_captcha(self, text: str) -> str:
        """Parses '12 + 7 = ?' and returns '19'."""
        m = re.search(r'(\d+)\s*([\+\-\*])\s*(\d+)', text)
        if m:
            a, op, b = int(m.group(1)), m.group(2), int(m.group(3))
            res = {'+': a + b, '-': a - b, '*': a * b}.get(op, a + b)
            return str(res)
        return ""

    def start_browser(self):
        """Launches the Chromium browser if not already open."""
        if self.browser and self.browser.is_connected() and self.page:
            try:
                if not self.page.is_closed():
                    return
            except Exception:
                pass

        self.log("Launching Chromium browser...")
        if self.page:
            try:
                self.page.close()
            except Exception:
                pass
            self.page = None
        if self.context:
            try:
                self.context.close()
            except Exception:
                pass
            self.context = None
        if self.browser:
            try:
                self.browser.close()
            except Exception:
                pass
            self.browser = None

        if self.playwright is None:
            self.playwright = sync_playwright().start()

        launch_args = ["--start-maximized", "--disable-blink-features=AutomationControlled"]
        self.browser = None

        # Universal Browser Resolution:
        # 1. Standard Chromium (if available in Playwright cache)
        # 2. Microsoft Edge (channel="msedge") - built into every Windows 10/11 installation!
        # 3. Google Chrome (channel="chrome")
        attempts = [
            ("Chromium", {"headless": self.headless, "slow_mo": 50, "args": launch_args}),
            ("Microsoft Edge", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}),
            ("Google Chrome", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args})
        ]

        last_err = None
        for browser_name, launch_opts in attempts:
            try:
                self.browser = self.playwright.chromium.launch(**launch_opts)
                self.log(f"Launched {browser_name} browser successfully.")
                break
            except Exception as e:
                last_err = e
                self.log(f"{browser_name} not available. Trying fallback...")

        if not self.browser:
            raise RuntimeError(
                f"Could not launch any web browser on this system!\n\n"
                f"Please ensure Microsoft Edge or Google Chrome is installed.\n"
                f"Details: {last_err}"
            )

        self.context = self.browser.new_context(
            no_viewport=True,
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
        )
        self.page = self.context.new_page()
        self.page.set_default_timeout(30000)

        # Automatically accept dialogs like "Jobcards Saved Successfully"
        self.page.on("dialog", self._handle_dialog)

    def _handle_dialog(self, dialog):
        msg = dialog.message
        self.log(f"Browser Dialog [{dialog.type}]: {msg}")
        try:
            dialog.accept()
        except Exception:
            pass

    def safe_goto(self, url: str, wait_until: str = "domcontentloaded", max_retries: int = 3, timeout: int = 30000) -> bool:
        """
        Navigates to a URL safely with retry handling for interrupted navigations.
        This handles cases where a concurrent background transition (e.g. ASP.NET postback redirect
        to SendJobcards.aspx?Type=1) interrupts the page.goto call.
        """
        for attempt in range(1, max_retries + 1):
            try:
                # If page is actively navigating or settling from a previous action, wait briefly
                try:
                    self.page.wait_for_load_state("domcontentloaded", timeout=3000)
                except Exception:
                    pass

                self.page.goto(url, wait_until=wait_until, timeout=timeout)
                return True
            except Exception as e:
                err_msg = str(e)
                if "interrupted by another navigation" in err_msg or "navigating" in err_msg.lower():
                    self.log(f"Navigation was interrupted by concurrent page transition ({attempt}/{max_retries}). Waiting for page to settle...")
                    time.sleep(2.5)
                    try:
                        self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                    except Exception:
                        pass
                    if attempt == max_retries:
                        time.sleep(1.5)
                        self.page.goto(url, wait_until=wait_until, timeout=timeout)
                        return True
                else:
                    raise

    def login(self, otp: str = "") -> bool:
        """Performs full authentication flow into SLAM."""
        self.start_browser()
        self.log(f"Navigating to login page: {config.LOGIN_URL}...")
        self.safe_goto(config.LOGIN_URL, wait_until="domcontentloaded")

        # Check if already logged in (e.g. redirected to dashboard)
        if "login" not in self.page.url.lower():
            self.log(f"Already logged in. Current URL: {self.page.url}")
            self.is_logged_in = True
            return True

        self.log(f"Entering credentials for '{self.username}'...")
        user_input = self.page.locator("#ContentPlaceHolder1_txtUserName")
        user_input.wait_for(state="visible", timeout=15000)
        user_input.fill(self.username)

        pwd_input = self.page.locator("#ContentPlaceHolder1_txtPassword")
        pwd_input.fill(self.password)

        # Solve CAPTCHA
        cap_el = self.page.locator("#ContentPlaceHolder1_captchaEquation")
        if cap_el.count() > 0 and cap_el.first.is_visible():
            cap_text = cap_el.first.inner_text().strip()
            answer = self._solve_captcha(cap_text)
            self.log(f"CAPTCHA: '{cap_text}' -> Solved: '{answer}'")
            self.page.locator("#ContentPlaceHolder1_txtCaptcha").fill(answer)

        # Click Get OTP
        self.log("Clicking 'Get OTP' button...")
        get_otp_btn = self.page.locator("#ContentPlaceHolder1_BtnGetOtp")
        get_otp_btn.click()

        # Wait for OTP field to appear
        otp_input = self.page.locator("#ContentPlaceHolder1_txtOTP")
        otp_input.wait_for(state="visible", timeout=15000)
        self.log("OTP requested successfully!")

        used_cached_otp = False
        if not otp:
            cached = config.get_daily_otp()
            if cached:
                self.log(f"Using saved daily OTP: {cached}")
                otp = cached
                used_cached_otp = True
            elif self.otp_callback:
                self.log("Waiting for user to enter OTP via prompt...")
                otp = self.otp_callback()
            else:
                otp = input("[ACTION REQUIRED] Enter the OTP received on mobile/WhatsApp: ").strip()

        if not otp:
            self.log("No OTP provided. Login aborted.")
            return False

        self.log("Submitting Password and OTP...")
        # ASP.NET clears password on postback; re-enter it
        pwd_input.fill(self.password)
        otp_input.fill(str(otp).strip())

        # Click Login
        login_btn = self.page.locator("#ContentPlaceHolder1_btnSignIn")
        login_btn.click()

        # Wait for redirect away from Login.aspx
        self.log("Waiting for dashboard redirect...")
        start_t = time.time()
        while time.time() - start_t < 15:
            if "login" not in self.page.url.lower():
                self.log(f"Login SUCCESSFUL! Redirected to: {self.page.url}")
                self.is_logged_in = True
                config.set_daily_otp(otp)
                return True
            time.sleep(1)

        # If cached OTP failed, prompt user for fresh OTP and retry
        if used_cached_otp:
            self.log("Saved daily OTP was rejected or expired. Requesting fresh OTP from user...")
            if self.otp_callback:
                otp = self.otp_callback()
            else:
                otp = input("[ACTION REQUIRED] Enter new OTP received on mobile/WhatsApp: ").strip()

            if otp:
                self.log(f"Retrying login with fresh OTP: {otp}...")
                pwd_input.fill(self.password)
                otp_input.fill(str(otp).strip())
                login_btn.click()
                start_t = time.time()
                while time.time() - start_t < 15:
                    if "login" not in self.page.url.lower():
                        self.log(f"Login SUCCESSFUL with fresh OTP! Redirected to: {self.page.url}")
                        self.is_logged_in = True
                        config.set_daily_otp(otp)
                        return True
                    time.sleep(1)

        self.log("Login failed: Still on login page.")
        return False

    def navigate_to_jobcard_create(self):
        """Navigates to the Jobcard Create screen."""
        self.start_browser()
        self.log(f"Navigating to Jobcard Create page: {config.JOBCARD_CREATE_URL}...")
        self.safe_goto(config.JOBCARD_CREATE_URL, wait_until="domcontentloaded")
        if "login" in self.page.url.lower():
            self.log("Session expired or redirected to Login. Re-authenticating...")
            self.is_logged_in = False
            if not self.login():
                raise RuntimeError("Failed to re-authenticate to SLAM portal.")
            self.safe_goto(config.JOBCARD_CREATE_URL, wait_until="domcontentloaded")
        self.page.wait_for_selector("#Content1_ddlLoco", timeout=20000)
        self.log("Arrived at Jobcard Create screen.")

    def select_locomotive(self, loco_number: str) -> bool:
        """Selects the locomotive in the dropdown, which fires postback."""
        loco_clean = str(loco_number).strip()
        self.log(f"Locating Locomotive '{loco_clean}' in dropdown...")

        loco_select = self.page.locator("#Content1_ddlLoco")
        loco_select.wait_for(state="visible", timeout=15000)

        # Check options
        options = self.page.evaluate("""() => {
            const el = document.getElementById('Content1_ddlLoco');
            return el ? Array.from(el.options).map(o => ({ text: o.text, value: o.value })) : [];
        }""")

        target_opt = next((o for o in options if o["text"] == loco_clean), None)
        if not target_opt:
            available = [o["text"] for o in options if o["text"] != "--Select--"]
            raise ValueError(f"Locomotive '{loco_clean}' not found in SLAM active loco list! Available: {available}")

        self.log(f"Selecting Loco '{loco_clean}'...")
        # Select option
        self.page.evaluate(f"""() => {{
            const el = document.getElementById('Content1_ddlLoco');
            el.value = {repr(target_opt['value'])};
            el.onchange();
        }}""")

        # Wait for ASP.NET postback / UpdatePanel to settle
        time.sleep(3)
        self.page.wait_for_load_state("domcontentloaded")
        self.log(f"Locomotive '{loco_clean}' loaded successfully.")
        return True

    def configure_form_settings(self):
        """Configures Jobcard Type, Equipment Required, Source, and Kind."""
        self.log(f"Configuring form: Type='{self.jobcard_type}', EqReq='{self.eqt_required}', Source='{self.obs_type}', Kind='{self.job_kind}'...")

        type_val = config.JOBCARD_TYPE_VALUES.get(self.jobcard_type, "2")
        eqt_val = "2" if "without" in self.eqt_required.lower() else "1"
        obs_val = config.OBS_TYPE_VALUES.get(self.obs_type, "2")
        kind_val = config.JOB_KIND_VALUES.get(self.job_kind, "2")

        res = self.page.evaluate(f"""() => {{
            // 1. Jobcard Type
            const ddlJobcardType = document.getElementById('Content1_ddlJobcardType');
            if (ddlJobcardType) {{
                ddlJobcardType.value = '{type_val}';
                $(ddlJobcardType).trigger('change');
            }}

            // 2. Eqt Required
            const ddlEqReq = document.getElementById('Content1_EQtRequired');
            if (ddlEqReq) {{
                ddlEqReq.value = '{eqt_val}';
                SectionEnable(ddlEqReq.id, 'SectionSingleSelectDiv', 'SectionMultiCheckDiv');
            }}

            // 3. Obs Type (Source)
            const ddlObsType = document.getElementById('Content1_ddlObsType');
            if (ddlObsType) {{
                ddlObsType.value = '{obs_val}';
                $(ddlObsType).trigger('change');
            }}

            // 4. Job Kind
            const ddlJobKind = document.getElementById('Content1_ddlJobKind');
            if (ddlJobKind) {{
                ddlJobKind.value = '{kind_val}';
                $(ddlJobKind).trigger('change');
            }}

            return {{
                jobcardType: ddlJobcardType ? ddlJobcardType.options[ddlJobcardType.selectedIndex].text : null,
                eqtReq: ddlEqReq ? ddlEqReq.options[ddlEqReq.selectedIndex].text : null,
                obsType: ddlObsType ? ddlObsType.options[ddlObsType.selectedIndex].text : null,
                jobKind: ddlJobKind ? ddlJobKind.options[ddlJobKind.selectedIndex].text : null
            }};
        }}""")
        self.log(f"Form configured: {res}")

    def add_observation(
        self,
        section_codes: List[str],
        text: str,
        obs_type: Optional[str] = None,
        kind: Optional[str] = None
    ) -> bool:
        """Adds a single observation with its assigned section(s) and obs_type via AJAX."""
        if not section_codes or not text.strip():
            return False

        target_obs = obs_type or self.obs_type
        obs_val = config.OBS_TYPE_VALUES.get(target_obs, "2")

        target_kind = kind or self.job_kind
        kind_val = config.JOB_KIND_VALUES.get(target_kind, "2")

        res = self.page.evaluate(f"""async () => {{
            // 1. Set Obs Type for this specific entry
            const ddlObsType = document.getElementById('Content1_ddlObsType');
            if (ddlObsType && ddlObsType.value !== '{obs_val}') {{
                ddlObsType.value = '{obs_val}';
                $(ddlObsType).trigger('change');
            }}

            // 2. Set Job Kind for this specific entry
            const ddlJobKind = document.getElementById('Content1_ddlJobKind');
            if (ddlJobKind && ddlJobKind.value !== '{kind_val}') {{
                ddlJobKind.value = '{kind_val}';
                $(ddlJobKind).trigger('change');
            }}

            function selectSections(sectionCodes) {{
                const div = document.getElementById('ctl00$Content1$ddlSection1Parent');
                const dd = document.getElementsByName('ctl00$Content1$ddlSection1')[0];
                const hidden = document.getElementsByName('ctl00$Content1$ddlSection1_hidden')[0];
                if (!div || !dd || !hidden) return;
                
                const checkboxes = div.getElementsByTagName('input');
                let returnArray = [];
                let returnSelectedValues = [];
                
                for (let i = 0; i < checkboxes.length; i++) {{
                    const chk = checkboxes[i];
                    if (chk.type === 'checkbox' && chk.title !== 'Check/Uncheck All' && chk.title !== '--Select--') {{
                        if (sectionCodes.includes(chk.title)) {{
                            chk.checked = true;
                            if (chk.parentNode) chk.parentNode.style.background = '#FDFDCD';
                            returnArray.push(chk.title);
                            returnSelectedValues.push(chk.value);
                        }} else {{
                            chk.checked = false;
                            if (chk.parentNode) chk.parentNode.style.background = 'transparent';
                        }}
                    }}
                }}
                hidden.value = returnSelectedValues.length > 0 ? returnSelectedValues.join(',') : '0';
                dd.value = returnArray.join(';');
            }}

            selectSections({section_codes});
            const txt = document.getElementById('Content1_txtRemarks');
            txt.value = {repr(text.strip())};
            const initialRows = document.getElementById('tbldata1') ? document.getElementById('tbldata1').rows.length : 0;
            
            // Trigger Add To Job Observations
            StoreInfo();
            
            // Wait for tbldata1 to update
            for (let i = 0; i < 40; i++) {{
                await new Promise(r => setTimeout(r, 200));
                const currentRows = document.getElementById('tbldata1') ? document.getElementById('tbldata1').rows.length : 0;
                if (currentRows > initialRows) {{
                    return {{ success: true, currentRows: currentRows }};
                }}
            }}
            return {{ success: false, timeout: true }};
        }}""")

        return res.get("success", False)

    def get_staged_observations(self) -> List[Dict[str, Any]]:
        """Extracts all staged rows from #tbldata1 table on the portal."""
        if not self.page:
            return []
        try:
            rows = self.page.evaluate("""() => {
                const tbl = document.getElementById('tbldata1');
                if (!tbl) return [];
                const res = [];
                for (let i = 1; i < tbl.rows.length; i++) {
                    const r = tbl.rows[i];
                    res.push({
                        job_no: r.cells[1] ? r.cells[1].innerText.trim() : '',
                        section: r.cells[2] ? r.cells[2].innerText.trim() : '',
                        sno: r.cells[3] ? r.cells[3].innerText.trim() : '',
                        desc: r.cells[4] ? r.cells[4].innerText.trim() : '',
                        obs_type: r.cells[5] ? r.cells[5].innerText.trim() : '',
                        kind: r.cells[6] ? r.cells[6].innerText.trim() : ''
                    });
                }
                return res;
            }""")
            return rows or []
        except Exception as e:
            self.log(f"Note: Could not query staged observations table: {e}")
            return []

    def save_jobcards(self) -> bool:
        """Clicks Save, confirms the modal dialog, and verifies success."""
        self.log("Opening save confirmation dialog via SaveData()...")

        # 1. Trigger SaveData() which opens the popup and reveals #Content1_btnSave
        self.page.evaluate("""() => {
            if (typeof SaveData !== 'undefined') {
                SaveData();
            } else {
                const btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.trim() === 'Save');
                if (btn) btn.click();
            }
        }""")

        # Wait for the modal dialog to appear
        time.sleep(1.5)
        self.log("Confirming save in dialog ('Do You Want Save Job Card To Section??')...")

        # 2. Click the confirm Save button #Content1_btnSave
        save_btn = self.page.locator("#Content1_btnSave")
        if save_btn.is_visible():
            save_btn.click()
        else:
            self.page.evaluate("""() => {
                const btn = document.getElementById('Content1_btnSave');
                if (btn) btn.click();
            }""")

        self.log("Save submitted! Awaiting server response and dispatch confirmation...")

        # 3. Wait for the server to process and redirect away from JobcardCreate (typically to SendJobcards.aspx?Type=1)
        try:
            self.page.wait_for_url(lambda u: "jobcardcreate" not in u.lower(), timeout=30000)
            self.page.wait_for_load_state("domcontentloaded", timeout=15000)
            self.log(f"Server save confirmed! Current page: {self.page.url}")
        except Exception:
            try:
                self.page.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            time.sleep(3)

        self.log("🎉 Jobcards Saved Successfully!")
        return True

    def process_excel(
        self,
        excel_path: str = "",
        parsed_data: Optional[Dict[str, Any]] = None,
        otp: str = ""
    ) -> Dict[str, Any]:
        """Complete pipeline: Parse Excel (or use edited parsed_data), Login, Select Loco, Fill, and Save."""
        self._stop_requested = False
        if parsed_data is not None:
            parsed = parsed_data
        elif excel_path:
            parsed = parse_excel_bookings(excel_path)
        else:
            raise ValueError("No Excel file or parsed bookings provided!")

        loco_no = parsed["loco_number"]
        bookings = parsed["bookings"]
        total = len(bookings)

        file_name = parsed.get("file_name") or (os.path.basename(excel_path) if excel_path else "Custom Data")
        self.log(f"============================================================")
        self.log(f"Processing File: {file_name}")
        self.log(f"Locomotive: {loco_no} | Date: {parsed.get('date', '--')} | Schedule: {parsed.get('schedule', '--')}")
        self.log(f"Total Bookings to enter: {total}")
        self.log(f"============================================================")

        if total == 0:
            raise ValueError("No bookings found in the Excel / loaded data!")

        if not self.is_logged_in:
            if not self.login(otp=otp):
                raise RuntimeError("Failed to authenticate to SLAM portal.")

        if self._stop_requested:
            return {"status": "cancelled"}

        self.navigate_to_jobcard_create()

        if self._stop_requested:
            return {"status": "cancelled"}

        self.select_locomotive(loco_no)
        self.configure_form_settings()

        entered_count = 0
        failed_items = []

        for idx, item in enumerate(bookings, start=1):
            if self._stop_requested:
                self.log("Stopping batch upon user request.")
                break

            sno = item["sno"]
            desc = item["description"]
            secs = item["sections"]

            if not secs:
                self.log(f"[{idx}/{total}] Warning: S.No {sno} has no valid sections (raw: '{item['raw_section']}'). Skipping.")
                failed_items.append({"sno": sno, "desc": desc, "reason": f"No valid section ({item['raw_section']})"})
                continue

            item_obs = item.get("obs_type") or self.obs_type
            item_kind = item.get("kind") or self.job_kind
            self.log(f"[{idx}/{total}] S.No {sno} -> Sections {secs} [{item_obs}]: '{desc[:50]}...'")
            success = self.add_observation(secs, desc, obs_type=item_obs, kind=item_kind)

            if success:
                entered_count += 1
            else:
                self.log(f" -> Failed to add observation for S.No {sno}!")
                failed_items.append({"sno": sno, "desc": desc, "reason": "AJAX timeout or rejected"})

            if self.progress_callback:
                self.progress_callback(idx, total, f"Added S.No {sno} ({idx}/{total})")

            time.sleep(0.2)

        self.log(f"\nAll observations added ({entered_count}/{total} successful).")

        # Verification & Save Jobcards
        if entered_count > 0:
            if self.verify_before_save:
                self.log("Verification stage: All observations staged. Awaiting confirmation before saving...")
                staged_rows = self.get_staged_observations()
                meta = {
                    "loco_number": loco_no,
                    "schedule": parsed.get("schedule", ""),
                    "date": parsed.get("date", ""),
                    "total_staged": len(staged_rows) or entered_count,
                    "jobcard_type": self.jobcard_type,
                    "eqt_required": self.eqt_required,
                    "obs_type": self.obs_type,
                    "job_kind": self.job_kind
                }
                if self.verify_callback:
                    try:
                        proceed = self.verify_callback(staged_rows, meta)
                    except TypeError:
                        try:
                            proceed = self.verify_callback(staged_rows)
                        except TypeError:
                            proceed = self.verify_callback()

                    if not proceed:
                        self.log("Save aborted or deferred by user during verification.")
                        return {
                            "status": "verification_deferred",
                            "loco_number": loco_no,
                            "total": total,
                            "entered": entered_count,
                            "failed_items": failed_items,
                            "staged_rows": staged_rows
                        }
                else:
                    try:
                        ans = input("Observations staged into table. Proceed to Save? (Y/n): ").strip().lower()
                        if ans not in ("", "y", "yes"):
                            self.log("Save aborted by user.")
                            return {
                                "status": "verification_deferred",
                                "loco_number": loco_no,
                                "total": total,
                                "entered": entered_count,
                                "failed_items": failed_items,
                                "staged_rows": staged_rows
                            }
                    except Exception:
                        pass

            self.save_jobcards()
            status = "completed"
        else:
            status = "failed"

        return {
            "status": status,
            "loco_number": loco_no,
            "total": total,
            "entered": entered_count,
            "failed_items": failed_items
        }

    def close(self):
        """Closes browser and cleans up Playwright resources."""
        self.log("Closing browser session...")
        if self.page:
            try:
                self.page.close()
            except Exception:
                pass
            self.page = None
        if self.context:
            try:
                self.context.close()
            except Exception:
                pass
            self.context = None
        if self.browser:
            try:
                self.browser.close()
            except Exception:
                pass
            self.browser = None
        if self.playwright:
            try:
                self.playwright.stop()
            except Exception:
                pass
            self.playwright = None
        self.is_logged_in = False
        self.log("Browser closed.")
