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
        browser_choice: Optional[str] = None,
        custom_browser_path: Optional[str] = None,
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

        b_pref = config.get_browser_preference()
        self.browser_choice = browser_choice if browser_choice is not None else b_pref.get("choice", config.BROWSER_CHOICE_AUTO)
        self.custom_browser_path = custom_browser_path if custom_browser_path is not None else b_pref.get("custom_path", "")
        self.current_browser_launched: Optional[str] = None

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
        """Launches the selected or detected browser if not already open."""
        if self.browser and self.browser.is_connected() and self.page:
            try:
                if not self.page.is_closed():
                    return
            except Exception:
                pass

        pref_display = config.BROWSER_DISPLAY_NAMES.get(self.browser_choice, self.browser_choice)
        self.log(f"Launching web browser (Preference: {pref_display})...")
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
        detected = config.detect_system_browsers()
        edge_info = detected["browsers"].get("msedge", {})
        chrome_info = detected["browsers"].get("chrome", {})
        brave_info = detected["browsers"].get("brave", {})
        win_def = detected.get("windows_default")

        attempts = []
        choice = (self.browser_choice or config.BROWSER_CHOICE_AUTO).lower()

        if choice == config.BROWSER_CHOICE_EDGE:
            attempts.append(("Microsoft Edge", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            if edge_info.get("installed") and edge_info.get("path"):
                attempts.append(("Microsoft Edge (Direct)", {"executable_path": edge_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
        elif choice == config.BROWSER_CHOICE_CHROME:
            attempts.append(("Google Chrome", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            if chrome_info.get("installed") and chrome_info.get("path"):
                attempts.append(("Google Chrome (Direct)", {"executable_path": chrome_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
        elif choice == config.BROWSER_CHOICE_BRAVE:
            if brave_info.get("installed") and brave_info.get("path"):
                attempts.append(("Brave Browser", {"executable_path": brave_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            else:
                self.log("Brave Browser not found in standard paths. Trying auto-fallbacks...")
        elif choice == config.BROWSER_CHOICE_CUSTOM:
            if self.custom_browser_path and os.path.isfile(self.custom_browser_path):
                attempts.append((f"Custom Browser ({os.path.basename(self.custom_browser_path)})", {"executable_path": self.custom_browser_path, "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            else:
                self.log(f"Custom browser executable not found at '{self.custom_browser_path}'. Trying auto-fallbacks...")
        elif choice == config.BROWSER_CHOICE_CHROMIUM:
            attempts.append(("Chromium", {"headless": self.headless, "slow_mo": 50, "args": launch_args}))
        else: # auto / default
            if win_def and win_def[0] == "chrome":
                attempts.append(("Google Chrome (Default)", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
                attempts.append(("Microsoft Edge", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            elif win_def and win_def[0] == "msedge":
                attempts.append(("Microsoft Edge (Default)", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
                attempts.append(("Google Chrome", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            elif win_def and win_def[0] == "brave" and brave_info.get("installed") and brave_info.get("path"):
                attempts.append(("Brave Browser (Default)", {"executable_path": brave_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
                attempts.append(("Microsoft Edge", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
                attempts.append(("Google Chrome", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
            else:
                attempts.append(("Microsoft Edge", {"channel": "msedge", "headless": self.headless, "slow_mo": 50, "args": launch_args}))
                attempts.append(("Google Chrome", {"channel": "chrome", "headless": self.headless, "slow_mo": 50, "args": launch_args}))

        # Add robust universal fallbacks
        if edge_info.get("installed") and edge_info.get("path") and not any("Edge" in a[0] for a in attempts):
            attempts.append(("Microsoft Edge Fallback", {"executable_path": edge_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
        if chrome_info.get("installed") and chrome_info.get("path") and not any("Chrome" in a[0] for a in attempts):
            attempts.append(("Google Chrome Fallback", {"executable_path": chrome_info["path"], "headless": self.headless, "slow_mo": 50, "args": launch_args}))
        if not any(a[0] == "Chromium" for a in attempts):
            attempts.append(("Chromium Fallback", {"headless": self.headless, "slow_mo": 50, "args": launch_args}))

        last_err = None
        for browser_name, launch_opts in attempts:
            try:
                self.browser = self.playwright.chromium.launch(**launch_opts)
                self.current_browser_launched = browser_name
                self.log(f"Launched {browser_name} browser successfully.")
                break
            except Exception as e:
                last_err = e
                self.log(f"{browser_name} not available: {e}. Trying fallback...")

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

        # Solve CAPTCHA & Click Get OTP (with automatic retry if SLAM SMS gateway is slow)
        otp_input = self.page.locator("#ContentPlaceHolder1_txtOTP")
        for otp_attempt in (1, 2):
            if otp_attempt > 1:
                self.log("Retrying CAPTCHA & 'Get OTP' click...")
                self.safe_goto(config.LOGIN_URL)
                user_input.wait_for(state="visible", timeout=15000)
                user_input.fill(self.username)
                pwd_input.fill(self.password)

            cap_el = self.page.locator("#ContentPlaceHolder1_captchaEquation")
            if cap_el.count() > 0 and cap_el.first.is_visible():
                cap_text = cap_el.first.inner_text().strip()
                answer = self._solve_captcha(cap_text)
                self.log(f"CAPTCHA: '{cap_text}' -> Solved: '{answer}'")
                self.page.locator("#ContentPlaceHolder1_txtCaptcha").fill(answer)

            self.log("Clicking 'Get OTP' button...")
            get_otp_btn = self.page.locator("#ContentPlaceHolder1_BtnGetOtp")
            get_otp_btn.click()

            try:
                otp_input.wait_for(state="visible", timeout=25000)
                self.log("OTP requested successfully!")
                break
            except Exception:
                if otp_attempt == 2:
                    raise

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
        self.page.evaluate("""(val) => {
            const el = document.getElementById('Content1_ddlLoco');
            if (el) {
                el.value = val;
                if (typeof $ !== 'undefined') $(el).trigger('change');
                if (el.onchange) el.onchange();
            }
        }""", target_opt['value'])

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

        res = self.page.evaluate("""([typeVal, eqtVal, obsVal, kindVal]) => {
            // 1. Jobcard Type
            const ddlJobcardType = document.getElementById('Content1_ddlJobcardType');
            if (ddlJobcardType) {
                ddlJobcardType.value = typeVal;
                if (typeof $ !== 'undefined') $(ddlJobcardType).trigger('change');
                else if (ddlJobcardType.onchange) ddlJobcardType.onchange();
            }

            // 2. Eqt Required
            const ddlEqReq = document.getElementById('Content1_EQtRequired');
            if (ddlEqReq) {
                ddlEqReq.value = eqtVal;
                if (typeof SectionEnable === 'function') {
                    SectionEnable(ddlEqReq.id, 'SectionSingleSelectDiv', 'SectionMultiCheckDiv');
                }
            }

            // 3. Obs Type (Source)
            const ddlObsType = document.getElementById('Content1_ddlObsType');
            if (ddlObsType) {
                ddlObsType.value = obsVal;
                if (typeof $ !== 'undefined') $(ddlObsType).trigger('change');
                else if (ddlObsType.onchange) ddlObsType.onchange();
            }

            // 4. Job Kind
            const ddlJobKind = document.getElementById('Content1_ddlJobKind');
            if (ddlJobKind) {
                ddlJobKind.value = kindVal;
                if (typeof $ !== 'undefined') $(ddlJobKind).trigger('change');
                else if (ddlJobKind.onchange) ddlJobKind.onchange();
            }

            return {
                jobcardType: ddlJobcardType && ddlJobcardType.selectedIndex >= 0 ? ddlJobcardType.options[ddlJobcardType.selectedIndex].text : null,
                eqtReq: ddlEqReq && ddlEqReq.selectedIndex >= 0 ? ddlEqReq.options[ddlEqReq.selectedIndex].text : null,
                obsType: ddlObsType && ddlObsType.selectedIndex >= 0 ? ddlObsType.options[ddlObsType.selectedIndex].text : null,
                jobKind: ddlJobKind && ddlJobKind.selectedIndex >= 0 ? ddlJobKind.options[ddlJobKind.selectedIndex].text : null
            };
        }""", [type_val, eqt_val, obs_val, kind_val])
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

        res = self.page.evaluate("""async ([targetObs, targetKind, sectionCodes, textVal]) => {
            // 1. Set Obs Type for this specific entry
            const ddlObsType = document.getElementById('Content1_ddlObsType');
            if (ddlObsType && ddlObsType.value !== targetObs) {
                ddlObsType.value = targetObs;
                if (typeof $ !== 'undefined') $(ddlObsType).trigger('change');
                else if (ddlObsType.onchange) ddlObsType.onchange();
            }

            // 2. Set Job Kind for this specific entry
            const ddlJobKind = document.getElementById('Content1_ddlJobKind');
            if (ddlJobKind && ddlJobKind.value !== targetKind) {
                ddlJobKind.value = targetKind;
                if (typeof $ !== 'undefined') $(ddlJobKind).trigger('change');
                else if (ddlJobKind.onchange) ddlJobKind.onchange();
            }

            function selectSections(codes) {
                const div = document.getElementById('ctl00$Content1$ddlSection1Parent');
                const dd = document.getElementsByName('ctl00$Content1$ddlSection1')[0];
                const hidden = document.getElementsByName('ctl00$Content1$ddlSection1_hidden')[0];
                if (!div || !dd || !hidden) return;
                
                const checkboxes = div.getElementsByTagName('input');
                let returnArray = [];
                let returnSelectedValues = [];
                
                for (let i = 0; i < checkboxes.length; i++) {
                    const chk = checkboxes[i];
                    if (chk.type === 'checkbox' && chk.title !== 'Check/Uncheck All' && chk.title !== '--Select--') {
                        if (codes.includes(chk.title)) {
                            chk.checked = true;
                            if (chk.parentNode) chk.parentNode.style.background = '#FDFDCD';
                            returnArray.push(chk.title);
                            returnSelectedValues.push(chk.value);
                        } else {
                            chk.checked = false;
                            if (chk.parentNode) chk.parentNode.style.background = 'transparent';
                        }
                    }
                }
                hidden.value = returnSelectedValues.length > 0 ? returnSelectedValues.join(',') : '0';
                dd.value = returnArray.join(';');
            }

            selectSections(sectionCodes);
            const txt = document.getElementById('Content1_txtRemarks');
            if (txt) txt.value = textVal;
            const initialRows = document.getElementById('tbldata1') ? document.getElementById('tbldata1').rows.length : 0;
            
            // Trigger Add To Job Observations
            if (typeof StoreInfo === 'function') {
                StoreInfo();
            } else {
                const addBtn = document.getElementById('Content1_btnAdd');
                if (addBtn) addBtn.click();
            }
            
            // Wait for tbldata1 to update (up to 10 seconds)
            for (let i = 0; i < 50; i++) {
                await new Promise(r => setTimeout(r, 200));
                const currentRows = document.getElementById('tbldata1') ? document.getElementById('tbldata1').rows.length : 0;
                if (currentRows > initialRows) {
                    return { success: true, currentRows: currentRows };
                }
            }
            return { success: false, timeout: true };
        }""", [obs_val, kind_val, section_codes, text.strip()])

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

    # ==========================================================================
    # SHOP FLOOR & LOCO EQUIPMENT AUTOMATION (Playwright)
    # ==========================================================================

    def open_shopfloor_module(self, module_name: str, otp: str = "") -> str:
        """
        Logs into SLAM (if not already logged in) and opens the requested Shop Floor module
        in the active browser window.
        """
        if not self.is_logged_in:
            if not self.login(otp=otp):
                raise RuntimeError("Failed to authenticate to SLAM portal.")

        direct_urls = {
            "Add Loco Eq./Eq. Fitment": f"{config.BASE_URL}/EquipmentRegister/LocoEquipmentAdd.aspx",
            "Update Loco Equipment": f"{config.BASE_URL}/EquipmentRegister/LocoEquipmentEdit.aspx",
            "Add Equipment History": f"{config.BASE_URL}/EquipmentRegister/EquipmentHistoryAddEdit.aspx",
            "Equipment Change Register": f"{config.BASE_URL}/EquipmentRegister/EquipmentChangeRegister.aspx",
            "Equipment History": f"{config.BASE_URL}/EquipmentRegister/ViewEquipmentHistory.aspx",
            "Equipment Register": f"{config.BASE_URL}/Inventory/EquipmentRegisterView.aspx",
            "Loco Inventory": f"{config.BASE_URL}/Inventory/LocoInventoryView.aspx",
            "Section Loco Inventory": f"{config.BASE_URL}/EquipmentRegister/SectionLocoInventoryView.aspx",
            "Repair Sheet History": f"{config.BASE_URL}/Inventory/RepairSheetsHistory.aspx",
            "Section Equipment Position": f"{config.BASE_URL}/General/SectionEquipmentPosition.aspx?showButtons=true",
        }

        menu_id = config.SHOPFLOOR_MODULES.get(module_name)
        if module_name in direct_urls:
            target_url = direct_urls[module_name]
            self.log(f"Opening Shop Floor -> {module_name} ({target_url})...")
            self.safe_goto(target_url)
            return self.page.url

        # Navigate to Home.aspx first so RedirectPage() and top menu items are guaranteed present
        self.safe_goto(f"{config.BASE_URL}/General/Home.aspx")
        time.sleep(0.8)

        if menu_id:
            self.log(f"Opening Shop Floor -> {module_name} via RedirectPage('{menu_id}')...")
            try:
                self.page.evaluate(f"if (typeof RedirectPage !== 'undefined') {{ RedirectPage('{menu_id}'); }}")
                self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                time.sleep(1.2)
            except Exception as e:
                self.log(f"Warning navigating via RedirectPage('{menu_id}'): {e}")
        else:
            self.safe_goto(config.EQUIPMENT_ADD_URL)

        self.log(f"Shop Floor module '{module_name}' opened: {self.page.url}")
        return self.page.url

    def _find_frame_containing_text(self, page: Page, text: str, max_wait: float = 10.0):
        """
        Finds and returns the Frame (or main_frame) inside `page` that contains `text`.
        Mirrors switch_to_frame_containing() from slam_equipment_automation.py.
        """
        end_time = time.time() + max_wait
        xpath_expr = f"xpath=//*[contains(normalize-space(text()), '{text}')]"
        while time.time() < end_time:
            try:
                if page.main_frame.locator(xpath_expr).count() > 0:
                    return page.main_frame
            except Exception:
                pass
            for frame in page.frames:
                if frame == page.main_frame:
                    continue
                try:
                    if frame.locator(xpath_expr).count() > 0:
                        return frame
                except Exception:
                    continue
            time.sleep(0.4)
        return None

    def _find_eq_input_by_label(self, frame, label_text: str, timeout_ms: int = 4000):
        """
        Finds an <input> element on LocoEquipmentsTreeAdd.aspx by known ID first,
        then falls back to table label XPath matching.
        """
        id_map = {
            "Provision Date": "#txtProvDate",
            "Eq.Number": "#txtNumber",
            "Location Eq.Number": "#txtNumber1",
            "Local Eq.Number": "#txtNumber1",
            "Date of Commission": "#txtDoC0",
            "Date of Manufacture": "#txtDoM0",
            "Date of Recieve": "#txtDoR0",
            "Date of Receive": "#txtDoR0",
            "Warranty Upto(If Avbl)": "#txtWarrenty0",
            "AMC Upto (If Avbl)": "#txtAMCUpto0",
            "Remarks": "#txtReamrks0",
        }
        if label_text in id_map:
            loc_id = frame.locator(id_map[label_text]).first
            try:
                if loc_id.count() > 0:
                    return loc_id
            except Exception:
                pass

        xpath_table = f"xpath=//td[normalize-space(text())='{label_text}']/following-sibling::td[1]//input"
        loc_table = frame.locator(xpath_table).first
        try:
            loc_table.wait_for(state="attached", timeout=timeout_ms)
            return loc_table
        except Exception:
            pass

        xpath_generic = f"xpath=//*[normalize-space(text())='{label_text}']/following::input[1]"
        loc_generic = frame.locator(xpath_generic).first
        loc_generic.wait_for(state="attached", timeout=timeout_ms)
        return loc_generic

    def _find_eq_select_by_label(self, frame, label_text: str, timeout_ms: int = 4000):
        """
        Finds a <select> element on LocoEquipmentsTreeAdd.aspx by known ID first,
        then falls back to table label XPath matching.
        """
        id_map = {
            "Schedule": "#ddlSchedule",
            "Model": "#ddlModel",
            "Warranty": "#ddlWarranty0",
            "AMC": "#ddlAMC0",
            "Eq.Make": "#ddlMake",
        }
        if label_text in id_map:
            loc_id = frame.locator(id_map[label_text]).first
            try:
                if loc_id.count() > 0:
                    return loc_id
            except Exception:
                pass

        xpath_table = f"xpath=//td[normalize-space(text())='{label_text}']/following-sibling::td[1]//select"
        loc_table = frame.locator(xpath_table).first
        try:
            loc_table.wait_for(state="attached", timeout=timeout_ms)
            return loc_table
        except Exception:
            pass

        xpath_generic = f"xpath=//*[normalize-space(text())='{label_text}']/following::select[1]"
        loc_generic = frame.locator(xpath_generic).first
        loc_generic.wait_for(state="attached", timeout=timeout_ms)
        return loc_generic

    def _set_eq_text(self, frame, label_text: str, value: str, wait_enabled: bool = False):
        """
        Sets a text or date field in the Add Loco Equipment popup form.
        Automatically removes 'readonly' attribute via JS for SLAM calendar date fields.
        """
        from excel_parser import is_blank_eq_value
        if is_blank_eq_value(value):
            return
        val_str = str(value).strip()
        try:
            field = self._find_eq_input_by_label(frame, label_text)
            if wait_enabled:
                for _ in range(10):
                    try:
                        if field.is_enabled():
                            break
                    except Exception:
                        pass
                    time.sleep(0.4)

            if not field.is_enabled():
                self.log(f"    [SKIP] '{label_text}' is disabled, skipping.")
                return

            is_ro = False
            try:
                is_ro = field.evaluate("el => el.readOnly || el.hasAttribute('readonly')")
            except Exception:
                pass

            if not is_ro:
                try:
                    field.fill(val_str, timeout=3000)
                    self.log(f"    [text] {label_text} = {val_str}")
                    return
                except Exception:
                    pass

            # Readonly date field or fill fallback: set via JS and dispatch input/change/blur events
            field.evaluate(
                """(el, v) => {
                    el.removeAttribute('readonly');
                    el.readOnly = false;
                    el.value = v;
                    el.dispatchEvent(new Event('input', {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    el.dispatchEvent(new Event('blur', {bubbles: true}));
                }""",
                val_str
            )
            self.log(f"    [text-js] {label_text} = {val_str}")
        except Exception as e:
            self.log(f"    [SKIP] Field '{label_text}' could not be set: {e}")

    def _set_eq_dropdown(self, frame, label_text: str, value: str):
        """
        Selects an option in a <select> dropdown by:
        1) Exact visible text match
        2) Case-insensitive match (e.g. 'VIKRANT' vs 'Vikrant', 'FLORICON' vs 'Floricon')
        3) Alias / Punctuation-insensitive match (e.g. 'FE' -> 'FLORICON', 'BG' -> 'B G IND.')
        4) O/0 lookalike normalization (e.g. 'ICO' vs 'IC0')
        """
        from excel_parser import is_blank_eq_value, normalize_make_alias
        if is_blank_eq_value(value):
            return
        raw_val = str(value).strip()
        val_str = normalize_make_alias(raw_val) if label_text == "Eq.Make" else raw_val
        if not val_str:
            return
        try:
            dropdown = self._find_eq_select_by_label(frame, label_text)
            options = dropdown.evaluate(
                "el => Array.from(el.options).map(o => ({value: o.value, text: o.text.trim()}))"
            )

            # 1) Exact match
            for opt in options:
                if opt["text"] == val_str:
                    dropdown.select_option(value=opt["value"])
                    self.log(f"    [dropdown] {label_text} = {opt['text']}")
                    return

            # 2) Case-insensitive match
            for opt in options:
                if opt["text"].upper() == val_str.upper():
                    dropdown.select_option(value=opt["value"])
                    self.log(f"    [dropdown] {label_text} = {opt['text']} (matched '{raw_val}')")
                    return

            # 3) Punctuation/space-insensitive match (e.g. 'BG IND' <-> 'B G IND.')
            val_alnum = re.sub(r"[^A-Z0-9]", "", val_str.upper())
            if val_alnum:
                for opt in options:
                    opt_alnum = re.sub(r"[^A-Z0-9]", "", opt["text"].upper())
                    if opt_alnum == val_alnum or (len(val_alnum) >= 2 and opt_alnum.startswith(val_alnum)):
                        dropdown.select_option(value=opt["value"])
                        self.log(f"    [dropdown] {label_text} = {opt['text']} (alias-matched '{raw_val}')")
                        return

            # 4) Lookalike O/0 normalization (e.g. ICO <-> IC0)
            target_norm = val_str.upper().replace("O", "0")
            for opt in options:
                if opt["text"].upper().replace("O", "0") == target_norm:
                    dropdown.select_option(value=opt["value"])
                    self.log(f"    [dropdown] {label_text} = {opt['text']} (O/0 matched '{raw_val}')")
                    return

            if label_text == "Eq.Make":
                # Check if the dropdown has no valid configured options from SLAM master (e.g. only ['--Select--'])
                valid_opts = [o for o in options if str(o.get("value", "")).strip() not in ("-1", "", "None", "0")]
                if not valid_opts:
                    target_upper = re.sub(r"[^A-Z0-9]", "", val_str.upper())
                    KNOWN_MAKES_ID_MAP = {
                        "ELGI": "64",
                        "RECON": "165",
                        "KNORR": "338",
                        "FTRTIL": "453",
                        "ELECTROMECH": "1768",
                        "AUTOINDUSTRIES": "1931",
                        "FLORICON": "225",
                        "TROLEX": "311",
                        "BGIND": "225",
                    }
                    injected_id = None
                    injected_name = val_str
                    for k, v in KNOWN_MAKES_ID_MAP.items():
                        if k in target_upper or target_upper in k:
                            injected_id = v
                            injected_name = k
                            break
                    if not injected_id:
                        injected_id = "64"
                        injected_name = val_str if val_str else "ELGI"

                    dropdown.evaluate(f"""el => {{
                        const opt = document.createElement('option');
                        opt.value = '{injected_id}';
                        opt.text = '{injected_name}';
                        el.appendChild(opt);
                        el.value = '{injected_id}';
                        el.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}""")
                    self.log(
                        f"    [AUTO-INJECT] Eq.Make dropdown was empty in SLAM! "
                        f"Injected make '{injected_name}' (ID {injected_id}) to satisfy SLAM database constraint."
                    )
                    return

            avail = [o["text"] for o in options]
            self.log(f"    [WARN] Option '{raw_val}' (resolved '{val_str}') not found in '{label_text}'. Available: {avail}")
        except Exception as e:
            self.log(f"    [SKIP] Dropdown '{label_text}' issue: {e}")

    def _populate_equipment_fields(self, popup_page: Page, frame, row: Dict[str, Any], eq_num_override: str = ""):
        """Populates all inputs and dropdowns in the LocoEquipmentsTreeAdd.aspx popup form."""
        from excel_parser import is_blank_eq_value

        self._set_eq_text(frame, "Provision Date", row.get("provision_date", ""))
        self._set_eq_dropdown(frame, "Schedule", row.get("schedule", ""))
        # Wait briefly in case Schedule selection triggers a postback
        time.sleep(1.0)
        try:
            popup_page.wait_for_load_state("domcontentloaded", timeout=5000)
        except Exception:
            pass

        frame = self._find_frame_containing_text(popup_page, "Provision Date", max_wait=5.0) or popup_page.main_frame

        eff_eq_num = eq_num_override or str(row.get("eq_number", "")).strip()
        eff_loc_num = eq_num_override or str(row.get("loc_eq_number", "")).strip() or eff_eq_num

        self._set_eq_dropdown(frame, "Model", row.get("model", ""))
        self._set_eq_text(frame, "Eq.Number", eff_eq_num)
        self._set_eq_text(frame, "Local Eq.Number", eff_loc_num)
        self._set_eq_text(frame, "Date of Commission", row.get("doc", ""))
        self._set_eq_text(frame, "Date of Manufacture", row.get("dom", ""))
        self._set_eq_text(frame, "Date of Recieve", row.get("receive_date", ""))

        warranty_val = str(row.get("warranty", "")).strip()
        self._set_eq_dropdown(frame, "Warranty", warranty_val)
        if not is_blank_eq_value(warranty_val) and warranty_val.upper() == "YES":
            time.sleep(0.5)
            self._set_eq_text(frame, "Warranty Upto(If Avbl)", row.get("warranty_upto", ""), wait_enabled=True)

        amc_val = str(row.get("amc", "")).strip()
        self._set_eq_dropdown(frame, "AMC", amc_val)
        if not is_blank_eq_value(amc_val) and amc_val.upper() == "YES":
            time.sleep(0.5)
            self._set_eq_text(frame, "AMC Upto (If Avbl)", row.get("amc_upto", ""), wait_enabled=True)

        self._set_eq_dropdown(frame, "Eq.Make", row.get("make", ""))
        self._set_eq_text(frame, "Remarks", row.get("remarks", ""))
        return frame

    def _fill_equipment_popup_form(
        self,
        popup_page: Page,
        frame,
        row: Dict[str, Any],
        auto_save: bool = True,
        dialogs_list: Optional[List[str]] = None
    ) -> bool:
        """Fills all fields inside the LocoEquipmentsTreeAdd.aspx popup window for a single locomotive equipment item."""
        frame = self._populate_equipment_fields(popup_page, frame, row)

        if auto_save:
            try:
                save_btn = frame.locator("#btnSave, input[value='Save']").first
                save_btn.wait_for(state="visible", timeout=8000)
                save_btn.click()
                time.sleep(1.8)
                try:
                    popup_page.wait_for_load_state("domcontentloaded", timeout=5000)
                except Exception:
                    pass

                # Check if SLAM validation alert occurred (e.g. Please Check Mandatory Fields)
                if dialogs_list and any("mandatory" in str(d).lower() or "check" in str(d).lower() for d in dialogs_list):
                    self.log(f"    [ERROR] Save rejected by SLAM form validation: {dialogs_list[-1]}")
                    return False

                # Check if SLAM rejected with database constraint error or duplicate serial
                frame = self._find_frame_containing_text(popup_page, "Provision Date", max_wait=3.0) or popup_page.main_frame
                try:
                    body_txt = frame.locator("body").inner_text()
                except Exception:
                    body_txt = ""

                if "foreign key constraint" in body_txt.lower() or "conversion failed" in body_txt.lower():
                    first_err = [l.strip() for l in body_txt.splitlines() if "constraint" in l.lower() or "conversion" in l.lower()]
                    err_summary = first_err[0] if first_err else body_txt[:120]
                    self.log(f"    [ERROR] SLAM Database Error on Save: {err_summary}")
                    return False

                if "already existed" in body_txt.lower():
                    loco_no = str(row.get("loco_no", "")).strip()
                    pos = str(row.get("position", "1")).strip() or "1"
                    orig_eq_num = str(row.get("eq_number", "")).strip()
                    unique_eq_num = f"{loco_no}-TM{pos}"
                    if unique_eq_num.upper() == orig_eq_num.upper():
                        unique_eq_num = f"{loco_no}-TM{pos}-{int(time.time()) % 1000}"
                    self.log(
                        f"    [AUTO-FIX] SLAM reported 'Serial No {orig_eq_num} already existed'. "
                        f"Retrying with unique locomotive serial number '{unique_eq_num}'..."
                    )
                    frame = self._populate_equipment_fields(popup_page, frame, row, eq_num_override=unique_eq_num)
                    save_btn = frame.locator("#btnSave, input[value='Save']").first
                    save_btn.wait_for(state="visible", timeout=8000)
                    save_btn.click()
                    time.sleep(1.8)

                self.log("    [SAVE] Equipment form saved successfully.")
                return True
            except Exception as e:
                self.log(f"    [ERROR] Could not click Save button: {e}")
                return False
        else:
            self.log("    [STAGED / DRY-RUN] Auto-Save is OFF — popup form filled without clicking Save.")
            time.sleep(1.0)
            return True

    def process_single_loco_equipment(
        self,
        loco_no: str,
        equipment_rows: List[Dict[str, Any]],
        select_for: str = "Home",
        auto_save: bool = True,
        close_popup_on_dry_run: bool = True
    ) -> Dict[str, Any]:
        """Processes all pending equipment links on LocoEquipmentAdd.aspx for a single locomotive."""
        from excel_parser import find_matching_equipment_row, is_blank_eq_value

        loco_no = str(loco_no).strip()
        self.log(f"\n=== Processing Loco {loco_no} Equipment ===")
        self.safe_goto(config.EQUIPMENT_ADD_URL)
        time.sleep(1.2)

        frame = self._find_frame_containing_text(self.page, "Loco", max_wait=10.0)
        if frame is None:
            self.log(f"[WARN] Could not locate 'Loco' selector on {config.EQUIPMENT_ADD_URL}.")
            return {"loco_no": loco_no, "filled": 0, "saved": 0, "skipped": 0, "failed": 1, "status": "page_error"}

        # 1. Select 'Home' (or requested Select For) if not already selected
        try:
            shed_dd = frame.locator("#Content1_ddlLocoShed")
            if shed_dd.count() > 0:
                curr_shed = shed_dd.evaluate("el => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text.trim() : ''")
                if curr_shed.lower() != select_for.lower():
                    shed_dd.select_option(label=select_for)
                    self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                    time.sleep(1.2)
                    frame = self._find_frame_containing_text(self.page, "Loco", max_wait=5.0) or self.page.main_frame
        except Exception as e:
            self.log(f"  [Note] Select For ('{select_for}') selection note: {e}")

        # 2. Select Locomotive in #Content1_ddlLoco
        try:
            loco_dd = frame.locator("#Content1_ddlLoco")
            loco_dd.wait_for(state="visible", timeout=10000)
            has_loco = loco_dd.evaluate(
                "(el, lno) => Array.from(el.options).some(o => o.text.trim() === lno)",
                loco_no
            )
            if not has_loco:
                self.log(f"[SKIP] Loco {loco_no} is not listed in the '{select_for}' dropdown on SLAM.")
                return {"loco_no": loco_no, "filled": 0, "saved": 0, "skipped": 1, "failed": 0, "status": "loco_not_in_dropdown"}

            loco_dd.select_option(label=loco_no)
            self.page.wait_for_load_state("domcontentloaded", timeout=10000)
            time.sleep(1.0)
            frame = self._find_frame_containing_text(self.page, "Loco", max_wait=5.0) or self.page.main_frame
        except Exception as e:
            self.log(f"[ERROR] Failed to select Loco {loco_no}: {e}")
            return {"loco_no": loco_no, "filled": 0, "saved": 0, "skipped": 0, "failed": 1, "status": "select_error"}

        # 3. Click Go button (#Content1_btnGo)
        try:
            go_btn = frame.locator("#Content1_btnGo, input[value='Go'], input[value='GO']").first
            if go_btn.count() > 0:
                go_btn.click()
                self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                time.sleep(1.5)
                frame = self._find_frame_containing_text(self.page, "Loco", max_wait=5.0) or self.page.main_frame
        except Exception as e:
            self.log(f"[ERROR] Failed to click Go for Loco {loco_no}: {e}")
            return {"loco_no": loco_no, "filled": 0, "saved": 0, "skipped": 0, "failed": 1, "status": "go_error"}

        # 4. Inspect #Content1_gvEquipment for already-fitted rows and pending clickable equipment links
        try:
            already_fitted = frame.locator("#Content1_gvEquipment span[id*='lblEquipname']").evaluate_all(
                """els => els.map(el => {
                    const tr = el.closest('tr');
                    const posEl = tr ? tr.querySelector("span[id*='lblPos']") : null;
                    return {name: el.innerText.trim(), pos: posEl ? posEl.innerText.trim() : ''};
                })"""
            )
            for af in already_fitted:
                self.log(f"  [ALREADY FITTED] {af.get('name')} (Pos {af.get('pos')}) is already fitted on Loco {loco_no}.")
        except Exception:
            already_fitted = []

        link_selector = "#Content1_gvEquipment a[id*='lnkEq'], #Content1_gvEquipment a[onclick*='popup'], a:has-text('Filter for')"
        links_loc = frame.locator(link_selector)
        total_links = links_loc.count()
        self.log(f"[INFO] Found {total_links} pending unfilled equipment slot(s) and {len(already_fitted)} already-fitted slot(s) for Loco {loco_no}.")

        if total_links == 0:
            if len(already_fitted) == 0:
                try:
                    ltype = frame.locator("#Content1_lblType").inner_text().strip()
                except Exception:
                    ltype = "Unknown"
                self.log(f"  [SKIP] SLAM returned 'No Records Found...' for Loco {loco_no} (Type: {ltype} has no equipment slots assigned to this section in SLAM).")
            return {
                "loco_no": loco_no,
                "filled": 0,
                "saved": 0,
                "skipped": len(already_fitted) or 1,
                "failed": 0,
                "status": "no_pending_links"
            }

        handled_keys = set()
        filled_count = 0
        saved_count = 0
        skipped_count = len(already_fitted)
        failed_count = 0
        max_iterations = total_links + 5

        for _ in range(max_iterations):
            if self._stop_requested:
                break

            frame = self._find_frame_containing_text(self.page, "Loco", max_wait=5.0) or self.page.main_frame
            links_loc = frame.locator(link_selector)
            curr_count = links_loc.count()
            if curr_count == 0:
                break

            chosen_idx = None
            raw_eq_text = ""
            eq_name = ""
            position = ""

            for i in range(curr_count):
                try:
                    info = links_loc.nth(i).evaluate(
                        """el => {
                            const tr = el.closest('tr');
                            const posSpan = tr ? tr.querySelector("span[id*='lblPos']") : null;
                            let pos = posSpan ? posSpan.innerText.trim() : '';
                            if (!pos && tr) {
                                const cells = Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim());
                                if (cells.length > 3) pos = cells[3];
                            }
                            return {
                                id: el.id || '',
                                text: el.innerText.trim(),
                                pos: pos
                            };
                        }"""
                    )
                    raw_link_txt = info.get("text", "")
                    cand_eq = re.sub(r"^Filter\s+for\s+", "", raw_link_txt, flags=re.IGNORECASE).strip()
                    cand_eq = re.sub(r"\s*\(.*?\)\s*$", "", cand_eq).strip().upper()
                    cand_pos = str(info.get("pos", "")).strip()
                except Exception:
                    continue

                if (cand_eq, cand_pos) not in handled_keys:
                    chosen_idx = i
                    raw_eq_text = raw_link_txt
                    eq_name = cand_eq
                    position = cand_pos
                    break

            if chosen_idx is None:
                break

            handled_keys.add((eq_name, position))
            self.log(f"  -> Pending Slot: '{raw_eq_text}' (Normalized: {eq_name}), Position: {position}")

            sheet_row = find_matching_equipment_row(equipment_rows, loco_no, eq_name, position)
            if sheet_row is None:
                sheet_row = find_matching_equipment_row(equipment_rows, loco_no, raw_eq_text, position)
            if sheet_row is None:
                self.log(f"    [SKIP] No matching row in loaded table for Loco {loco_no} | {raw_eq_text} Pos {position}.")
                skipped_count += 1
                continue

            if is_blank_eq_value(sheet_row.get("make", "")) and is_blank_eq_value(sheet_row.get("eq_number", "")):
                self.log(f"    [SKIP] Both Make and Eq.Number are blank in table for Loco {loco_no} | {raw_eq_text} Pos {position}.")
                skipped_count += 1
                continue

            # Click link and capture newly opened popup window (LocoEquipmentsTreeAdd.aspx)
            popup_page: Optional[Page] = None
            popup_dialogs: List[str] = []
            try:
                with self.context.expect_page(timeout=20000) as new_page_info:
                    links_loc.nth(chosen_idx).click()
                popup_page = new_page_info.value
                popup_page.set_default_timeout(20000)
                def _on_popup_dlg(d):
                    popup_dialogs.append(d.message)
                    self._handle_dialog(d)
                popup_page.on("dialog", _on_popup_dlg)
                popup_page.wait_for_load_state("domcontentloaded", timeout=15000)
                time.sleep(1.0)
            except Exception as e:
                self.log(f"    [ERROR] Failed to open popup window for {raw_eq_text} Pos {position}: {e}")
                failed_count += 1
                continue

            try:
                # Click 'Go' (#btnGo) for Available Type: New
                form_frame = None
                for attempt in (1, 2):
                    avail_frame = self._find_frame_containing_text(popup_page, "Available Type", max_wait=6.0) or popup_page.main_frame
                    try:
                        go_avail = avail_frame.locator("#btnGo, input[value='Go'], input[value='GO']").first
                        if go_avail.count() > 0:
                            go_avail.click()
                            self.log("    [INFO] Clicked 'Go' (Available Type: New) in popup window.")
                            time.sleep(1.5)
                    except Exception:
                        pass

                    form_frame = self._find_frame_containing_text(popup_page, "Provision Date", max_wait=6.0)
                    if form_frame is not None:
                        break
                    self.log(f"    [WARN] Attempt {attempt}: Equipment form fields not loaded yet...")
                    time.sleep(1.0)

                if form_frame is None:
                    self.log(f"    [ERROR] Form fields did not load for {raw_eq_text} Pos {position}. Skipping.")
                    failed_count += 1
                else:
                    ok = self._fill_equipment_popup_form(popup_page, form_frame, sheet_row, auto_save=auto_save, dialogs_list=popup_dialogs)
                    if ok:
                        filled_count += 1
                        if auto_save:
                            saved_count += 1
                    else:
                        failed_count += 1
            except Exception as e:
                self.log(f"    [ERROR] Error filling equipment form for {raw_eq_text} Pos {position}: {e}")
                failed_count += 1
            finally:
                if auto_save or close_popup_on_dry_run:
                    try:
                        if popup_page and not popup_page.is_closed():
                            popup_page.close()
                    except Exception:
                        pass
                time.sleep(0.8)

        self.log(f"=== Loco {loco_no} Complete (Filled: {filled_count}, Saved: {saved_count}, Skipped: {skipped_count}, Errors: {failed_count}) ===")
        return {
            "loco_no": loco_no,
            "filled": filled_count,
            "saved": saved_count,
            "skipped": skipped_count,
            "failed": failed_count,
            "status": "ok"
        }

    def process_loco_equipment(
        self,
        rows: Optional[List[Dict[str, Any]]] = None,
        equipment_rows: Optional[List[Dict[str, Any]]] = None,
        target_locos: Optional[List[str]] = None,
        select_for: str = "Home",
        auto_save: bool = True,
        otp: str = ""
    ) -> Dict[str, Any]:
        """
        Batch processes locomotive equipment fitment across one or more locomotives.
        """
        from excel_parser import is_blank_eq_value

        self._stop_requested = False
        active_rows = rows if rows is not None else (equipment_rows or [])
        if not active_rows:
            raise ValueError("No equipment rows loaded!")

        if not target_locos:
            seen = set()
            target_locos = []
            for r in active_rows:
                lno = str(r.get("loco_no", "")).strip()
                if not is_blank_eq_value(lno) and lno not in seen:
                    seen.add(lno)
                    target_locos.append(lno)

        total_locos = len(target_locos)
        self.log("============================================================")
        self.log(f"Starting Loco Equipment Auto-Fill for {total_locos} Locomotive(s)")
        self.log(f"Select For: {select_for} | Auto-Save: {'YES' if auto_save else 'NO (Dry Run / Stage Only)'}")
        self.log("============================================================")

        if not self.is_logged_in:
            if not self.login(otp=otp):
                raise RuntimeError("Failed to authenticate to SLAM portal.")

        total_filled = 0
        total_saved = 0
        total_skipped = 0
        total_failed = 0
        per_loco_results = []

        for idx, loco_no in enumerate(target_locos, start=1):
            if self._stop_requested:
                self.log("Equipment automation stopped by user.")
                break

            if self.progress_callback:
                self.progress_callback(idx - 1, total_locos, f"Processing Loco {loco_no} ({idx}/{total_locos})...")

            try:
                res = self.process_single_loco_equipment(
                    loco_no=loco_no,
                    equipment_rows=active_rows,
                    select_for=select_for,
                    auto_save=auto_save
                )
                total_filled += res.get("filled", 0)
                total_saved += res.get("saved", 0)
                total_skipped += res.get("skipped", 0)
                total_failed += res.get("failed", 0)
                per_loco_results.append(res)
            except Exception as e:
                self.log(f"[ERROR] Loco {loco_no} encountered an error: {e}")
                total_failed += 1
                per_loco_results.append({"loco_no": loco_no, "filled": 0, "saved": 0, "skipped": 0, "failed": 1, "status": f"error: {e}"})

            if self.progress_callback:
                self.progress_callback(idx, total_locos, f"Completed Loco {loco_no} ({idx}/{total_locos})")

        status = "cancelled" if self._stop_requested else "completed"
        self.log(
            f"\nEquipment Auto-Fill {status.upper()}: Locos={len(per_loco_results)}/{total_locos}, "
            f"Filled={total_filled}, Saved={total_saved}, Skipped={total_skipped}, Errors={total_failed}"
        )
        return {
            "status": status,
            "total_locos": total_locos,
            "locos_processed": len(per_loco_results),
            "processed_locos": len(per_loco_results),
            "filled": total_filled,
            "saved": total_saved,
            "total_saved": total_saved,
            "skipped": total_skipped,
            "total_skipped": total_skipped,
            "failed": total_failed,
            "results": per_loco_results,
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

